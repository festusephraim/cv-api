require("dotenv").config();

const express = require("express");
const cors = require("cors");
const fs = require("fs");
const path = require("path");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const { OpenAI } = require("openai");

const app = express();
app.disable("x-powered-by");

app.use(cors());
app.use(express.json({ limit: "2mb" }));

const PORT = Number(process.env.PORT || 3001);
const BASE_URL = process.env.BASE_URL || `http://localhost:${PORT}`;
const NODE_ENV = process.env.NODE_ENV || "development";

const TEMPLATE_PATH = path.join(__dirname, "templates", "cv-template.docx");
const OUTPUT_DIR = path.join(__dirname, "generated");

const FILE_RETENTION_HOURS = Number(process.env.FILE_RETENTION_HOURS || 24);
const CLEANUP_INTERVAL_MINUTES = Number(process.env.CLEANUP_INTERVAL_MINUTES || 60);

if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

if (!process.env.OPENAI_API_KEY) {
  console.error("Missing OPENAI_API_KEY in environment variables.");
  process.exit(1);
}

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

/**
 * ----------------------------------------
 * HELPERS
 * ----------------------------------------
 */
function safeString(value) {
  if (value === null || value === undefined) return "";

  const cleaned = String(value)
    .replace(/\u00A0/g, " ")
    .replace(/[“”]/g, '"')
    .replace(/[‘’]/g, "'")
    .replace(/\s+/g, " ")
    .trim();

  if (
    cleaned.toLowerCase() === "null" ||
    cleaned.toLowerCase() === "undefined"
  ) {
    return "";
  }

  return cleaned;
}

function safeArray(value) {
  if (!Array.isArray(value)) return [];
  return value.filter((item) => item !== null && item !== undefined);
}

function clampArray(arr, max) {
  return Array.isArray(arr) ? arr.slice(0, max) : [];
}

function normaliseReferenceChoice(value) {
  const cleaned = safeString(value).toLowerCase();

  if (cleaned === "include full references in my cv") return "included";

  if (
    cleaned === "use ‘references available upon request’" ||
    cleaned === "use 'references available upon request'" ||
    cleaned === "use references available upon request" ||
    cleaned === "references available upon request" ||
    cleaned === "available"
  ) {
    return "available";
  }

  if (cleaned === "none") return "none";

  if (cleaned === "included" || cleaned === "available" || cleaned === "none") {
    return cleaned;
  }

  return "available";
}

function toSingleLine(value) {
  return safeString(value)
    .replace(/\s+/g, " ")
    .replace(/\s+([,.;:!?])/g, "$1")
    .trim();
}

function buildContactLine(data) {
  const parts = [
    safeString(data.address),
    safeString(data.phone),
    safeString(data.email),
    safeString(data.linkedin),
  ].filter(Boolean);

  return parts.join(" | ");
}

function buildSkillsLine(skills) {
  return safeArray(skills)
    .map((item) => safeString(item))
    .filter(Boolean)
    .join(" • ");
}

function endOrPresent(value) {
  const cleaned = safeString(value);
  return cleaned || "Present";
}

function normaliseBulletArray(items, max = 4) {
  return clampArray(safeArray(items), max)
    .map((item) => safeString(item))
    .filter(Boolean);
}

function splitLinesToArray(value, max = 5) {
  return clampArray(
    safeString(value)
      .split(/\r?\n|•|;/)
      .map((item) => safeString(item))
      .filter(Boolean),
    max
  );
}

function splitSkills(value) {
  if (Array.isArray(value)) {
    return clampArray(
      value.map((item) => safeString(item)).filter(Boolean),
      12
    );
  }

  return clampArray(
    safeString(value)
      .split(/,|\n|\||•|;/)
      .map((item) => safeString(item))
      .filter(Boolean),
    12
  );
}

function cleanReferenceEntries(entries) {
  return clampArray(safeArray(entries), 3)
    .map((item) => ({
      name: safeString(item?.name),
      position: safeString(item?.position),
      organization: safeString(item?.organization),
      location: safeString(item?.location),
      email: safeString(item?.email),
      phone: safeString(item?.phone),
    }))
    .filter(
      (item) =>
        item.name ||
        item.position ||
        item.organization ||
        item.location ||
        item.email ||
        item.phone
    );
}

function buildReferenceDetailsFromEntries(entries) {
  const cleanedEntries = cleanReferenceEntries(entries);

  return cleanedEntries
    .map((entry) => {
      const line1 = [entry.name, entry.position].filter(Boolean).join(", ");
      const line2 = [entry.organization, entry.location].filter(Boolean).join(", ");
      const line3 = [
        entry.email ? `Email: ${entry.email}` : "",
        entry.phone ? `Phone: ${entry.phone}` : "",
      ]
        .filter(Boolean)
        .join(", ");

      return [line1, line2, line3].filter(Boolean).join("\n");
    })
    .join("\n\n");
}

function buildReferenceText(referenceChoice, referenceDetails) {
  switch (referenceChoice) {
    case "included":
      return safeString(referenceDetails);
    case "available":
      return "References available upon request";
    case "none":
      return "";
    default:
      return "References available upon request";
  }
}

function cleanDisplayName(value) {
  const cleaned = safeString(value)
    .toLowerCase()
    .replace(/[^a-z0-9\s-]/g, "")
    .replace(/\s+/g, " ")
    .trim();

  if (!cleaned) return "Applicant";

  const upperList = ["cv", "ats", "ngo", "api", "sql", "html", "css"];

  return cleaned
    .split(" ")
    .filter(Boolean)
    .map((word) => {
      if (upperList.includes(word)) {
        return word.toUpperCase();
      }

      return word.charAt(0).toUpperCase() + word.slice(1);
    })
    .join(" ");
}

function generateUniqueFileName(fullName) {
  const cleanName = cleanDisplayName(fullName);
  const ext = ".docx";
  const baseName = `${cleanName} CV`;

  let fileName = `${baseName}${ext}`;
  let counter = 1;

  while (fs.existsSync(path.join(OUTPUT_DIR, fileName))) {
    fileName = `${baseName} (${counter})${ext}`;
    counter++;
  }

  return fileName;
}

function parseRequestBody(reqBody) {
  if (
    reqBody?.raw_submission_json &&
    typeof reqBody.raw_submission_json === "object" &&
    !Array.isArray(reqBody.raw_submission_json)
  ) {
    return reqBody.raw_submission_json;
  }

  if (typeof reqBody?.raw_submission_json === "string") {
    try {
      return JSON.parse(reqBody.raw_submission_json);
    } catch (error) {
      const customError = new Error("Invalid saved raw_submission_json");
      customError.details = error.message;
      customError.statusCode = 400;
      throw customError;
    }
  }

  return reqBody;
}

function normalizeIncomingPayload(body) {
  const basicInfo = body?.basic_information || {};
  const workExperience = clampArray(safeArray(body?.work_experience), 3);
  const education = clampArray(safeArray(body?.education), 3);
  const projects = clampArray(safeArray(body?.projects_research), 3);

  const referenceEntries = cleanReferenceEntries(body?.references?.reference_entries);
  const builtReferenceDetails = buildReferenceDetailsFromEntries(referenceEntries);

  let reference_choice = normaliseReferenceChoice(
    body?.references_section_preference
  );

  if (reference_choice === "included" && !builtReferenceDetails) {
    reference_choice = "available";
  }

  const mappedExperience = workExperience.map((item) => ({
    title: safeString(item?.job_title),
    company: safeString(item?.company),
    location: safeString(item?.location),
    start: safeString(item?.start_date),
    end: item?.currently_working_here ? "" : safeString(item?.end_date),
    role_summary: "",
    tasks: splitLinesToArray(item?.what_did_you_do_in_this_role, 5),
  }));

  const mappedEducation = education.map((item) => ({
    degree: safeString(item?.degree_qualification),
    school: safeString(item?.school),
    location: safeString(item?.location),
    start: safeString(item?.start_date),
    end: item?.currently_studying_here ? "" : safeString(item?.end_date),
    edu_detail: safeString(item?.grade_result),
  }));

  const mappedProjects = projects.map((item) => ({
    project_title: safeString(item?.project_title),
    project_description: safeString(item?.project_description),
    start: safeString(item?.start_date),
    end: item?.currently_working_on_this_project ? "" : safeString(item?.end_date),
    project_tasks: splitLinesToArray(item?.what_did_you_do_in_this_project, 5),
  }));

  const extra_sections = [];
  if (safeString(body?.additional_information)) {
    extra_sections.push({
      section_title: "",
      section_content: safeString(body.additional_information),
    });
  }

  return {
    document_type: safeString(body?.document_type),
    document_purpose: safeString(body?.document_purpose),

    full_name: safeString(basicInfo?.full_name),
    address: safeString(basicInfo?.location),
    phone: safeString(basicInfo?.phone_number),
    email: safeString(basicInfo?.email_address),
    linkedin: safeString(basicInfo?.linkedin_profile),
    job_description: safeString(basicInfo?.job_description),

    professional_summary: safeString(body?.professional_summary),

    skills: splitSkills(body?.skills),

    experience: mappedExperience,
    projects: mappedProjects,
    education: mappedEducation,

    certifications: safeString(body?.certifications_awards)
      ? safeString(body.certifications_awards)
          .split(/\r?\n|,/)
          .map((item) => safeString(item))
          .filter(Boolean)
      : [],

    extra_sections,

    reference_choice,
    reference_details: builtReferenceDetails,
    reference_entries: referenceEntries,
  };
}

function cleanExperienceArray(experience) {
  return clampArray(safeArray(experience), 5)
    .map((item) => ({
      title: safeString(item?.title),
      company: safeString(item?.company),
      location: safeString(item?.location),
      start: safeString(item?.start),
      end: safeString(item?.end),
      end_or_present: endOrPresent(item?.end),
      role_summary: safeString(item?.role_summary),
      tasks: normaliseBulletArray(item?.tasks, 5),
    }))
    .filter((item) => item.title || item.company || item.tasks.length);
}

function cleanProjectsArray(projects) {
  return clampArray(safeArray(projects), 4)
    .map((item) => ({
      project_title: safeString(item?.project_title),
      project_description: safeString(item?.project_description),
      start: safeString(item?.start),
      end: safeString(item?.end),
      end_or_present: endOrPresent(item?.end),
      project_tasks: normaliseBulletArray(item?.project_tasks, 4),
    }))
    .filter(
      (item) =>
        item.project_title ||
        item.project_description ||
        item.start ||
        item.end ||
        item.project_tasks.length
    );
}

function cleanEducationArray(education) {
  return clampArray(safeArray(education), 3)
    .map((item) => ({
      degree: safeString(item?.degree),
      school: safeString(item?.school),
      location: safeString(item?.location),
      start: safeString(item?.start),
      end: safeString(item?.end),
      end_or_present: endOrPresent(item?.end),
      edu_detail: safeString(item?.edu_detail),
    }))
    .filter((item) => item.degree || item.school);
}

function cleanCertificationsArray(certifications) {
  return clampArray(safeArray(certifications), 8)
    .map((item) => safeString(item))
    .filter(Boolean);
}

function cleanExtraSections(extraSections) {
  return clampArray(safeArray(extraSections), 6)
    .map((item) => ({
      section_title: safeString(item?.section_title),
      section_content: safeString(item?.section_content),
    }))
    .filter((item) => item.section_content);
}

function cleanStructuredData(data) {
  return {
    full_name: safeString(data.full_name).toUpperCase(),
    address: safeString(data.address),
    phone: safeString(data.phone),
    email: safeString(data.email),
    linkedin: safeString(data.linkedin),
    job_description: safeString(data.job_description),
    professional_summary: toSingleLine(data.professional_summary),

    skills: clampArray(
      safeArray(data.skills)
        .map((item) => safeString(item))
        .filter(Boolean),
      8
    ),

    experience: cleanExperienceArray(data.experience),
    projects: cleanProjectsArray(data.projects),
    education: cleanEducationArray(data.education),
    certifications: cleanCertificationsArray(data.certifications),
    extra_sections: cleanExtraSections(data.extra_sections),

    reference_choice: normaliseReferenceChoice(data.reference_choice),
    reference_details: safeString(data.reference_details),
  };
}

function preserveSectionDatesFromRawInput(parsed, rawInput) {
  const parsedExperience = safeArray(parsed?.experience);
  const parsedEducation = safeArray(parsed?.education);
  const parsedProjects = safeArray(parsed?.projects);

  const rawExperience = safeArray(rawInput?.experience);
  const rawEducation = safeArray(rawInput?.education);
  const rawProjects = safeArray(rawInput?.projects);

  parsed.experience = parsedExperience.map((item, index) => ({
    ...item,
    start: safeString(rawExperience[index]?.start),
    end: safeString(rawExperience[index]?.end),
  }));

  parsed.education = parsedEducation.map((item, index) => ({
    ...item,
    start: safeString(rawEducation[index]?.start),
    end: safeString(rawEducation[index]?.end),
  }));

  parsed.projects = parsedProjects.map((item, index) => ({
    ...item,
    start: safeString(rawProjects[index]?.start),
    end: safeString(rawProjects[index]?.end),
  }));

  return parsed;
}

function preserveReferencesFromRawInput(parsed, rawInput) {
  parsed.reference_choice = rawInput.reference_choice;
  parsed.reference_details = rawInput.reference_details;
  return parsed;
}

function validateIncomingBody(body) {
  if (!body || typeof body !== "object" || Array.isArray(body)) {
    return "Invalid JSON body";
  }

  const basicInfo = body?.basic_information || {};
  const name = safeString(basicInfo?.full_name);
  const email = safeString(basicInfo?.email_address);
  const phone = safeString(basicInfo?.phone_number);

  if (!name) {
    return "basic_information.full_name is required";
  }

  if (!email && !phone) {
    return "At least one contact field is required: basic_information.email_address or basic_information.phone_number";
  }

  return null;
}

function buildPrompt(rawInput) {
  return `
You are a world-class ATS CV writer, senior HR reviewer, recruiter, and CV structuring engine.

Your task is to convert raw user input into a highly professional, ATS-optimised CV structure suitable for real job applications.

IMPORTANT CONTEXT:
- This CV may be used for real job applications
- Many users submit rough, poorly written, badly capitalised, misspelled, repetitive, incomplete, or inconsistently formatted input
- Your job is to clean, refine, and professionalise the writing without changing the truth of the information
- If a job_description is provided, tailor the CV toward that role
- Extract important role keywords naturally from the job description
- Do NOT copy the job description directly
- Do NOT invent qualifications, dates, employers, tools, industries, achievements, certifications, numbers, or metrics not supported by the input

HR AND RECRUITER QUALITY STANDARD:
- Write like a strong HR professional preparing a candidate for screening
- Make the candidate sound clear, employable, and credible
- Prioritise clarity, relevance, evidence, and professionalism
- Remove weak, casual, vague, or repetitive wording
- Use action-driven wording that shows contribution, not just duty
- Keep the candidate’s level realistic; do not inflate entry-level experience into senior-level claims
- Where the role target is clear, align the summary, skills, experience, and projects to that target
- Where the input is limited, use modest professional language instead of exaggeration

INPUT CLEAN-UP RULES:
- Correct obvious spelling mistakes in normal English words
- Correct poor capitalisation throughout
- Convert messy text into proper sentence case where appropriate
- Preserve acronyms and known professional abbreviations in correct form, such as CV, ATS, NGO, UNICEF, WHO, Excel, SQL, DHIS2, NHLMIS, HTML, CSS, API
- Remove needless repetition across sections
- Rewrite awkward or poorly written user input into polished professional English
- Improve grammar, punctuation, spacing, and readability
- Where user input is fragmentary, convert it into proper professional phrasing without inventing facts
- Where multiple entries repeat the same idea, keep the strongest and cleanest version
- Do not produce messy, casual, chat-style, or informal wording
- Ensure final wording is suitable for a professional CV

STRICT RULES:
- Use clear, simple, professional English with correct spelling and punctuation
- Use British English
- No tables, no columns, no graphics
- ATS-friendly wording only
- Maintain correct capitalisation, sentence case, and professional formatting across all sections
- Each experience task must begin with a strong action verb where appropriate
- Avoid weak phrases like "Responsible for"
- Each task should show action, contribution, scope, or outcome where possible
- Keep professional_summary concise, polished, and recruiter-friendly
- No personal pronouns such as "I", "my", or "me"
- Ensure dates are consistent in style
- Preserve exact dates as supplied within their correct sections
- Never transfer dates from projects to work experience or from education to projects
- If project dates are provided, keep them attached to the project entries
- Full name must be uppercase
- If information is missing, return empty strings or empty arrays
- Return only the schema fields
- Do not include markdown
- Do not include commentary
- Do not rewrite or fabricate reference details

TRUTH AND ACCURACY RULE:
- Improve language quality without changing factual meaning
- Never create false claims, false industries, false tools, false achievements, or false qualifications
- Do not exaggerate responsibilities
- Do not add seniority that the input does not support
- Where evidence is limited, use modest but professional wording

METRICS AND IMPACT RULE:
- Use measurable details only when they are explicitly stated or clearly supported by the input
- Preserve real numbers, counts, frequencies, tools, timelines, workloads, team sizes, customer volumes, or output volumes when provided
- If the input suggests scale or frequency but gives no exact figures, write realistic impact language without inventing numbers
- Good examples of safe phrasing without fabricated figures include:
  - "Handled high-volume customer interactions"
  - "Maintained accurate records across daily operations"
  - "Supported timely report preparation and documentation"
  - "Coordinated routine administrative tasks in a fast-paced environment"
- Do NOT create percentages, revenue figures, growth rates, rankings, time savings, or exact counts unless the user input supports them
- Do NOT exaggerate achievements

REFERENCE RULE:
- "included" means use reference_details
- "available" means references available upon request
- "none" means blank reference section

FINAL QUALITY CHECK BEFORE RETURNING JSON:
- Ensure language is polished and professional
- Ensure capitalisation is clean and consistent
- Ensure spelling and punctuation are corrected
- Ensure there is no unnecessary repetition
- Ensure output will look clean when inserted into a CV template
- Ensure all content remains faithful to the original user information

USER INPUT:
${JSON.stringify(rawInput, null, 2)}
`.trim();
}

function ensureTemplateExists() {
  return fs.existsSync(TEMPLATE_PATH);
}

function cleanupOldGeneratedFiles() {
  try {
    if (!fs.existsSync(OUTPUT_DIR)) return;

    const files = fs.readdirSync(OUTPUT_DIR);
    const now = Date.now();
    const maxAgeMs = FILE_RETENTION_HOURS * 60 * 60 * 1000;

    for (const file of files) {
      const filePath = path.join(OUTPUT_DIR, file);

      try {
        const stat = fs.statSync(filePath);
        const ageMs = now - stat.mtimeMs;

        if (stat.isFile() && ageMs > maxAgeMs) {
          fs.unlinkSync(filePath);
        }
      } catch (fileError) {
        console.error(`Failed to inspect/delete file: ${filePath}`, fileError.message);
      }
    }
  } catch (error) {
    console.error("Cleanup process failed:", error.message);
  }
}

/**
 * ----------------------------------------
 * STRUCTURED OUTPUT SCHEMA
 * ----------------------------------------
 */
const CV_JSON_SCHEMA = {
  name: "ats_cv_output",
  strict: true,
  schema: {
    type: "object",
    additionalProperties: false,
    properties: {
      full_name: { type: "string" },
      address: { type: "string" },
      phone: { type: "string" },
      email: { type: "string" },
      linkedin: { type: "string" },
      job_description: { type: "string" },
      professional_summary: { type: "string" },

      skills: {
        type: "array",
        items: { type: "string" },
      },

      experience: {
        type: "array",
        items: {
          type: "object",
          additionalProperties: false,
          properties: {
            title: { type: "string" },
            company: { type: "string" },
            location: { type: "string" },
            start: { type: "string" },
            end: { type: "string" },
            role_summary: { type: "string" },
            tasks: {
              type: "array",
              items: { type: "string" },
            },
          },
          required: [
            "title",
            "company",
            "location",
            "start",
            "end",
            "role_summary",
            "tasks",
          ],
        },
      },

      projects: {
        type: "array",
        items: {
          type: "object",
          additionalProperties: false,
          properties: {
            project_title: { type: "string" },
            project_description: { type: "string" },
            start: { type: "string" },
            end: { type: "string" },
            project_tasks: {
              type: "array",
              items: { type: "string" },
            },
          },
          required: ["project_title", "project_description", "start", "end", "project_tasks"],
        },
      },

      education: {
        type: "array",
        items: {
          type: "object",
          additionalProperties: false,
          properties: {
            degree: { type: "string" },
            school: { type: "string" },
            location: { type: "string" },
            start: { type: "string" },
            end: { type: "string" },
            edu_detail: { type: "string" },
          },
          required: ["degree", "school", "location", "start", "end", "edu_detail"],
        },
      },

      certifications: {
        type: "array",
        items: { type: "string" },
      },

      extra_sections: {
        type: "array",
        items: {
          type: "object",
          additionalProperties: false,
          properties: {
            section_title: { type: "string" },
            section_content: { type: "string" },
          },
          required: ["section_title", "section_content"],
        },
      },

      reference_choice: {
        type: "string",
        enum: ["none", "available", "included"],
      },

      reference_details: { type: "string" },
    },
    required: [
      "full_name",
      "address",
      "phone",
      "email",
      "linkedin",
      "job_description",
      "professional_summary",
      "skills",
      "experience",
      "projects",
      "education",
      "certifications",
      "extra_sections",
      "reference_choice",
      "reference_details",
    ],
  },
};

/**
 * ----------------------------------------
 * ROUTES
 * ----------------------------------------
 */

app.get("/", (req, res) => {
  return res.status(200).json({
    success: true,
    message: "CV API is running",
    environment: NODE_ENV,
    template_exists: ensureTemplateExists(),
  });
});

app.get("/api/health", (req, res) => {
  return res.status(200).json({
    success: true,
    message: "Server is healthy",
    environment: NODE_ENV,
    template_exists: ensureTemplateExists(),
  });
});

app.get("/download/:file", (req, res) => {
  try {
    const fileName = path.basename(decodeURIComponent(req.params.file));
    const filePath = path.join(OUTPUT_DIR, fileName);

    if (!fs.existsSync(filePath)) {
      return res.status(404).json({
        success: false,
        error: "File not found",
      });
    }

    return res.download(filePath, fileName);
  } catch (error) {
    return res.status(500).json({
      success: false,
      error: "Download failed",
    });
  }
});

app.post("/generate-cv", async (req, res) => {
  try {
    cleanupOldGeneratedFiles();

    let requestBody;

    try {
      requestBody = parseRequestBody(req.body);
    } catch (parseError) {
      return res.status(parseError.statusCode || 400).json({
        success: false,
        error: parseError.message || "Invalid request body",
        details: parseError.details || "",
      });
    }

    const incomingError = validateIncomingBody(requestBody);
    if (incomingError) {
      return res.status(400).json({
        success: false,
        error: incomingError,
      });
    }

    if (!ensureTemplateExists()) {
      return res.status(500).json({
        success: false,
        error: "Template file not found: templates/cv-template.docx",
      });
    }

    const rawInput = normalizeIncomingPayload(requestBody);
    const prompt = buildPrompt(rawInput);

    let completion;
    try {
      completion = await openai.chat.completions.create({
        model: "gpt-4.1-mini",
        temperature: 0.2,
        response_format: {
          type: "json_schema",
          json_schema: CV_JSON_SCHEMA,
        },
        messages: [
          {
            role: "developer",
            content:
              "Return only valid JSON matching the provided schema. No markdown. No commentary.",
          },
          {
            role: "user",
            content: prompt,
          },
        ],
      });
    } catch (openaiError) {
      console.error("OpenAI request failed:", openaiError?.message || openaiError);

      const statusCode =
        typeof openaiError?.status === "number" && openaiError.status >= 400
          ? 502
          : 500;

      return res.status(statusCode).json({
        success: false,
        error: "AI generation request failed",
        details: openaiError?.message || "Unknown OpenAI error",
      });
    }

    const content = completion.choices?.[0]?.message?.content;

    if (!content) {
      return res.status(500).json({
        success: false,
        error: "Empty AI response",
      });
    }

    let parsed;
    try {
      parsed = JSON.parse(content);
    } catch (parseError) {
      console.error("Structured output parse failure:", content);
      return res.status(500).json({
        success: false,
        error: "AI returned unreadable JSON",
      });
    }

    parsed = preserveSectionDatesFromRawInput(parsed, rawInput);
    parsed = preserveReferencesFromRawInput(parsed, rawInput);

    const data = cleanStructuredData(parsed);
    const referenceText = buildReferenceText(
      rawInput.reference_choice,
      rawInput.reference_details
    );

    const renderData = {
      FULL_NAME: data.full_name || "",
      CONTACT_LINE: buildContactLine(data) || "",
      PROFESSIONAL_SUMMARY: data.professional_summary || "",
      SKILLS_LINE: buildSkillsLine(data.skills) || "",

      HAS_EXPERIENCE: data.experience.length > 0,
      experience: data.experience,

      HAS_PROJECTS: data.projects.length > 0,
      projects: data.projects,

      HAS_EDUCATION: data.education.length > 0,
      education: data.education,

      HAS_CERTIFICATIONS: data.certifications.length > 0,
      certifications: data.certifications,

      HAS_EXTRA: data.extra_sections.length > 0,
      extra_sections: data.extra_sections,

      HAS_REFERENCE: Boolean(referenceText),
      REFERENCE_SECTION: referenceText || "",

      HAS_REFERENCES_LIST: rawInput.reference_entries.length > 0,
      references_list: rawInput.reference_entries,
    };

    if (NODE_ENV !== "production") {
      console.log("NORMALIZED INPUT:");
      console.dir(rawInput, { depth: null });
      console.log("RENDER DATA:");
      console.dir(renderData, { depth: null });
    }

    let buffer;
    try {
      const binaryTemplate = fs.readFileSync(TEMPLATE_PATH, "binary");
      const zip = new PizZip(binaryTemplate);

      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
        nullGetter() {
          return "";
        },
      });

      doc.render(renderData);

      buffer = doc.getZip().generate({
        type: "nodebuffer",
        compression: "DEFLATE",
      });
    } catch (docError) {
      console.error("Document render failed:", docError?.message || docError);

      return res.status(500).json({
        success: false,
        error: "CV document rendering failed",
        details: docError?.message || "Template render error",
      });
    }

    const fileName = generateUniqueFileName(data.full_name);
    const filePath = path.join(OUTPUT_DIR, fileName);

    try {
      fs.writeFileSync(filePath, buffer);
    } catch (writeError) {
      console.error("Failed to save generated file:", writeError?.message || writeError);

      return res.status(500).json({
        success: false,
        error: "Failed to save generated CV file",
      });
    }

    const protocol = req.headers["x-forwarded-proto"] || req.protocol;
    const host = req.get("host");
    const fullBaseUrl = `${protocol}://${host}`;

    return res.status(200).json({
      success: true,
      message: "CV generated successfully",
      file_name: fileName,
      download_url: `${fullBaseUrl}/download/${encodeURIComponent(fileName)}`,
      reference_text: referenceText,
      preview: renderData,
    });
  } catch (error) {
    console.error("CV generation failed:", error);

    return res.status(500).json({
      success: false,
      error: "CV generation failed",
      details: error?.message || "Unknown error",
    });
  }
});

/**
 * ----------------------------------------
 * JSON PARSE ERROR HANDLER
 * ----------------------------------------
 */
app.use((err, req, res, next) => {
  if (err instanceof SyntaxError && err.status === 400 && "body" in err) {
    return res.status(400).json({
      success: false,
      error: "Invalid JSON body",
    });
  }

  return next(err);
});

/**
 * ----------------------------------------
 * START SERVER
 * ----------------------------------------
 */
cleanupOldGeneratedFiles();

setInterval(() => {
  cleanupOldGeneratedFiles();
}, CLEANUP_INTERVAL_MINUTES * 60 * 1000);

app.listen(PORT, "0.0.0.0", () => {
  console.log(`CV API running on ${BASE_URL}`);
  console.log(`Template exists: ${ensureTemplateExists()}`);
  console.log(
    `File cleanup: every ${CLEANUP_INTERVAL_MINUTES} minute(s), retention ${FILE_RETENTION_HOURS} hour(s)`
  );
});