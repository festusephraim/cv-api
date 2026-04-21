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

/**
 * Cleanup settings
 * DELETE files older than this many hours
 */
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
  return String(value).trim();
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
  return safeString(value).replace(/\s+/g, " ");
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
      .split(",")
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

function slugifyFileName(value) {
  return safeString(value)
    .toLowerCase()
    .replace(/[^a-z0-9\s-]/g, "")
    .replace(/\s+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "");
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
You are an expert ATS CV writer and CV structuring engine.

Your task is to convert raw user input into a highly professional, ATS-optimised CV structure.

IMPORTANT CONTEXT:
- This CV may be used for real job applications
- If a job_description is provided, tailor the CV to it
- Extract and align important keywords from the job description
- Do NOT copy the job description directly
- Naturally integrate relevant keywords into the professional summary, skills, role summaries, project descriptions, and experience tasks
- Do NOT invent information, qualifications, dates, tools, industries, achievements, or metrics not supported by the input

STRICT RULES:
- Use clear, simple, professional English. Strictly British English
- No tables, no columns, no graphics
- ATS-friendly wording only
- Each experience task must begin with a strong action verb
- Avoid weak phrases like "Responsible for"
- Each task should show action, contribution, scope, or outcome where possible
- Keep professional_summary concise and recruiter-friendly
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

METRICS AND IMPACT RULE:
- Use measurable details when they are explicitly stated or clearly supported by the input
- Preserve real numbers, counts, frequencies, tools, timelines, workloads, team sizes, customer volumes, or output volumes when provided
- If the input suggests scale or frequency but gives no exact figures, write realistic impact language without inventing numbers
- Good examples of safe phrasing without fabricated figures include:
  - "Handled high-volume customer interactions"
  - "Maintained accurate records across daily operations"
  - "Supported timely report preparation and documentation"
  - "Coordinated routine administrative tasks in a fast-paced environment"
- Do NOT create percentages, revenue figures, growth rates, rankings, time savings, or exact counts unless the user input supports them
- Do NOT exaggerate achievements
- Where evidence is limited, prefer truthful contribution-focused wording over artificial quantified claims

REFERENCE RULE:
- "included" means use reference_details
- "available" means references available upon request
- "none" means blank reference section

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

function generateFileName(fullName) {
  const safeName = slugifyFileName(fullName) || "cv";
  const timestamp = Date.now();
  return `${safeName}-cv-${timestamp}.docx`;
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

// Health check
app.get("/", (req, res) => {
  return res.status(200).json({
    success: true,
    message: "CV API is running",
    environment: NODE_ENV,
    template_exists: ensureTemplateExists(),
  });
});

// Download route
app.get("/download/:file", (req, res) => {
  try {
    const fileName = path.basename(req.params.file);
    const filePath = path.join(OUTPUT_DIR, fileName);

    if (!fs.existsSync(filePath)) {
      return res.status(404).json({
        success: false,
        error: "File not found",
      });
    }

    return res.download(filePath);
  } catch (error) {
    return res.status(500).json({
      success: false,
      error: "Download failed",
    });
  }
});

// Main generation route
app.post("/generate-cv", async (req, res) => {
  try {
    cleanupOldGeneratedFiles();

    const incomingError = validateIncomingBody(req.body);
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

    const rawInput = normalizeIncomingPayload(req.body);
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

    const fileName = generateFileName(data.full_name);
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

    return res.status(200).json({
      success: true,
      message: "CV generated successfully",
      file_name: fileName,
      download_url: `${BASE_URL}/download/${fileName}`,
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
}, CLEANUP_INTERVAL_MINUTES * 60 * 60 * 1000);

app.listen(PORT, () => {
  console.log(`CV API running on ${BASE_URL}`);
  console.log(`Template exists: ${ensureTemplateExists()}`);
  console.log(
    `File cleanup: every ${CLEANUP_INTERVAL_MINUTES} minute(s), retention ${FILE_RETENTION_HOURS} hour(s)`
  );
});