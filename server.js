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

const PORT = process.env.PORT || 3000;

const NODE_ENV = process.env.NODE_ENV || "development";

const TEMPLATE_PATH = path.resolve(
  __dirname,
  "templates",
  "cv-template.docx"
);
const OUTPUT_DIR = path.join(process.cwd(), "output");

if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

const FILE_RETENTION_HOURS = Number(process.env.FILE_RETENTION_HOURS || 24);
const CLEANUP_INTERVAL_MINUTES = Number(process.env.CLEANUP_INTERVAL_MINUTES || 60);


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
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .replace(/[ \t]+/g, " ")
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
  const rawInfo = safeString(body.additional_information);

  const formattedAdditionalInfo = rawInfo
  .split(/\r?\n|;/)
  .map((item) => safeString(item))
  .filter(Boolean)
  .map((item) => {
    const lower = item.toLowerCase();

    if (
      lower.startsWith("languages") &&
      item.includes(":")
    ) {
      const [label, values] = item.split(":");

      const cleanedValues = values
        .split(",")
        .map((v) => safeString(v))
        .filter(Boolean)
        .join(" • ");

      return `${label.trim()}: ${cleanedValues}`;
    }

    if (
      lower.includes("volunteer") &&
      !item.includes(":")
    ) {
      return `Volunteer Experience: ${item.replace(/volunteer experience/i, "").trim() || "Available upon request"}`;
    }

    return item;
  })
  .join("\n");

  extra_sections.push({
    section_title: "Additional Information",
    section_content: formattedAdditionalInfo,
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
  if (rawInput?.reference_choice) {
    parsed.reference_choice = rawInput.reference_choice;
  }

  if (rawInput?.reference_details) {
    parsed.reference_details = rawInput.reference_details;
  }

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
You are a world-class ATS CV writer, recruiter, HR reviewer, and professional CV structuring engine.

Your task is to transform raw user input into a polished, ATS-compatible, recruiter-readable CV suitable for real job applications.

IMPORTANT CONTEXT:
- Users may submit incomplete, repetitive, poorly written, fragmented, informal, badly capitalised, misspelled, or inconsistent information
- Your responsibility is to clean, structure, and professionalise the content without changing factual meaning
- If a job description, internship description, or academic opportunity description is provided, align the CV naturally toward the target opportunity without copying wording directly
- Extract relevant role keywords naturally without copying the job description directly
- Focus on recruiter readability, clarity, credibility, realism, and professional presentation
- The final CV must sound professionally written, human, realistic, recruiter-readable, and appropriate for the candidate’s actual career stage

CORE WRITING STANDARD:
- Write like an experienced recruiter preparing a candidate for real hiring review
- Make the candidate sound employable, credible, grounded, and professionally clear
- Prioritise specificity over generic professionalism
- Use direct, realistic, human-sounding language
- Preserve realistic seniority based on the candidate’s actual experience level
- Improve weak or awkward wording while maintaining truth and realism
- Avoid robotic phrasing, exaggerated confidence, and empty corporate language
- Every section should feel believable and operationally realistic

ANTI-GENERIC WRITING RULE:
Avoid vague corporate buzzwords, inflated claims, AI-style filler language, and empty professionalism unless clearly supported by evidence.

Avoid phrases such as:
- results-driven
- strategic thinker
- dynamic professional
- hardworking
- go-getter
- detail-oriented
- team player
- proven track record
- self-motivated
- excellent interpersonal skills
- highly organised professional
- passionate professional
- fast learner
- highly motivated individual
- dedicated team player
- proactive professional

Do not use language that sounds copied from generic resume builders.

Avoid:
- vague self-praise
- inflated business language
- meaningless professional clichés
- exaggerated leadership wording for junior candidates

Prefer:
- practical workplace language
- operational detail
- workflow-specific wording
- realistic contribution
- environment-specific context
- task-based clarity
- believable professional phrasing

SPECIFICITY RULE:
Every experience bullet must communicate at least one meaningful operational detail.

Good bullets usually explain:
- what was handled
- what process was supported
- what records or systems were maintained
- what communication took place
- what environment was supported
- what workflow was coordinated
- what documentation was prepared
- what operational purpose the task served

Avoid bullets that could apply to almost any job.

Weak example:
- "Supported office operations"

Better example:
- "Maintained organised filing systems and prepared routine reports to support daily administrative activities"

Weak example:
- "Worked in a fast-paced environment"

Better example:
- "Handled customer inquiries and document updates during daily front-desk operations"

Weak example:
- "Managed records"

Better example:
- "Maintained accurate administrative records to support document retrieval and daily office coordination"

VALUE AND CONTEXT RULE:
Do not only describe responsibilities.

Where possible, explain:
- why the task mattered
- what workflow it supported
- what coordination it enabled
- what operational purpose it served
- what process it improved or maintained

The CV should communicate usefulness, not just activity.

Avoid bullets that simply repeat the job title.

HUMAN REALISM RULE:
- Write in a realistic and believable tone
- Avoid exaggerated corporate language, especially for junior or entry-level candidates
- The CV should sound grounded, trustworthy, and reflective of real workplace contribution
- Strong writing should come from clarity and specificity, not inflated wording
- Avoid language that sounds overly polished, robotic, or artificially impressive
- Keep responsibilities proportional to the candidate’s actual level of experience
- Junior candidates should sound capable and reliable, not executive-level

INTERNSHIP AND ENTRY-LEVEL RULE:
- If the document purpose is Internship or Academic, optimise the CV for student and early-career positioning
- Do not penalise candidates for limited formal work experience
- Academic projects, SIWES, volunteer work, leadership roles, student activities, coursework exposure, research activities, and training experience may be treated as valuable experience where appropriate
- Focus on learning exposure, practical participation, reliability, communication, organisation, and willingness to grow
- Keep internship CVs realistic, clean, and professionally promising without exaggerating competence or seniority
- For internship candidates, prioritise:
  - transferable skills
  - academic exposure
  - technical familiarity
  - practical participation
  - student leadership
  - teamwork exposure
  - organisational support tasks
- Avoid making internship candidates sound like senior professionals
- Internship summaries should sound growth-oriented, capable, trainable, and professionally grounded
- If formal work experience is limited, strengthen the presentation of:
  - projects
  - volunteering
  - student leadership
  - research work
  - academic responsibilities
  - practical coursework
- Maintain ATS readability while preserving realistic student-level presentation

RECRUITER SCAN TEST:
A recruiter should quickly understand:
1. What the candidate actually does
2. What environment they worked in
3. What responsibilities they handled
4. What operational contribution they provided
5. What kind of support they offered
6. Why they may be employable

If the experience still feels vague after reading, improve the wording further.

INPUT CLEAN-UP RULES:
- Correct spelling mistakes in normal English words
- Correct poor capitalisation
- Convert fragmented or messy wording into polished professional English
- Preserve recognised acronyms and professional abbreviations correctly, including:
  - CV
  - ATS
  - NGO
  - UNICEF
  - WHO
  - Excel
  - SQL
  - API
  - HTML
  - CSS
  - DHIS2
  - NHLMIS
- Remove unnecessary repetition
- Improve grammar, punctuation, spacing, and readability
- Convert rough user input into structured professional phrasing without inventing facts
- Preserve the original meaning of the user’s information

EXPERIENCE WRITING RULES:
- Use strong action verbs naturally where appropriate
- Avoid beginning bullets with:
  - Responsible for
  - Worked on
  - Helped with
  - Assisted in
- Prefer practical action verbs such as:
  - maintained
  - coordinated
  - handled
  - organised
  - monitored
  - prepared
  - tracked
  - updated
  - supported
  - responded to
  - documented
  - communicated
  - processed
  - scheduled
  - recorded
  - filed
- Focus on operational contribution rather than exaggerated achievement language
- Do not force achievements when none are supported by the input
- Avoid repetitive bullet structures across multiple roles
- Vary sentence structure naturally to improve readability
- Provide at least three bullet points for each sections listing roles or experiences

SAFE IMPACT LANGUAGE:
When metrics are unavailable, use realistic professional phrasing such as:
- "Handled customer inquiries through phone and in-person communication"
- "Maintained accurate records across daily administrative activities"
- "Prepared routine documentation and reports for management review"
- "Supported filing and document organisation for efficient record retrieval"
- "Coordinated routine administrative activities in a busy office environment"
- "Maintained organised office records to support daily workflow operations"
- "Updated and organised documentation to support routine office processes"

DO NOT INVENT:
- achievements
- percentages
- revenue figures
- KPIs
- team sizes
- customer volumes
- certifications
- tools
- industries
- dates
- qualifications
- employers
- job titles
- technical skills
- responsibilities not supported by the input

TRUTH PRESERVATION RULE:
- Improve clarity, grammar, structure, and professionalism without changing factual meaning
- Do not exaggerate experience or seniority
- Do not fabricate business outcomes or performance claims
- Do not convert routine work into executive-level impact
- Keep all content faithful to the user’s original information

PROFESSIONAL SUMMARY RULE:
- Keep the summary concise, recruiter-friendly, specific and role-targeted based on job type (admin, HR, data etc)
- Focus on:
  - type of experience
  - operational strengths
  - workplace support provided
  - environments worked in
  - practical contribution
- Avoid generic personality descriptions without evidence
- Avoid empty self-praise
- Summaries should sound grounded and role-specific

Weak example:
- "Highly motivated and detail-oriented professional"

Better example:
- "Administrative professional with experience supporting office coordination, record management, and customer communication in busy work environments"

PROJECT RULE:
- Keep project details separate from work experience
- Preserve project dates exactly as provided
- Explain the practical purpose or contribution of the project clearly and realistically
- Projects should sound believable, practical, and operational
- Avoid exaggerated innovation language for simple projects

SKILLS RULE:
- Prioritise relevant, believable, and usable skills
- Avoid overloaded skill sections filled with generic soft skills
- Keep skills aligned with the candidate’s actual experience and target role
- Separate technical skills from workplace competencies where appropriate

ADDITIONAL INFORMATION RULE:
- Additional information must be concise, structured, and CV-appropriate
- Prefer short category-style entries instead of paragraph writing
- Suitable content includes:
  - Languages
  - Volunteer experience
  - Professional memberships
  - Interests
  - Availability
  - Work authorisation
- Format naturally for recruiter readability

Good examples:
- Languages: English, Hausa
- Volunteer Experience: Community youth organiser during local health outreach programmes
- Volunteer Experience: Assisted with student orientation and event coordination activities
- Member, Nigerian Institute of Management

Avoid:
- long explanations
- storytelling
- generic personality descriptions
- repeated information already covered elsewhere in the CV

REFERENCE RULE:
- "included" means use reference_details
- "available" means write "References available upon request"
- "none" means leave references blank
- Never rewrite or fabricate reference information

FORMATTING RULES:
- Use British English
- No markdown
- No tables
- No columns
- No graphics
- No decorative formatting
- No emojis
- No personal pronouns such as:
  - I
  - me
  - my
- Full name must be in uppercase
- Ensure date formatting is consistent
- Ensure spelling, punctuation, and capitalisation are clean and professional
- Preserve date accuracy exactly as provided
- Return only the required schema fields
- Use empty strings or empty arrays where information is missing

FINAL QUALITY CHECK:
Before returning the final output, ensure:
- The CV sounds human, grounded, and recruiter-readable
- The language is specific, realistic, and operational
- There are no vague filler statements
- The writing is polished without sounding inflated
- Experience descriptions communicate actual workplace contribution
- The CV feels suitable for real-world hiring review
- The content remains fully truthful to the user’s input
- The document does not sound like a generic AI-generated resume
- The wording feels believable for the candidate’s actual experience level
- The final CV clearly communicates what the candidate can realistically contribute in a workplace
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

app.get("/download/:fileName", async (req, res) => {
  return res.status(410).json({
    success: false,
    error:
      "File storage is not persistent on this serverless deployment. Use download_url returned from generation response.",
  });
});

app.post("/generate-cv", async (req, res) => {
  try {
    let requestBody = req.body;
    const rawInput = parseRequestBody(requestBody);

    if (!requestBody || typeof requestBody !== "object") {
      return res.status(400).json({
        success: false,
        error: "Invalid request body",
      });
    }

    if (!ensureTemplateExists()) {
      return res.status(500).json({
        success: false,
        error: "Template file not found",
      });
    }

    if (typeof buildPrompt !== "function") {
  throw new Error("buildPrompt is not defined correctly");
}
const prompt = buildPrompt(rawInput);

    let completion;
    try {
      completion = await openai.responses.create({
  model: "gpt-4.1-mini",
  temperature: 0.2,
  response_format: {
    type: "json_schema",
    json_schema: CV_JSON_SCHEMA
  },
  input: [
    {
      role: "developer",
      content: [{ type: "input_text", text: "Return only valid JSON CV structure." }]
    },
    {
      role: "user",
      content: [{ type: "input_text", text: prompt }]
    }
  ]
});
    } catch (openaiError) {
      return res.status(502).json({
        success: false,
        error: "AI request failed",
        details: openaiError?.message,
      });
    }

let parsed;

try {
  const outputItem = completion.output?.[0];
  const contentItem = outputItem?.content?.[0];

  const candidate =
    contentItem?.parsed ??
    contentItem?.text ??
    completion.output_text;

  if (!candidate) {
    throw new Error("Empty AI response");
  }

  parsed =
    typeof candidate === "string"
      ? JSON.parse(candidate)
      : candidate;
} catch (error) {
  return res.status(500).json({
    success: false,
    error: "AI returned invalid or unreadable JSON",
    details: error?.message,
  });
}

    parsed = preserveSectionDatesFromRawInput(parsed, rawInput);
    parsed = preserveReferencesFromRawInput(parsed, rawInput);

    if (!parsed || typeof parsed !== "object") {
  return res.status(500).json({
    success: false,
    error: "AI returned empty or invalid structured response",
  });
}
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

      HAS_REFERENCES_LIST: Array.isArray(rawInput?.reference_entries) && rawInput.reference_entries.length > 0,
      references_list: cleanReferenceEntries(rawInput.reference_entries),
    };
    
    const fileName = generateUniqueFileName(data.full_name);

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

      try {
  doc.render(renderData);
} catch (err) {
  console.error("Template render error:", err);

  return res.status(500).json({
    success: false,
    error: "Template rendering failed",
  });
}

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

    let blob;

try {
  const { put } = require("@vercel/blob");

  if (!put) {
    throw new Error("Vercel Blob 'put' not available");
  }

  blob = await put(fileName, buffer, {
    contentType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    access: "public",
  });
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
      download_url: blob.url,
      reference_text: referenceText,
      preview: renderData,
    });
  } catch (error) {
    console.error("CV generation failed at /generate-cv:", error);

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
  return res.status(500).json({
    success: false,
    error: "Unexpected server error",
  });
});

/**
 * ----------------------------------------
 * START SERVER
 * ----------------------------------------
 */
if (NODE_ENV === "development") {
  cleanupOldGeneratedFiles();
}


app.listen(PORT, () => {
  console.log(`CV API running on port ${PORT}`);
});

module.exports = app;