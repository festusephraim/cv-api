require("dotenv").config();

const express = require("express");
const cors = require("cors");
const fs = require("fs");
const path = require("path");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const { OpenAI } = require("openai");

const axios = require("axios");
const pdfParse = require("pdf-parse");
const mammoth = require("mammoth");

const { S3Client, PutObjectCommand, GetObjectCommand } = require("@aws-sdk/client-s3");
const { getSignedUrl } = require("@aws-sdk/s3-request-presigner");

const app = express();

app.set("trust proxy", 1);

app.disable("x-powered-by");

process.on("uncaughtException", (err) => {
  console.error("UNCAUGHT EXCEPTION:", err);
});

process.on("unhandledRejection", (reason) => {
  console.error("UNHANDLED REJECTION:", reason);
});

app.use(cors());
app.use(express.json({ limit: "10mb" }));

const rateLimit = require("express-rate-limit");

const apiLimiter = rateLimit({
  windowMs: 60 * 1000, // 1 minute
  max: 10, // limit each IP to 10 requests per minute
  standardHeaders: true,
  legacyHeaders: false,
});

app.use(apiLimiter);

const PORT = Number(process.env.PORT || 3001);
const BASE_URL = process.env.BASE_URL || `http://localhost:${PORT}`;
const NODE_ENV = process.env.NODE_ENV || "development";

function getTemplatePath(candidateLevel, experienceCount) {

  if (candidateLevel === "executive candidate") {
    return path.join(
      process.cwd(),
      "templates",
      "template_executive.docx"
    );
  }

  if (candidateLevel === "professional candidate") {
    return path.join(
      process.cwd(),
      "templates",
      "template_professional.docx"
    );
  }

  return path.join(
    process.cwd(),
    "templates",
    "cv-template.docx"
  );
}


if (!process.env.OPENAI_API_KEY) {
  console.error("Missing OPENAI_API_KEY in environment variables.");
  process.exit(1);
}

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

const s3 = new S3Client({
  region: process.env.AWS_REGION,
  credentials: {
    accessKeyId: process.env.AWS_ACCESS_KEY_ID,
    secretAccessKey: process.env.AWS_SECRET_ACCESS_KEY,
  },
});

function safeString(value) {
  if (value === null || value === undefined) return "";

  const cleaned = String(value)
    .replace(/\u00A0/g, " ")
    .replace(/[“”]/g, '"')
    .replace(/[‘’]/g, "'")

    // normalize line endings
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")

    // remove excessive line breaks
    .replace(/\n{2,}/g, "\n")

    // remove tabs/spaces
    .replace(/[ \t]+/g, " ")

    // trim spaces around line breaks
    .replace(/ *\n */g, "\n")

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

function normaliseBulletArray(items) {
  return safeArray(items)
    .map((item) => safeString(item))
    .filter(Boolean);
}

function splitLinesToArray(value) {
  return safeString(value)
    .split(/\r?\n|•|;/)
    .map((item) => safeString(item))
    .filter(Boolean);
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
      const line1 = [entry.name, entry.position]
        .filter(Boolean)
        .join(", ");

      const line2 = [entry.organization, entry.location]
        .filter(Boolean)
        .join(", ");

      let contactLine = "";

      if (entry.email && entry.phone) {
        contactLine = `Email: ${entry.email}, Phone: ${entry.phone}`;
      } else if (entry.email) {
        contactLine = `Email: ${entry.email}`;
      } else if (entry.phone) {
        contactLine = `Phone: ${entry.phone}`;
      }

      return [line1, line2, contactLine]
        .filter(Boolean)
        .join(" | ");
    })
    .join("\n\n");
}

function extractReferencesFromCvText(cvText) {

  console.log(
    "extractReferencesFromCvText CALLED"
  );

  const text = safeString(cvText);

  if (!text) {
    return [];
  }

  const referencesSectionMatch =
  text.match(
    /(references|referees)([\s\S]*)$/i
  );

if (!referencesSectionMatch) {
  return [];
}

const referencesText =
  referencesSectionMatch?.[2] || "";
console.log(
  "REFERENCES MATCH:",
  !!referencesSectionMatch
);

  console.log(
  "REFERENCES TEXT:"
);

console.log(
  referencesText
);

console.log(
  "REFERENCES TEXT:"
);

console.log(
  referencesText.substring(0, 1000)
);

console.log(
  "LOOKING FOR REFERENCES SECTION"
);

  const phoneRegex =
    /(\+?\d[\d\s()-]{7,})/g;

  const phones =
    [...referencesText.matchAll(phoneRegex)];

  const entries = [];

  for (let i = 0; i < phones.length; i++) {
    const phone =
      phones[i][0].trim();

    const startIndex =
      i === 0
        ? 0
        : phones[i - 1].index +
          phones[i - 1][0].length;

    const endIndex =
      phones[i].index;

    const block =
      referencesText
        .substring(startIndex, endIndex)
        .trim();

    const lines =
      block
        .split("\n")
        .map((line) => line.trim())
        .filter(Boolean);

   let name = "";
let position = "";
let organization = "";
let location = "";

const combinedText = lines
  .join(" ")
  .replace(/\s+/g, " ")
  .trim();

const parts = combinedText
  .split(",")
  .map((p) => p.trim())
  .filter(Boolean);

name = parts[0] || "";
position = parts[1] || "";

let remaining = parts.slice(2).join(", ");

const knownLocations = [
  "Port Harcourt",
  "Lagos",
  "Abuja",
  "Warri",
  "Eket",
  "Yenagoa",
  "Bonny",
  "Onne",
  "Owerri",
  "Uyo"
];

const matchedLocation = knownLocations.find(
  (city) =>
    remaining
      .toLowerCase()
      .endsWith(city.toLowerCase())
);

if (matchedLocation) {
  location = matchedLocation;

  organization = remaining
    .replace(
      new RegExp(matchedLocation, "i"),
      ""
    )
    .replace(/\s+/g, " ")
    .trim();
} else {
  organization = remaining;
}

if (!position && lines[1]) {
  position = lines[1];
}

if (!organization && lines[2]) {
  organization = lines[2];
}

console.log(
  "REFERENCE PARSED:"
);

console.log({
  name,
  position,
  organization,
  location,
  email: "",
  phone,
});

entries.push({
  name,
  position,
  organization,
  location,
  email: "",
  phone,
});

  }
  console.log(
  "RAW REFERENCE ENTRIES:"
);

console.log(entries);

  return cleanReferenceEntries(entries);
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

let fileCounter = 0;

function generateUniqueFileName(fullName) {
  const cleanName = cleanDisplayName(fullName);

  let fileName;

  if (fileCounter === 0) {
    fileName = `${cleanName} CV.docx`;
  } else {
    fileName = `${cleanName} CV (${fileCounter}).docx`;
  }

  fileCounter++;

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

  console.log(
    "ENTERED normalizeIncomingPayload"
  );

  console.dir(body, {
  depth: null
});
  const basicInfo = body?.basic_information || {};

  console.log("NORMALIZE AFTER BASICINFO");

  const candidateLevel = safeString(
  body?.candidate_level
).toLowerCase();

const workExperience = safeArray(
  body?.work_experience
).filter(
  (item) =>
    safeString(item?.job_title) ||
    safeString(item?.company) ||
    safeString(item?.what_did_you_do_in_this_role)
);

const education = safeArray(
  body?.education
).filter(
  (item) =>
    safeString(item?.degree_qualification) ||
    safeString(item?.school)
);

const projects = safeArray(
  body?.projects_research
).filter(
  (item) =>
    safeString(item?.project_title) ||
    safeString(item?.project_description)
);

console.log("BEFORE REFERENCE CLEAN");

  let referenceEntries =
  cleanReferenceEntries(
    body?.references?.reference_entries
  );

console.log("AFTER REFERENCE CLEAN");

console.log(
  "REFERENCE COUNT BEFORE EXTRACTION:",
  referenceEntries.length
);

console.log(
  "HAS uploaded_cv_text:",
  !!body?.uploaded_cv_text
);

if (
  referenceEntries.length === 0 &&
  body?.uploaded_cv_text
) {
  console.log("ABOUT TO EXTRACT REFERENCES");
  referenceEntries =
    extractReferencesFromCvText(
      body.uploaded_cv_text
    );
}

const builtReferenceDetails =
  buildReferenceDetailsFromEntries(
    referenceEntries
  );

  console.log(
  "REFERENCE DETAILS AFTER NORMALIZATION:"
);
console.log(builtReferenceDetails);

console.log(
  "REFERENCE ENTRIES AFTER EXTRACTION:"
);
console.log(referenceEntries);

let reference_choice =
  normaliseReferenceChoice(
    body?.references_section_preference
  );

if (
  reference_choice === "included" &&
  referenceEntries.length === 0
) {
  reference_choice = "available";
}
  const mappedExperience = workExperience.map((item) => ({
  title: safeString(item?.job_title),
  company: safeString(item?.company),
  location: safeString(item?.location),
  start: safeString(item?.start_date),

  end: item?.currently_working_here
    ? "Present"
    : safeString(item?.end_date),

  role_summary: "",
  tasks: splitLinesToArray(
    item?.what_did_you_do_in_this_role,
    5
  ),
}));

const shouldUseEduCompetencies =
  mappedExperience.length < 1;

const mappedEducation = education.map((item) => ({
  degree: safeString(item?.degree_qualification),
  school: safeString(item?.school),
  location: safeString(item?.location),
  start: safeString(item?.start_date),

  end: item?.currently_studying_here
    ? "Present"
    : safeString(item?.end_date),

  edu_detail: safeString(item?.grade_result),
  edu_competencies: [],
}));

const mappedProjects = projects.map((item) => ({
  project_title: safeString(item?.project_title),
  project_description: safeString(item?.project_description),
  start: safeString(item?.start_date),

  end: item?.currently_working_on_this_project
    ? "Present"
    : safeString(item?.end_date),

  project_tasks: splitLinesToArray(
    item?.what_did_you_do_in_this_project,
    5
  ),
}));

  const extra_sections = [];

if (safeString(body?.additional_information)) {
  const rawInfo = safeString(body.additional_information);

  // ---------- LANGUAGES ----------
  const languageMatches = rawInfo.match(
    /(english|hausa|yoruba|igbo|french|kuteb|jukun-takum)/gi
  );

  if (languageMatches?.length) {
    const languageItems = [...new Set(languageMatches)].map((lang) => {
      const cleanLang = cleanDisplayName(lang);

      let level = "Conversational";

      if (cleanLang.toLowerCase() === "english") {
        level = "Fluent";
      }

      return `${cleanLang} (${level})`;
    });

    extra_sections.push({
      section_title: "LANGUAGES",
      items: languageItems,
      section_content: "",
    });
  }

  // ---------- VOLUNTEER EXPERIENCE ----------
  if (
    rawInfo.toLowerCase().includes("volunteer") ||
    rawInfo.toLowerCase().includes("campaign") ||
    rawInfo.toLowerCase().includes("outreach")
  ) {
    extra_sections.push({
      section_title: "VOLUNTEER EXPERIENCE",
      items: [
        "Participated in peace building campaign activities in southern Taraba State",
      ],
      section_content: "",
    });
  }
}

  return {
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
          .split(/\r?\n+|•|;/)
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
  return safeArray(experience)
    .map((item) => ({
      title: safeString(item?.title),
      company: safeString(item?.company),
      location: safeString(item?.location),
      start: safeString(item?.start),
      end: safeString(item?.end),
      end_or_present: endOrPresent(item?.end),
      role_summary: safeString(item?.role_summary),
      tasks: normaliseBulletArray(item?.tasks, 20),
    }))
    .filter(
      (item) =>
        item.title ||
        item.company ||
        item.tasks.length
    );
}

function cleanProjectsArray(projects) {
  return safeArray(projects)
    .map((item) => ({
      project_title: safeString(item?.project_title),
      project_description: safeString(item?.project_description),
      start: safeString(item?.start),
      end: safeString(item?.end),
      end_or_present: endOrPresent(item?.end),
      project_tasks: normaliseBulletArray(
        item?.project_tasks,
        20
      ),
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
  return safeArray(education)
    .map((item) => ({
      degree: safeString(item?.degree),
      school: safeString(item?.school),
      location: safeString(item?.location),
      start: safeString(item?.start),
      end: safeString(item?.end),
      end_or_present: endOrPresent(item?.end),
      edu_detail: safeString(item?.edu_detail),

      edu_competencies: normaliseBulletArray(
        item?.edu_competencies,
        20
      ),
    }))
    .filter(
      (item) =>
        item.degree ||
        item.school
    );
}

function cleanCertificationsArray(certifications) {
  return safeArray(certifications)
    .map((item) => safeString(item))
    .filter(Boolean);
}

function cleanExtraSections(extraSections) {
  return safeArray(extraSections)
    .map((item) => ({
      section_title: safeString(
       item?.section_title
        ).toUpperCase(),
      items: normaliseBulletArray(
        item?.items,
        50
      ),
      section_content: safeString(
        item?.section_content
      ),
    }))
    .filter(
      (item) =>
        item.section_content ||
        (
          Array.isArray(item.items) &&
          item.items.length > 0
        )
    );
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
references_list: cleanReferenceEntries(data.references_list),
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
  start:
    safeString(rawExperience[index]?.start) ||
    safeString(item?.start),

  end:
    safeString(rawExperience[index]?.end) ||
    safeString(item?.end),
}));

  parsed.education = parsedEducation.map((item, index) => ({
  ...item,
  start:
    safeString(rawEducation[index]?.start) ||
    safeString(item?.start),

  end:
    safeString(rawEducation[index]?.end) ||
    safeString(item?.end),
}));

  parsed.projects = parsedProjects.map((item, index) => ({
  ...item,
  start:
    safeString(rawProjects[index]?.start) ||
    safeString(item?.start),

  end:
    safeString(rawProjects[index]?.end) ||
    safeString(item?.end),
}));

  return parsed;
}

function preserveReferencesFromRawInput(parsed, rawInput) {

  if (
    !safeString(parsed.reference_details)
  ) {
    parsed.reference_details =
      rawInput.reference_details;
  }

  if (
    !safeString(parsed.reference_choice)
  ) {
    parsed.reference_choice =
      rawInput.reference_choice;
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

Your task is to transform raw user input into a polished humanly written, ATS-compatible, recruiter-readable CV suitable for real job applications.

IMPORTANT CONTEXT:
- Users may submit incomplete, repetitive, poorly written, fragmented, informal, badly capitalised, misspelled, or inconsistent information
- Your responsibility is to clean, structure, and professionalise the content keeping it highly human without changing factual meaning
- If a job description, internship description, or academic opportunity description is provided, align the CV naturally toward the target opportunity without copying wording directly. Also ensure all words and sentences reveal a human behind the text.
- Extract relevant role keywords naturally without copying the job description directly, mirroring how best a human would think things through and connects the dots. 
- Focus on recruiter readability, clarity, credibility, realism, and professional presentation that is highly human and practical. 
- The final CV must sound professionally written, human, realistic, recruiter-readable, and appropriate for the candidate’s actual career stage
- If there are no experience, prioritise stronger edu_competencies generation
- If experience entry is 1 or more, edu_competencies should be empty

- If references or referees are present in the uploaded CV, extract them and include them in reference_details.
- Do not replace existing referee information with "References available upon request" when referee details exist.
- Preserve referee names, positions, organizations, phone numbers and email addresses whenever available.

TEMPLATE AUTHORITY RULE

The application supplies a candidate_level value.

candidate_level is already determined before CV generation.

You are not responsible for deciding whether the candidate is entry, professional, or executive.

You must strictly follow the supplied candidate_level.

Do not downgrade a professional candidate to entry level.

Do not upgrade an entry-level candidate to professional.

Do not override candidate_level based on your own interpretation.

ENTRY LEVEL TEMPLATE
Use this when:
- student, intern, SIWES, NYSC, fresh graduate
- Has no formal work experience

PROFESSIONAL TEMPLATE
Use this when:
- clear job history exists
- candidate has real workplace responsibility
- multiple roles or structured employment experience

STRICT RULE
- Only ONE template is active per CV
- Do NOT mix tones
- Do NOT “upgrade” entry-level candidates
- Do NOT “simplify” professional candidates

ENTRY LEVEL TEMPLATE

Used only when candidate_level = entry.

PROFESSIONAL TEMPLATE

Used only when candidate_level = professional.

EXECUTIVE TEMPLATE

Used only when candidate_level = executive.

IMPORTANT:
This layer executes BEFORE CV writing begins.

This is not a formatting phase.
This is not a grammar correction phase.

This is the strategic interpretation and recruiter-alignment phase.
The objective is to determine:
* what role the candidate is actually competing for,
* how recruiters will interpret the profile,
* how ATS systems will classify relevance,
* and how the candidate should be strategically positioned before writing starts.

If candidate_level = professional,
maintain professional-level structure and content throughout the document.

If candidate_level = executive,
maintain executive-level structure and content throughout the document.

Do NOT immediately generate CV content.

The system must first perform recruiter-style interpretation logic internally.

TEMPLATE CONSISTENCY CHECK

Before generating output:

If candidate_level = professional:

- edu_competencies must be empty
- work experience must be preserved
- technical skills must be preserved
- professional certifications must be preserved

If candidate_level = executive:

- edu_competencies must be empty
- work experience must be preserved
- executive-level experience must not be compressed into entry-level wording

If candidate_level = entry:

- edu_competencies may be generated
- education may be expanded appropriately
---

## CORE PRINCIPLE

Expert ATS CV work is NOT primarily writing.

It is:

* interpretation skill,
* hiring logic understanding,
* role alignment analysis,
* ATS relevance engineering,
* recruiter psychology,
* and strategic positioning.

Writing is only the final execution layer.

The system must therefore think like:

* a recruiter,
* a hiring manager,
* an ATS parser,
* and a career strategist simultaneously.

---

## STEP 1 — JOB DECONSTRUCTION

If a job description, internship description, scholarship opportunity, or target role is provided:

Internally decode the opportunity before writing.

Extract and analyse:

A. CORE SKILLS
Identify:

* mandatory technical skills,
* operational responsibilities,
* workflow expectations,
* role-critical competencies.

B. SUPPORTING SKILLS
Identify:

* secondary capabilities,
* coordination expectations,
* communication requirements,
* reporting responsibilities,
* administrative or collaborative functions.

C. ATS TRIGGER TERMS
Extract:

* repeated keywords,
* systems,
* software,
* platforms,
* certifications,
* tools,
* operational phrases,
* industry terminology.

IMPORTANT:
Do NOT simply copy keywords.
Understand their operational meaning first.

D. SENIORITY SIGNALS
Determine whether the role is:

* internship,
* entry-level,
* junior,
* mid-level,
* specialist,
* supervisory,
* managerial,
* executive.

Never inflate seniority beyond realistic real-life evidence.

E. INDUSTRY LANGUAGE PATTERN
Observe:
* how the employer communicates,
* vocabulary style,
* operational tone,
* performance expectations,
* role framing logic.

Mirror language naturally without copying the job description directly and ensure it is highly human and makes complete sense.

---

## STEP 2 — CANDIDATE MAPPING

Map the candidate to the opportunity using 3 layers:

LAYER 1 — DIRECT ALIGNMENT
Identify:

* matching responsibilities,
* same tools,
* same workflows,
* same industry exposure,
* directly relevant tasks.

LAYER 2 — TRANSFERABLE ALIGNMENT
Identify:

* adjacent operational experience,
* similar coordination logic,
* related systems exposure,
* transferable technical capability,
* similar reporting or workflow structures.

LAYER 3 — HIDDEN VALUE
Identify:

* reliability signals,
* communication ability,
* adaptability,
* learning exposure,
* leadership indicators,
* organisational contribution,
* operational awareness,
* analytical support,
* customer interaction,
* workflow continuity support.

Interpret intelligently while preserving truth and keep it practically human.

---

## STEP 3 — POSITIONING STRATEGY

Determine the strongest believable professional identity for the candidate.

Examples:

* Production Maintenance Engineer
* Administrative Officer
* Graduate Trainee
* Customer Support Assistant
* Front Desk Administrator
* Laboratory Assistant
* Data Entry Clerk

Avoid vague identities such as:

* hardworking individual
* team player
* proactive professional
* dynamic candidate

The positioning identity must be:

* believable,
* recruiter-friendly,
* ATS-relevant,
* evidence-supported,
* and appropriate for the candidate’s actual level.

IMPORTANT:
The system must position candidates based on employability reality, not aspirational exaggeration.

---

## STEP 4 — EXPERIENCE PRIORITISATION

Not all experiences carry equal strategic value.

Prioritise experiences based on:

* Selected Achievements - USE ALL THE SELECTED ACHIEVEMENTS BEFORE ADDING OTHER RESPONSIBILITIES IF NECESSARY 
* relevance to target role,
* operational contribution,
* continuity,
* ATS relevance density,
* credibility,
* and positioning strength.

Older experiences may be summarised ONLY when doing so does not remove candidate evidence, achievements, certifications, technical skills, leadership experience, or relevant accomplishments.

The objective is relevance density, human writing, not document length.

More information does not automatically equal stronger positioning. Improve clarity without removing relevant candidate evidence.

---

## STEP 5 — ATS KEYWORD ENGINEERING

Do NOT keyword stuff.

Expert ATS optimisation means:

* naturally integrating relevant terms,
* aligning phrasing with hiring language,
* improving contextual keyword relevance,
* and increasing recruiter interpretation clarity.

Prioritise keyword placement inside:

* professional headline,
* professional summary,
* first 3 bullets of each role,
* skills section,
* role titles where appropriate.

ATS systems evaluate:

* keyword frequency,
* placement,
* contextual relevance,
* semantic relationship,
* and role alignment.

Keywords must feel:

* natural,
* readable,
* and operationally believable.

Never produce robotic repetition.

---

## STEP 6 — HUMAN SCAN OPTIMISATION

After ATS screening, recruiters typically scan a CV in approximately 6–12 seconds.

The document must therefore communicate quickly:

* role fit,
* employability,
* operational relevance,
* workplace exposure,
* and professional clarity.

Optimise for:

* clean readability,
* low cognitive friction,
* visible relevance,
* concise structure,
* and scanning ease.

Ensure:

* first lines establish role alignment quickly,
* bullets begin with strong practical verbs,
* metrics remain visible when available,
* important information is not buried inside text walls.

The CV must satisfy TWO audiences:

1. ATS systems
2. Human recruiters

If conflict exists:
prioritise HUMAN RECRUITER READABILITY while preserving ATS compatibility naturally.

---

## STEP 7 — PROOF ARCHITECTURE

Premium-quality CV positioning requires proof logic.

Where evidence exists, strengthen credibility using:

* measurable outputs,
* operational scope,
* workflow responsibility,
* reporting ownership,
* process coordination,
* maintenance responsibilities,
* productivity contribution,
* customer interaction volume,
* support impact,
* continuity support,
* technical contribution.

Where exact metrics are unavailable:
use realistic operational framing without fabrication and ensure the output reflects real human writings.

Never invent:

* achievements,
* percentages,
* KPIs,
* revenue,
* team size,
* certifications,
* technical expertise,
* or unsupported business impact.

---

CORE WRITING STANDARD
- Write like a real human recruiter preparing a real CV for hiring
- Make the CV sound highly human, natural, grounded, believable and make complete practical sense
- Prioritise clarity over “impressive wording”
- Never sound robotic or templated, keep it 100% human. 
- Avoid exaggeration or inflated professionalism, mirror real human form of writing
- Keep tone consistent with selected template and ensure it reflects real human crafted work

HUMAN SOUNDING RULE (IMPORTANT)
The CV must NOT feel:
 - repetitive
 - overly structured like AI output
 - overly polished or artificial
 - filled with generic phrases
Instead:
 - vary sentence structure humanly and natural
 - avoid repeated openings in bullets
 - avoid filler adjectives
 - keep writing practical, real-world based and classic human believable writings

STRICT STYLE SEPARATION RULE
- NEVER mix entry-level tone with professional-level tone
- NEVER upgrade entry-level candidates into senior-sounding professionals
- NEVER downgrade professional candidates into overly simple student-like language
- Maintain correct tone consistency throughout the CV

TEMPLATE SELECTION IS MANDATORY.

The candidate_level supplied by the application is the final authority.

candidate_level values:

- entry
- professional
- executive

Selection rules:

- entry → ENTRY LEVEL TEMPLATE
- professional → PROFESSIONAL TEMPLATE
- executive → EXECUTIVE TEMPLATE

Do not override the supplied candidate_level based on your own interpretation.

Work experience should influence writing style, not template selection.

ANTI-GENERIC WRITING RULE:
Avoid vague corporate buzzwords, inflated claims, AI-style filler language, and empty professionalism unless clearly supported by evidence.

Avoid unsupported self-praise and generic corporate clichés.

Avoid repetitive bullet openings and weak generic phrasing.
Use varied and natural action verbs appropriate for the candidate's experience level.
For internship and student roles, supportive verbs such as:
- assisted
- supported
- participated
- coordinated
- prepared
- documented
are acceptable when used naturally in a way that reflects a real human writings with vivid picture of real work.

Do not use language that sounds copied from generic resume builders.

Avoid:
- vague self-praise
- inflated business language
- meaningless professional clichés. 
- exaggerated leadership wording for junior candidates

Prefer:
- practical workplace language
- operational detail. 
- workflow-specific wording
- realistic contribution
- environment-specific context
- task-based clarity
- believable professional phrasing
See yourself as a human who distaste robotic phrases and writing styles

SPECIFICITY RULE:
Experience bullets should be specific, realistic, relevant to the candidate’s actual level and makes complete sentence. Never cut sentence when the whole practical picture is not satisfactorily painted.

For experienced candidates:
- include operational detail and workplace contribution in a clear manner that any recruiter can easily envision the work in real time.

For students, interns, SIWES, and entry-level candidates:
- focus on practical exposure
- participation
- learning support
- administrative contribution
- laboratory/technical familiarity
- communication
- teamwork
- reliability
without forcing artificial complexity.
EDUCATION COMPETENCIES RULE:
- For entry-level, internship, SIWES, NYSC, and fresh graduate candidates with limited work experience, generate edu_competencies inside each education entry and ensure it is human and fit the course of study
- edu_competencies should contain practical academic strengths, technical exposure, laboratory familiarity, coursework relevance, research exposure, software familiarity, communication skills, or analytical abilities supported by the candidate’s field of study and ensure it relates to the purpose of the cv selected by the user
- Keep competencies realistic, highly human and well, grounded
- Do not invent advanced expertise
- Return 3 to 4 concise bullet-style competencies per education entry where appropriate
- If the candidate already has at least 1 work experience, edu_competencies should be an empty array

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


CANDIDATE-SPECIFICITY RULE

Experience bullets must prioritise information unique to the candidate.

When candidate-provided duties exist:
- rewrite and improve them
- reorganise them
- remove duplicate wording but preserve distinct evidence.

Do NOT replace them with generic role descriptions.

The model must prefer:
candidate evidence > inferred duties > generic industry duties.

If original responsibilities are available,
at least 70% of generated bullets should be directly derived from those responsibilities.

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
- Convert rough user input into structured professional phrasing
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
Bullets must explain:
- what was done
- what it supported
- what process it contributed to
- Avoid vague statements like: “worked in a fast-paced environment”
PROFESSIONAL SUMMARY RULE:
- Must match selected template and be highly human heavy writing style
- Must reflect actual level of experience that one can easily envision and relate with
- No self-praise language
- Must sound realistic, role-specific and strictly what an experienced humans can easily write
Example style: “Administrative professional with experience managing documentation, coordinating day-to-day office activities, maintaining employee and operational records, and serving as a point of contact for internal and external enquiries.”
Or:
“Administrative officer with experience handling document control, correspondence management, scheduling activities, and maintaining accurate organisational records within busy office environments.”
Or:
“Administrative support professional experienced in coordinating office documentation, maintaining records, preparing routine reports, and supporting daily operational activities.
”


Write practical, role-specific, evidence-based bullets that clearly explain:
- what was done
- what process was supported
- what operational purpose was served

PROFESSIONAL SUMMARY NATURAL LANGUAGE RULE

Professional summaries must read as a short professional introduction, not a list of competencies.

Avoid writing summaries that simply list:
- skills
- duties
- competencies
- responsibilities

Instead, summaries should naturally explain:

1. Who the candidate is.
2. What environment they have worked in.
3. What work they perform.
4. What value they bring.

A recruiter should be able to understand the candidate's professional identity and likely workplace contribution within the first two sentences.

- Keep the summary concise, recruiter-friendly, specific and role-targeted based on job type (admin, HR, data etc)
- Focus on:
  - type of experience
  - operational strengths
  - workplace support provided
  - environments worked in
  - practical contribution
Avoid exaggerated personality claims, but allow realistic professional qualities suitable for the candidate’s experience level.
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
- Keep skills relevant and realistic
- Avoid overloading with generic soft skills
- Only include skills supported by input
- Separate technical skills where needed

SKILLS ENHANCEMENT RULE

When generating the Core Competencies or Skills section:

- Do not simply copy generic user-provided skills if stronger evidence exists elsewhere in the profile.
- Review the candidate's education, experience, projects, certifications, research work, volunteer activities, achievements, and training records.
- Identify practical competencies that are clearly demonstrated by the information provided.
- Convert generic skill labels into recruiter-readable professional competencies where appropriate.
- consolidate only truly duplicate skills while preserving competency breadth.
- Prioritize role-relevant and industry-relevant competencies over generic soft skills.
- Ensure all skills remain fully supported by the candidate's information.
Examples:
-Instead of:
 - Communication
 - Leadership
 - Time Management
- Use:
  - Clinical Documentation
  - Team Coordination
  - Patient Communication
  - Scientific Reporting
  - Quality Control Awareness
  - Project Coordination
  - Administrative Support
  - Community Outreach
  where supported by the candidate's background.

TRUTH PRESERVATION FOR SKILLS
- Never invent skills.
- Never create competencies that are not supported by the supplied information.
- Every skill included must be traceable to education, experience, projects, certifications, achievements, training, or user-provided skills.
- If evidence is insufficient, retain the original skill rather than fabricating a stronger one.

SKILL PRIORITIZATION RULE
Order skills by relevance and professional value.
Priority:
- Technical / Occupational Skills
- Industry-Specific Competencies
- Functional Professional Skills
- Software / Tools
- Soft Skills
Avoid producing skill sections dominated by generic soft skills when stronger evidence exists elsewhere in the profile.

SKILL REFINEMENT EXAMPLES
Medical Laboratory Graduate
Weak:
 - Communication
 - Leadership
 - Excel
Better:
 - Specimen Collection & Handling
 - Laboratory Safety & Quality Control
 - Scientific Documentation
 - Microsoft Office (Word & Excel)
 - Data Recording & Reporting

Nurse / Public Health Professional

Weak:
 - Leadership
 - Communication
 - Teamwork
Better:
 - Clinical Team Leadership
 - Infection Prevention & Control
 - Patient Advocacy
 - Community Health Outreach
 - Monitoring & Reporting

Administrative Professional

Weak:
- Communication
- Time Management
- Leadership

Better:
- Record Management
- Administrative Coordination
- Office Documentation
- Customer Communication
- Microsoft Office Applications

Ensure to observe the same skill pattern above for users from other field of work and studies. 

ADDITIONAL INFORMATION RULE
- Additional information must be structured, short, and CV-ready. No paragraphs or storytelling.
- This section must follow fixed formats depending on content type:
- LANGUAGES FORMAT
 - Always use this exact style:
   - English (Fluent)
   - Hausa (Conversational)
- Rules:
- Each language on a new line
- Always include proficiency level in brackets
- No explanations, no extra words
- VOLUNTEER EXPERIENCE FORMAT
  - Always use action-based bullet style:
   - Assisted in community health outreach activities
   - Participated in student-led laboratory awareness campaigns
   - Supported administrative tasks during outreach programmes
- Rules:
- Each bullet must describe a real task/activity
- No storytelling or explanations
- No “during which”, “where”, or descriptive sentences
- MEMBERSHIP / AFFILIATION FORMAT
  - Use strict identity format only:
   - Member, Nigerian Institute of Management
   - Member, Student Research Association
-Rules:
- No descriptions
- No sentences
- Format must be: “Member, Organisation Name”

ADDITIONAL INFORMATION SECTION MUST FOLLOW STRICT OUTPUT SHAPE RULE:

LANGUAGES:
- One line per language
- Format: Language (Proficiency)
- No explanations
- No extra words

VOLUNTEER EXPERIENCE:
- Bullet only format
- Each line must begin with action verb
- No context phrases
- No storytelling
- Each bullet must be independent

MEMBERSHIP:
- One line per entry only
- Format: Member, Organisation Name
- No verbs
- No punctuation changes allowed

DO NOT MIX FORMATS BETWEEN CATEGORIES
GENERAL RULE
- Keep entries short and structured
- No paragraphs or long explanations
- No repetition of information already in other CV sections
- No narrative writing
- Output must look like a real CV, not a generated text summary

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

EVIDENCE PRESERVATION RULE

Before improving, optimising, aligning, condensing, or rewriting:

Preserve candidate evidence.

Candidate evidence includes:

- employment history
- achievements
- accomplishments
- certifications
- projects
- technical skills
- software skills
- professional strengths
- memberships
- leadership positions
- research
- awards
- publications
- target roles
- references

Optimisation may improve wording.

Optimisation must not remove evidence.

Responsibilities and achievements serve different purposes.

Responsibilities explain work performed.
Achievements explain outcomes achieved.

Both should be preserved whenever available.

EXPERIENCED CANDIDATE RULE

If candidate has 3+ years of professional experience:

- Education should remain concise.
- Do not generate education competency bullets.
- Do not generate coursework bullets.
- Do not generate laboratory exposure bullets.
- Do not generate academic project bullets unless explicitly supplied by the candidate.
- Prioritise professional experience, achievements, certifications, and technical skills.

CRITICAL CV PRESERVATION RULE:

Do not reduce, delete, or simplify candidate information when restructuring CV content.

1. TECHNICAL SKILLS PRESERVATION
All technical tools, software, and engineering systems (e.g. AutoCAD, SolidWorks, Excel, maintenance tools, ERP systems) MUST be preserved and included in a dedicated Technical Skills section OR Core Competencies if no technical section exists in the template.

Never remove technical tools unless they are clearly irrelevant to the candidate’s field.

2. ACHIEVEMENTS INTEGRATION RULE
Any content labeled or interpreted as:
- Selected Achievements
- Key Achievements
- Career Highlights
- Impact Statements

must NOT be converted into a separate standalone section unless the template explicitly requires it.

Instead:
- Integrate achievements into the relevant PROFESSIONAL EXPERIENCE role under the same employer
- Convert them into bullet points under that role
- Do NOT create a new heading such as “Selected Achievements”

3. SENIORITY PRESERVATION RULE
Maintain the candidate’s true professional seniority based on evidence:
- Job titles
- Years of experience
- Leadership responsibility
- Scope of work

Do not soften language unnecessarily using terms such as “assisted”, “supported”, or “helped” when the original CV indicates ownership, coordination, leadership, or execution responsibility.

Use strong but truthful verbs such as:
- Led
- Coordinated
- Managed
- Executed
- Conducted
- Oversaw
- Implemented
- Directed

Match verb strength to real responsibility level.

4. CONTENT DENSITY RULE
Preserve ALL meaningful sections from the original CV, including but not limited to:
- Technical Skills
- Professional Strengths
- Achievements / Highlights
- Certifications
- Memberships
- Tools & Software
- Awards
- Languages
- Training

If the template does not have a dedicated field, place them in extra sections WITHOUT merging into unrelated sections like EXPERIENCE or EDUCATION.

5. NO VALUE LOSS RULE
Under no condition should restructuring result in:
- removal of tools
- removal of achievements
- dilution of seniority
- flattening of responsibilities

The goal is restructuring, not simplification..


USER INPUT:
${JSON.stringify(rawInput, null, 2)}
`.trim();
}

function ensureTemplateExists(templatePath) {
  return fs.existsSync(templatePath);
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

              edu_competencies: {
              type: "array",
              items: { type: "string" },
},
          },
          required: [
  "degree",
  "school",
  "location",
  "start",
  "end",
  "edu_detail",
  "edu_competencies",
],
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

  items: {
    type: "array",
    items: { type: "string" },
  },

  section_content: { type: "string" },
},
required: ["section_title", "items", "section_content"],
        },
      },

      reference_choice: {
        type: "string",
        enum: ["none", "available", "included"],
      },

      reference_details: { type: "string" },

references_list: {
  type: "array",
  items: {
    type: "object",
    additionalProperties: false,
    properties: {
      name: { type: "string" },
      position: { type: "string" },
      organization: { type: "string" },
      location: { type: "string" },
      email: { type: "string" },
      phone: { type: "string" }
    },
    required: [
      "name",
      "position",
      "organization",
      "location",
      "email",
      "phone"
    ]
  }
},
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
  "references_list"
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
    template_exists: {
  entry: fs.existsSync(
    path.join(process.cwd(), "templates", "cv-template.docx")
  ),
  professional: fs.existsSync(
    path.join(process.cwd(), "templates", "template_professional.docx")
  ),
},
  });
});

app.get("/api/health", (req, res) => {
  return res.status(200).json({
    success: true,
    message: "Server is healthy",
    environment: NODE_ENV,
    template_exists: {
  entry: fs.existsSync(
    path.join(process.cwd(), "templates", "cv-template.docx")
  ),
  professional: fs.existsSync(
    path.join(process.cwd(), "templates", "template_professional.docx")
  ),
  executive: fs.existsSync(
    path.join(process.cwd(), "templates", "template_executive.docx")
  ),
},
  });
});

app.post("/generate-cv", async (req, res) => {
  try {

    let requestBody;
    let cvUrl = "";
    let extractedCvText = "";

    try {
      console.log("FULL BODY:");
      console.log(req.body);

      console.log("BODY TYPE:");
      console.log(typeof req.body);

      console.log("RAW SUBMISSION JSON:");
      console.log(req.body.raw_submission_json);

      console.log("RAW SUBMISSION JSON TYPE:");
      console.log(typeof req.body.raw_submission_json);

      requestBody = parseRequestBody(req.body);

      const uploadedCvUrl =
        requestBody.uploaded_cv_url;

      cvUrl = uploadedCvUrl;

      if (
        cvUrl &&
        cvUrl.startsWith("//")
      ) {
        cvUrl = "https:" + cvUrl;
      }

    } catch (parseError) {
      return res.status(parseError.statusCode || 400).json({
        success: false,
        error: parseError.message || "Invalid request body",
        details: parseError.details || "",
      });
    }

    if (cvUrl) {

      const response = await axios.get(
        cvUrl,
        {
          responseType: "arraybuffer"
        }
      );

      const fileBuffer = Buffer.from(
        response.data
      );

      if (
        cvUrl.toLowerCase().includes(".pdf")
      ) {

        const pdfData =
          await pdfParse(fileBuffer);

        extractedCvText =
          pdfData.text;

      } else if (
        cvUrl.toLowerCase().includes(".docx")
      ) {

        const result =
          await mammoth.extractRawText({
            buffer: fileBuffer
          });

        extractedCvText =
          result.value;
      }
    }

    console.log(
      "EXTRACTED CV TEXT:"
    );

    console.log(
      extractedCvText.substring(0, 1000)
    );
console.log(
  "CV TEXT LENGTH:",
  extractedCvText.length
);

console.log(
  "LAST 2000 CHARS:"
);

console.log(
  extractedCvText.slice(-2000)
);

console.log("STEP 1");

    requestBody.uploaded_cv_text =
  extractedCvText;

  console.log("STEP 2");

    const incomingError = validateIncomingBody(requestBody);

    console.log("STEP 3");

    if (incomingError) {
      return res.status(400).json({
        success: false,
        error: incomingError,
      });
    }

    
let rawInput;

try {
  console.log("ABOUT TO NORMALIZE");

  rawInput = normalizeIncomingPayload(requestBody);

  console.log("NORMALIZATION SUCCESS");
} catch (err) {
  console.error("NORMALIZATION ERROR:");
  console.error(err);
  console.error(err.stack);

  return res.status(500).json({
    success: false,
    error: err.message,
  });
}

console.log("STEP 4");

console.dir(
  rawInput.reference_entries,
  { depth: null }
);

const templatePath = getTemplatePath(
  requestBody?.candidate_level
);

console.log("CANDIDATE LEVEL:", requestBody?.candidate_level);
console.log("EXPERIENCE COUNT:", rawInput.experience.length);
console.log("TEMPLATE PATH:", templatePath);
  
if (!ensureTemplateExists(templatePath)) {
  return res.status(500).json({
    success: false,
    error: `Template file not found: ${templatePath}`,
  });
}

if (
  rawInput.reference_entries.length === 0 &&
  extractedCvText
) {
  rawInput.reference_entries =
    extractReferencesFromCvText(
      extractedCvText
    );

  rawInput.reference_details =
    buildReferenceDetailsFromEntries(
      rawInput.reference_entries
    );

  if (
    rawInput.reference_entries.length > 0
  ) {
    rawInput.reference_choice =
      "included";
  }
}

  const prompt = buildPrompt(rawInput) + `

UPLOADED CV CONTENT:

${extractedCvText}

CRITICAL INSTRUCTIONS:

The uploaded CV is the PRIMARY source of truth.

If the uploaded CV contains work experience, education, certifications, trainings, memberships, leadership positions, projects, achievements, research, publications, awards, references, languages, competencies or any other career information, extract and use them.

Do NOT ignore information simply because it is not present in the form submission.

Do NOT reduce the candidate's experience level if the uploaded CV shows substantial professional experience.

Preserve as many valid positions, certifications, trainings, projects and education records as are present in the uploaded CV.

If the uploaded CV contains references, extract them and populate reference_details.

If the uploaded CV contains leadership positions, memberships, research, languages, workshops, trainings or additional professional information, include them in extra_sections.

If the uploaded CV contains research projects, dissertations, publications, thesis topics or academic research work etc, preserve them in extra_sections under the title "Research".

Do not discard research topics even if they are not professional projects.

Preserve software skills, computer skills and technical tools.
Examples include Microsoft Word, Excel, PowerPoint, AutoCAD, SAP, MATLAB, SolidWorks (2D & 3D Design), Python and similar tools.
Store them under either Skills or an extra section titled "Software Proficiency".
For education, preserve the full degree title, field of study, institution name, country and graduation year whenever available.
Do not shorten degree names if the original CV provides more detail.

Preserve all leadership positions, professional memberships, awards, honours and affiliations found in the uploaded CV.
Store them as separate extra_sections rather than discarding them.

Only use form data to supplement or fill gaps where the uploaded CV does not provide information.

The uploaded CV should override conflicting information from the form whenever professional history is more complete in the uploaded CV.

EXTRACTION RULES:

- Preserve every employment record found in the uploaded CV.
- Preserve every certification found in the uploaded CV.
- Preserve all research topics.
- Preserve all professional memberships.
- Preserve all leadership positions.
- Preserve all software and computer skills.
- Preserve all references if present.
- Do not merge jobs together.
- Do not invent dates.
- If a date is missing, leave it blank.
- If information exists in the uploaded CV but not in the form, use the uploaded CV.
SECTION PRESERVATION RULE

When revamping an uploaded CV:

Do not remove existing CV sections simply because they are not required by a standard resume format.

If the uploaded CV contains:
- Professional Strengths
- Key Achievements
- Selected Achievements
- Career Highlights
- Professional Memberships
- Technical Skills
- Software Proficiency
- Target Roles
- Awards
- Languages
- Leadership Experience
- Volunteer Experience
- Research
- Publications
- Training

preserve them by placing them into extra_sections when no dedicated schema field exists.

ACHIEVEMENT PRESERVATION RULE

If achievements appear under a specific employment record, preserve them within that employment record.

Do not move role-specific achievements into extra_sections.

Role achievements should remain attached to the relevant position and should be incorporated into the experience content where appropriate.

Only place achievements in extra_sections when they are presented as a standalone section in the uploaded CV and are not tied to a specific role.

The objective is professional improvement, not content reduction.

Prefer preserving and restructuring information rather than discarding it.

DATE PRESERVATION RULES:

When preserving employment history from the uploaded CV:

- Never invent employment dates.
- Never assume a position is current unless the uploaded CV explicitly states "Present", "Current", or an ongoing date range.
- If only an end year is available, preserve only the end year.
- If both start and end dates are missing, leave both blank.
- Do not output "Present" when no start date exists.
- Do not create date ranges using guessed dates.
EMPLOYMENT QUALITY RULES:

Before returning employment records:

- Verify that each position contains a valid title and employer.
- If dates are unclear or missing, preserve the position but leave the date fields blank.
- Never fabricate employment periods.

REFERENCES EXTRACTION RULES

When a References or Referees section exists, create one references_list object per referee.

For each referee extract:

- name
- position
- organization
- location
- email
- phone

Do not combine fields.

Do not place phone numbers inside organization.

Do not place the word "Phone" inside organization.

Do not place name and position in the same field.

If a referee appears as:

Engr. Morrison Chukuka
Regional Manager, Montego Nigeria Limited
Port Harcourt
Phone: +2347034034456

Return:

{
  "name": "Engr. Morrison Chukuka",
  "position": "Regional Manager",
  "organization": "Montego Nigeria Limited",
  "location": "Port Harcourt",
  "email": "",
  "phone": "+2347034034456"
}

If location cannot be determined, leave it blank.

Never return an entire referee as a single string.
`;

console.log(
  "PROMPT LENGTH:",
  prompt.length
);

console.log(
  "REFERENCE ENTRIES:",
  rawInput.reference_entries
);

console.log(
  "REFERENCE CHOICE:",
  rawInput.reference_choice
);

console.log(
  "REFERENCE DETAILS:",
  rawInput.reference_details
);

console.log(
  "REFERENCES EXTRACTED FROM CV:"
);

console.log(
  rawInput.reference_entries
);
    let completion;
    try {
      completion = await openai.responses.create({
  model: "gpt-4.1-mini",
  temperature: 0.2,
  text: {
    format: {
      type: "json_schema",
      name: CV_JSON_SCHEMA.name,
      strict: true,
      schema: CV_JSON_SCHEMA.schema,
    },
  },
  input: [
    {
      role: "developer",
      content: [
        {
          type: "input_text",
          text: "Return only valid JSON matching the provided schema. No markdown. No commentary.",
        },
      ],
    },
    {
      role: "user",
      content: [
        {
          type: "input_text",
          text: prompt,
        },
      ],
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

    const content = completion.output_text;

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

console.log("REFERENCE TEXT:");
console.log(referenceText);

    const renderData = {
      FULL_NAME: data.full_name || "",
      CONTACT_LINE: buildContactLine(data) || "",
      PROFESSIONAL_SUMMARY: data.professional_summary || "",
      
      HAS_SKILLS: data.skills.length > 0,
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
  const candidateLevel =
  (requestBody?.candidate_level || "")
    .trim()
    .toLowerCase();

const templatePath = getTemplatePath(
  candidateLevel,
  rawInput.experience.length
);

const binaryTemplate = fs.readFileSync(templatePath, "binary");
      const zip = new PizZip(binaryTemplate);

      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: false,
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

const bucket = process.env.AWS_BUCKET_NAME;

console.log("BUCKET VALUE:", bucket);

const fileName = generateUniqueFileName(data.full_name);
const s3Key = `generated-cv/${fileName}`;

try {
  await s3.send(
    new PutObjectCommand({
      Bucket: bucket,
      Key: s3Key,
      Body: buffer,
      ContentType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    })
  );
} catch (uploadError) {
  console.error("S3 upload failed:", uploadError);

  return res.status(500).json({
    success: false,
    error: "Failed to upload CV",
  });
}

const command = new GetObjectCommand({
  Bucket: bucket,
  Key: s3Key,
});

const downloadUrl = await getSignedUrl(s3, command, {
  expiresIn: 600,
});

return res.status(200).json({
  success: true,
  message: "CV generated successfully",
  file_name: fileName,
  download_url: downloadUrl,
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

app.listen(PORT, () => {
  console.log(`CV API running on port ${PORT}`);
});

module.exports = app;