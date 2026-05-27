require("dotenv").config();

const express = require("express");
const cors = require("cors");
const fs = require("fs");
const path = require("path");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const { OpenAI } = require("openai");

const { S3Client, PutObjectCommand, GetObjectCommand } = require("@aws-sdk/client-s3");
const { getSignedUrl } = require("@aws-sdk/s3-request-presigner");
const { ListObjectsV2Command } = require("@aws-sdk/client-s3");

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

async function generateFileName(s3, bucket, fullName) {
  const cleanName = cleanDisplayName(fullName);
  const baseKeyPrefix = `generated-cv/${cleanName} CV`;

  const list = await s3.send(
    new ListObjectsV2Command({
      Bucket: bucket,
      Prefix: baseKeyPrefix,
    })
  );

  const existing = (list.Contents || [])
    .map((obj) => obj.Key)
    .filter(Boolean);

  if (existing.length === 0) {
    return `${cleanName} CV.docx`;
  }

  let maxIndex = 0;

  for (const key of existing) {
    const match = key.match(/\((\d+)\)\.docx$/);
    if (match) {
      maxIndex = Math.max(maxIndex, parseInt(match[1], 10));
    }
  }

  return `${cleanName} CV (${maxIndex + 1}).docx`;
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
You are a world-class ATS CV writer, recruiter, HR reviewer, and professional CV structuring engine.

Your task is to transform raw user input into a polished, ATS-compatible, recruiter-readable CV suitable for real job applications.

IMPORTANT CONTEXT:
- Users may submit incomplete, repetitive, poorly written, fragmented, informal, badly capitalised, misspelled, or inconsistent information
- Your responsibility is to clean, structure, and professionalise the content without changing factual meaning
- If a job description, internship description, or academic opportunity description is provided, align the CV naturally toward the target opportunity without copying wording directly
- Extract relevant role keywords naturally without copying the job description directly
- Focus on recruiter readability, clarity, credibility, realism, and professional presentation
- The final CV must sound professionally written, human, realistic, recruiter-readable, and appropriate for the candidate’s actual career stage

ENTRY LEVEL TEMPLATE
Use this when:
- student, intern, SIWES, NYSC, fresh graduate
- limited or no formal work experience
- experience is mostly academic, volunteer, training, or project-based

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
- Used for students, interns, SIWES, fresh graduates, or limited experience candidates
- Writing must be SIMPLE, CLEAR, and LEARNING-ORIENTED
- Do NOT exaggerate responsibility or seniority
- Focus on:
  - exposure to tasks
  - participation in activities
  - administrative or academic support
  - teamwork and learning environments
- Experience bullets must sound like TRAINING OR SUPPORT ROLES, not independent authority roles
- Language must remain grounded and non-impressive
- Avoid over-structured corporate phrasing

ENTRY LEVEL TEMPLATE STYLE RULES
- Keep language simple, clear, and learning-focused
- Emphasise exposure, participation, and support work
- Avoid strong authority tone
- Experience must reflect:
  - assistance
  - learning
  - observation
  - supervision-based tasks
  - academic or internship involvement
- No inflated responsibility claims

PROFESSIONAL TEMPLATE
- Used for candidates with real job experience
- Writing must reflect INDEPENDENT WORK, RESPONSIBILITY, AND WORKFLOW OWNERSHIP
- Focus on:
  - coordination of tasks
  - management of processes
  - execution of responsibilities
  - workplace contribution
- Experience bullets should show STRUCTURE, DECISION SUPPORT, AND OPERATIONAL INVOLVEMENT
- Language can be slightly more advanced but must remain realistic and non-inflated

PROFESSIONAL TEMPLATE STYLE RULES
- Reflect independent work and responsibility
- Show workflow ownership and coordination
- Emphasise execution of tasks within real systems
- Experience must reflect:
  - responsibility
  - coordination
  - reporting
  - operational contribution
- Language can be structured but must remain realistic

CORE WRITING STANDARD
- Write like a recruiter preparing a real CV for hiring
- Make the CV sound natural, grounded, and believable
- Prioritise clarity over “impressive wording”
- Never sound robotic or templated
- Avoid exaggeration or inflated professionalism
- Keep tone consistent with selected template

HUMAN SOUNDING RULE (IMPORTANT)
The CV must NOT feel:
 - repetitive
 - overly structured like AI output
 - overly polished or artificial
 - filled with generic phrases
Instead:
 - vary sentence structure naturally
 - avoid repeated openings in bullets
 - avoid filler adjectives
 - keep writing practical and real-world based

STRICT STYLE SEPARATION RULE
- NEVER mix entry-level tone with professional-level tone
- NEVER upgrade entry-level candidates into senior-sounding professionals
- NEVER downgrade professional candidates into overly simple student-like language
- Maintain correct tone consistency throughout the CV

TEMPLATE SELECTION IS MANDATORY AND MUST BE BASED ONLY ON EXPLICIT USER INPUT.
- If user input does not clearly show employment history:
→ Default to ENTRY LEVEL TEMPLATE.

- If user input contains 2 or more structured job roles with responsibilities:
→ PROFESSIONAL TEMPLATE ONLY.

- If unclear:
→ Always choose ENTRY LEVEL TEMPLATE.

CORE WRITING STANDARD:
- Write like an experienced recruiter preparing a candidate for real hiring review
- Make the candidate sound employable, credible, grounded, and professionally clear
- Prioritise specificity over generic professionalism
- Use direct, realistic, human-sounding language
- Preserve realistic seniority based on the candidate’s actual experience level
- Improve weak or awkward wording while maintaining truth and realism
- Avoid robotic phrasing, exaggerated confidence, and empty corporate language
- Every section should feel believable and operationally realistic
- Tone must always match the selected template level and never drift into generic corporate phrasing.

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
Avoid repetitive bullet openings and weak generic phrasing.
Use varied and natural action verbs appropriate for the candidate's experience level.
For internship and student roles, supportive verbs such as:
- assisted
- supported
- participated
- coordinated
- prepared
- documented
are acceptable when used naturally.

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
Experience bullets should be specific, realistic, and relevant to the candidate’s actual level.

For experienced candidates:
- include operational detail and workplace contribution.

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

HUMAN SOUNDING CONTROL:
Every CV must pass this internal check:
- If removed formatting, it should still read like a human wrote it in real workplace context
- No sentence should feel like a template
- No repeated sentence starters in same section
- No over-consistent grammar patterns across bullets
- Avoid symmetry in writing style across bullets

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
Bullets must explain:
- what was done
- what it supported
- what process it contributed to
- Avoid vague statements like: “worked in a fast-paced environment”

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
- Never invent:
 - achievements, KPIs, tools, employers, dates, qualifications
 - Do not upgrade responsibilities beyond input
 - Keep all content strictly based on user data
 - Improve clarity only, not facts

PROFESSIONAL SUMMARY RULE:
- Must match selected template
- Must reflect actual level of experience
- No self-praise language
- Must sound realistic and role-specific
Example style: “Administrative professional with experience supporting office coordination, record management, and customer communication in structured work environments”
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
- Each line must start with a strong action verb:
- Assisted, Supported, Participated, Coordinated, Organized, Documented
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

FINAL OUTPUT STABILITY RULE:
- Do not change section order once CV structure is created
- Do not merge sections
- Do not skip sections even if empty
- If missing data, return empty string or empty array
- Never generate extra sections not requested

DUAL LAYER OUTPUT MODE:

Every CV must satisfy two simultaneous layers:

1. ATS LAYER:
- keyword clarity
- structured formatting
- role-relevant terms
- clean action verbs

2. HUMAN RECRUITER LAYER:
- natural sentence variation
- realistic tone
- believable workplace phrasing
- no robotic structure

If both layers conflict:
→ prioritize HUMAN RECRUITER LAYER while preserving ATS keywords naturally.

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

function determineTemplateType(rawInput) {
  const experience = safeArray(rawInput?.experience);
  const hasExperience = experience.length > 0;

  const purpose = safeString(rawInput?.document_purpose).toLowerCase();

  // Strong signals for entry level
  const entryLevelKeywords = [
    "internship",
    "siwes",
    "student",
    "graduate",
    "fresh graduate",
    "entry",
    "no experience"
  ];

  const isEntryByPurpose = entryLevelKeywords.some((k) =>
    purpose.includes(k)
  );

  // Rule logic (simple but effective)
  if (!hasExperience || isEntryByPurpose) {
    return "entry";
  }

  return "professional";
}

app.get("/", (req, res) => {
  return res.status(200).json({
    success: true,
    message: "CV API is running",
    environment: NODE_ENV,
    template_exists: {
  entry: fs.existsSync(
    path.join(process.cwd(), "templates", "template_entry_level.docx")
  ),
  professional: fs.existsSync(
    path.join(process.cwd(), "templates", "template_professional.docx")
  ),
}
  });
});

app.get("/api/health", (req, res) => {
  return res.status(200).json({
    success: true,
    message: "Server is healthy",
    environment: NODE_ENV,
    template_exists: {
  entry: fs.existsSync(
    path.join(process.cwd(), "templates", "template_entry_level.docx")
  ),
  professional: fs.existsSync(
    path.join(process.cwd(), "templates", "template_professional.docx")
  ),
}
  });
});

app.post("/generate-cv", async (req, res) => {
  try {

    console.log("STEP 1: Route hit");
    console.log("RAW BODY:", JSON.stringify(req.body, null, 2));

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

    console.log("STEP 2: Before validation");

const incomingError = validateIncomingBody(requestBody);

console.log("STEP 3: After validation");

if (incomingError) {
  return res.status(400).json({
    success: false,
    error: incomingError,
  });
}

console.log("STEP 4: Before normalization");

const rawInput = normalizeIncomingPayload(requestBody);

console.log("STEP 5: After normalization");

const rawInput = normalizeIncomingPayload(requestBody);

console.log("STEP 2: Payload normalized");

const templateType = determineTemplateType(rawInput);

const TEMPLATE_PATH =
  templateType === "entry"
    ? path.join(process.cwd(), "templates", "template_entry_level.docx")
    : path.join(process.cwd(), "templates", "template_professional.docx");

    if (!fs.existsSync(TEMPLATE_PATH)) {
      return res.status(500).json({
        success: false,
        error: "Template file not found: " + TEMPLATE_PATH,
      });
    }

    const prompt = buildPrompt(rawInput);
    
    console.log("STEP 3: Prompt built");

    console.log("STEP 4: Sending OpenAI request");

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

  console.log("STEP 5: OpenAI response received");

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

    const content = completion?.output_text || completion?.output?.[0]?.content?.[0]?.text;

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

    console.log("STEP 6: JSON parsed");

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

    if (NODE_ENV !== "development") {
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

      console.log("STEP 7: Rendering DOCX");

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

console.log("STEP 8: DOCX rendered");

    console.log("STEP 9: Uploading to S3");

const bucket = process.env.AWS_BUCKET_NAME;

console.log("BUCKET VALUE:", bucket);

const fileName = await generateFileName(s3, bucket, data.full_name);
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

  console.log("STEP 10: S3 upload successful");
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