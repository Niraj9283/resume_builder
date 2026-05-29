import cors from "cors";
import dotenv from "dotenv";
import express from "express";
import bcrypt from "bcryptjs";
import jwt from "jsonwebtoken";
import multer from "multer";
import ImageKit from "imagekit";
import pg from "pg";
import { GoogleGenerativeAI } from "@google/generative-ai";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRoot = path.resolve(__dirname, "../..");
dotenv.config({ path: path.join(projectRoot, ".env") });
dotenv.config({ path: path.join(projectRoot, "server", ".env"), override: true });

const app = express();
const PORT = process.env.PORT || 5000;
const CLIENT_URL = process.env.CLIENT_URL || process.env.RENDER_EXTERNAL_URL || "http://localhost:5173";
const SERVER_URL = process.env.SERVER_URL || process.env.RENDER_EXTERNAL_URL || `http://localhost:${PORT}`;
const JWT_SECRET = process.env.JWT_SECRET || "development_secret_change_me";

const uploadDir = path.join(projectRoot, "uploads");
const dataDir = path.join(projectRoot, "server", "data");
const localDbPath = path.join(dataDir, "local-db.json");
const clientDistDir = path.join(projectRoot, "client", "dist");
const { Pool } = pg;

const corsOrigins = [
  ...CLIENT_URL.split(",").map((item) => item.trim()),
  "http://localhost:5173",
  "http://127.0.0.1:5173"
].filter(Boolean);

app.use(
  cors({
    origin(origin, callback) {
      if (!origin || corsOrigins.includes(origin)) return callback(null, true);
      callback(new Error(`CORS blocked origin: ${origin}`));
    },
    credentials: true
  })
);
app.use(express.json({ limit: "4mb" }));
app.use("/uploads", express.static(uploadDir));

await fs.mkdir(uploadDir, { recursive: true });
await fs.mkdir(dataDir, { recursive: true });

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 5 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    if (!file.mimetype.startsWith("image/")) {
      cb(new Error("Only image uploads are supported"));
      return;
    }
    cb(null, true);
  }
});

const usePostgres = Boolean(
  process.env.DATABASE_URL ||
    process.env.PGHOST ||
    process.env.PGUSER ||
    process.env.PGDATABASE
);
let pool;
let localDb = { users: [], resumes: [], jobs: [] };
const DEFAULT_SECTION_ORDER = ["summary", "experience", "projects", "education", "skills"];

if (usePostgres) {
  pool = new Pool({
    connectionString: process.env.DATABASE_URL || undefined,
    ssl: process.env.PGSSL === "true" ? { rejectUnauthorized: false } : undefined
  });
  await initPostgres();
} else {
  try {
    localDb = { ...localDb, ...(JSON.parse(await fs.readFile(localDbPath, "utf8")) || {}) };
    localDb.users ||= [];
    localDb.resumes ||= [];
    localDb.jobs ||= [];
  } catch {
    await saveLocalDb();
  }
}

const imagekit =
  process.env.IMAGEKIT_PUBLIC_KEY &&
  process.env.IMAGEKIT_PRIVATE_KEY &&
  process.env.IMAGEKIT_URL_ENDPOINT
    ? new ImageKit({
        publicKey: process.env.IMAGEKIT_PUBLIC_KEY,
        privateKey: process.env.IMAGEKIT_PRIVATE_KEY,
        urlEndpoint: process.env.IMAGEKIT_URL_ENDPOINT
      })
    : null;

const genAI = process.env.GEMINI_API_KEY
  ? new GoogleGenerativeAI(process.env.GEMINI_API_KEY)
  : null;

function id() {
  return crypto.randomBytes(12).toString("hex");
}

async function saveLocalDb() {
  await fs.mkdir(dataDir, { recursive: true });
  await fs.writeFile(localDbPath, JSON.stringify(localDb, null, 2));
}

async function initPostgres() {
  await pool.query(`
    CREATE TABLE IF NOT EXISTS users (
      id TEXT PRIMARY KEY,
      name TEXT NOT NULL,
      email TEXT NOT NULL UNIQUE,
      password TEXT NOT NULL,
      email_verified BOOLEAN NOT NULL DEFAULT false,
      verification_token TEXT,
      reset_token TEXT,
      reset_expires TIMESTAMPTZ,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );

    CREATE TABLE IF NOT EXISTS resumes (
      id TEXT PRIMARY KEY,
      user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
      title TEXT NOT NULL,
      personal_info JSONB NOT NULL DEFAULT '{}'::jsonb,
      professional_summary TEXT NOT NULL DEFAULT '',
      experience JSONB NOT NULL DEFAULT '[]'::jsonb,
      education JSONB NOT NULL DEFAULT '[]'::jsonb,
      project JSONB NOT NULL DEFAULT '[]'::jsonb,
      skills JSONB NOT NULL DEFAULT '[]'::jsonb,
      section_order JSONB NOT NULL DEFAULT '["summary","experience","projects","education","skills"]'::jsonb,
      font_family TEXT NOT NULL DEFAULT 'Inter',
      density TEXT NOT NULL DEFAULT 'comfortable',
      template TEXT NOT NULL DEFAULT 'minimal-image',
      accent_color TEXT NOT NULL DEFAULT '#14B8A6',
      public BOOLEAN NOT NULL DEFAULT false,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );

    CREATE INDEX IF NOT EXISTS idx_resumes_user_updated ON resumes(user_id, updated_at DESC);
    CREATE INDEX IF NOT EXISTS idx_resumes_public ON resumes(id) WHERE public = true;

    CREATE TABLE IF NOT EXISTS jobs (
      id TEXT PRIMARY KEY,
      user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
      resume_id TEXT REFERENCES resumes(id) ON DELETE SET NULL,
      title TEXT NOT NULL,
      company TEXT NOT NULL DEFAULT '',
      location TEXT NOT NULL DEFAULT '',
      status TEXT NOT NULL DEFAULT 'Saved',
      description TEXT NOT NULL DEFAULT '',
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );

    CREATE INDEX IF NOT EXISTS idx_jobs_user_updated ON jobs(user_id, updated_at DESC);
  `);

  await pool.query(`
    ALTER TABLE users ADD COLUMN IF NOT EXISTS email_verified BOOLEAN NOT NULL DEFAULT false;
    ALTER TABLE users ADD COLUMN IF NOT EXISTS verification_token TEXT;
    ALTER TABLE users ADD COLUMN IF NOT EXISTS reset_token TEXT;
    ALTER TABLE users ADD COLUMN IF NOT EXISTS reset_expires TIMESTAMPTZ;
    ALTER TABLE resumes ADD COLUMN IF NOT EXISTS section_order JSONB NOT NULL DEFAULT '["summary","experience","projects","education","skills"]'::jsonb;
    ALTER TABLE resumes ADD COLUMN IF NOT EXISTS font_family TEXT NOT NULL DEFAULT 'Inter';
    ALTER TABLE resumes ADD COLUMN IF NOT EXISTS density TEXT NOT NULL DEFAULT 'comfortable';
  `);
}

function now() {
  return new Date().toISOString();
}

function normalizeDoc(doc) {
  if (!doc) return null;
  const plain = typeof doc.toObject === "function" ? doc.toObject() : { ...doc };
  if (plain.id && !plain._id) plain._id = plain.id;
  if (plain.user_id && !plain.userId) plain.userId = plain.user_id;
  if (plain.created_at && !plain.createdAt) plain.createdAt = plain.created_at;
  if (plain.updated_at && !plain.updatedAt) plain.updatedAt = plain.updated_at;
  plain._id = String(plain._id);
  if (plain.userId) plain.userId = String(plain.userId);
  return plain;
}

function toPgResume(row) {
  if (!row) return null;
  return normalizeDoc({
    _id: row.id,
    userId: row.user_id,
    title: row.title,
    personal_info: row.personal_info || {},
    professional_summary: row.professional_summary || "",
    experience: row.experience || [],
    education: row.education || [],
    project: row.project || [],
    skills: row.skills || [],
    section_order: row.section_order || DEFAULT_SECTION_ORDER,
    font_family: row.font_family || "Inter",
    density: row.density || "comfortable",
    template: row.template || "minimal-image",
    accent_color: row.accent_color || "#14B8A6",
    public: Boolean(row.public),
    createdAt: row.created_at,
    updatedAt: row.updated_at
  });
}

function publicUser(user) {
  const plain = normalizeDoc(user);
  return {
    _id: plain._id,
    name: plain.name,
    email: plain.email,
    emailVerified: Boolean(plain.emailVerified ?? plain.email_verified)
  };
}

function defaultResume(userId, title = "Untitled Resume") {
  return {
    userId,
    title,
    public: false,
    template: "minimal-image",
    accent_color: "#14B8A6",
    personal_info: {
      full_name: "",
      email: "",
      phone: "",
      location: "",
      linkedin: "",
      website: "",
      profession: "",
      image: ""
    },
    professional_summary: "",
    experience: [],
    education: [],
    project: [],
    skills: [],
    section_order: DEFAULT_SECTION_ORDER,
    font_family: "Inter",
    density: "comfortable"
  };
}

async function findUserByEmail(email) {
  const normalized = email.toLowerCase().trim();
  if (usePostgres) {
    const { rows } = await pool.query("SELECT * FROM users WHERE email = $1 LIMIT 1", [normalized]);
    return normalizeDoc(rows[0]);
  }
  return localDb.users.find((user) => user.email === normalized) || null;
}

async function findUserById(userId) {
  if (usePostgres) {
    const { rows } = await pool.query("SELECT * FROM users WHERE id = $1 LIMIT 1", [userId]);
    return normalizeDoc(rows[0]);
  }
  return localDb.users.find((user) => user._id === userId) || null;
}

async function createUser(data) {
  const payload = { ...data, email: data.email.toLowerCase().trim() };
  if (usePostgres) {
    const userId = id();
    const { rows } = await pool.query(
      `INSERT INTO users (id, name, email, password, email_verified, verification_token)
       VALUES ($1, $2, $3, $4, $5, $6)
       RETURNING *`,
      [userId, payload.name, payload.email, payload.password, Boolean(payload.emailVerified), payload.verificationToken || null]
    );
    return normalizeDoc(rows[0]);
  }
  const user = { _id: id(), ...payload, createdAt: now(), updatedAt: now() };
  localDb.users.push(user);
  await saveLocalDb();
  return user;
}

async function findUserByVerificationToken(token) {
  if (!token) return null;
  if (usePostgres) {
    const { rows } = await pool.query("SELECT * FROM users WHERE verification_token = $1 LIMIT 1", [token]);
    return normalizeDoc(rows[0]);
  }
  return localDb.users.find((user) => user.verificationToken === token) || null;
}

async function findUserByResetToken(token) {
  if (!token) return null;
  if (usePostgres) {
    const { rows } = await pool.query(
      "SELECT * FROM users WHERE reset_token = $1 AND reset_expires > NOW() LIMIT 1",
      [token]
    );
    return normalizeDoc(rows[0]);
  }
  const user = localDb.users.find((item) => item.resetToken === token);
  if (!user || !user.resetExpires || new Date(user.resetExpires).getTime() < Date.now()) return null;
  return user;
}

async function verifyUserEmail(token) {
  const user = await findUserByVerificationToken(token);
  if (!user) return null;
  const userId = normalizeDoc(user)._id;

  if (usePostgres) {
    const { rows } = await pool.query(
      `UPDATE users
       SET email_verified = true,
           verification_token = NULL,
           updated_at = NOW()
       WHERE id = $1
       RETURNING *`,
      [userId]
    );
    return normalizeDoc(rows[0]);
  }

  const index = localDb.users.findIndex((item) => item._id === userId);
  localDb.users[index] = { ...localDb.users[index], emailVerified: true, verificationToken: "", updatedAt: now() };
  await saveLocalDb();
  return localDb.users[index];
}

async function createPasswordReset(email) {
  const user = await findUserByEmail(email);
  if (!user) return null;

  const resetToken = id();
  const resetExpires = new Date(Date.now() + 1000 * 60 * 30).toISOString();
  const userId = normalizeDoc(user)._id;

  if (usePostgres) {
    await pool.query(
      `UPDATE users
       SET reset_token = $2,
           reset_expires = $3,
           updated_at = NOW()
       WHERE id = $1`,
      [userId, resetToken, resetExpires]
    );
    return { resetToken, resetExpires };
  }

  const index = localDb.users.findIndex((item) => item._id === userId);
  localDb.users[index] = { ...localDb.users[index], resetToken, resetExpires, updatedAt: now() };
  await saveLocalDb();
  return { resetToken, resetExpires };
}

async function resetPassword(token, password) {
  const user = await findUserByResetToken(token);
  if (!user) return null;
  const hashed = await bcrypt.hash(password, 10);
  const userId = normalizeDoc(user)._id;

  if (usePostgres) {
    const { rows } = await pool.query(
      `UPDATE users
       SET password = $2,
           reset_token = NULL,
           reset_expires = NULL,
           updated_at = NOW()
       WHERE id = $1
       RETURNING *`,
      [userId, hashed]
    );
    return normalizeDoc(rows[0]);
  }

  const index = localDb.users.findIndex((item) => item._id === userId);
  localDb.users[index] = { ...localDb.users[index], password: hashed, resetToken: "", resetExpires: "", updatedAt: now() };
  await saveLocalDb();
  return localDb.users[index];
}

async function createResume(data) {
  if (usePostgres) {
    const resumeId = id();
    const { rows } = await pool.query(
      `INSERT INTO resumes (
        id, user_id, title, personal_info, professional_summary, experience,
        education, project, skills, section_order, font_family, density, template, accent_color, public
      )
      VALUES ($1, $2, $3, $4::jsonb, $5, $6::jsonb, $7::jsonb, $8::jsonb, $9::jsonb, $10::jsonb, $11, $12, $13, $14, $15)
      RETURNING *`,
      [
        resumeId,
        data.userId,
        data.title,
        JSON.stringify(data.personal_info || {}),
        data.professional_summary || "",
        JSON.stringify(data.experience || []),
        JSON.stringify(data.education || []),
        JSON.stringify(data.project || []),
        JSON.stringify(data.skills || []),
        JSON.stringify(data.section_order || DEFAULT_SECTION_ORDER),
        data.font_family || "Inter",
        data.density || "comfortable",
        data.template || "minimal-image",
        data.accent_color || "#14B8A6",
        Boolean(data.public)
      ]
    );
    return toPgResume(rows[0]);
  }
  const resume = { _id: id(), ...data, createdAt: now(), updatedAt: now() };
  localDb.resumes.push(resume);
  await saveLocalDb();
  return resume;
}

async function listUserResumes(userId) {
  if (usePostgres) {
    const { rows } = await pool.query(
      "SELECT * FROM resumes WHERE user_id = $1 ORDER BY updated_at DESC",
      [userId]
    );
    return rows.map(toPgResume);
  }
  return localDb.resumes
    .filter((resume) => resume.userId === userId)
    .sort((a, b) => String(b.updatedAt).localeCompare(String(a.updatedAt)));
}

async function getResumeForUser(resumeId, userId) {
  if (usePostgres) {
    const { rows } = await pool.query(
      "SELECT * FROM resumes WHERE id = $1 AND user_id = $2 LIMIT 1",
      [resumeId, userId]
    );
    return toPgResume(rows[0]);
  }
  return localDb.resumes.find((resume) => resume._id === resumeId && resume.userId === userId) || null;
}

async function getPublicResume(resumeId) {
  if (usePostgres) {
    const { rows } = await pool.query(
      "SELECT * FROM resumes WHERE id = $1 AND public = true LIMIT 1",
      [resumeId]
    );
    return toPgResume(rows[0]);
  }
  return localDb.resumes.find((resume) => resume._id === resumeId && resume.public) || null;
}

async function updateResumeForUser(resumeId, userId, resumeData) {
  const clean = { ...resumeData, updatedAt: now() };
  delete clean._id;
  delete clean.userId;
  delete clean.createdAt;
  delete clean.updatedAt;

  if (usePostgres) {
    const { rows } = await pool.query(
      `UPDATE resumes
       SET title = $3,
           personal_info = $4::jsonb,
           professional_summary = $5,
           experience = $6::jsonb,
           education = $7::jsonb,
           project = $8::jsonb,
           skills = $9::jsonb,
           section_order = $10::jsonb,
           font_family = $11,
           density = $12,
           template = $13,
           accent_color = $14,
           public = $15,
           updated_at = NOW()
       WHERE id = $1 AND user_id = $2
       RETURNING *`,
      [
        resumeId,
        userId,
        clean.title || "Untitled Resume",
        JSON.stringify(clean.personal_info || {}),
        clean.professional_summary || "",
        JSON.stringify(clean.experience || []),
        JSON.stringify(clean.education || []),
        JSON.stringify(clean.project || []),
        JSON.stringify(clean.skills || []),
        JSON.stringify(clean.section_order || DEFAULT_SECTION_ORDER),
        clean.font_family || "Inter",
        clean.density || "comfortable",
        clean.template || "minimal-image",
        clean.accent_color || "#14B8A6",
        Boolean(clean.public)
      ]
    );
    return toPgResume(rows[0]);
  }

  const index = localDb.resumes.findIndex((resume) => resume._id === resumeId && resume.userId === userId);
  if (index === -1) return null;
  localDb.resumes[index] = { ...localDb.resumes[index], ...clean, updatedAt: now() };
  await saveLocalDb();
  return localDb.resumes[index];
}

async function deleteResumeForUser(resumeId, userId) {
  if (usePostgres) {
    const result = await pool.query("DELETE FROM resumes WHERE id = $1 AND user_id = $2", [resumeId, userId]);
    return result.rowCount > 0;
  }
  const before = localDb.resumes.length;
  localDb.resumes = localDb.resumes.filter((resume) => !(resume._id === resumeId && resume.userId === userId));
  await saveLocalDb();
  return before !== localDb.resumes.length;
}

function toPgJob(row) {
  if (!row) return null;
  return normalizeDoc({
    _id: row.id,
    userId: row.user_id,
    resumeId: row.resume_id,
    title: row.title,
    company: row.company,
    location: row.location,
    status: row.status,
    description: row.description,
    createdAt: row.created_at,
    updatedAt: row.updated_at
  });
}

async function listJobs(userId) {
  if (usePostgres) {
    const { rows } = await pool.query("SELECT * FROM jobs WHERE user_id = $1 ORDER BY updated_at DESC", [userId]);
    return rows.map(toPgJob);
  }
  return localDb.jobs
    .filter((job) => job.userId === userId)
    .sort((a, b) => String(b.updatedAt).localeCompare(String(a.updatedAt)));
}

async function createJob(userId, data) {
  const payload = {
    title: String(data.title || "Untitled job").trim() || "Untitled job",
    company: String(data.company || "").trim(),
    location: String(data.location || "").trim(),
    status: String(data.status || "Saved").trim() || "Saved",
    description: String(data.description || ""),
    resumeId: data.resumeId || null
  };

  if (usePostgres) {
    const jobId = id();
    const { rows } = await pool.query(
      `INSERT INTO jobs (id, user_id, resume_id, title, company, location, status, description)
       VALUES ($1, $2, $3, $4, $5, $6, $7, $8)
       RETURNING *`,
      [jobId, userId, payload.resumeId, payload.title, payload.company, payload.location, payload.status, payload.description]
    );
    return toPgJob(rows[0]);
  }

  const job = { _id: id(), userId, ...payload, createdAt: now(), updatedAt: now() };
  localDb.jobs.push(job);
  await saveLocalDb();
  return job;
}

async function updateJob(userId, jobId, data) {
  if (usePostgres) {
    const { rows } = await pool.query(
      `UPDATE jobs
       SET resume_id = $3,
           title = $4,
           company = $5,
           location = $6,
           status = $7,
           description = $8,
           updated_at = NOW()
       WHERE id = $1 AND user_id = $2
       RETURNING *`,
      [
        jobId,
        userId,
        data.resumeId || null,
        String(data.title || "Untitled job").trim() || "Untitled job",
        String(data.company || "").trim(),
        String(data.location || "").trim(),
        String(data.status || "Saved").trim() || "Saved",
        String(data.description || "")
      ]
    );
    return toPgJob(rows[0]);
  }

  const index = localDb.jobs.findIndex((job) => job._id === jobId && job.userId === userId);
  if (index === -1) return null;
  localDb.jobs[index] = {
    ...localDb.jobs[index],
    ...data,
    title: String(data.title || localDb.jobs[index].title || "Untitled job").trim() || "Untitled job",
    updatedAt: now()
  };
  await saveLocalDb();
  return localDb.jobs[index];
}

async function deleteJob(userId, jobId) {
  if (usePostgres) {
    const result = await pool.query("DELETE FROM jobs WHERE id = $1 AND user_id = $2", [jobId, userId]);
    return result.rowCount > 0;
  }
  const before = localDb.jobs.length;
  localDb.jobs = localDb.jobs.filter((job) => !(job._id === jobId && job.userId === userId));
  await saveLocalDb();
  return before !== localDb.jobs.length;
}

function signToken(userId) {
  return jwt.sign({ id: String(userId) }, JWT_SECRET, { expiresIn: "7d" });
}

async function auth(req, res, next) {
  try {
    const header = req.headers.authorization || "";
    const token = header.startsWith("Bearer ") ? header.slice(7) : header;
    if (!token) return res.status(401).json({ message: "Authentication required" });

    const decoded = jwt.verify(token, JWT_SECRET);
    const user = await findUserById(decoded.id);
    if (!user) return res.status(401).json({ message: "Invalid token" });

    req.userId = String(normalizeDoc(user)._id);
    req.user = user;
    next();
  } catch {
    res.status(401).json({ message: "Invalid or expired token" });
  }
}

function asyncRoute(handler) {
  return (req, res, next) => Promise.resolve(handler(req, res, next)).catch(next);
}

async function uploadImage(file) {
  if (!file) return "";

  const safeName = `${Date.now()}-${file.originalname.replace(/[^a-zA-Z0-9._-]/g, "-")}`;

  if (imagekit) {
    const result = await imagekit.upload({
      file: file.buffer.toString("base64"),
      fileName: safeName,
      folder: "/resume-builder"
    });
    return result.url;
  }

  await fs.writeFile(path.join(uploadDir, safeName), file.buffer);
  return `${SERVER_URL}/uploads/${safeName}`;
}

async function askGemini(prompt) {
  if (!genAI) return "";
  const model = genAI.getGenerativeModel({
    model: process.env.GEMINI_MODEL || "gemini-1.5-flash"
  });
  const result = await model.generateContent(prompt);
  return result.response.text().trim();
}

function enhanceLocally(section, content = "", context = {}) {
  const role = context.role || context.profession || "the target role";
  const base = String(content || "").trim();

  if (section === "summary") {
    if (!base) {
      return `Results-driven ${role} with hands-on experience building reliable products, collaborating across teams, and turning business requirements into polished user outcomes. Skilled at balancing clean execution with measurable impact.`;
    }
    return `${base.replace(/\s+/g, " ")} Demonstrates strong ownership, clear communication, and a practical focus on measurable outcomes.`;
  }

  if (section === "experience") {
    const line = base || `Built and improved production-ready solutions for ${role}.`;
    return line
      .split(/\n|\. /)
      .filter(Boolean)
      .map((item) => item.trim().replace(/^\W+/, ""))
      .map((item) => `Delivered ${item.charAt(0).toLowerCase()}${item.slice(1).replace(/\.$/, "")}.`)
      .slice(0, 4)
      .join("\n");
  }

  if (section === "project") {
    return base
      ? `${base.replace(/\s+/g, " ")} Focused on usability, maintainability, and clear technical execution.`
      : "Designed and built a production-style project with clear user workflows, reusable components, and a maintainable full stack architecture.";
  }

  if (section === "skills") {
    return base || "JavaScript, React, Node.js, Express, PostgreSQL, REST APIs, Git, Tailwind CSS, Problem Solving, Communication";
  }

  return base;
}

function extractJson(text) {
  const match = String(text).match(/\{[\s\S]*\}/);
  if (!match) throw new Error("No JSON object returned");
  return JSON.parse(match[0]);
}

function parseResumeLocally(title, resumeText) {
  const text = String(resumeText || "");
  const lines = text.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
  const email = text.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i)?.[0] || "";
  const phone = text.match(/(\+?\d[\d\s().-]{8,}\d)/)?.[0] || "";
  const website = text.match(/https?:\/\/[^\s]+/i)?.[0] || "";
  const skillBank = [
    "JavaScript",
    "React",
    "Node.js",
    "Express",
    "PostgreSQL",
    "Python",
    "SQL",
    "TypeScript",
    "Tailwind",
    "Git",
    "AWS",
    "Docker",
    "REST APIs",
    "HTML",
    "CSS"
  ];
  const lower = text.toLowerCase();
  const skills = skillBank.filter((skill) => lower.includes(skill.toLowerCase()));

  return {
    ...defaultResume("", title || "Uploaded Resume"),
    title: title || "Uploaded Resume",
    personal_info: {
      ...defaultResume("").personal_info,
      full_name: lines[0] || "",
      email,
      phone,
      website,
      profession: lines.find((line) => /developer|engineer|designer|analyst|manager/i.test(line)) || ""
    },
    professional_summary:
      lines.find((line) => line.length > 90) ||
      text.slice(0, 320).replace(/\s+/g, " "),
    skills: skills.length ? skills : ["Communication", "Problem Solving", "Teamwork"],
    experience: [],
    education: [],
    project: []
  };
}

const STOP_WORDS = new Set([
  "and",
  "the",
  "for",
  "with",
  "from",
  "that",
  "this",
  "you",
  "your",
  "are",
  "will",
  "have",
  "has",
  "was",
  "were",
  "our",
  "their",
  "they",
  "job",
  "role",
  "work",
  "team",
  "using",
  "including",
  "experience",
  "ability",
  "skills"
]);

function resumeToText(resume = {}) {
  const info = resume.personal_info || {};
  const chunks = [
    resume.title,
    info.full_name,
    info.profession,
    info.email,
    info.phone,
    info.location,
    info.linkedin,
    info.website,
    resume.professional_summary,
    ...(resume.skills || []),
    ...(resume.experience || []).flatMap((item) => [item.position, item.company, item.description]),
    ...(resume.project || []).flatMap((item) => [item.name, item.type, item.description]),
    ...(resume.education || []).flatMap((item) => [item.institution, item.degree, item.field, item.gpa])
  ];
  return chunks.filter(Boolean).join(" ");
}

function keywordCounts(text = "") {
  const counts = new Map();
  const matches = String(text).toLowerCase().match(/[a-z][a-z0-9+#.-]{2,}/g) || [];
  for (const raw of matches) {
    let word = raw.replace(/^[^a-z0-9]+|[^a-z0-9+#.]+$/g, "");
    if (word.endsWith(".js")) word = word.replace(".js", "");
    if (word === "apis") word = "api";
    if (word.length > 5 && word.endsWith("s")) word = word.slice(0, -1);
    if (!word || STOP_WORDS.has(word) || word.length < 3) continue;
    counts.set(word, (counts.get(word) || 0) + 1);
  }
  return counts;
}

function topKeywords(text = "", limit = 30) {
  return [...keywordCounts(text).entries()]
    .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
    .slice(0, limit)
    .map(([keyword]) => keyword);
}

function sectionChecks(resume = {}) {
  const info = resume.personal_info || {};
  return [
    {
      label: "Contact details",
      ok: Boolean(info.full_name && info.email && info.phone),
      detail: "Add full name, email, and phone number."
    },
    {
      label: "Professional summary",
      ok: String(resume.professional_summary || "").trim().length >= 70,
      detail: "Use a focused 2-4 line summary with target role, strengths, and impact."
    },
    {
      label: "Work experience",
      ok: (resume.experience || []).some((item) => item.position && item.company && item.description),
      detail: "Add at least one role with company, title, dates, and measurable bullets."
    },
    {
      label: "Skills",
      ok: (resume.skills || []).length >= 6,
      detail: "List at least 6 relevant hard and soft skills."
    },
    {
      label: "Education",
      ok: (resume.education || []).some((item) => item.institution || item.degree),
      detail: "Add education, certification, or training details."
    },
    {
      label: "Projects",
      ok: (resume.project || []).some((item) => item.name && item.description),
      detail: "Add projects when experience is limited or to prove technical ability."
    }
  ];
}

function analyzeResume(resume = {}, jobDescription = "") {
  const checks = sectionChecks(resume);
  const missingSections = checks.filter((check) => !check.ok).map((check) => check.label);
  const resumeText = resumeToText(resume);
  const resumeKeywords = new Set(topKeywords(resumeText, 80));
  const jobKeywords = topKeywords(jobDescription, 35);
  const matchedKeywords = jobKeywords.filter((keyword) => resumeKeywords.has(keyword));
  const missingKeywords = jobKeywords.filter((keyword) => !resumeKeywords.has(keyword)).slice(0, 16);
  const keywordScore = jobKeywords.length ? Math.round((matchedKeywords.length / jobKeywords.length) * 35) : 25;
  const sectionScore = Math.round((checks.filter((check) => check.ok).length / checks.length) * 45);
  const length = resumeText.trim().split(/\s+/).filter(Boolean).length;
  const lengthScore = length >= 260 && length <= 900 ? 20 : length >= 160 ? 14 : 8;
  const score = Math.max(0, Math.min(100, sectionScore + keywordScore + lengthScore));
  const suggestions = [
    ...checks.filter((check) => !check.ok).map((check) => check.detail),
    ...(missingKeywords.length ? [`Add important job keywords naturally: ${missingKeywords.slice(0, 8).join(", ")}.`] : []),
    ...(length < 260 ? ["Expand short sections with achievement-focused bullet points and measurable results."] : []),
    ...(length > 900 ? ["Trim older or less relevant details so the resume stays focused."] : [])
  ].slice(0, 8);

  return {
    score,
    checks,
    missingSections,
    matchedKeywords,
    missingKeywords,
    resumeKeywords: [...resumeKeywords].slice(0, 25),
    jobKeywords,
    suggestions
  };
}

function coverLetterLocally(resume = {}, job = {}) {
  const info = resume.personal_info || {};
  const name = info.full_name || "Your Name";
  const role = job.jobTitle || info.profession || "the role";
  const company = job.company || "your team";
  const skills = (resume.skills || []).slice(0, 5).join(", ") || "relevant technical and communication skills";
  const strongest = (resume.experience || [])[0]?.description || resume.professional_summary || "I have built practical projects, collaborated across teams, and delivered polished outcomes.";

  return `Dear Hiring Manager,

I am excited to apply for ${role} at ${company}. My background aligns well with this opportunity because I bring hands-on experience with ${skills}, along with a practical focus on clear execution and measurable results.

${String(strongest).replace(/\s+/g, " ").slice(0, 420)}

I would welcome the opportunity to discuss how my experience can support ${company}'s goals. Thank you for your time and consideration.

Sincerely,
${name}`;
}

function guidedTipsLocally(resume = {}, section = "summary") {
  const analysis = analyzeResume(resume);
  const tips = {
    summary: [
      "Start with the target role and years or type of experience.",
      "Mention 2-3 strongest skills that match the job.",
      "End with the kind of impact you create."
    ],
    experience: [
      "Begin bullets with action verbs such as Built, Improved, Reduced, Led, or Automated.",
      "Add numbers wherever possible: users, speed, cost, revenue, accuracy, or time saved.",
      "Keep each bullet tied to a business or user outcome."
    ],
    skills: [
      "Put the most job-relevant skills first.",
      "Split broad skills into specific tools or technologies.",
      "Avoid skills you cannot discuss in an interview."
    ],
    default: [
      "Use concise, concrete wording.",
      "Mirror important job description language naturally.",
      "Remove repeated ideas and make every line earn its place."
    ]
  };
  return [...(tips[section] || tips.default), ...analysis.suggestions.slice(0, 3)].slice(0, 6);
}

app.get("/api/health", (req, res) => {
  res.json({ ok: true, storage: usePostgres ? "postgresql" : "local-json" });
});

app.post(
  "/api/users/verify-email",
  asyncRoute(async (req, res) => {
    const user = await verifyUserEmail(req.body.token);
    if (!user) return res.status(400).json({ message: "Invalid verification token" });
    res.json({ message: "Email verified", user: publicUser(user) });
  })
);

app.post(
  "/api/users/forgot-password",
  asyncRoute(async (req, res) => {
    if (!req.body.email) return res.status(400).json({ message: "Email is required" });
    const reset = await createPasswordReset(req.body.email);
    const payload = { message: "If that email exists, a reset link has been prepared." };
    if (reset && process.env.NODE_ENV !== "production") {
      payload.resetToken = reset.resetToken;
      payload.resetUrl = `${CLIENT_URL.split(",")[0]}/auth?reset=${reset.resetToken}`;
      payload.expiresAt = reset.resetExpires;
    }
    res.json(payload);
  })
);

app.post(
  "/api/users/reset-password",
  asyncRoute(async (req, res) => {
    const { token, password } = req.body;
    if (!token || !password) return res.status(400).json({ message: "Token and password are required" });
    if (String(password).length < 6) return res.status(400).json({ message: "Password must be at least 6 characters" });
    const user = await resetPassword(token, password);
    if (!user) return res.status(400).json({ message: "Invalid or expired reset token" });
    res.json({ message: "Password updated" });
  })
);

app.post(
  "/api/users/:action",
  asyncRoute(async (req, res) => {
    const { action } = req.params;
    const { name, email, password } = req.body;

    if (!["login", "signup"].includes(action)) {
      return res.status(404).json({ message: "Unknown auth action" });
    }

    if (!email || !password) {
      return res.status(400).json({ message: "Email and password are required" });
    }

    if (action === "signup") {
      if (!name) return res.status(400).json({ message: "Name is required" });
      const existing = await findUserByEmail(email);
      if (existing) return res.status(409).json({ message: "Email already exists" });

      const hashed = await bcrypt.hash(password, 10);
      const verificationToken = id();
      const user = await createUser({ name, email, password: hashed, emailVerified: false, verificationToken });
      const userData = publicUser(user);
      return res.status(201).json({
        message: "Account created successfully",
        token: signToken(userData._id),
        user: userData,
        verificationToken,
        verificationUrl: `${CLIENT_URL.split(",")[0]}/auth?verify=${verificationToken}`
      });
    }

    const user = await findUserByEmail(email);
    if (!user) return res.status(401).json({ message: "Invalid email or password" });
    const valid = await bcrypt.compare(password, normalizeDoc(user).password);
    if (!valid) return res.status(401).json({ message: "Invalid email or password" });

    const userData = publicUser(user);
    res.json({ message: "Logged in successfully", token: signToken(userData._id), user: userData });
  })
);

app.get(
  "/api/users/data",
  auth,
  asyncRoute(async (req, res) => {
    res.json({ user: publicUser(req.user) });
  })
);

app.get(
  "/api/users/resumes",
  auth,
  asyncRoute(async (req, res) => {
    res.json({ resumes: await listUserResumes(req.userId) });
  })
);

app.post(
  "/api/resumes/create",
  auth,
  asyncRoute(async (req, res) => {
    const resume = await createResume(defaultResume(req.userId, req.body.title || "Untitled Resume"));
    res.status(201).json({ message: "Resume created", resume });
  })
);

app.get(
  "/api/resumes/get/:resumeId",
  auth,
  asyncRoute(async (req, res) => {
    const resume = await getResumeForUser(req.params.resumeId, req.userId);
    if (!resume) return res.status(404).json({ message: "Resume not found" });
    res.json({ resume });
  })
);

app.put(
  "/api/resumes/update",
  auth,
  upload.single("image"),
  asyncRoute(async (req, res) => {
    const resumeId = req.body.resumeId || req.body._id;
    if (!resumeId) return res.status(400).json({ message: "Resume id is required" });

    const resumeData =
      typeof req.body.resumeData === "string"
        ? JSON.parse(req.body.resumeData)
        : req.body.resumeData || req.body;

    if (req.file) {
      const imageUrl = await uploadImage(req.file);
      resumeData.personal_info = { ...(resumeData.personal_info || {}), image: imageUrl };
    }

    const resume = await updateResumeForUser(resumeId, req.userId, resumeData);
    if (!resume) return res.status(404).json({ message: "Resume not found" });
    res.json({ message: "Changes saved", resume });
  })
);

app.delete(
  "/api/resumes/delete/:resumeId",
  auth,
  asyncRoute(async (req, res) => {
    const deleted = await deleteResumeForUser(req.params.resumeId, req.userId);
    if (!deleted) return res.status(404).json({ message: "Resume not found" });
    res.json({ message: "Resume deleted" });
  })
);

app.get(
  "/api/resumes/public/:resumeId",
  asyncRoute(async (req, res) => {
    const resume = await getPublicResume(req.params.resumeId);
    if (!resume) return res.status(404).json({ message: "Resume not found" });
    res.json({ resume });
  })
);

app.get(
  "/api/jobs",
  auth,
  asyncRoute(async (req, res) => {
    res.json({ jobs: await listJobs(req.userId) });
  })
);

app.post(
  "/api/jobs",
  auth,
  asyncRoute(async (req, res) => {
    const job = await createJob(req.userId, req.body);
    res.status(201).json({ message: "Job saved", job });
  })
);

app.put(
  "/api/jobs/:jobId",
  auth,
  asyncRoute(async (req, res) => {
    const job = await updateJob(req.userId, req.params.jobId, req.body);
    if (!job) return res.status(404).json({ message: "Job not found" });
    res.json({ message: "Job updated", job });
  })
);

app.delete(
  "/api/jobs/:jobId",
  auth,
  asyncRoute(async (req, res) => {
    const deleted = await deleteJob(req.userId, req.params.jobId);
    if (!deleted) return res.status(404).json({ message: "Job not found" });
    res.json({ message: "Job deleted" });
  })
);

app.post(
  "/api/ai/ats-score",
  auth,
  asyncRoute(async (req, res) => {
    const result = analyzeResume(req.body.resume || {}, req.body.jobDescription || "");
    res.json(result);
  })
);

app.post(
  "/api/ai/job-match",
  auth,
  asyncRoute(async (req, res) => {
    const result = analyzeResume(req.body.resume || {}, req.body.jobDescription || "");
    res.json({
      score: result.score,
      matchedKeywords: result.matchedKeywords,
      missingKeywords: result.missingKeywords,
      jobKeywords: result.jobKeywords,
      suggestions: result.suggestions
    });
  })
);

app.post(
  "/api/ai/cover-letter",
  auth,
  asyncRoute(async (req, res) => {
    const resume = req.body.resume || {};
    const job = {
      jobTitle: req.body.jobTitle,
      company: req.body.company,
      jobDescription: req.body.jobDescription
    };
    let coverLetter = "";
    if (genAI) {
      try {
        coverLetter = await askGemini(`Write a concise, professional cover letter. Return plain text only.

Resume:
${resumeToText(resume)}

Target job:
${JSON.stringify(job)}`);
      } catch (error) {
        console.warn("Gemini cover letter failed:", error.message);
      }
    }
    res.json({ coverLetter: coverLetter || coverLetterLocally(resume, job) });
  })
);

app.post(
  "/api/ai/guided-help",
  auth,
  asyncRoute(async (req, res) => {
    const resume = req.body.resume || {};
    const section = req.body.section || "summary";
    let tips = [];
    if (genAI) {
      try {
        const response = await askGemini(`Return JSON only in this shape: {"tips":["tip"]}.
Give practical resume writing tips for the "${section}" section based on this resume:
${JSON.stringify(resume)}`);
        tips = extractJson(response).tips || [];
      } catch (error) {
        console.warn("Gemini guided help failed:", error.message);
      }
    }
    res.json({ tips: tips.length ? tips.slice(0, 6) : guidedTipsLocally(resume, section) });
  })
);

app.post(
  "/api/ai/enhance",
  auth,
  asyncRoute(async (req, res) => {
    const { section, content, context } = req.body;
    const prompt = `Improve this resume ${section} for a professional resume. Return only the rewritten content, no markdown.

Context: ${JSON.stringify(context || {})}
Content:
${content || ""}`;

    let text = "";
    try {
      text = await askGemini(prompt);
    } catch (error) {
      console.warn("Gemini enhance failed:", error.message);
    }

    res.json({ text: text || enhanceLocally(section, content, context) });
  })
);

app.post(
  "/api/ai/upload-resume",
  auth,
  asyncRoute(async (req, res) => {
    const { title, resumeText } = req.body;
    let parsed = null;

    if (genAI && resumeText) {
      try {
        const response = await askGemini(`Extract this resume into JSON matching this shape:
{
  "title": "string",
  "personal_info": {"full_name": "", "email": "", "phone": "", "location": "", "linkedin": "", "website": "", "profession": "", "image": ""},
  "professional_summary": "",
  "skills": [],
  "experience": [{"company": "", "position": "", "start_date": "YYYY-MM", "end_date": "YYYY-MM", "description": "", "is_current": false}],
  "education": [{"institution": "", "degree": "", "field": "", "graduation_date": "YYYY-MM", "gpa": ""}],
  "project": [{"name": "", "type": "", "description": ""}]
}
Return only JSON.

Resume text:
${resumeText}`);
        parsed = extractJson(response);
      } catch (error) {
        console.warn("Gemini parse failed:", error.message);
      }
    }

    const fallback = parseResumeLocally(title, resumeText);
    const resume = await createResume({
      ...fallback,
      ...(parsed || {}),
      userId: req.userId,
      title: parsed?.title || title || fallback.title,
      template: parsed?.template || "minimal-image",
      accent_color: parsed?.accent_color || "#14B8A6",
      public: false
    });

    res.status(201).json({ message: "Resume imported", resumeId: resume._id });
  })
);

if (process.env.NODE_ENV === "production") {
  app.use(express.static(clientDistDir));
  app.get(/^\/(?!api|uploads).*/, (req, res) => {
    res.sendFile(path.join(clientDistDir, "index.html"));
  });
}

app.use((error, req, res, next) => {
  console.error(error);
  res.status(error.status || 500).json({ message: error.message || "Something went wrong" });
});

app.listen(PORT, () => {
  console.log(`Resume Builder API running on ${SERVER_URL}`);
  console.log(`Storage: ${usePostgres ? "PostgreSQL" : "local JSON"}`);
});
