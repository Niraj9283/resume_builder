import { createContext, useContext, useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import { Link, Navigate, Route, Routes, useNavigate, useParams, useSearchParams } from "react-router-dom";
import axios from "axios";
import { Document, HeadingLevel, Packer, Paragraph, TextRun } from "docx";
import toast, { Toaster } from "react-hot-toast";
import html2canvas from "html2canvas";
import jsPDF from "jspdf";
import * as pdfjsLib from "pdfjs-dist";
import pdfWorker from "pdfjs-dist/build/pdf.worker.mjs?url";
import {
  ArrowLeft,
  ArrowRight,
  ArrowDown,
  ArrowUp,
  Briefcase,
  Bot,
  Check,
  ChevronLeft,
  ChevronRight,
  ClipboardCheck,
  Download,
  Edit3,
  FileDown,
  FileText,
  Globe,
  GraduationCap,
  GripVertical,
  ImagePlus,
  KeyRound,
  LayoutDashboard,
  Link2,
  Loader2,
  Lock,
  LogOut,
  Mail,
  Menu,
  Palette,
  Phone,
  Plus,
  Save,
  Send,
  Settings,
  ShieldCheck,
  Sparkles,
  Trash2,
  Upload,
  User,
  Wand2,
  X
} from "lucide-react";
import logo from "./assets/assets/logo.svg";
import dummyProfile from "./assets/assets/dummy_profile.png";
import { dummyResumeData } from "./assets/assets/assets.js";
import ClassicTemplate from "./assets/assets/templates/ClassicTemplate.jsx";
import MinimalImageTemplate from "./assets/assets/templates/MinimalImageTemplate.jsx";
import MinimalTemplate from "./assets/assets/templates/MinimalTemplate.jsx";
import ModernTemplate from "./assets/assets/templates/ModernTemplate.jsx";

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;

const API_URL = import.meta.env.VITE_API_URL || (import.meta.env.PROD ? "" : "http://localhost:5000");
const api = axios.create({ baseURL: API_URL });
const AuthContext = createContext(null);
const DEFAULT_SECTION_ORDER = ["summary", "experience", "projects", "education", "skills"];

const templates = [
  { id: "classic", name: "Classic", preview: "Traditional resume with strong section hierarchy." },
  { id: "modern", name: "Modern", preview: "Color-forward header with clean editorial spacing." },
  { id: "minimal-image", name: "Minimal Image", preview: "Profile image layout with a calm sidebar." },
  { id: "minimal", name: "Minimal", preview: "Whitespace-first layout for simple content." },
  { id: "executive", name: "Executive", preview: "Dense leadership layout with a bold nameplate." },
  { id: "compact", name: "Compact", preview: "ATS-friendly single column with tight spacing." }
];

const accents = ["#14B8A6", "#6366F1", "#3B82F6", "#9333EA", "#D97706", "#DC2626", "#0284C7", "#16A34A"];
const fontOptions = ["Inter", "Arial", "Georgia", "Times New Roman", "Verdana"];
const densityOptions = [
  { id: "comfortable", label: "Comfortable" },
  { id: "compact", label: "Compact" }
];

function emptyResume(title = "Untitled Resume") {
  return {
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

function authHeader(token) {
  return { Authorization: token };
}

function cleanResumeForSave(resume) {
  const copy = JSON.parse(JSON.stringify({
    ...resume,
    personal_info: {
      ...resume.personal_info,
      image: typeof resume.personal_info?.image === "string" ? resume.personal_info.image : ""
    }
  }));
  return copy;
}

function formatDate(value) {
  if (!value) return "Recently";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "Recently";
  return date.toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" });
}

function getFileImageUrl(value) {
  if (!value) return "";
  if (typeof value === "string") return value;
  return URL.createObjectURL(value);
}

function safeFileName(value, fallback = "resume") {
  return String(value || fallback)
    .replace(/[\\/:*?"<>|]+/g, "-")
    .replace(/\s+/g, " ")
    .trim() || fallback;
}

function formatMonth(value) {
  if (!value) return "";
  const [year, month] = String(value).split("-");
  if (!year || !month) return value;
  return new Date(Number(year), Number(month) - 1).toLocaleDateString("en-US", { month: "short", year: "numeric" });
}

function alphaColor(hex, alpha = 0.1) {
  const clean = String(hex || "#14B8A6").replace("#", "");
  const value = clean.length === 3 ? clean.split("").map((item) => item + item).join("") : clean;
  const number = Number.parseInt(value, 16);
  if (Number.isNaN(number)) return `rgba(20, 184, 166, ${alpha})`;
  const r = (number >> 16) & 255;
  const g = (number >> 8) & 255;
  const b = number & 255;
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

function waitForImages(element) {
  const images = [...element.querySelectorAll("img")];
  return Promise.all(
    images.map((image) => {
      if (image.complete) return Promise.resolve();
      return new Promise((resolve) => {
        image.onload = resolve;
        image.onerror = resolve;
      });
    })
  );
}

async function renderResumeForExport(resume) {
  const host = document.createElement("div");
  host.style.cssText = "position:fixed;left:-10000px;top:0;width:816px;background:#fff;z-index:-1;";
  document.body.appendChild(host);
  const root = createRoot(host);
  root.render(<ResumePreview data={resume} classes="border-0 shadow-none" exportMode />);
  await new Promise((resolve) => setTimeout(resolve, 250));
  await document.fonts?.ready;
  const element = host.querySelector("[data-resume-preview]");
  if (!element) throw new Error("Resume preview could not be prepared for export");
  await waitForImages(element);
  return { host, root, element };
}

function docxHeading(text) {
  return new Paragraph({
    text: String(text || "").toUpperCase(),
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 80 }
  });
}

function docxText(text, options = {}) {
  return new Paragraph({
    children: [new TextRun({ text: String(text || ""), bold: options.bold })],
    bullet: options.bullet ? { level: 0 } : undefined,
    spacing: { after: options.after ?? 80 }
  });
}

async function exportResumeDocx(resume) {
  const info = resume.personal_info || {};
  const children = [
    new Paragraph({
      children: [new TextRun({ text: info.full_name || resume.title || "Resume", bold: true, size: 34 })],
      spacing: { after: 80 }
    }),
    docxText([info.email, info.phone, info.location, info.linkedin, info.website].filter(Boolean).join(" | "), { after: 160 })
  ];

  if (resume.professional_summary) {
    children.push(docxHeading("Professional Summary"), docxText(resume.professional_summary, { after: 160 }));
  }

  if ((resume.experience || []).length) {
    children.push(docxHeading("Experience"));
    resume.experience.forEach((item) => {
      children.push(docxText(`${item.position || "Role"}${item.company ? `, ${item.company}` : ""}`, { bold: true, after: 20 }));
      children.push(docxText(`${formatMonth(item.start_date)} - ${item.is_current ? "Present" : formatMonth(item.end_date)}`.trim(), { after: 40 }));
      String(item.description || "")
        .split(/\r?\n/)
        .filter(Boolean)
        .forEach((line) => children.push(docxText(line.trim(), { bullet: true })));
    });
  }

  if ((resume.project || []).length) {
    children.push(docxHeading("Projects"));
    resume.project.forEach((item) => {
      children.push(docxText(item.name || "Project", { bold: true, after: 20 }));
      if (item.description) children.push(docxText(item.description));
    });
  }

  if ((resume.education || []).length) {
    children.push(docxHeading("Education"));
    resume.education.forEach((item) => {
      children.push(docxText([item.degree, item.field, item.institution].filter(Boolean).join(", "), { bold: true, after: 20 }));
      if (item.graduation_date || item.gpa) children.push(docxText([formatMonth(item.graduation_date), item.gpa && `GPA: ${item.gpa}`].filter(Boolean).join(" | ")));
    });
  }

  if ((resume.skills || []).length) {
    children.push(docxHeading("Skills"), docxText((resume.skills || []).join(", ")));
  }

  const doc = new Document({ sections: [{ children }] });
  const blob = await Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `${safeFileName(resume.title)}.docx`;
  link.click();
  URL.revokeObjectURL(url);
}

function useAuth() {
  return useContext(AuthContext);
}

function AuthProvider({ children }) {
  const [token, setToken] = useState(() => localStorage.getItem("token") || "");
  const [user, setUser] = useState(null);
  const [checking, setChecking] = useState(Boolean(token));

  useEffect(() => {
    if (!token) {
      setChecking(false);
      return;
    }

    api
      .get("/api/users/data", { headers: authHeader(token) })
      .then(({ data }) => setUser(data.user))
      .catch(() => {
        localStorage.removeItem("token");
        setToken("");
        setUser(null);
      })
      .finally(() => setChecking(false));
  }, [token]);

  const value = useMemo(
    () => ({
      token,
      user,
      checking,
      login(payload) {
        localStorage.setItem("token", payload.token);
        setToken(payload.token);
        setUser(payload.user);
      },
      logout() {
        localStorage.removeItem("token");
        setToken("");
        setUser(null);
      }
    }),
    [checking, token, user]
  );

  return <AuthContext.Provider value={value}>{children}</AuthContext.Provider>;
}

function ProtectedRoute({ children }) {
  const { token, checking } = useAuth();
  if (checking) return <PageLoader />;
  if (!token) return <Navigate to="/auth?state=login" replace />;
  return children;
}

function App() {
  return (
    <AuthProvider>
      <Toaster position="top-right" />
      <Routes>
        <Route path="/" element={<Home />} />
        <Route path="/auth" element={<AuthPage />} />
        <Route
          path="/app"
          element={
            <ProtectedRoute>
              <Dashboard />
            </ProtectedRoute>
          }
        />
        <Route
          path="/app/builder/:resumeId"
          element={
            <ProtectedRoute>
              <Builder />
            </ProtectedRoute>
          }
        />
        <Route path="/resume/:resumeId" element={<PublicResume />} />
        <Route path="*" element={<Navigate to="/" replace />} />
      </Routes>
    </AuthProvider>
  );
}

function Header() {
  const { token, logout } = useAuth();
  const [open, setOpen] = useState(false);

  const nav = [
    { label: "Home", href: "/" },
    { label: "Features", href: "/#features" },
    { label: "Pricing", href: "/#pricing" },
    { label: "About", href: "/#about" }
  ];

  return (
    <header className="sticky top-0 z-50 border-b border-slate-200/80 bg-white/90 backdrop-blur">
      <div className="mx-auto flex h-16 max-w-7xl items-center justify-between px-4">
        <Link to="/" className="flex items-center gap-3">
          <img src={logo} alt="Resume Builder" className="h-8 w-auto" />
        </Link>

        <nav className="hidden items-center gap-8 text-sm font-medium text-slate-600 md:flex">
          {nav.map((item) => (
            <a key={item.label} href={item.href} className="transition hover:text-slate-950">
              {item.label}
            </a>
          ))}
        </nav>

        <div className="hidden items-center gap-3 md:flex">
          {token ? (
            <>
              <Link className="rounded-full bg-green-500 px-6 py-2 text-sm font-medium text-white transition hover:bg-green-600" to="/app">
                Dashboard
              </Link>
              <button className="rounded-full p-2 text-slate-500 transition hover:bg-slate-100 hover:text-slate-900" onClick={logout} title="Logout">
                <LogOut className="size-5" />
              </button>
            </>
          ) : (
            <>
              <Link className="rounded-full px-5 py-2 text-sm font-medium text-slate-700 transition hover:text-slate-950" to="/auth?state=login">
                Login
              </Link>
              <Link className="rounded-full bg-green-500 px-6 py-2 text-sm font-medium text-white transition hover:bg-green-600" to="/auth?state=signup">
                Get started
              </Link>
            </>
          )}
        </div>

        <button className="rounded-lg p-2 md:hidden" onClick={() => setOpen(true)} aria-label="Open menu">
          <Menu className="size-6" />
        </button>
      </div>

      {open && (
        <div className="fixed inset-0 z-50 bg-slate-950/50 md:hidden">
          <div className="ml-auto flex h-full w-72 flex-col bg-white p-6 shadow-xl">
            <button className="self-end rounded-lg p-2 text-slate-500" onClick={() => setOpen(false)} aria-label="Close menu">
              <X className="size-5" />
            </button>
            <div className="mt-6 grid gap-4 text-lg font-medium text-slate-700">
              {nav.map((item) => (
                <a key={item.label} href={item.href} onClick={() => setOpen(false)}>
                  {item.label}
                </a>
              ))}
              <Link to={token ? "/app" : "/auth?state=login"} onClick={() => setOpen(false)} className="mt-4 rounded-full bg-green-500 px-6 py-3 text-center text-white">
                {token ? "Dashboard" : "Login"}
              </Link>
            </div>
          </div>
        </div>
      )}
    </header>
  );
}

function Home() {
  const { token } = useAuth();
  const demoResume = { ...dummyResumeData[0], personal_info: { ...dummyResumeData[0].personal_info, image: dummyProfile } };

  return (
    <div className="min-h-screen bg-white text-slate-900">
      <Header />
      <main>
        <section className="relative overflow-hidden bg-[radial-gradient(circle_at_top,#dcfce7_0,#ffffff_38%,#ffffff_100%)]">
          <div className="mx-auto grid min-h-[calc(100vh-4rem)] max-w-7xl items-center gap-12 px-4 py-16 lg:grid-cols-[1fr_0.9fr]">
            <div className="max-w-3xl">
              <div className="mb-5 inline-flex items-center gap-2 rounded-full bg-green-100 px-4 py-2 text-sm font-medium text-green-700">
                <Sparkles className="size-4" />
                Used by 10,000+ users
              </div>
              <h1 className="text-5xl font-semibold leading-tight tracking-normal text-slate-950 md:text-6xl">
                Land your dream job with <span className="text-green-600">AI-powered</span> resumes.
              </h1>
              <p className="mt-6 max-w-xl text-lg leading-8 text-slate-600">
                Create, edit, share, and download professional resumes with a full stack builder that keeps every section live and editable.
              </p>
              <div className="mt-8 flex flex-wrap items-center gap-4">
                <Link to={token ? "/app" : "/auth?state=signup"} className="inline-flex h-12 items-center gap-2 rounded-full bg-green-500 px-8 font-medium text-white ring-1 ring-green-400 ring-offset-2 transition hover:bg-green-600">
                  Get started
                  <ArrowRight className="size-4" />
                </Link>
                <a href="#demo" className="inline-flex h-12 items-center gap-2 rounded-full border border-slate-300 px-8 font-medium text-slate-700 transition hover:border-slate-400 hover:bg-slate-50">
                  Try demo
                </a>
              </div>
            </div>
            <div id="demo" className="relative mx-auto w-full max-w-[430px] rounded-lg border border-slate-200 bg-white p-3 shadow-2xl">
              <ResumePreview data={demoResume} scale="small" />
            </div>
          </div>
        </section>

        <section id="features" className="mx-auto max-w-7xl px-4 py-20">
          <div className="mx-auto max-w-2xl text-center">
            <div className="mx-auto mb-4 inline-flex items-center gap-2 rounded-full bg-green-50 px-4 py-2 text-sm font-medium text-green-700">
              <Bot className="size-4" />
              Simple Process
            </div>
            <h2 className="text-3xl font-semibold text-slate-950 md:text-4xl">Build your resume</h2>
            <p className="mt-4 text-slate-600">A focused workflow for creating, improving, and sharing polished resumes in minutes.</p>
          </div>
          <div className="mt-14 grid gap-6 md:grid-cols-3">
            {[
              { icon: FileText, title: "Live resume editor", text: "Edit every section while the final resume updates beside the form." },
              { icon: Wand2, title: "AI content help", text: "Enhance summaries, job descriptions, project descriptions, and skills." },
              { icon: Globe, title: "Shareable links", text: "Publish resumes publicly and send a direct link to recruiters." }
            ].map((item) => (
              <div key={item.title} className="rounded-lg border border-slate-200 bg-white p-6 shadow-sm">
                <item.icon className="mb-5 size-8 text-green-600" />
                <h3 className="text-lg font-semibold text-slate-900">{item.title}</h3>
                <p className="mt-3 text-sm leading-6 text-slate-600">{item.text}</p>
              </div>
            ))}
          </div>
        </section>

        <section id="pricing" className="border-y border-slate-200 bg-slate-50">
          <div className="mx-auto max-w-7xl px-4 py-16">
            <div className="grid items-center gap-8 md:grid-cols-[1fr_auto]">
              <div>
                <h2 className="text-3xl font-semibold text-slate-950">Free local project</h2>
                <p className="mt-3 max-w-2xl text-slate-600">Run it locally, connect your PostgreSQL, Gemini, and ImageKit keys, then deploy it like any full stack app.</p>
              </div>
              <Link to="/auth?state=signup" className="inline-flex h-12 items-center justify-center rounded-full bg-slate-950 px-8 font-medium text-white transition hover:bg-slate-800">
                Create a Resume
              </Link>
            </div>
          </div>
        </section>
      </main>
      <Footer />
    </div>
  );
}

function Footer() {
  return (
    <footer id="about" className="bg-white">
      <div className="mx-auto flex max-w-7xl flex-col gap-6 px-4 py-10 text-sm text-slate-500 md:flex-row md:items-center md:justify-between">
        <div className="flex items-center gap-3">
          <img src={logo} alt="Resume Builder" className="h-7 w-auto" />
          <span>(c) 2025 Resume Builder</span>
        </div>
        <div className="flex flex-wrap gap-5">
          <a href="#features" className="hover:text-slate-900">Features</a>
          <a href="#pricing" className="hover:text-slate-900">Pricing</a>
          <a href="mailto:support@example.com" className="hover:text-slate-900">Contact</a>
        </div>
      </div>
    </footer>
  );
}

function AuthPage() {
  const [params, setParams] = useSearchParams();
  const state = params.get("state") === "signup" ? "signup" : "login";
  const resetToken = params.get("reset") || "";
  const verifyToken = params.get("verify") || "";
  const [form, setForm] = useState({ name: "", email: "", password: "" });
  const [loading, setLoading] = useState(false);
  const [notice, setNotice] = useState("");
  const auth = useAuth();
  const navigate = useNavigate();

  useEffect(() => {
    if (!verifyToken) return;
    api
      .post("/api/users/verify-email", { token: verifyToken })
      .then(({ data }) => {
        setNotice(data.message);
        toast.success(data.message);
        setParams({ state: "login" });
      })
      .catch((error) => toast.error(error.response?.data?.message || error.message));
  }, [verifyToken, setParams]);

  async function submit(event) {
    event.preventDefault();
    setLoading(true);
    try {
      if (resetToken) {
        const { data } = await api.post("/api/users/reset-password", { token: resetToken, password: form.password });
        toast.success(data.message);
        setParams({ state: "login" });
        setForm({ name: "", email: "", password: "" });
        return;
      }
      const { data } = await api.post(`/api/users/${state}`, form);
      auth.login(data);
      toast.success(data.message);
      navigate("/app");
    } catch (error) {
      toast.error(error.response?.data?.message || error.message);
    } finally {
      setLoading(false);
    }
  }

  async function requestPasswordReset() {
    if (!form.email) return toast.error("Enter your email first");
    setLoading(true);
    try {
      const { data } = await api.post("/api/users/forgot-password", { email: form.email });
      setNotice(data.resetToken ? `Reset token: ${data.resetToken}` : data.message);
      toast.success(data.message);
    } catch (error) {
      toast.error(error.response?.data?.message || error.message);
    } finally {
      setLoading(false);
    }
  }

  return (
    <div className="flex min-h-screen items-center justify-center bg-gray-50 px-4">
      <form onSubmit={submit} className="w-full max-w-[370px] rounded-2xl border border-gray-300/70 bg-white px-8 py-10 text-center shadow-sm">
        <Link to="/" className="mb-8 inline-flex justify-center">
          <img src={logo} alt="Resume Builder" className="h-8 w-auto" />
        </Link>
        <h1 className="text-3xl font-medium text-gray-900">{resetToken ? "Reset password" : state === "login" ? "Login" : "Sign up"}</h1>
        <p className="mt-2 text-sm text-gray-500">{resetToken ? "Choose a new password for your account" : `Please ${state === "login" ? "sign in" : "create an account"} to continue`}</p>
        {notice && <p className="mt-4 rounded-lg bg-blue-50 px-3 py-2 text-left text-xs text-blue-700">{notice}</p>}

        <div className="mt-8 space-y-4 text-left">
          {state === "signup" && !resetToken && (
            <Field label="Full name" icon={User} value={form.name} onChange={(event) => setForm({ ...form, name: event.target.value })} required />
          )}
          {!resetToken && <Field label="Email" icon={Mail} type="email" value={form.email} onChange={(event) => setForm({ ...form, email: event.target.value })} required />}
          <Field label="Password" icon={Lock} type="password" value={form.password} onChange={(event) => setForm({ ...form, password: event.target.value })} required />
        </div>

        <button disabled={loading} className="mt-7 inline-flex h-11 w-full items-center justify-center gap-2 rounded-full bg-green-500 font-medium text-white transition hover:bg-green-600 disabled:opacity-70">
          {loading && <Loader2 className="size-4 animate-spin" />}
          {resetToken ? "Update password" : state === "login" ? "Login" : "Create account"}
        </button>

        {state === "login" && !resetToken && (
          <button type="button" onClick={requestPasswordReset} className="mt-4 inline-flex items-center gap-2 text-sm font-medium text-slate-500 hover:text-slate-900">
            <KeyRound className="size-4" />
            Forgot password?
          </button>
        )}

        <p className="mt-6 text-sm text-gray-500">
          {state === "login" ? "Don't have an account?" : "Already have an account?"}{" "}
          <button
            type="button"
            className="font-medium text-green-600 hover:text-green-700"
            onClick={() => setParams({ state: state === "login" ? "signup" : "login" })}
          >
            {state === "login" ? "Sign up" : "Login"}
          </button>
        </p>
      </form>
    </div>
  );
}

function Dashboard() {
  const { token, user, logout } = useAuth();
  const [resumes, setResumes] = useState([]);
  const [loading, setLoading] = useState(true);
  const [createOpen, setCreateOpen] = useState(false);
  const [uploadOpen, setUploadOpen] = useState(false);
  const [title, setTitle] = useState("");
  const [file, setFile] = useState(null);
  const [busy, setBusy] = useState(false);
  const navigate = useNavigate();

  async function loadResumes() {
    setLoading(true);
    try {
      const { data } = await api.get("/api/users/resumes", { headers: authHeader(token) });
      setResumes(data.resumes);
    } catch (error) {
      toast.error(error.response?.data?.message || error.message);
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    loadResumes();
  }, []);

  async function createResume(event) {
    event.preventDefault();
    setBusy(true);
    try {
      const { data } = await api.post("/api/resumes/create", { title: title || "Untitled Resume" }, { headers: authHeader(token) });
      toast.success(data.message);
      navigate(`/app/builder/${data.resume._id}`);
    } catch (error) {
      toast.error(error.response?.data?.message || error.message);
    } finally {
      setBusy(false);
    }
  }

  async function uploadResume(event) {
    event.preventDefault();
    if (!file) return toast.error("Choose a file first");
    setBusy(true);
    try {
      const resumeText = await extractResumeText(file);
      const { data } = await api.post("/api/ai/upload-resume", { title: title || file.name.replace(/\.[^.]+$/, ""), resumeText }, { headers: authHeader(token) });
      toast.success(data.message);
      navigate(`/app/builder/${data.resumeId}`);
    } catch (error) {
      toast.error(error.response?.data?.message || error.message);
    } finally {
      setBusy(false);
    }
  }

  async function renameResume(resume) {
    const next = window.prompt("Edit Resume Title", resume.title);
    if (!next || next === resume.title) return;
    try {
      const { data } = await api.put("/api/resumes/update", { resumeId: resume._id, resumeData: { ...resume, title: next } }, { headers: authHeader(token) });
      setResumes((items) => items.map((item) => (item._id === resume._id ? data.resume : item)));
      toast.success("Title updated");
    } catch (error) {
      toast.error(error.response?.data?.message || error.message);
    }
  }

  async function deleteResume(resumeId) {
    if (!window.confirm("Are you sure you want to delete this resume?")) return;
    try {
      const { data } = await api.delete(`/api/resumes/delete/${resumeId}`, { headers: authHeader(token) });
      setResumes((items) => items.filter((item) => item._id !== resumeId));
      toast.success(data.message);
    } catch (error) {
      toast.error(error.response?.data?.message || error.message);
    }
  }

  return (
    <div className="min-h-screen bg-slate-50">
      <div className="border-b border-slate-200 bg-white">
        <div className="mx-auto flex h-16 max-w-7xl items-center justify-between px-4">
          <Link to="/" className="flex items-center gap-3">
            <img src={logo} alt="Resume Builder" className="h-8 w-auto" />
          </Link>
          <div className="flex items-center gap-3">
            <span className="hidden text-sm text-slate-500 sm:block">Welcome, {user?.name || "Joe Doe"}</span>
            <button onClick={logout} className="rounded-full p-2 text-slate-500 transition hover:bg-slate-100 hover:text-slate-950" title="Logout">
              <LogOut className="size-5" />
            </button>
          </div>
        </div>
      </div>

      <main className="mx-auto max-w-7xl px-4 py-8">
        <p className="mb-6 text-2xl font-medium text-slate-700 sm:hidden">Welcome, {user?.name || "Joe Doe"}</p>
        <div className="mb-8 flex flex-col gap-4 sm:flex-row">
          <button onClick={() => setCreateOpen(true)} className="flex min-h-36 flex-1 items-center justify-center gap-3 rounded-lg border border-green-200 bg-green-50 p-6 text-green-700 transition hover:border-green-300 hover:bg-green-100">
            <Plus className="size-6" />
            Create a Resume
          </button>
          <button onClick={() => setUploadOpen(true)} className="flex min-h-36 flex-1 items-center justify-center gap-3 rounded-lg border border-blue-200 bg-blue-50 p-6 text-blue-700 transition hover:border-blue-300 hover:bg-blue-100">
            <Upload className="size-6" />
            Upload Existing
          </button>
        </div>

        {loading ? (
          <PageLoader />
        ) : resumes.length ? (
          <div className="grid gap-5 sm:grid-cols-2 lg:grid-cols-3">
            {resumes.map((resume) => (
              <article key={resume._id} className="rounded-lg border border-slate-200 bg-white p-5 shadow-sm">
                <div className="flex items-start justify-between gap-4">
                  <div>
                    <h2 className="line-clamp-2 text-lg font-semibold text-slate-900">{resume.title}</h2>
                    <p className="mt-1 text-sm text-slate-500">{formatDate(resume.updatedAt || resume.createdAt)}</p>
                  </div>
                  <span className={`rounded-full px-3 py-1 text-xs font-medium ${resume.public ? "bg-green-100 text-green-700" : "bg-slate-100 text-slate-600"}`}>
                    {resume.public ? "Public" : "Private"}
                  </span>
                </div>
                <div className="mt-6 flex items-center justify-between">
                  <button onClick={() => navigate(`/app/builder/${resume._id}`)} className="inline-flex items-center gap-2 rounded-full bg-slate-950 px-4 py-2 text-sm font-medium text-white">
                    <Edit3 className="size-4" />
                    Edit
                  </button>
                  <div className="flex items-center gap-1">
                    <button onClick={() => renameResume(resume)} className="rounded-full p-2 text-slate-500 hover:bg-slate-100 hover:text-slate-900" title="Edit Resume Title">
                      <FileText className="size-4" />
                    </button>
                    <button onClick={() => deleteResume(resume._id)} className="rounded-full p-2 text-red-500 hover:bg-red-50" title="Delete">
                      <Trash2 className="size-4" />
                    </button>
                  </div>
                </div>
              </article>
            ))}
          </div>
        ) : (
          <div className="rounded-lg border border-dashed border-slate-300 bg-white p-10 text-center text-slate-500">
            No resumes yet.
          </div>
        )}

        <JobTracker token={token} resumes={resumes} />
      </main>

      <Modal open={createOpen} onClose={() => setCreateOpen(false)} title="Create Resume">
        <form onSubmit={createResume} className="space-y-4">
          <Field label="Resume title" value={title} onChange={(event) => setTitle(event.target.value)} autoFocus />
          <button disabled={busy} className="inline-flex h-11 w-full items-center justify-center gap-2 rounded-lg bg-green-500 font-medium text-white hover:bg-green-600 disabled:opacity-70">
            {busy && <Loader2 className="size-4 animate-spin" />}
            Create Resume
          </button>
        </form>
      </Modal>

      <Modal open={uploadOpen} onClose={() => setUploadOpen(false)} title="Upload Resume">
        <form onSubmit={uploadResume} className="space-y-4">
          <Field label="Resume title" value={title} onChange={(event) => setTitle(event.target.value)} />
          <label className="flex cursor-pointer flex-col items-center justify-center rounded-lg border border-dashed border-slate-300 bg-slate-50 p-8 text-center text-sm text-slate-500 hover:bg-slate-100">
            <Upload className="mb-3 size-8 text-blue-500" />
            {file ? file.name : "Upload PDF, TXT, or MD file"}
            <input type="file" accept=".pdf,.txt,.md" className="hidden" onChange={(event) => setFile(event.target.files?.[0] || null)} />
          </label>
          <button disabled={busy} className="inline-flex h-11 w-full items-center justify-center gap-2 rounded-lg bg-blue-500 font-medium text-white hover:bg-blue-600 disabled:opacity-70">
            {busy && <Loader2 className="size-4 animate-spin" />}
            Upload Resume
          </button>
        </form>
      </Modal>
    </div>
  );
}

function JobTracker({ token, resumes }) {
  const [jobs, setJobs] = useState([]);
  const [open, setOpen] = useState(false);
  const [busy, setBusy] = useState(false);
  const [form, setForm] = useState({ title: "", company: "", location: "", status: "Saved", resumeId: "", description: "" });
  const statuses = ["Saved", "Applied", "Interview", "Offer", "Rejected"];

  async function loadJobs() {
    try {
      const { data } = await api.get("/api/jobs", { headers: authHeader(token) });
      setJobs(data.jobs || []);
    } catch (error) {
      toast.error(error.response?.data?.message || error.message);
    }
  }

  useEffect(() => {
    loadJobs();
  }, [token]);

  async function createJob(event) {
    event.preventDefault();
    setBusy(true);
    try {
      const { data } = await api.post("/api/jobs", form, { headers: authHeader(token) });
      setJobs((items) => [data.job, ...items]);
      setForm({ title: "", company: "", location: "", status: "Saved", resumeId: "", description: "" });
      setOpen(false);
      toast.success(data.message);
    } catch (error) {
      toast.error(error.response?.data?.message || error.message);
    } finally {
      setBusy(false);
    }
  }

  async function updateStatus(job, status) {
    try {
      const { data } = await api.put(`/api/jobs/${job._id}`, { ...job, status }, { headers: authHeader(token) });
      setJobs((items) => items.map((item) => (item._id === job._id ? data.job : item)));
    } catch (error) {
      toast.error(error.response?.data?.message || error.message);
    }
  }

  async function removeJob(jobId) {
    try {
      await api.delete(`/api/jobs/${jobId}`, { headers: authHeader(token) });
      setJobs((items) => items.filter((item) => item._id !== jobId));
      toast.success("Job deleted");
    } catch (error) {
      toast.error(error.response?.data?.message || error.message);
    }
  }

  return (
    <section className="mt-10 rounded-lg border border-slate-200 bg-white p-5 shadow-sm">
      <div className="flex flex-wrap items-center justify-between gap-3">
        <div>
          <h2 className="flex items-center gap-2 text-lg font-semibold text-slate-900">
            <Briefcase className="size-5 text-green-600" />
            Job Search Workflow
          </h2>
          <p className="mt-1 text-sm text-slate-500">Track jobs, attach a resume, and keep matching notes in one place.</p>
        </div>
        <button onClick={() => setOpen(true)} className="inline-flex items-center gap-2 rounded-lg bg-slate-950 px-4 py-2 text-sm font-medium text-white">
          <Plus className="size-4" />
          Add Job
        </button>
      </div>

      <div className="mt-5 grid gap-3 md:grid-cols-2 xl:grid-cols-3">
        {jobs.map((job) => (
          <article key={job._id} className="rounded-lg border border-slate-200 p-4">
            <div className="flex items-start justify-between gap-3">
              <div>
                <h3 className="font-semibold text-slate-900">{job.title}</h3>
                <p className="text-sm text-slate-500">{[job.company, job.location].filter(Boolean).join(" - ") || "No company added"}</p>
              </div>
              <button onClick={() => removeJob(job._id)} className="rounded-full p-2 text-red-500 hover:bg-red-50" title="Delete job">
                <Trash2 className="size-4" />
              </button>
            </div>
            <select value={job.status} onChange={(event) => updateStatus(job, event.target.value)} className="mt-4 h-10 w-full rounded-lg border border-slate-300 px-3 text-sm outline-none">
              {statuses.map((status) => (
                <option key={status}>{status}</option>
              ))}
            </select>
            {job.description && <p className="mt-3 line-clamp-3 text-sm leading-6 text-slate-600">{job.description}</p>}
          </article>
        ))}
      </div>
      {!jobs.length && <p className="mt-5 rounded-lg border border-dashed border-slate-300 p-6 text-center text-sm text-slate-500">No jobs tracked yet.</p>}

      <Modal open={open} onClose={() => setOpen(false)} title="Add Job">
        <form onSubmit={createJob} className="space-y-4">
          <Field label="Job title" value={form.title} onChange={(event) => setForm({ ...form, title: event.target.value })} required />
          <Field label="Company" value={form.company} onChange={(event) => setForm({ ...form, company: event.target.value })} />
          <Field label="Location" value={form.location} onChange={(event) => setForm({ ...form, location: event.target.value })} />
          <label className="block text-sm">
            <span className="mb-1.5 block font-medium text-gray-600">Attach resume</span>
            <select value={form.resumeId} onChange={(event) => setForm({ ...form, resumeId: event.target.value })} className="h-11 w-full rounded-lg border border-gray-300 px-3 outline-none">
              <option value="">No resume</option>
              {resumes.map((resume) => (
                <option key={resume._id} value={resume._id}>{resume.title}</option>
              ))}
            </select>
          </label>
          <label className="block text-sm">
            <span className="mb-1.5 block font-medium text-gray-600">Job description</span>
            <textarea value={form.description} onChange={(event) => setForm({ ...form, description: event.target.value })} rows={5} className="w-full rounded-lg border border-gray-300 p-3 outline-none" />
          </label>
          <button disabled={busy} className="inline-flex h-11 w-full items-center justify-center gap-2 rounded-lg bg-green-500 font-medium text-white hover:bg-green-600 disabled:opacity-70">
            {busy && <Loader2 className="size-4 animate-spin" />}
            Save Job
          </button>
        </form>
      </Modal>
    </section>
  );
}

async function extractResumeText(file) {
  if (file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf")) {
    const data = new Uint8Array(await file.arrayBuffer());
    const pdf = await pdfjsLib.getDocument({ data }).promise;
    const chunks = [];
    for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
      const page = await pdf.getPage(pageNumber);
      const content = await page.getTextContent();
      chunks.push(content.items.map((item) => item.str).join(" "));
    }
    return chunks.join("\n");
  }
  return file.text();
}

function Builder() {
  const { resumeId } = useParams();
  const { token } = useAuth();
  const navigate = useNavigate();
  const [resume, setResume] = useState(null);
  const [section, setSection] = useState(0);
  const [saving, setSaving] = useState(false);
  const [enhancing, setEnhancing] = useState("");
  const [downloading, setDownloading] = useState(false);
  const [toolsLoading, setToolsLoading] = useState("");
  const [jobDescription, setJobDescription] = useState("");
  const [atsResult, setAtsResult] = useState(null);
  const [matchResult, setMatchResult] = useState(null);
  const [coverLetter, setCoverLetter] = useState("");
  const [coverInfo, setCoverInfo] = useState({ jobTitle: "", company: "" });
  const [writingTips, setWritingTips] = useState([]);
  const previewRef = useRef(null);

  const sections = [
    { id: "personal", name: "Personal Info", icon: User },
    { id: "summary", name: "Summary", icon: FileText },
    { id: "experience", name: "Experience", icon: LayoutDashboard },
    { id: "education", name: "Education", icon: GraduationCap },
    { id: "projects", name: "Projects", icon: Sparkles },
    { id: "skills", name: "Skills", icon: Check },
    { id: "tools", name: "Career Tools", icon: ClipboardCheck },
    { id: "customize", name: "Customize", icon: Settings }
  ];

  const current = sections[section];

  useEffect(() => {
    api
      .get(`/api/resumes/get/${resumeId}`, { headers: authHeader(token) })
      .then(({ data }) => {
        const base = emptyResume();
        setResume({ ...base, ...data.resume, personal_info: { ...base.personal_info, ...(data.resume.personal_info || {}) }, section_order: data.resume.section_order || base.section_order });
        document.title = data.resume.title || "Resume Builder";
      })
      .catch((error) => {
        toast.error(error.response?.data?.message || error.message);
        navigate("/app");
      });

    return () => {
      document.title = "Resume Builder - Create your Resume in Minutes";
    };
  }, [resumeId, token, navigate]);

  function patch(update) {
    setResume((prev) => ({ ...prev, ...update }));
  }

  async function saveChanges() {
    setSaving(true);
    try {
      const form = new FormData();
      const image = resume.personal_info?.image;
      form.append("resumeId", resumeId);
      form.append("resumeData", JSON.stringify(cleanResumeForSave(resume)));
      if (image && typeof image !== "string") form.append("image", image);
      const { data } = await api.put("/api/resumes/update", form, {
        headers: { ...authHeader(token), "Content-Type": "multipart/form-data" }
      });
      const base = emptyResume();
      setResume({ ...base, ...data.resume, personal_info: { ...base.personal_info, ...(data.resume.personal_info || {}) }, section_order: data.resume.section_order || base.section_order });
      toast.success(data.message);
    } catch (error) {
      toast.error(error.response?.data?.message || error.message);
    } finally {
      setSaving(false);
    }
  }

  async function enhance(sectionName, content, apply) {
    setEnhancing(sectionName);
    try {
      const { data } = await api.post(
        "/api/ai/enhance",
        {
          section: sectionName,
          content,
          context: {
            role: resume.personal_info?.profession,
            name: resume.personal_info?.full_name,
            skills: resume.skills
          }
        },
        { headers: authHeader(token) }
      );
      apply(data.text);
      toast.success("AI Enhance complete");
    } catch (error) {
      toast.error(error.response?.data?.message || error.message);
    } finally {
      setEnhancing("");
    }
  }

  async function downloadPdf() {
    setDownloading(true);
    let exportRender = null;
    try {
      exportRender = await renderResumeForExport(resume);
      const element = exportRender.element;
      const canvas = await html2canvas(element, { scale: 2, useCORS: true, backgroundColor: "#ffffff" });
      const imgData = canvas.toDataURL("image/png");
      const pdf = new jsPDF("p", "pt", "letter");
      const pageWidth = pdf.internal.pageSize.getWidth();
      const pageHeight = pdf.internal.pageSize.getHeight();
      const imgHeight = (canvas.height * pageWidth) / canvas.width;
      let y = 0;
      pdf.addImage(imgData, "PNG", 0, y, pageWidth, imgHeight);
      let remaining = imgHeight - pageHeight;
      while (remaining > 0) {
        y -= pageHeight;
        pdf.addPage();
        pdf.addImage(imgData, "PNG", 0, y, pageWidth, imgHeight);
        remaining -= pageHeight;
      }
      pdf.save(`${safeFileName(resume.title)}.pdf`);
    } catch (error) {
      toast.error(error.message);
    } finally {
      exportRender?.root?.unmount();
      exportRender?.host?.remove();
      setDownloading(false);
    }
  }

  async function downloadDocx() {
    setToolsLoading("docx");
    try {
      await exportResumeDocx(resume);
      toast.success("DOCX exported");
    } catch (error) {
      toast.error(error.message);
    } finally {
      setToolsLoading("");
    }
  }

  async function runAtsScore() {
    setToolsLoading("ats");
    try {
      const { data } = await api.post("/api/ai/ats-score", { resume, jobDescription }, { headers: authHeader(token) });
      setAtsResult(data);
      toast.success("ATS score updated");
    } catch (error) {
      toast.error(error.response?.data?.message || error.message);
    } finally {
      setToolsLoading("");
    }
  }

  async function runJobMatch() {
    if (!jobDescription.trim()) return toast.error("Paste a job description first");
    setToolsLoading("match");
    try {
      const { data } = await api.post("/api/ai/job-match", { resume, jobDescription }, { headers: authHeader(token) });
      setMatchResult(data);
      toast.success("Job match complete");
    } catch (error) {
      toast.error(error.response?.data?.message || error.message);
    } finally {
      setToolsLoading("");
    }
  }

  async function buildCoverLetter() {
    setToolsLoading("cover");
    try {
      const { data } = await api.post(
        "/api/ai/cover-letter",
        { resume, jobDescription, jobTitle: coverInfo.jobTitle, company: coverInfo.company },
        { headers: authHeader(token) }
      );
      setCoverLetter(data.coverLetter);
      toast.success("Cover letter generated");
    } catch (error) {
      toast.error(error.response?.data?.message || error.message);
    } finally {
      setToolsLoading("");
    }
  }

  async function getWritingTips(targetSection = current.id) {
    setToolsLoading("tips");
    try {
      const { data } = await api.post("/api/ai/guided-help", { resume, section: targetSection }, { headers: authHeader(token) });
      setWritingTips(data.tips || []);
      toast.success("Writing tips ready");
    } catch (error) {
      toast.error(error.response?.data?.message || error.message);
    } finally {
      setToolsLoading("");
    }
  }

  async function copyShareLink() {
    if (!resume.public) {
      toast.error("Set the resume to Public first");
      return;
    }
    const link = `${window.location.origin}/resume/${resume._id}`;
    await navigator.clipboard.writeText(link);
    toast.success("Public link copied");
  }

  if (!resume) return <PageLoader />;

  return (
    <div className="min-h-screen bg-slate-100">
      <div className="mx-auto max-w-7xl px-4 py-6">
        <Link to="/app" className="inline-flex items-center gap-2 text-slate-500 transition hover:text-slate-800">
          <ArrowLeft className="size-4" />
          Back to Dashboard
        </Link>
      </div>

      <main className="mx-auto grid max-w-7xl gap-8 px-4 pb-10 lg:grid-cols-12">
        <section className="relative lg:col-span-5">
          <div className="rounded-lg border border-gray-200 bg-white p-6 pt-2 shadow-sm">
            <hr className="absolute left-4 right-4 top-0 border-2 border-gray-200" />
            <hr className="absolute left-4 top-0 h-1 border-none bg-gradient-to-r from-green-500 to-green-600 transition-all duration-700" style={{ width: `${(section * 100) / (sections.length - 1)}%` }} />

            <div className="mb-6 flex flex-wrap items-center justify-between gap-3 border-b border-gray-200 py-3">
              <div className="flex items-center gap-2">
                <TemplatePicker selected={resume.template} onChange={(template) => patch({ template })} />
                <ColorPicker selected={resume.accent_color} onChange={(accent_color) => patch({ accent_color })} />
              </div>
              <div className="flex items-center gap-1">
                {section !== 0 && (
                  <button onClick={() => setSection((value) => Math.max(value - 1, 0))} className="flex items-center gap-1 rounded-lg p-3 text-sm font-medium text-gray-600 transition hover:bg-gray-50">
                    <ChevronLeft className="size-4" />
                    Previous
                  </button>
                )}
                <button disabled={section === sections.length - 1} onClick={() => setSection((value) => Math.min(value + 1, sections.length - 1))} className="flex items-center gap-1 rounded-lg p-3 text-sm font-medium text-gray-600 transition hover:bg-gray-50 disabled:opacity-50">
                  Next
                  <ChevronRight className="size-4" />
                </button>
              </div>
            </div>

            <div className="mb-5 flex items-center gap-2 text-lg font-semibold text-gray-900">
              <current.icon className="size-5 text-green-600" />
              {current.name}
            </div>

            <div className="space-y-6">
              {current.id === "personal" && <PersonalSection data={resume.personal_info} onChange={(personal_info) => patch({ personal_info })} />}
              {current.id === "summary" && (
                <SummarySection
                  value={resume.professional_summary}
                  loading={enhancing === "summary"}
                  onChange={(professional_summary) => patch({ professional_summary })}
                  onEnhance={() => enhance("summary", resume.professional_summary, (text) => patch({ professional_summary: text }))}
                />
              )}
              {current.id === "experience" && (
                <ExperienceSection
                  data={resume.experience || []}
                  loading={enhancing}
                  onChange={(experience) => patch({ experience })}
                  onEnhance={(item, index) =>
                    enhance("experience", item.description, (text) => {
                      const next = [...(resume.experience || [])];
                      next[index] = { ...next[index], description: text };
                      patch({ experience: next });
                    })
                  }
                />
              )}
              {current.id === "education" && <EducationSection data={resume.education || []} onChange={(education) => patch({ education })} />}
              {current.id === "projects" && (
                <ProjectSection
                  data={resume.project || []}
                  loading={enhancing}
                  onChange={(project) => patch({ project })}
                  onEnhance={(item, index) =>
                    enhance("project", item.description, (text) => {
                      const next = [...(resume.project || [])];
                      next[index] = { ...next[index], description: text };
                      patch({ project: next });
                    })
                  }
                />
              )}
              {current.id === "skills" && (
                <SkillsSection
                  data={resume.skills || []}
                  loading={enhancing === "skills"}
                  onChange={(skills) => patch({ skills })}
                  onEnhance={() => enhance("skills", (resume.skills || []).join(", "), (text) => patch({ skills: text.split(",").map((item) => item.trim()).filter(Boolean) }))}
                />
              )}
              {current.id === "tools" && (
                <CareerToolsSection
                  jobDescription={jobDescription}
                  setJobDescription={setJobDescription}
                  atsResult={atsResult}
                  matchResult={matchResult}
                  coverLetter={coverLetter}
                  setCoverLetter={setCoverLetter}
                  coverInfo={coverInfo}
                  setCoverInfo={setCoverInfo}
                  writingTips={writingTips}
                  loading={toolsLoading}
                  onAts={runAtsScore}
                  onMatch={runJobMatch}
                  onCover={buildCoverLetter}
                  onTips={getWritingTips}
                />
              )}
              {current.id === "customize" && (
                <CustomizeSection
                  resume={resume}
                  onChange={patch}
                />
              )}
            </div>

            <button onClick={saveChanges} disabled={saving} className="mt-7 flex w-full items-center justify-center gap-2 rounded-lg bg-gradient-to-br from-green-100 to-green-200 px-6 py-3 text-sm font-medium text-green-700 ring-green-300 transition hover:ring disabled:opacity-70">
              {saving ? <Loader2 className="size-4 animate-spin" /> : <Save className="size-4" />}
              Save Changes
            </button>
          </div>
        </section>

        <section className="lg:col-span-7">
          <div className="sticky top-6">
            <div className="mb-4 flex flex-wrap items-center justify-between gap-3 rounded-lg border border-gray-200 bg-white p-4 shadow-sm">
              <div className="flex items-center gap-3">
                <label className="relative inline-flex cursor-pointer items-center gap-3 text-sm font-medium text-gray-700">
                  <input type="checkbox" className="peer sr-only" checked={resume.public} onChange={() => patch({ public: !resume.public })} />
                  <span className="h-5 w-9 rounded-full bg-slate-300 transition peer-checked:bg-green-600" />
                  <span className="absolute left-1 top-1 h-3 w-3 rounded-full bg-white transition-transform peer-checked:translate-x-4" />
                  {resume.public ? "Public" : "Private"}
                </label>
                <button onClick={copyShareLink} className="flex items-center gap-2 rounded-lg bg-blue-50 px-4 py-2 text-xs font-medium text-blue-700 ring-blue-200 transition hover:ring">
                  <Link2 className="size-4" />
                  Share
                </button>
              </div>
              <div className="flex items-center gap-2">
                <button onClick={downloadDocx} disabled={toolsLoading === "docx"} className="flex items-center gap-2 rounded-lg bg-slate-50 px-4 py-2 text-xs font-medium text-slate-700 ring-slate-200 transition hover:ring disabled:opacity-70">
                  {toolsLoading === "docx" ? <Loader2 className="size-4 animate-spin" /> : <FileDown className="size-4" />}
                  DOCX
                </button>
                <button onClick={downloadPdf} disabled={downloading} className="flex items-center gap-2 rounded-lg bg-green-50 px-5 py-2 text-xs font-medium text-green-700 ring-green-200 transition hover:ring disabled:opacity-70">
                  {downloading ? <Loader2 className="size-4 animate-spin" /> : <Download className="size-4" />}
                  PDF
                </button>
              </div>
            </div>
            <div ref={previewRef} className="overflow-auto rounded-lg bg-slate-200 p-4">
              <ResumePreview data={resume} />
            </div>
          </div>
        </section>
      </main>
    </div>
  );
}

function PersonalSection({ data, onChange }) {
  function update(field, value) {
    onChange({ ...data, [field]: value });
  }

  const imageUrl = getFileImageUrl(data?.image) || dummyProfile;

  return (
    <div className="space-y-5">
      <div className="flex items-center gap-4">
        <img src={imageUrl} alt="Profile" className="size-16 rounded-full border border-slate-200 object-cover" />
        <label className="flex cursor-pointer items-center gap-2 rounded-full border border-slate-300 px-4 py-2 text-sm text-slate-600 hover:bg-slate-50">
          <ImagePlus className="size-4" />
          upload user image
          <input type="file" accept="image/jpeg,image/png,image/webp" className="hidden" onChange={(event) => update("image", event.target.files?.[0] || data.image)} />
        </label>
      </div>
      {data?.image && typeof data.image === "object" && (
        <div className="flex flex-col gap-1 pl-4 text-sm">
          <p>Remove Background</p>
          <label className="relative inline-flex cursor-pointer items-center gap-3 text-gray-900">
            <input type="checkbox" className="peer sr-only" onChange={() => toast("Currently unavailable")} />
            <span className="h-5 w-9 rounded-full bg-slate-300 transition peer-checked:bg-green-600" />
            <span className="absolute left-1 top-1 h-3 w-3 rounded-full bg-white transition-transform peer-checked:translate-x-4" />
          </label>
        </div>
      )}
      <div className="grid gap-4 sm:grid-cols-2">
        <Field label="Full name" value={data.full_name || ""} onChange={(event) => update("full_name", event.target.value)} />
        <Field label="Profession" value={data.profession || ""} onChange={(event) => update("profession", event.target.value)} />
        <Field label="Email" value={data.email || ""} onChange={(event) => update("email", event.target.value)} />
        <Field label="Phone" value={data.phone || ""} onChange={(event) => update("phone", event.target.value)} />
        <Field label="Location" value={data.location || ""} onChange={(event) => update("location", event.target.value)} />
        <Field label="LinkedIn" value={data.linkedin || ""} onChange={(event) => update("linkedin", event.target.value)} />
        <Field label="Website" value={data.website || ""} onChange={(event) => update("website", event.target.value)} />
      </div>
    </div>
  );
}

function SummarySection({ value, onChange, onEnhance, loading }) {
  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between gap-4">
        <div>
          <h3 className="text-lg font-semibold text-gray-900">Professional Summary</h3>
          <p className="text-sm text-gray-500">Add summary for your resume here</p>
        </div>
        <AIButton loading={loading} onClick={onEnhance} />
      </div>
      <textarea
        value={value || ""}
        onChange={(event) => onChange(event.target.value)}
        rows={7}
        className="w-full rounded-lg border border-gray-300 p-3 text-sm outline-none transition focus:border-blue-500 focus:ring focus:ring-blue-100"
        placeholder="Write a compelling professional summary that highlights your key strengths and career objectives..."
      />
      <p className="mx-auto max-w-[80%] text-center text-xs text-gray-500">Tip: Keep it concise and focus on relevant achievements and skills.</p>
    </div>
  );
}

function ExperienceSection({ data, onChange, onEnhance, loading }) {
  function add() {
    onChange([...data, { company: "", position: "", start_date: "", end_date: "", description: "", is_current: false }]);
  }

  function update(index, field, value) {
    const next = [...data];
    next[index] = { ...next[index], [field]: value };
    onChange(next);
  }

  function remove(index) {
    onChange(data.filter((_, itemIndex) => itemIndex !== index));
  }

  return (
    <ListSection empty="No work experience added yet." addLabel="Add Experience" onAdd={add}>
      {data.map((item, index) => (
        <div key={index} className="rounded-lg border border-gray-200 p-4">
          <div className="mb-4 flex items-center justify-between">
            <h3 className="font-medium text-gray-800">Professional Experience</h3>
            <button onClick={() => remove(index)} className="rounded-full p-2 text-red-500 hover:bg-red-50">
              <Trash2 className="size-4" />
            </button>
          </div>
          <div className="grid gap-4 sm:grid-cols-2">
            <Field label="Company" value={item.company} onChange={(event) => update(index, "company", event.target.value)} />
            <Field label="Job Title" value={item.position} onChange={(event) => update(index, "position", event.target.value)} />
            <Field label="Start Date" type="month" value={item.start_date} onChange={(event) => update(index, "start_date", event.target.value)} />
            <Field label="End Date" type="month" disabled={item.is_current} value={item.end_date} onChange={(event) => update(index, "end_date", event.target.value)} />
          </div>
          <label className="mt-4 flex items-center gap-2 text-sm text-gray-600">
            <input type="checkbox" checked={item.is_current} onChange={(event) => update(index, "is_current", event.target.checked)} />
            Currently working here
          </label>
          <div className="mt-4">
            <div className="mb-2 flex items-center justify-between">
              <label className="text-sm font-medium text-gray-600">Job Description</label>
              <AIButton loading={loading === "experience"} onClick={() => onEnhance(item, index)} />
            </div>
            <textarea className="w-full rounded-lg border border-gray-300 p-3 text-sm outline-none focus:border-blue-500 focus:ring focus:ring-blue-100" rows={5} value={item.description} onChange={(event) => update(index, "description", event.target.value)} />
          </div>
        </div>
      ))}
    </ListSection>
  );
}

function EducationSection({ data, onChange }) {
  function add() {
    onChange([...data, { institution: "", degree: "", field: "", graduation_date: "", gpa: "" }]);
  }

  function update(index, field, value) {
    const next = [...data];
    next[index] = { ...next[index], [field]: value };
    onChange(next);
  }

  return (
    <ListSection empty="No education added yet." addLabel="Add Education" onAdd={add}>
      {data.map((item, index) => (
        <div key={index} className="rounded-lg border border-gray-200 p-4">
          <div className="mb-4 flex items-center justify-between">
            <h3 className="font-medium text-gray-800">Education</h3>
            <button onClick={() => onChange(data.filter((_, itemIndex) => itemIndex !== index))} className="rounded-full p-2 text-red-500 hover:bg-red-50">
              <Trash2 className="size-4" />
            </button>
          </div>
          <div className="grid gap-4 sm:grid-cols-2">
            <Field label="Institution" value={item.institution} onChange={(event) => update(index, "institution", event.target.value)} />
            <Field label="Degree" value={item.degree} onChange={(event) => update(index, "degree", event.target.value)} />
            <Field label="Field" value={item.field} onChange={(event) => update(index, "field", event.target.value)} />
            <Field label="Graduation Date" type="month" value={item.graduation_date} onChange={(event) => update(index, "graduation_date", event.target.value)} />
            <Field label="GPA" value={item.gpa} onChange={(event) => update(index, "gpa", event.target.value)} />
          </div>
        </div>
      ))}
    </ListSection>
  );
}

function ProjectSection({ data, onChange, onEnhance, loading }) {
  function add() {
    onChange([...data, { name: "", type: "", description: "" }]);
  }

  function update(index, field, value) {
    const next = [...data];
    next[index] = { ...next[index], [field]: value };
    onChange(next);
  }

  return (
    <ListSection empty="Add your projects" addLabel="Add Project" onAdd={add}>
      {data.map((item, index) => (
        <div key={index} className="rounded-lg border border-gray-200 p-4">
          <div className="mb-4 flex items-center justify-between">
            <h3 className="font-medium text-gray-800">Projects</h3>
            <button onClick={() => onChange(data.filter((_, itemIndex) => itemIndex !== index))} className="rounded-full p-2 text-red-500 hover:bg-red-50">
              <Trash2 className="size-4" />
            </button>
          </div>
          <div className="grid gap-4 sm:grid-cols-2">
            <Field label="Project Name" value={item.name} onChange={(event) => update(index, "name", event.target.value)} />
            <Field label="Type" value={item.type} onChange={(event) => update(index, "type", event.target.value)} />
          </div>
          <div className="mt-4">
            <div className="mb-2 flex items-center justify-between">
              <label className="text-sm font-medium text-gray-600">Description</label>
              <AIButton loading={loading === "project"} onClick={() => onEnhance(item, index)} />
            </div>
            <textarea className="w-full rounded-lg border border-gray-300 p-3 text-sm outline-none focus:border-blue-500 focus:ring focus:ring-blue-100" rows={5} value={item.description} onChange={(event) => update(index, "description", event.target.value)} />
          </div>
        </div>
      ))}
    </ListSection>
  );
}

function SkillsSection({ data, onChange, onEnhance, loading }) {
  const [skill, setSkill] = useState("");

  function add() {
    const value = skill.trim();
    if (!value) return;
    onChange([...new Set([...data, value])]);
    setSkill("");
  }

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between">
        <div>
          <h3 className="text-lg font-semibold text-gray-900">Skills</h3>
          <p className="text-sm text-gray-500">Add your technical and soft skills above.</p>
        </div>
        <AIButton loading={loading} onClick={onEnhance} />
      </div>
      <div className="flex gap-2">
        <Field label="Skill" value={skill} onChange={(event) => setSkill(event.target.value)} onKeyDown={(event) => event.key === "Enter" && (event.preventDefault(), add())} />
        <button onClick={add} className="mt-6 rounded-lg bg-slate-950 px-4 text-white">
          <Plus className="size-5" />
        </button>
      </div>
      <div className="flex flex-wrap gap-2">
        {data.map((item) => (
          <span key={item} className="inline-flex items-center gap-2 rounded-full bg-green-50 px-3 py-1 text-sm text-green-700">
            {item}
            <button onClick={() => onChange(data.filter((skillItem) => skillItem !== item))}>
              <X className="size-3" />
            </button>
          </span>
        ))}
      </div>
      {!data.length && <p className="text-sm text-gray-500">No skills added yet.</p>}
    </div>
  );
}

function ScorePanel({ result }) {
  if (!result) return <p className="rounded-lg border border-dashed border-slate-300 p-4 text-center text-sm text-slate-500">Run a check to see your score, missing sections, and improvement ideas.</p>;
  return (
    <div className="rounded-lg border border-slate-200 p-4">
      <div className="flex items-center justify-between gap-4">
        <div>
          <p className="text-sm font-medium text-slate-500">Resume Score</p>
          <p className="text-4xl font-semibold text-slate-950">{result.score}<span className="text-lg text-slate-400">/100</span></p>
        </div>
        <ShieldCheck className="size-10 text-green-600" />
      </div>
      <div className="mt-4 space-y-2">
        {(result.checks || []).map((check) => (
          <div key={check.label} className="flex items-start gap-2 text-sm">
            <span className={`mt-1 size-2 rounded-full ${check.ok ? "bg-green-500" : "bg-red-500"}`} />
            <span className="text-slate-700">{check.label}: <span className="text-slate-500">{check.ok ? "Looks good" : check.detail}</span></span>
          </div>
        ))}
      </div>
      {!!(result.suggestions || []).length && (
        <div className="mt-4 rounded-lg bg-amber-50 p-3 text-sm text-amber-800">
          {(result.suggestions || []).slice(0, 4).map((item) => <p key={item}>{item}</p>)}
        </div>
      )}
    </div>
  );
}

function KeywordChips({ title, items, tone = "green" }) {
  if (!items?.length) return null;
  const color = tone === "red" ? "bg-red-50 text-red-700" : "bg-green-50 text-green-700";
  return (
    <div>
      <p className="mb-2 text-sm font-medium text-slate-700">{title}</p>
      <div className="flex flex-wrap gap-2">
        {items.slice(0, 18).map((item) => (
          <span key={item} className={`rounded-full px-3 py-1 text-xs ${color}`}>{item}</span>
        ))}
      </div>
    </div>
  );
}

function CareerToolsSection({
  jobDescription,
  setJobDescription,
  atsResult,
  matchResult,
  coverLetter,
  setCoverLetter,
  coverInfo,
  setCoverInfo,
  writingTips,
  loading,
  onAts,
  onMatch,
  onCover,
  onTips
}) {
  return (
    <div className="space-y-6">
      <div className="rounded-lg border border-slate-200 p-4">
        <div className="mb-3 flex items-center justify-between gap-3">
          <div>
            <h3 className="text-lg font-semibold text-slate-900">ATS Checker</h3>
            <p className="text-sm text-slate-500">Score the resume and find missing sections.</p>
          </div>
          <button onClick={onAts} disabled={loading === "ats"} className="inline-flex items-center gap-2 rounded-lg bg-green-50 px-4 py-2 text-sm font-medium text-green-700 ring-green-200 hover:ring disabled:opacity-70">
            {loading === "ats" ? <Loader2 className="size-4 animate-spin" /> : <ClipboardCheck className="size-4" />}
            Check
          </button>
        </div>
        <ScorePanel result={atsResult} />
      </div>

      <div className="rounded-lg border border-slate-200 p-4">
        <h3 className="text-lg font-semibold text-slate-900">Job Description Matcher</h3>
        <textarea
          value={jobDescription}
          onChange={(event) => setJobDescription(event.target.value)}
          rows={7}
          className="mt-3 w-full rounded-lg border border-gray-300 p-3 text-sm outline-none focus:border-blue-500 focus:ring focus:ring-blue-100"
          placeholder="Paste the target job description here..."
        />
        <button onClick={onMatch} disabled={loading === "match"} className="mt-3 inline-flex items-center gap-2 rounded-lg bg-blue-50 px-4 py-2 text-sm font-medium text-blue-700 ring-blue-200 hover:ring disabled:opacity-70">
          {loading === "match" ? <Loader2 className="size-4 animate-spin" /> : <Briefcase className="size-4" />}
          Match Keywords
        </button>
        {matchResult && (
          <div className="mt-4 space-y-4">
            <p className="text-sm font-medium text-slate-700">Match score: <span className="text-lg text-slate-950">{matchResult.score}/100</span></p>
            <KeywordChips title="Matched keywords" items={matchResult.matchedKeywords} />
            <KeywordChips title="Missing keywords" items={matchResult.missingKeywords} tone="red" />
          </div>
        )}
      </div>

      <div className="rounded-lg border border-slate-200 p-4">
        <h3 className="text-lg font-semibold text-slate-900">Cover Letter Builder</h3>
        <div className="mt-3 grid gap-3 sm:grid-cols-2">
          <Field label="Job title" value={coverInfo.jobTitle} onChange={(event) => setCoverInfo({ ...coverInfo, jobTitle: event.target.value })} />
          <Field label="Company" value={coverInfo.company} onChange={(event) => setCoverInfo({ ...coverInfo, company: event.target.value })} />
        </div>
        <button onClick={onCover} disabled={loading === "cover"} className="mt-3 inline-flex items-center gap-2 rounded-lg bg-purple-50 px-4 py-2 text-sm font-medium text-purple-700 ring-purple-200 hover:ring disabled:opacity-70">
          {loading === "cover" ? <Loader2 className="size-4 animate-spin" /> : <Send className="size-4" />}
          Generate Cover Letter
        </button>
        <textarea
          value={coverLetter}
          onChange={(event) => setCoverLetter(event.target.value)}
          rows={9}
          className="mt-3 w-full rounded-lg border border-gray-300 p-3 text-sm outline-none focus:border-blue-500 focus:ring focus:ring-blue-100"
          placeholder="Your generated cover letter will appear here..."
        />
      </div>

      <div className="rounded-lg border border-slate-200 p-4">
        <div className="flex items-center justify-between gap-3">
          <div>
            <h3 className="text-lg font-semibold text-slate-900">Guided Writing Help</h3>
            <p className="text-sm text-slate-500">Get practical suggestions based on the resume content.</p>
          </div>
          <button onClick={() => onTips("summary")} disabled={loading === "tips"} className="inline-flex items-center gap-2 rounded-lg bg-amber-50 px-4 py-2 text-sm font-medium text-amber-700 ring-amber-200 hover:ring disabled:opacity-70">
            {loading === "tips" ? <Loader2 className="size-4 animate-spin" /> : <Wand2 className="size-4" />}
            Tips
          </button>
        </div>
        {!!writingTips.length && (
          <ul className="mt-4 space-y-2 text-sm text-slate-700">
            {writingTips.map((tip) => <li key={tip} className="rounded-lg bg-slate-50 p-3">{tip}</li>)}
          </ul>
        )}
      </div>
    </div>
  );
}

function CustomizeSection({ resume, onChange }) {
  const [dragging, setDragging] = useState("");
  const order = resume.section_order?.length ? resume.section_order : DEFAULT_SECTION_ORDER;
  const labels = {
    summary: "Professional Summary",
    experience: "Experience",
    projects: "Projects",
    education: "Education",
    skills: "Skills"
  };

  function reorder(nextOrder) {
    onChange({ section_order: nextOrder });
  }

  function move(id, direction) {
    const index = order.indexOf(id);
    const nextIndex = index + direction;
    if (index < 0 || nextIndex < 0 || nextIndex >= order.length) return;
    const next = [...order];
    [next[index], next[nextIndex]] = [next[nextIndex], next[index]];
    reorder(next);
  }

  function dropOn(target) {
    if (!dragging || dragging === target) return;
    const next = order.filter((item) => item !== dragging);
    next.splice(next.indexOf(target), 0, dragging);
    reorder(next);
    setDragging("");
  }

  return (
    <div className="space-y-6">
      <div className="rounded-lg border border-slate-200 p-4">
        <h3 className="text-lg font-semibold text-slate-900">Templates</h3>
        <div className="mt-3 grid gap-3">
          {templates.map((template) => (
            <button
              key={template.id}
              onClick={() => onChange({ template: template.id })}
              className={`rounded-lg border p-3 text-left transition ${resume.template === template.id ? "border-green-400 bg-green-50" : "border-slate-200 hover:bg-slate-50"}`}
            >
              <div className="flex items-center justify-between">
                <span className="font-medium text-slate-900">{template.name}</span>
                {resume.template === template.id && <Check className="size-4 text-green-600" />}
              </div>
              <p className="mt-1 text-xs text-slate-500">{template.preview}</p>
            </button>
          ))}
        </div>
      </div>

      <div className="rounded-lg border border-slate-200 p-4">
        <h3 className="text-lg font-semibold text-slate-900">Font and Density</h3>
        <div className="mt-3 grid gap-3 sm:grid-cols-2">
          <label className="block text-sm">
            <span className="mb-1.5 block font-medium text-gray-600">Font</span>
            <select value={resume.font_family || "Inter"} onChange={(event) => onChange({ font_family: event.target.value })} className="h-11 w-full rounded-lg border border-gray-300 px-3 outline-none">
              {fontOptions.map((font) => <option key={font}>{font}</option>)}
            </select>
          </label>
          <label className="block text-sm">
            <span className="mb-1.5 block font-medium text-gray-600">Spacing</span>
            <select value={resume.density || "comfortable"} onChange={(event) => onChange({ density: event.target.value })} className="h-11 w-full rounded-lg border border-gray-300 px-3 outline-none">
              {densityOptions.map((item) => <option key={item.id} value={item.id}>{item.label}</option>)}
            </select>
          </label>
        </div>
      </div>

      <div className="rounded-lg border border-slate-200 p-4">
        <h3 className="text-lg font-semibold text-slate-900">Section Order</h3>
        <div className="mt-3 space-y-2">
          {order.map((item) => (
            <div
              key={item}
              draggable
              onDragStart={() => setDragging(item)}
              onDragOver={(event) => event.preventDefault()}
              onDrop={() => dropOn(item)}
              className="flex items-center justify-between gap-3 rounded-lg border border-slate-200 bg-white p-3"
            >
              <div className="flex items-center gap-3 text-sm font-medium text-slate-800">
                <GripVertical className="size-4 text-slate-400" />
                {labels[item]}
              </div>
              <div className="flex items-center gap-1">
                <button onClick={() => move(item, -1)} className="rounded-full p-1.5 text-slate-500 hover:bg-slate-100" title="Move up">
                  <ArrowUp className="size-4" />
                </button>
                <button onClick={() => move(item, 1)} className="rounded-full p-1.5 text-slate-500 hover:bg-slate-100" title="Move down">
                  <ArrowDown className="size-4" />
                </button>
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

function ListSection({ children, empty, addLabel, onAdd }) {
  const hasChildren = Array.isArray(children) ? children.length > 0 : Boolean(children);
  return (
    <div className="space-y-4">
      {hasChildren ? children : <p className="rounded-lg border border-dashed border-gray-300 p-5 text-center text-sm text-gray-500">{empty}</p>}
      <button onClick={onAdd} className="flex items-center gap-2 rounded-lg bg-slate-950 px-4 py-2 text-sm font-medium text-white">
        <Plus className="size-4" />
        {addLabel}
      </button>
    </div>
  );
}

function TemplatePicker({ selected, onChange }) {
  const [open, setOpen] = useState(false);
  return (
    <div className="relative">
      <button onClick={() => setOpen((value) => !value)} className="flex items-center gap-2 rounded-lg bg-blue-50 px-3 py-2 text-sm font-medium text-blue-700 ring-blue-200 transition hover:ring">
        <LayoutDashboard className="size-4" />
        Template
      </button>
      {open && (
        <div className="absolute left-0 top-full z-20 mt-2 w-72 rounded-lg border border-gray-200 bg-white p-3 shadow-xl">
          {templates.map((template) => (
            <button
              key={template.id}
              onClick={() => {
                onChange(template.id);
                setOpen(false);
              }}
              className={`mb-2 w-full rounded-lg border p-3 text-left transition ${selected === template.id ? "border-blue-400 bg-blue-50" : "border-gray-200 hover:bg-gray-50"}`}
            >
              <div className="flex items-center justify-between">
                <h4 className="font-medium text-gray-900">{template.name}</h4>
                {selected === template.id && <Check className="size-4 text-blue-500" />}
              </div>
              <p className="mt-1 text-xs text-gray-500">{template.preview}</p>
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

function ColorPicker({ selected, onChange }) {
  const [open, setOpen] = useState(false);
  return (
    <div className="relative">
      <button onClick={() => setOpen((value) => !value)} className="flex items-center gap-2 rounded-lg bg-purple-50 px-3 py-2 text-sm font-medium text-purple-700 ring-purple-200 transition hover:ring">
        <Palette className="size-4" />
        Accent
      </button>
      {open && (
        <div className="absolute left-0 top-full z-20 mt-2 grid w-44 grid-cols-4 gap-2 rounded-lg border border-gray-200 bg-white p-3 shadow-xl">
          {accents.map((color) => (
            <button key={color} title={color} onClick={() => onChange(color)} className="grid size-8 place-items-center rounded-full border border-white shadow" style={{ backgroundColor: color }}>
              {selected === color && <Check className="size-4 text-white" />}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

function orderedSections(data) {
  const stored = Array.isArray(data.section_order) ? data.section_order : [];
  return [...new Set([...stored, ...DEFAULT_SECTION_ORDER])].filter((item) => DEFAULT_SECTION_ORDER.includes(item));
}

function ResumeSectionBlock({ id, data, accentColor, compact = false }) {
  const titleMap = {
    summary: "Professional Summary",
    experience: "Professional Experience",
    projects: "Projects",
    education: "Education",
    skills: "Skills"
  };
  const heading = (
    <h2 className={`${compact ? "mb-2 text-sm" : "mb-3 text-base"} font-bold uppercase tracking-normal`} style={{ color: accentColor }}>
      {titleMap[id]}
    </h2>
  );

  if (id === "summary" && data.professional_summary) {
    return <section className={compact ? "mb-4" : "mb-6"}>{heading}<p className="whitespace-pre-line text-slate-700">{data.professional_summary}</p></section>;
  }

  if (id === "experience" && data.experience?.length) {
    return (
      <section className={compact ? "mb-4" : "mb-6"}>
        {heading}
        <div className={compact ? "space-y-3" : "space-y-4"}>
          {data.experience.map((item, index) => (
            <div key={index}>
              <div className="flex items-start justify-between gap-4">
                <div>
                  <h3 className="font-semibold text-slate-950">{item.position || "Role"}</h3>
                  {item.company && <p className="font-medium text-slate-700">{item.company}</p>}
                </div>
                <p className="shrink-0 text-right text-xs text-slate-500">{formatMonth(item.start_date)} - {item.is_current ? "Present" : formatMonth(item.end_date)}</p>
              </div>
              {item.description && <p className="mt-1 whitespace-pre-line text-slate-700">{item.description}</p>}
            </div>
          ))}
        </div>
      </section>
    );
  }

  if (id === "projects" && data.project?.length) {
    return (
      <section className={compact ? "mb-4" : "mb-6"}>
        {heading}
        <div className={compact ? "space-y-2" : "space-y-3"}>
          {data.project.map((item, index) => (
            <div key={index}>
              <h3 className="font-semibold text-slate-950">{item.name || "Project"}</h3>
              {item.description && <p className="whitespace-pre-line text-slate-700">{item.description}</p>}
            </div>
          ))}
        </div>
      </section>
    );
  }

  if (id === "education" && data.education?.length) {
    return (
      <section className={compact ? "mb-4" : "mb-6"}>
        {heading}
        <div className={compact ? "space-y-2" : "space-y-3"}>
          {data.education.map((item, index) => (
            <div key={index} className="flex items-start justify-between gap-4">
              <div>
                <h3 className="font-semibold text-slate-950">{[item.degree, item.field].filter(Boolean).join(", ") || "Education"}</h3>
                {item.institution && <p className="text-slate-700">{item.institution}</p>}
              </div>
              <p className="shrink-0 text-xs text-slate-500">{formatMonth(item.graduation_date)}</p>
            </div>
          ))}
        </div>
      </section>
    );
  }

  if (id === "skills" && data.skills?.length) {
    return (
      <section className={compact ? "mb-4" : "mb-6"}>
        {heading}
        <div className="flex flex-wrap gap-2">
          {data.skills.map((skill) => (
            <span key={skill} className="rounded-full px-3 py-1 text-xs font-medium" style={{ backgroundColor: alphaColor(accentColor, 0.1), color: accentColor }}>{skill}</span>
          ))}
        </div>
      </section>
    );
  }

  return null;
}

function OrderedResumeSections({ data, accentColor, compact }) {
  return orderedSections(data).map((id) => <ResumeSectionBlock key={id} id={id} data={data} accentColor={accentColor} compact={compact} />);
}

function ExecutiveTemplate({ data, accentColor }) {
  const info = data.personal_info || {};
  return (
    <div className="mx-auto max-w-4xl bg-white p-8 text-slate-800">
      <header className="border-b-4 pb-5" style={{ borderColor: accentColor }}>
        <p className="text-sm font-semibold uppercase tracking-normal" style={{ color: accentColor }}>{info.profession || "Professional"}</p>
        <h1 className="mt-1 text-4xl font-bold text-slate-950">{info.full_name || "Your Name"}</h1>
        <div className="mt-4 flex flex-wrap gap-x-4 gap-y-2 text-sm text-slate-600">
          {[info.email, info.phone, info.location, info.linkedin, info.website].filter(Boolean).map((item) => <span key={item} className="break-all">{item}</span>)}
        </div>
      </header>
      <main className="pt-6">
        <OrderedResumeSections data={data} accentColor={accentColor} />
      </main>
    </div>
  );
}

function CompactTemplate({ data, accentColor }) {
  const info = data.personal_info || {};
  return (
    <div className="mx-auto max-w-4xl bg-white p-7 text-[13px] leading-relaxed text-slate-800">
      <header className="mb-5 flex flex-col gap-2 border-b pb-4 sm:flex-row sm:items-end sm:justify-between" style={{ borderColor: accentColor }}>
        <div>
          <h1 className="text-3xl font-bold text-slate-950">{info.full_name || "Your Name"}</h1>
          {info.profession && <p className="font-medium" style={{ color: accentColor }}>{info.profession}</p>}
        </div>
        <p className="max-w-md text-right text-xs text-slate-600">{[info.email, info.phone, info.location, info.linkedin, info.website].filter(Boolean).join(" | ")}</p>
      </header>
      <OrderedResumeSections data={data} accentColor={accentColor} compact />
    </div>
  );
}

function ResumePreview({ data, scale = "normal", classes = "" }) {
  const template = data.template || "classic";
  const accentColor = data.accent_color || "#14B8A6";
  const base = emptyResume();
  const previewData = {
    ...base,
    ...data,
    personal_info: { ...base.personal_info, ...(data.personal_info || {}) },
    section_order: data.section_order || base.section_order
  };

  const content =
    template === "executive" ? (
      <ExecutiveTemplate data={previewData} accentColor={accentColor} />
    ) : template === "compact" ? (
      <CompactTemplate data={previewData} accentColor={accentColor} />
    ) : template === "modern" ? (
      <ModernTemplate data={previewData} accentColor={accentColor} />
    ) : template === "minimal" ? (
      <MinimalTemplate data={previewData} accentColor={accentColor} />
    ) : template === "minimal-image" ? (
      <MinimalImageTemplate data={previewData} accentColor={accentColor} />
    ) : (
      <ClassicTemplate data={previewData} accentColor={accentColor} />
    );

  return (
    <div className={`w-full bg-gray-100 ${scale === "small" ? "preview-small" : ""}`} style={{ fontFamily: previewData.font_family || "Inter" }}>
      <div data-resume-preview="true" className={`border border-gray-200 bg-white print:border-none print:shadow-none ${previewData.density === "compact" ? "text-[13px]" : ""} ${classes}`}>
        {content}
      </div>
    </div>
  );
}

function PublicResume() {
  const { resumeId } = useParams();
  const [loading, setLoading] = useState(true);
  const [resume, setResume] = useState(null);

  useEffect(() => {
    api
      .get(`/api/resumes/public/${resumeId}`)
      .then(({ data }) => setResume(data.resume))
      .catch(() => setResume(null))
      .finally(() => setLoading(false));
  }, [resumeId]);

  if (loading) return <PageLoader />;

  return resume ? (
    <div className="min-h-screen bg-slate-100 px-4 py-10">
      <div className="mx-auto max-w-4xl">
        <ResumePreview data={resume} classes="py-4" />
      </div>
    </div>
  ) : (
    <div className="flex min-h-screen flex-col items-center justify-center px-4 text-center">
      <p className="text-5xl font-medium text-slate-400">Resume not found</p>
      <Link to="/" className="mt-6 inline-flex h-10 items-center gap-2 rounded-full bg-green-500 px-6 text-white">
        <ArrowLeft className="size-4" />
        go to home page
      </Link>
    </div>
  );
}

function Field({ label, icon: Icon, className = "", ...props }) {
  return (
    <label className={`block w-full text-sm ${className}`}>
      <span className="mb-1.5 flex items-center gap-2 font-medium text-gray-600">
        {Icon && <Icon className="size-4" />}
        {label}
      </span>
      <input className="h-11 w-full rounded-lg border border-gray-300 px-3 text-sm outline-none transition focus:border-blue-500 focus:ring focus:ring-blue-100 disabled:bg-gray-100" {...props} />
    </label>
  );
}

function AIButton({ loading, onClick }) {
  return (
    <button onClick={onClick} disabled={loading} className="flex items-center gap-2 rounded px-3 py-1 text-sm text-purple-700 transition hover:bg-purple-100 disabled:opacity-60">
      {loading ? <Loader2 className="size-4 animate-spin" /> : <Wand2 className="size-4" />}
      {loading ? "Enhancing..." : "AI Enhance"}
    </button>
  );
}

function Modal({ open, onClose, title, children }) {
  if (!open) return null;
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/50 p-4">
      <div className="w-full max-w-md rounded-lg bg-white p-6 shadow-xl">
        <div className="mb-5 flex items-center justify-between">
          <h2 className="text-xl font-semibold text-slate-900">{title}</h2>
          <button onClick={onClose} className="rounded-full p-2 text-slate-500 hover:bg-slate-100">
            <X className="size-5" />
          </button>
        </div>
        {children}
      </div>
    </div>
  );
}

function PageLoader() {
  return (
    <div className="grid min-h-[50vh] place-items-center">
      <Loader2 className="size-8 animate-spin text-green-500" />
    </div>
  );
}

export default App;
