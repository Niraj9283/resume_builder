# AI Resume Builder

This is a full stack AI resume builder project based on the GreatStack resume builder tutorial and live preview.

## Features

- User sign up and login with JWT auth
- Dashboard to create, rename, delete, and manage multiple resumes
- Upload an existing resume text file and turn it into editable resume data
- Live resume builder with personal info, summary, experience, education, projects, skills, templates, accents, public/private toggle, and image upload
- Multiple professional templates from the provided project assets
- ATS score, missing-section checks, job description keyword matching, and guided writing tips
- Cover letter builder with Gemini AI when configured and a local fallback when not
- DOCX export plus fixed-width PDF export that renders the selected template before download
- Extra Executive and Compact templates with font, density, accent, and section-order controls
- Job search workflow tracker for saved jobs, application status, descriptions, and attached resumes
- Password reset and email verification endpoints for production account flows
- Public share links for published resumes
- Gemini AI enhancement endpoint with a local fallback when no API key is set
- ImageKit-ready upload flow with local upload fallback
- PostgreSQL backend using JSONB resume sections, with a JSON fallback for quick local demos

## Setup

1. Install dependencies:

```bash
npm run install-all
```

2. Copy the environment examples:

```bash
copy server\.env.example server\.env
copy client\.env.example client\.env
```

3. Start both apps:

```bash
npm run dev
```

Client: `http://localhost:5173`

Server: `http://localhost:5000`

Set `DATABASE_URL` in `server\.env` to use PostgreSQL, for example:

```bash
DATABASE_URL=postgresql://postgres:password@localhost:5432/resume_builder
```

The server creates the tables automatically on startup. The same schema is available at `server/schema.sql` if you want to run it manually. When `DATABASE_URL` is empty, the backend stores demo data in `server/data/local-db.json`.

## API keys

Put your Gemini API key in `server\.env`:

```bash
GEMINI_API_KEY=your_api_key_here
GEMINI_MODEL=gemini-2.5-flash
```

Restart the server after changing `.env`. ATS scoring and keyword matching work locally without a key; Gemini improves AI enhancement, guided writing, resume parsing, and cover letter generation.

## Deploy Online

The easiest deployment is one Node web service that builds the React app and serves it from Express.

Recommended Render settings:

- Root directory: repository root
- Build command: `npm run install-all && npm run build`
- Start command: `npm start`
- Environment variables:
  - `NODE_ENV=production`
  - `DATABASE_URL=your_postgresql_connection_string`
  - `PGSSL=true` when your PostgreSQL provider requires SSL
  - `JWT_SECRET=a_long_random_secret`
  - `GEMINI_API_KEY=your_gemini_key`
  - `GEMINI_MODEL=gemini-2.5-flash`
  - `IMAGEKIT_PUBLIC_KEY=your_imagekit_public_key`
  - `IMAGEKIT_PRIVATE_KEY=your_imagekit_private_key`
  - `IMAGEKIT_URL_ENDPOINT=https://ik.imagekit.io/your_imagekit_id`

For production, use PostgreSQL and ImageKit. The local JSON database and local `uploads` folder are only for local testing.
