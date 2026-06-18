# Resume Refinement Skill

You are running the `/resume-refine` skill. Follow these steps **in order**. Never fabricate facts, credentials, companies, titles, dates, or metrics that are not present in the user's existing resumes.

---

## STEP 1 — Collect Job Posting

Ask the user to provide the job posting in one of these ways:
- Paste the full text directly into the chat
- Provide a file path to a `.txt`, `.pdf`, or `.docx` file containing the posting

Extract and present a structured summary of:
- **Company name**
- **Job title / requisition title**
- **Key required skills** (hard skills, tools, languages, platforms)
- **Key preferred skills**
- **Seniority level & years of experience required**
- **Key responsibilities / deliverables**
- **Industry or domain context**
- **Any notable keywords / buzzwords used repeatedly**

Ask the user to confirm or correct this summary before proceeding.

---

## STEP 2 — Scan Resume Inventory & Auto-Select Baseline

Scan all files inside `/Users/dmanh/Projects/Project One/inventory/resumes/` recursively (`.docx`, `.doc`, `.pdf`, `.txt`, `.md`).

For each file:
1. Read its content (use Read tool; note if a file can't be parsed)
2. Identify: target role, industry, key skills highlighted, years of experience, most recent companies

Score each resume against the job posting:
| Criterion | Weight |
|---|---|
| Role/title alignment | 30% |
| Skills overlap | 35% |
| Industry/domain fit | 20% |
| Seniority match | 15% |

**Auto-select the top-scoring resume as the baseline. Do not ask for confirmation.**

Announce the selection clearly:
> **Baseline resume selected:** `{filename}` — tailored for {Company/Role context}, chosen because {one-line reason}.

Then proceed directly to Step 3.

If the inventory is empty, stop and tell the user to add at least one resume file to `inventory/resumes/` before continuing.

---

## STEP 3 — Analyze Gaps & Propose Revisions

With the baseline resume loaded, perform a section-by-section gap analysis against the job posting.

For each section (Summary, Experience, Skills, etc.), list:
- **Keep as-is**: content that already aligns well
- **Rephrase**: accurate content that uses different language than the posting — show original → suggested rewrite side by side
- **Reorder/promote**: strong content that should be moved up for prominence
- **Add**: keywords or phrasing the posting uses that are *already supported by existing experience* but not yet mentioned
- **Flag**: requirements in the posting NOT covered in the resume — do not invent anything, just flag for the user to decide

**Hard rule**: Never add new companies, titles, dates, degrees, or achievements not already in the resume.

Present the full gap analysis and wait for the user to approve, modify, or reject specific suggestions before generating the output document.

---

## STEP 4 — Scan Cover Letter Inventory & Auto-Select Baseline

Scan all files inside `/Users/dmanh/Projects/Project One/inventory/cover-letters/` recursively.

For each cover letter:
1. Extract: target role, company, tone/style markers (formal vs. conversational, sentence length, paragraph structure, opening hook style), approximate word count

Score against the current job posting:
| Criterion | Weight |
|---|---|
| Role/industry alignment | 40% |
| Tone match to company culture (inferred from posting) | 30% |
| Structural quality | 30% |

**Auto-select the top-scoring cover letter as the baseline. Do not ask for confirmation.**

Announce the selection clearly:
> **Baseline cover letter selected:** `{filename}` — written for {Company}, {Role}.

Then proceed directly to Step 5.

If the inventory is empty, draft a new cover letter from scratch using the user's writing style inferred from the resume.

---

## STEP 5 — Draft Cover Letter

Using the selected base as a style template (not a copy — do not reuse specific sentences verbatim):

1. **Preserve**: sentence rhythm, paragraph length, tone, level of formality, opening structure, closing style
2. **Adapt**: all company/role-specific references and accomplishment highlights to match the target job
3. **Structure**: Opening hook → Why this company/role → Specific value you bring (2–3 examples from the resume) → Forward-looking close

Draft the full cover letter and present it in the chat. Wait for user feedback and iterate until approved.

**Target length**: 3–4 paragraphs, 250–400 words.

---

## STEP 6 — Create Output Folder & Save Job Posting

### Folder name format
```
{Company Name} - {Job Title}
```
Sanitize: replace `/`, `\`, `:`, `*`, `?`, `"`, `<`, `>`, `|` with `-`. Trim extra spaces.

Full path: `/Users/dmanh/Projects/Project One/applications/{Company Name} - {Job Title}/`

Create that folder.

### Job posting file
Save to:
```
/Users/dmanh/Projects/Project One/applications/{Company Name} - {Job Title}/job-posting.txt
```
Include the full original posting text with the extracted summary prepended, clearly marked:
```
--- EXTRACTED SUMMARY ---
...
--- END SUMMARY ---

{full original posting text}
```

---

## STEP 7 — Generate DOCX & PDF Documents

### File naming convention

Follow the exact naming pattern used in the existing resume inventory — natural spacing, no underscores, words capitalised as in the original files:

- Resume: `David Manh {Job Title} Resume.docx` / `.pdf`
- Cover Letter: `David Manh {Job Title} Cover Letter.docx` / `.pdf`

Examples matching the existing inventory style:
- `David Manh Senior Technical Program Manager Resume.docx`
- `David Manh Senior Technical Program Manager Cover Letter.docx`
- `David Manh Program Manager Resume.docx`
- `David Manh Program Manager Cover Letter.docx`
- `David Manh Lead Technical Program Manager Resume.docx`

Use the job title from the posting (condensed/normalised if very long) as the role label in the filename.

### Build resume data JSON

```json
{
  "type": "resume",
  "output_docx": "/Users/dmanh/Projects/Project One/applications/{Company} - {Job Title}/David Manh {Job Title} Resume.docx",
  "output_pdf":  "/Users/dmanh/Projects/Project One/applications/{Company} - {Job Title}/David Manh {Job Title} Resume.pdf",
  "data": {
    "contact": {"name": "...", "phone": "...", "email": "...", "linkedin": "...", "location": "..."},
    "summary": "...",
    "experience": [
      {"title": "...", "company": "...", "dates": "...", "location": "...", "bullets": ["..."]}
    ],
    "skills": {"Category Name": ["skill1", "skill2"]},
    "education": [
      {"degree": "...", "school": "...", "dates": "...", "details": "..."}
    ],
    "certifications": ["..."],
    "sections": []
  }
}
```

### Build cover letter data JSON

```json
{
  "type": "cover_letter",
  "output_docx": "/Users/dmanh/Projects/Project One/applications/{Company} - {Job Title}/David Manh {Job Title} Cover Letter.docx",
  "output_pdf":  "/Users/dmanh/Projects/Project One/applications/{Company} - {Job Title}/David Manh {Job Title} Cover Letter.pdf",
  "data": {
    "sender": {"name": "...", "address": "...", "phone": "...", "email": "..."},
    "date": "...",
    "recipient": {"name": "Hiring Manager", "title": "Hiring Manager", "company": "...", "address": ""},
    "subject": "Re: {Job Title} — Application",
    "salutation": "Dear Hiring Manager,",
    "body_paragraphs": ["paragraph 1", "paragraph 2", "paragraph 3", "closing paragraph"],
    "closing": "Sincerely,",
    "signature": "David Manh"
  }
}
```

### Run the generator

Save each JSON to `/tmp/` and run:
```bash
python3 "/Users/dmanh/Projects/Project One/tools/doc_generator.py" /tmp/resume_input.json
python3 "/Users/dmanh/Projects/Project One/tools/doc_generator.py" /tmp/cl_input.json
```

Confirm each file was created with `ls -lh`.

---

## STEP 8 — Auto-Update Inventory + Index

### Copy back to inventory (self-learning loop)

Copy the finalised DOCX files into inventory using company-prefixed natural spacing — no underscores anywhere:

```
inventory/resumes/{Company} David Manh {Job Title} Resume.docx
inventory/cover-letters/{Company} David Manh {Job Title} Cover Letter.docx
```

Example:
```
inventory/resumes/Airbnb David Manh Senior Program Manager Resume.docx
inventory/cover-letters/Airbnb David Manh Senior Program Manager Cover Letter.docx
```

### Update the master index

Append a row to `/Users/dmanh/Projects/Project One/applications/INDEX.md`.

Create the file with this header if it doesn't exist:
```
# Application History Index

| Date | Company | Job Title | Output Folder | Base Resume | Base Cover Letter |
|------|---------|-----------|---------------|-------------|-------------------|
```

Then append:
```
| {YYYY-MM-DD} | {Company} | {Job Title} | [{Company} - {Job Title}](./{Company} - {Job Title}/) | {inventory/resumes/filename} | {inventory/cover-letters/filename} |
```

### Final summary to user

```
Application package created: {Job Title} at {Company}
Folder: applications/{Company} - {Job Title}/

  ✓ job-posting.txt
  ✓ David Manh {Job Title} Resume.docx
  ✓ David Manh {Job Title} Resume.pdf
  ✓ David Manh {Job Title} Cover Letter.docx
  ✓ David Manh {Job Title} Cover Letter.pdf

Baseline resume:       inventory/resumes/{filename}
Baseline cover letter: inventory/cover-letters/{filename}

Gaps flagged (not in resume):
  - {gap 1}
  - {gap 2}

Inventory updated. Both documents are now candidates for future applications.
```

---

## RULES

1. **No fabrication** — never add companies, titles, dates, degrees, or achievements not in the existing resume
2. **Gap analysis first** — always present and get approval on revisions before generating documents
3. **Cover letter draft first** — always present and get approval on the draft before generating documents
4. **Preserve voice** — the cover letter must sound like the user, not a template
5. **Natural file names** — proper spacing, capitalised words, "Cover Letter" not "Coverletter", no underscores in the document name itself (underscores only used as the company prefix separator in inventory filenames)
6. **Both formats always** — DOCX + PDF for every document; if PDF fails, report the error and give the DOCX path
7. **Inventory always grows** — every completed application feeds back into the scoring pool
