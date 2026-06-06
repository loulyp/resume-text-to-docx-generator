# resume-text-to-docx-generator

## How This Dataset Was Generated

### Overview

The dataset was created in two main phases:

1. **Filtering** – Selecting ROWS resumes from the original dataset 
2. **Generation** – Converting raw resume text into formatted DOCX files using GPT-4o-mini

---

### Phase 1: Filtering the Original Dataset

The original dataset `cnamuangtoun/resume-job-description-fit` contains 6,241 resume-job pairs. To create a balanced dataset, I filtered it as follows:

| Step | Description |
|------|-------------|
| **Input** | 6,241 rows, 280 unique job descriptions, 642 unique resumes |
| **Filter criteria** | Each job description must have ≥5 Good Fit, ≥8 Potential Fit, ≥7 No Fit resumes |
| **Result** | 11 job descriptions met the criteria |
| **Selected** | 5 job descriptions covering diverse roles manually chosen (JD01, JD02, JD04, JD06, JD09) |
| **Sampling** | For each selected JD: 5 Good Fit + 8 Potential Fit + 7 No Fit = 20 resumes |
| **Total** | 5 JDs × 20 resumes = **100 resume-JD pairs** |

**Notebook used:** `explore_cnamuangtoun.ipynb`

The filtering notebook performs:
- Dataset exploration and label distribution analysis
- Per-JD label counts
- Filtering JDs meeting minimum requirements
- Random sampling (with seed=42) to extract exactly 5/8/7 per JD
- Saving to `selected_dataset.csv`

---

### Phase 2: Generating DOCX Files from Resume Text

The filtered dataset (`selected_dataset.csv`) contains raw `resume_text`. converted each resume into a professionally styled Word document using GPT-4o-mini.

#### Step 1: LLM Parsing (Text → JSON)

Each raw resume text is sent to OpenAI's **GPT-4o-mini** with a structured prompt. The prompt instructs the model to extract and structure the resume into JSON:

```json
{
  "name": "Full Name",
  "email": "email or empty string",
  "phone": "phone or empty string",
  "location": "city, state or empty string",
  "summary": "2-4 sentence professional summary paragraph",
  "skills": ["skill1", "skill2", "skill3"],
  "experience": [
    {
      "title": "Job Title",
      "company": "Company Name",
      "location": "City, State",
      "dates": "MM/YYYY - MM/YYYY or Present",
      "bullets": ["Achievement one", "Achievement two"]
    }
  ],
  "education": [
    {
      "degree": "Degree Name and Field",
      "institution": "University Name",
      "year": "Graduation year or expected"
    }
  ],
  "certifications": ["cert1", "cert2"],
  "languages": ["Language (level)"]
}
```

**LLM Settings:**
| Parameter | Value |
|-----------|-------|
| Model | `gpt-4o-mini` |
| Temperature | `0.1` (for consistent output) |

#### Step 2: DOCX Generation (JSON → Word Document)

The parsed JSON is rendered into a Microsoft Word (.docx) document using the `python-docx` library.

**Each DOCX contains these sections (when data is available):**
- Name (bold, accent color)
- Contact line (email | phone | location)
- Professional Summary
- Skills (grouped in rows of 4, separated by ` • `)
- Work Experience (title, company, location, dates, bullet points)
- Education
- Certifications
- Languages

#### Step 3: Random Styling Variation

To make the resumes look more natural, each DOCX receives **random style variations** from predefined pools:

| Style Element | Options |
|---------------|---------|
| Accent color | Navy, forest green, burgundy, teal, purple, brown, deep cyan, sage green |
| Separator symbol | `\|`, `•`, `—`, ` · ` |
| Margins | 0.75, 0.80, 0.90, 1.0 inch |
| Name font size | 16, 18, 20 pt |
| Body font size | 9.5, 10.0 pt |
| Section heading size | 10, 11 pt |
| Header alignment | Centered (2/3 chance) or left-aligned (1/3 chance) |
| Divider thickness | 6 or 8 |

#### Step 4: Output Files

The script produces:
| Output | Description |
|--------|-------------|
| `cv_files/` folder | 100 DOCX files named `{resume_id}.docx` |
| `final_dataset.csv` | CSV with metadata and file paths (renamed to `resume_metadata.csv`) |

**Script used:** `selected_cnamuangtoun_to_docx.py`

---

### Summary Diagram

```
Original Dataset (6,241 pairs)
        ↓
[Filtering] explore_cnamuangtoun.ipynb
        ↓
selected_dataset.csv (100 pairs: 5 JDs × 20 resumes)
        ↓
[Generation] selected_cnamuangtoun_to_docx.py
        ↓
    ┌───────────────┴───────────────┐
    ↓                               ↓
cv_files/ (100 DOCX files)    resume_metadata.csv
```

---

### Requirements to Reproduce

| Dependency | Version/Purpose |
|------------|-----------------|
| `pandas` | Data manipulation |
| `openai` | GPT-4o-mini API |
| `python-docx` | DOCX generation |
| `datasets` | Loading original dataset (filtering only) |

**Note:** An OpenAI API key is required to run the generation script.
