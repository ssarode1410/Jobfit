
# jobfit

Local, **zero-API-key** CLI tool: give it your base resume + one or more target designations. It collects live job openings from public sources and company ATS boards (Greenhouse, Lever, Ashby), checks location and visa eligibility for applicants in India, scores jobs against your experience, and exports formatted results to **Excel (`matches.xlsx`)** and **Markdown (`report.md`)**.

It never applies automatically and never rewrites your resume for you—it acts as your local job search pipeline and application tracker.

---

## Quick Start

### 1. Installation
```bash
git clone [https://github.com/ssarode1410/Jobfit.git](https://github.com/ssarode1410/Jobfit.git)
cd Jobfit

python3 -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt

```

**Note (Linux/macOS):** If `pypdf` strips whitespace from your PDF resume, convert it to plain text first:
 ```bash
 pdftotext Resume.pdf Resume.txt
 
 ```
 
---

## Usage

### Test Job Sources First

Verify which sources respond from your network and clean up any dead company slugs:

```bash
python3 jobfit.py sources --prune

```

### Standard Run (India + Worldwide Remote)

```bash
python3 jobfit.py run Resume.txt \
  --role "Cloud Systems Engineer, SRE, Site Reliability Engineer, DevOps Engineer, Cloud Engineer" \
  --years 6.5 \
  --no-relocation

```

### International / Visa Relocation Search

Separate your international job hunt into a dedicated output directory without overwriting your main India tracker:

```bash
python3 jobfit.py run Resume.txt \
  --role "Cloud Systems Engineer, SRE, Site Reliability Engineer, DevOps Engineer, Cloud Engineer" \
  --years 6.5 \
  --sources "boards,arbeitnow,remotive,remoteok,weworkremotely,jobicy,himalayas" \
  --work work_intl \
  --out jobfit_intl

```

### Score Pasted JDs (LinkedIn / Naukri / Workday)

For postings from sites without open APIs, paste JDs into a folder for maximum scoring accuracy:

```bash
python3 jobfit.py run Resume.txt --role "DevOps Engineer, SRE" --jd-dir ./jds

```

---

## Outputs

When `jobfit` finishes a run, it populates your output directory (default: `./jobfit_out/`):

1. **`matches.xlsx`** — An interactive Excel job application tracker featuring:
* Clickable **Direct Apply** URLs.
* Match Score, Location Tag, Company, and Role Title.
* Built-in **Status** dropdown menu (`To Apply`, `Applied`, `Interview`, `Rejected`, `Offer`, `Skipped`) to manage your daily workflow.
* Pre-formatted frozen headers and auto-filtered columns.


2. **`report.md`** — A detailed gap analysis for all matched postings (score ≥ 60%), including:
* Missing required & nice-to-have skills.
* Level and experience penalty breakdowns.
* Bullet point reordering and tailoring suggestions based on your resume.



---

## Demo

**Excel tracker (`matches.xlsx`):**


---

## Key CLI Flags

| Flag | Description | Default / Example |
| --- | --- | --- |
| `--role` | Target job designations (comma-separated strings) | Required |
| `--years` | Total years of relevant professional experience | `0.0` |
| `--min-match` | Minimum match percentage threshold for output reports | `60` |
| `--near` | Include near-miss roles within N percentage points of threshold | `15` |
| `--no-relocation` | Exclude international on-site roles offering visa/relocation | `False` |
| `--no-remote` | Exclude remote roles, keeping on-site/hybrid only | `False` |
| `--strict` | Drop remote jobs that do not explicitly state country eligibility | `False` |
| `--location` | Override location filter (`IN`, `WORLD`, `RELO`, or `any`) | Evaluates India |
| `--sources` | Specify feed sources (`boards`, `arbeitnow`, `remotive`, `remoteok`, `weworkremotely`, `jobicy`, `himalayas`) | All |
| `--work` | Directory for caching intermediate API response JSONs | `./work` |
| `--out` | Directory for exported `matches.xlsx` and `report.md` | `./jobfit_out` |

---

## Where Jobs Come From (No API Keys Required)

* **Remote Feeds:** Remotive, RemoteOK, We Work Remotely (RSS), Jobicy, Himalayas, Arbeitnow.
* **Direct Company ATS Feeds:** Parses public ATS endpoints listed in `boards.txt` (**Greenhouse**, **Lever**, **Ashby**). This is how India on-site and hybrid roles are fetched without scraper keys.
* **Manual JDs (`--jd-dir`):** Postings from LinkedIn, Naukri, Indeed, Workday, Darwinbox, Keka, etc. can be pasted directly into a target folder.

### Adding & Managing Company Boards

To track a new company's careers page:

```bash
python3 jobfit.py addboard [https://jobs.lever.co/](https://jobs.lever.co/)<company> [https://boards.greenhouse.io/](https://boards.greenhouse.io/)<company>

```

`addboard` identifies the underlying ATS, tests the endpoint, and appends valid slugs to `boards.txt`.

To test all listed boards and drop dead or renamed company endpoints:

```bash
python3 jobfit.py sources --prune

```

---

## "Can I take this job from India?" Location Logic

Every posting is evaluated for India applicant eligibility and assigned one of the following tags:

| Tag | Meaning | Kept by Default |
| --- | --- | --- |
| **`IN`** | Explicitly located in India (recognizes major Indian tech hubs & states). | Yes |
| **`WORLD`** | Remote and open to India or worldwide (APAC, Anywhere, Global, "We hire anywhere"). | Yes |
| **`RELO`** | On-site abroad, but the job posting explicitly mentions relocation assistance or visa sponsorship. | Yes *(Drop with `--no-relocation`)* |
| **`VERIFY`** | Remote role, but the description does not explicitly specify country boundaries. | Yes *(Drop with `--strict`)* |
| **`DROPPED`** | US/EU-only remote, US security clearance required, on-site abroad without relocation, or explicit "No Visa Sponsorship". | No |

---

## Scoring Formula & Algorithm

`jobfit` evaluates candidates against job descriptions using a multi-factor local scoring model:

$$\text{Score} = (0.55 \times \text{Skill Overlap}) + (0.25 \times \text{TF-IDF Similarity}) + (0.20 \times \text{Title Match}) - \text{Experience Penalties}$$

* **Skill Overlap (55%):** Weighted comparison of required vs. optional skills extracted from the JD against your resume.
* **TF-IDF Text Similarity (25%):** Cosine similarity between resume keyword distribution and JD description text.
* **Title Similarity (20%):** Standardized title normalization (e.g., mapping SRE ↔ DevOps ↔ Cloud Systems Engineer).
* **Experience Level Check:** Deducts up to 15 points if the JD demands significantly more years of experience than specified by `--years`, and up to 12 points for title level mismatches (e.g., Staff/Principal level if total experience < 9 years). Disable with `--no-level-penalty`.

---

## Honest Limitations

* **ATS Scope:** `jobfit` parses public feeds and Greenhouse/Lever/Ashby ATS APIs. Workday, Darwinbox, Keka, and custom enterprise portals do not expose open JSON endpoints; use `--jd-dir` for these.
* **Coverage:** Keyless sources cannot index every single role in India. Coverage depends on the companies listed in your `boards.txt` and active remote feeds.
* **Relocation Tags:** `RELO` matching is keyword-based. US postings mentioning standard domestic relocation benefits will match `RELO` unless excluded via `--no-relocation`. Always inspect recruiter details.

---

## Contributing

Contributions, ideas, and pull requests from developers, SREs, recruiters, and job seekers are very welcome!

### High-impact areas to improve

* **ATS & Job Sources:** Add support for additional public career boards or improve existing parsers.
* **Skill Taxonomy (`SKILLS`):** Expand canonical skills and aliases for non-DevOps domains (Data, Frontend, Backend, Security, Product, etc.).
* **Eligibility & Location Filters:** Improve regex/heuristics for detecting remote-eligible countries, visa sponsorship, and false positives.
* **Scoring & Matching:** Refine TF-IDF/skill weighting, add optional local embedding support (e.g., Sentence Transformers), or improve PDF text extraction.
* **Company Boards (`boards.txt`):** Submit verified Greenhouse, Lever, or Ashby slugs for companies actively hiring.

### How to contribute

1. Fork the repository.
2. Create a feature branch (`git checkout -b feat/improve-scoring`).
3. Make your changes and test them locally (`python3 jobfit.py sources` / `run`).
4. Commit your changes (`git commit -m "Add: description of improvement"`).
5. Push to your branch and open a **Pull Request** explaining what you changed and why.

Not a coder? You can still help by opening an **Issue** to report broken board slugs, suggest features, or share feedback on the workflow.

```
