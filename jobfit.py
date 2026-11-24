#!/usr/bin/env python3
"""
jobfit - local, no-API-key resume <-> job matcher and resume-gap advisor.

Stages (each writes JSON into ./work so you can re-run any one of them):
  parse   resume.pdf/.docx/.txt  -> work/resume.json
  fetch   designation + sources  -> work/jobs.json
  score   resume x jobs          -> work/matches.json  (numbered candidate table)
  report  matches >= --min-match -> jobfit_out/report.md & jobfit_out/matches.xlsx
  sources test which job sources respond from your network (--prune removes dead boards)
  addboard <url> add a company's Greenhouse/Lever/Ashby board to boards.txt
  run     all of the above in one command

It NEVER applies to jobs and NEVER edits your resume. Suggestions only select,
reorder or re-emphasise facts already in your resume. Skills you don't have are
listed as "add ONLY if true".
"""
from __future__ import annotations

import argparse
import html
import json
import math
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
from functools import partial
from pathlib import Path

# --------------------------------------------------------------------------
# Skill taxonomy: canonical name -> aliases. Extend freely (or add your own).
# --------------------------------------------------------------------------
SKILLS = {
    # cloud / infra
    "aws": ["aws", "amazon web services"], "azure": ["azure", "microsoft azure"],
    "gcp": ["gcp", "google cloud", "google cloud platform"],
    "ec2": ["ec2"], "s3": ["s3"], "eks": ["eks"], "ecs": ["ecs"], "lambda": ["aws lambda"],
    "cloudfront": ["cloudfront"], "cloudformation": ["cloudformation"], "route 53": ["route 53", "route53"],
    "iam": ["iam"], "vpc": ["vpc"], "rds": ["rds"],
    "terraform": ["terraform"], "ansible": ["ansible"], "pulumi": ["pulumi"], "helm": ["helm"],
    "kubernetes": ["kubernetes", "k8s"], "docker": ["docker"], "openshift": ["openshift"],
    "iac": ["infrastructure as code", "iac"], "linux": ["linux", "unix"], "nginx": ["nginx"],
    "networking": ["networking", "tcp/ip", "dns", "load balancing", "load balancer"],
    # ci/cd
    "ci/cd": ["ci/cd", "ci cd", "cicd", "continuous integration", "continuous delivery", "continuous deployment"],
    "jenkins": ["jenkins"], "github actions": ["github actions"],
    "gitlab ci": ["gitlab ci", "gitlab-ci", "gitlab ci/cd"], "azure devops": ["azure devops"],
    "circleci": ["circleci"], "codepipeline": ["codepipeline", "aws codepipeline"],
    "argocd": ["argocd", "argo cd"], "gitops": ["gitops"], "git": ["git", "github", "gitlab", "bitbucket"],
    # observability / reliability
    "prometheus": ["prometheus"], "grafana": ["grafana"], "datadog": ["datadog"], "splunk": ["splunk"],
    "elk": ["elk", "elasticsearch", "elastic stack", "opensearch"], "cloudwatch": ["cloudwatch"],
    "new relic": ["new relic", "newrelic"], "opentelemetry": ["opentelemetry", "otel"],
    "observability": ["observability"], "monitoring": ["monitoring", "alerting"],
    "sre": ["sre", "site reliability"],
    "incident management": ["incident management", "incident response", "on-call", "on call", "oncall"],
    "slo/sli": ["slo", "slos", "sli", "slis", "error budget", "error budgets"],
    # security
    "devsecops": ["devsecops"], "sonarqube": ["sonarqube"], "trivy": ["trivy"],
    "vault": ["hashicorp vault"], "owasp": ["owasp"], "security": ["security scanning", "vulnerability management"],
    # languages / frameworks
    "python": ["python"], "bash": ["bash", "shell scripting", "shell script", "shell scripts"],
    "go": ["golang"], "java": ["java"], "javascript": ["javascript"], "typescript": ["typescript"],
    "node.js": ["node.js", "nodejs", "node js"], "c#": ["c#"], ".net": [".net", "dotnet", "asp.net"],
    "sql": ["sql"], "powershell": ["powershell"], "ruby": ["ruby"], "rust": ["rust"], "php": ["php"],
    "react": ["react", "react.js", "reactjs"], "django": ["django"], "flask": ["flask"], "fastapi": ["fastapi"],
    # data
    "postgresql": ["postgresql", "postgres"], "mysql": ["mysql"], "mongodb": ["mongodb"], "redis": ["redis"],
    "kafka": ["kafka"], "rabbitmq": ["rabbitmq"], "spark": ["spark", "pyspark"], "airflow": ["airflow"],
    "snowflake": ["snowflake"], "pandas": ["pandas"],
    # ml
    "machine learning": ["machine learning"], "pytorch": ["pytorch"], "tensorflow": ["tensorflow"],
    "llm": ["llm", "llms", "large language model", "large language models"],
    # practices
    "agile": ["agile", "scrum", "kanban"], "jira": ["jira"], "microservices": ["microservices", "microservice"],
    "rest api": ["rest api", "rest apis", "restful"], "graphql": ["graphql"],
    "cost optimization": ["cost optimization", "cost optimisation", "finops"],
    "disaster recovery": ["disaster recovery", "high availability"],
}
SKILL_RE = {
    canon: re.compile(
        r"(?<![A-Za-z0-9])(?:%s)(?![A-Za-z0-9+#])"
        % "|".join(re.escape(a) for a in sorted(aliases, key=len, reverse=True)), re.I)
    for canon, aliases in SKILLS.items()
}

STOP = set("""the and for with you our are will your have has that this from not but all can use able work team role
experience years year job company skills such including etc who what when where how their they them also more
most other than then these those into out over any each per via was were been being its it's about across within
without ability strong good great new we're we'll we've looking join help build etc working must should would
could may might one two three""".split())
SENIORITY = {"senior", "sr", "jr", "junior", "lead", "principal", "staff", "associate", "i", "ii", "iii", "iv",
             "and", "of", "the", "remote", "m", "f", "d", "w"}
GENERIC_ROLE_WORDS = SENIORITY | {"engineer", "developer", "manager", "specialist", "consultant", "analyst"}
ROLE_SYNONYMS = {
    "devops": ["devops", "sre", "site reliability", "platform", "infrastructure", "cloud", "release", "reliability"],
    "sre": ["sre", "site reliability", "reliability", "devops", "platform", "infrastructure"],
    "cloud": ["cloud", "devops", "infrastructure", "platform", "sre"],
    "platform": ["platform", "devops", "infrastructure", "sre", "cloud"],
}
TITLE_RE = re.compile(
    r"\b(?:(?:senior|sr\.?|lead|principal|junior|associate)\s+)?"
    r"(?:devops|sre|site reliability|cloud\s+systems?|cloud|platform|software|backend|frontend|full[- ]?stack|data|ml|"
    r"machine learning|security|systems?|infrastructure|qa|test|release|build)\s+"
    r"(?:engineer|developer|architect|analyst|administrator|specialist|consultant|manager)\b", re.I)
BULLET_RE = re.compile(r"^\s*[-•*▪●◦·–]\s*(.*\S)")
YEARS_RE = re.compile(r"(\d{1,2}(?:\.\d)?)\s*\+?\s*(?:-|to|–)?\s*(?:\d{1,2}\s*\+?\s*)?(?:years?|yrs?)", re.I)
NICE_RE = re.compile(r"nice to have|preferred|bonus|good to have|a plus|desirable|optional|extra credit", re.I)


# --------------------------------------------------------------------------
# helpers
# --------------------------------------------------------------------------
def die(msg: str):
    print("ERROR: " + msg, file=sys.stderr)
    sys.exit(1)


def strip_html(raw: str) -> str:
    t = html.unescape(raw or "")
    t = re.sub(r"</(p|li|div|h\d|ul|ol|tr)>|<br\s*/?>", "\n", t, flags=re.I)
    t = re.sub(r"<li[^>]*>", "\n- ", t, flags=re.I)
    t = re.sub(r"<[^>]+>", " ", t)
    t = html.unescape(t)
    t = re.sub(r"[ \t]+", " ", t)
    return re.sub(r"\n\s*\n+", "\n", t).strip()


def tokens(text: str) -> list[str]:
    return [t for t in re.findall(r"[a-z][a-z0-9+#]{2,}", text.lower()) if t not in STOP]


def count_skills(text: str) -> Counter:
    c: Counter = Counter()
    for canon, rx in SKILL_RE.items():
        n = len(rx.findall(text))
        if n:
            c[canon] = n
    return c


def norm_title_tokens(t: str) -> set[str]:
    t = t.lower().replace("site reliability", "sre").replace("dev ops", "devops")
    return {w for w in re.findall(r"[a-z0-9+#]+", t) if w not in SENIORITY}


def role_keywords(role: str) -> set[str]:
    toks = {w for w in re.findall(r"[a-z0-9+#]+", role.lower()) if w not in GENERIC_ROLE_WORDS}
    toks = toks or set(re.findall(r"[a-z0-9+#]+", role.lower()))
    for w in list(toks):
        toks.update(ROLE_SYNONYMS.get(w, []))
    return toks


def split_roles(role: str) -> list[str]:
    """'DevOps Engineer, SRE, Cloud Engineer' -> three roles (split on , ; | or ' / ')."""
    return [r.strip() for r in re.split(r"[,;|]|\s/\s", role) if r.strip()] or [role]


NEG_TITLE_RE = re.compile(
    r"power platform|power apps|power bi|salesforce|servicenow|\bsap\b|trainer|instructor|teacher|tutor|"
    r"\bsales\b|marketing|recruit|talent|counsel|customer success|account (?:executive|manager)|legal|"
    r"finance|payroll|designer|\bux\b|content|writer|help ?desk|technical support|support specialist|"
    r"data scientist|\bintern(?:ship)?\b", re.I)


def title_relevant(role: str, title: str) -> bool:
    tl = (title or "").lower()
    if NEG_TITLE_RE.search(tl):
        return False
    return any(re.search(r"\b" + re.escape(k), tl) for k in role_keywords(role))


def save_json(path: Path, obj):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(obj, indent=2, ensure_ascii=False), encoding="utf-8")


def load_json(path: Path, what: str):
    if not path.exists():
        die(f"{path} not found - run the '{what}' stage first.")
    return json.loads(path.read_text(encoding="utf-8"))


# --------------------------------------------------------------------------
# preflight
# --------------------------------------------------------------------------
def preflight(args) -> None:
    problems = []
    resume = getattr(args, "resume", None)
    if resume:
        ext = Path(resume).suffix.lower()
        if ext == ".pdf":
            try:
                import pypdf  # noqa: F401
            except ImportError:
                problems.append("pypdf missing -> pip install pypdf")
        if ext == ".docx":
            try:
                import docx  # noqa: F401
            except ImportError:
                problems.append("python-docx missing -> pip install python-docx")
        if not Path(resume).exists():
            problems.append(f"resume file not found: {resume}")
    if getattr(args, "llm", None):
        try:
            urllib.request.urlopen("http://localhost:11434/api/tags", timeout=3)
        except Exception:
            problems.append("--llm given but Ollama isn't reachable on localhost:11434 "
                            "(start it, or drop --llm; the built-in scorer needs no LLM)")
    if problems:
        print("PREFLIGHT FAILED:")
        for p in problems:
            print("  - " + p)
        sys.exit(2)


# --------------------------------------------------------------------------
# stage 1: parse resume
# --------------------------------------------------------------------------
def read_resume_text(path: str) -> str:
    p = Path(path)
    ext = p.suffix.lower()
    if ext == ".pdf":
        from pypdf import PdfReader
        text = "\n".join((pg.extract_text() or "") for pg in PdfReader(str(p)).pages)
        if len(text.strip()) < 100:
            die("No text layer found in the PDF (scanned image?). Export a text PDF or use .docx/.txt.")
        return text
    if ext == ".docx":
        import docx
        d = docx.Document(str(p))
        parts = [para.text for para in d.paragraphs]
        for tbl in d.tables:
            for row in tbl.rows:
                parts.append(" | ".join(c.text for c in row.cells))
        return "\n".join(parts)
    return p.read_text(encoding="utf-8", errors="ignore")


def extract_bullets(text: str) -> list[str]:
    bullets: list[str] = []
    cur = None
    for line in text.splitlines():
        m = BULLET_RE.match(line)
        if m:
            if cur:
                bullets.append(cur)
            cur = m.group(1).strip()
        elif cur and line.strip() and line.strip()[0].islower():
            cur += " " + line.strip()  # wrapped continuation line
        else:
            if cur:
                bullets.append(cur)
            cur = None
    if cur:
        bullets.append(cur)
    bullets = [b for b in bullets if len(b) > 25]
    if len(bullets) < 3:  # no bullet markers (e.g. PDF text) -> use sentence-like lines
        bullets = [l.strip() for l in text.splitlines() if 40 <= len(l.strip()) <= 260]
    return bullets


def cmd_parse(args) -> dict:
    text = read_resume_text(args.resume)
    ym = YEARS_RE.search(text[:2000])
    years = args.years if args.years is not None else (float(ym.group(1)) if ym else None)
    titles = []
    for m in TITLE_RE.finditer(text):
        t = m.group(0).strip()
        if t.lower() not in [x.lower() for x in titles]:
            titles.append(t)
    resume = {
        "source_file": str(args.resume), "text": text, "skills": dict(count_skills(text)),
        "years": years, "titles": titles[:6], "bullets": extract_bullets(text),
    }
    save_json(Path(args.work) / "resume.json", resume)
    print(f"[parse] {len(resume['skills'])} skills, {len(resume['bullets'])} bullets, "
          f"years={years}, titles={resume['titles'] or '-'}")
    if years is None:
        print("        (couldn't detect years of experience - pass --years N for better scoring)")
    return resume


# --------------------------------------------------------------------------
# stage 2: fetch jobs (public, ToS-friendly sources only)
# --------------------------------------------------------------------------
UA = {"User-Agent": "jobfit/1.1 (personal job-search CLI)"}


def http_get(url: str, timeout: int = 20, accept: str = "application/json") -> str:
    req = urllib.request.Request(url, headers={**UA, "Accept": accept})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read().decode("utf-8", errors="ignore")


def http_json(url: str, timeout: int = 20):
    return json.loads(http_get(url, timeout))


def mk_job(source, jid, title, company, location, url, desc, remote=None, regions=None) -> dict:
    return {"id": f"{source}:{jid}", "source": source, "title": (title or "").strip(),
            "company": (company or "").strip(), "location": (location or "").strip(),
            "url": url or "", "description": strip_html(desc or ""),
            "remote": remote, "regions": regions}


def primary_keyword(role: str) -> str:
    words = [w for w in re.findall(r"[a-z0-9+#]+", role.lower()) if w not in GENERIC_ROLE_WORDS]
    return words[0] if words else role.lower()


# ---- remote job feeds (public, no key) -----------------------------------
def feed_remotive(role):
    q = urllib.parse.quote(role)
    for j in http_json(f"https://remotive.com/api/remote-jobs?search={q}&limit=100").get("jobs", []):
        yield mk_job("remotive", j.get("id"), j.get("title"), j.get("company_name"),
                     j.get("candidate_required_location") or "", j.get("url"), j.get("description"), remote=True)


def feed_remoteok():
    for j in http_json("https://remoteok.com/api"):
        if isinstance(j, dict) and j.get("position"):
            yield mk_job("remoteok", j.get("id"), j["position"], j.get("company"), j.get("location") or "",
                         j.get("apply_url") or j.get("url"), j.get("description"), remote=True)


WWR_CATEGORIES = ["remote-devops-sysadmin-jobs", "remote-back-end-programming-jobs",
                  "remote-full-stack-programming-jobs"]


def feed_wwr():
    n, err = 0, None
    for cat in WWR_CATEGORIES:
        try:
            root = ET.fromstring(http_get(f"https://weworkremotely.com/categories/{cat}.rss",
                                          accept="application/rss+xml"))
        except Exception as e:  # noqa: BLE001
            err = e
            continue
        for it in root.iter("item"):
            title = (it.findtext("title") or "").strip()
            company, sep, role = title.partition(": ")
            n += 1
            yield mk_job("weworkremotely", it.findtext("guid") or title, role if sep else title,
                         company if sep else "", (it.findtext("region") or "").strip(),
                         it.findtext("link"), it.findtext("description"), remote=True)
    if n == 0 and err:
        raise err


def feed_jobicy(role):
    q = urllib.parse.quote(primary_keyword(role))
    for j in http_json(f"https://jobicy.com/api/v2/remote-jobs?count=50&tag={q}").get("jobs", []):
        geo = j.get("jobGeo")
        geo = ", ".join(geo) if isinstance(geo, list) else (geo or "")
        yield mk_job("jobicy", j.get("id"), j.get("jobTitle"), j.get("companyName"), geo, j.get("url"),
                     j.get("jobDescription") or j.get("jobExcerpt"), remote=True)


def feed_himalayas(role):
    q = urllib.parse.quote(primary_keyword(role))
    for j in http_json(f"https://himalayas.app/jobs/api/search?q={q}").get("jobs", []):
        regs = j.get("locationRestrictions") or None
        if isinstance(regs, str):
            regs = [regs]
        regs = [str(r) for r in regs] if regs else None
        link = j.get("applicationLink") or j.get("guid") or ""
        yield mk_job("himalayas", link or j.get("title"), j.get("title"), j.get("companyName"),
                     ", ".join(regs) if regs else "", link, j.get("description") or j.get("excerpt"),
                     remote=True, regions=regs)


def feed_arbeitnow(pages=3):  # Germany-based board; opt-in only (has relocation-friendly listings)
    for p in range(1, pages + 1):
        for j in http_json(f"https://www.arbeitnow.com/api/job-board-api?page={p}").get("data", []):
            yield mk_job("arbeitnow", j.get("slug"), j.get("title"), j.get("company_name"),
                         j.get("location") or "", j.get("url"), j.get("description"),
                         remote=True if j.get("remote") else None)


# ---- company career boards (public ATS endpoints, no key) -----------------
def board_greenhouse(slug):
    d = http_json(f"https://boards-api.greenhouse.io/v1/boards/{slug}/jobs?content=true")
    for j in d.get("jobs", []):
        yield mk_job(f"greenhouse-{slug}", j.get("id"), j.get("title"), slug,
                     (j.get("location") or {}).get("name"), j.get("absolute_url"), j.get("content"))


def board_lever(slug):
    for j in http_json(f"https://api.lever.co/v0/postings/{slug}?mode=json"):
        parts = [j.get("descriptionPlain") or ""]
        for lst in j.get("lists", []) or []:
            parts.append((lst.get("text") or "") + "\n" + strip_html(lst.get("content") or ""))
        parts.append(j.get("additionalPlain") or "")
        wt = (j.get("workplaceType") or "").lower()
        remote = True if wt == "remote" else (False if wt in ("on-site", "onsite", "hybrid") else None)
        yield mk_job(f"lever-{slug}", j.get("id"), j.get("text"), slug,
                     (j.get("categories") or {}).get("location"), j.get("hostedUrl"), "\n".join(parts),
                     remote=remote)


def board_ashby(slug):
    d = http_json(f"https://api.ashbyhq.com/posting-api/job-board/{slug}")
    for j in d.get("jobs", []):
        locs = [j.get("location") or ""] + [s.get("location", "") for s in (j.get("secondaryLocations") or [])
                                              if isinstance(s, dict)]
        yield mk_job(f"ashby-{slug}", j.get("id"), j.get("title"), slug, "; ".join(l for l in locs if l),
                     j.get("jobUrl") or j.get("applyUrl"), j.get("descriptionHtml") or j.get("descriptionPlain"),
                     remote=j.get("isRemote"))


BOARD_FNS = {"greenhouse": board_greenhouse, "lever": board_lever, "ashby": board_ashby}


def load_boards(path) -> list[str]:
    p = Path(path) if path else None
    if not p or not p.exists():
        return []
    out = []
    for line in p.read_text(encoding="utf-8").splitlines():
        line = line.split("#")[0].strip()
        if re.fullmatch(r"(greenhouse|lever|ashby):[A-Za-z0-9_.%-]+", line):
            out.append(line)
    return out


def load_local_jds(folder: str):
    for p in sorted(Path(folder).glob("*")):
        if p.suffix.lower() not in {".txt", ".md"}:
            continue
        raw = p.read_text(encoding="utf-8", errors="ignore")
        meta = {"title": None, "company": "", "location": "", "url": ""}
        body = []
        for line in raw.splitlines():
            m = re.match(r"^(title|company|location|url)\s*:\s*(.+)$", line.strip(), re.I)
            if m and not body:
                meta[m.group(1).lower()] = m.group(2).strip()
            else:
                body.append(line)
        title = meta["title"] or next((l.strip() for l in body if l.strip()), p.stem)
        yield mk_job("local", p.stem, title, meta["company"], meta["location"], meta["url"], "\n".join(body))


def run_tasks(tasks, workers=12):
    """tasks: list of (kind, label, fn). Runs concurrently; never raises."""
    def one(t):
        kind, label, fn = t
        try:
            return kind, label, list(fn()), None
        except urllib.error.HTTPError as e:
            return kind, label, [], f"HTTP {e.code}"
        except Exception as e:  # noqa: BLE001
            return kind, label, [], str(getattr(e, "reason", e))[:90]
    with ThreadPoolExecutor(max_workers=workers) as ex:
        return list(ex.map(one, tasks))


def build_tasks(specs: list[str], roles: list[str], boards_file) -> list[tuple]:
    tasks: list[tuple] = []
    for spec in specs:
        name, _, arg = spec.partition(":")
        if name == "boards":
            for b in load_boards(boards_file):
                kind, slug = b.split(":", 1)
                tasks.append(("board", b, partial(BOARD_FNS[kind], slug)))
        elif name in BOARD_FNS and arg:
            tasks.append(("board", spec, partial(BOARD_FNS[name], arg)))
        elif name == "remotive":
            tasks += [("feed", f"remotive[{r}]", partial(feed_remotive, r)) for r in roles]
        elif name == "jobicy":
            tasks += [("feed", f"jobicy[{primary_keyword(r)}]", partial(feed_jobicy, r)) for r in roles]
        elif name == "himalayas":
            tasks += [("feed", f"himalayas[{primary_keyword(r)}]", partial(feed_himalayas, r)) for r in roles]
        elif name == "remoteok":
            tasks.append(("feed", "remoteok", feed_remoteok))
        elif name in ("weworkremotely", "wwr"):
            tasks.append(("feed", "weworkremotely", feed_wwr))
        elif name == "arbeitnow":
            tasks.append(("feed", "arbeitnow", feed_arbeitnow))
        else:
            print(f"[fetch] unknown source '{spec}' (boards, remotive, remoteok, weworkremotely, jobicy, "
                  f"himalayas, arbeitnow, greenhouse:<slug>, lever:<slug>, ashby:<slug>)")
    return tasks


# --------------------------------------------------------------------------
# eligibility: can someone living in India actually take this job?
# --------------------------------------------------------------------------
HOME_ALIASES = {
    "india": ["india", "bengaluru", "bangalore", "mumbai", "pune", "hyderabad", "chennai", "delhi", "gurgaon",
              "gurugram", "noida", "kolkata", "ahmedabad", "kochi", "nashik", "jaipur", "indore", "coimbatore",
              "thiruvananthapuram", "trivandrum", "navi mumbai", "thane", "chandigarh", "nagpur", "bhubaneswar",
              "mysuru", "mysore", "vadodara", "surat", "lucknow", "maharashtra", "karnataka", "telangana",
              "tamil nadu", "haryana", "uttar pradesh", "gujarat", "west bengal", "kerala"],
}
COUNTRIES = (r"us|u\.s\.|usa|united states|america|canada|uk|u\.k\.|united kingdom|britain|eu|europe|"
             r"european union|emea|latam|latin america|north america|germany|france|spain|italy|netherlands|"
             r"poland|ireland|australia|new zealand|brazil|mexico|singapore|uae|japan")
WORLDWIDE_RE = re.compile(r"worldwide|world-wide|anywhere|global(?:ly)?|work from anywhere|apac|asia", re.I)
REMOTE_RE = re.compile(r"\bremote\b|work from home|work-from-home|\bwfh\b|telecommut|fully distributed", re.I)
GENERIC_REMOTE_RE = re.compile(r"(?:100% |fully |global |distributed )?remote(?: job| work| - global)?|"
                               r"anywhere|worldwide|", re.I)
HARD_RES = [
    re.compile(rf"\b(?:{COUNTRIES})[- ](?:only|based|residents?|citizens?|timezones?)\b", re.I),
    re.compile(rf"\b(?:must|need to|needs to|required to|have to|should)\s+(?:currently\s+)?(?:be\s+)?"
               rf"(?:located|based|residing|reside|living|live|resident)\s+(?:in|within)\s+(?:the\s+)?"
               rf"(?:{COUNTRIES})\b", re.I),
    re.compile(rf"\b(?:authori[sz]ed|eligible|legally entitled|legal right|right)\s+to\s+work\s+in\s+"
               rf"(?:the\s+)?(?:{COUNTRIES})\b", re.I),
    re.compile(r"\bmust\s+(?:already\s+)?(?:be\s+)?(?:legally\s+)?(?:authori[sz]ed|eligible)\s+to\s+work\b", re.I),
]
SPONSOR_NEG_RE = re.compile(
    r"\bno\s+(?:visa\s+)?sponsorship\b|\b(?:unable|not able|cannot|can't|won't|will not|do not|don't)\s+(?:to\s+)?"
    r"(?:provide\s+|offer\s+)?(?:visa\s+)?sponsor|\bwithout\s+(?:the need for\s+)?(?:visa\s+)?sponsorship", re.I)
RELO_POS_RE = re.compile(
    r"visa\s+sponsorship\s+(?:is\s+)?(?:available|provided|offered)|\bwe\s+(?:will\s+)?sponsor\b|"
    r"sponsor(?:ship)?\s+(?:of\s+)?(?:work\s+)?visas?|relocation\s+(?:assistance|support|package|bonus|allowance|"
    r"benefits?|provided|offered)|(?:help|assist)(?:ance)?\s+(?:you\s+)?(?:with\s+)?reloca", re.I)
TZ_RE = re.compile(r"(?:working )?(?:hours|overlap|time ?zones?)[^.\n]{0,60}\b(?:est|edt|pst|pdt|cst|et|pt|cet|"
                   r"cest|gmt|utc[-+]?\d*)\b", re.I)
TEXT_WORLDWIDE_RE = re.compile(r"work from anywhere|anywhere in the world|any country|any location|"
                               r"hire (?:talent )?(?:from )?(?:anywhere|globally|worldwide)|"
                               r"remote[- ]first[^.\n]{0,40}(?:worldwide|globally)", re.I)


def home_regex(home: str):
    names = HOME_ALIASES.get(home.lower(), [home.lower()])
    return re.compile(r"\b(?:%s)\b" % "|".join(re.escape(n) for n in names), re.I)


def classify(job: dict, home_rx, allow_reloc: bool = True) -> tuple[str, str]:
    """-> (status, reason). Statuses: india, remote-worldwide, relocation, verify (all kept by default);
    remote-restricted, onsite-abroad (dropped by default); pasted (local JD, not checked)."""
    if job["source"] == "local":
        return "pasted", "pasted JD - location not checked"
    loc, text = job["location"], job["description"]
    remote = job.get("remote")
    if remote is None:
        remote = bool(REMOTE_RE.search(f"{job['title']} {loc}")) or bool(REMOTE_RE.search(text[:800]))
    hard = [m.group(0) for rx in HARD_RES for m in [rx.search(f"{loc}. {text}")] if m]
    sponsor_neg = SPONSOR_NEG_RE.search(text)
    tz_note = " (note: posting mentions timezone/working-hours overlap)" if TZ_RE.search(text) else ""

    if home_rx.search(loc):
        return "india", ("remote within India" if remote else "on-site/hybrid in India") + tz_note
    regions = job.get("regions")
    if regions:
        rl = " ".join(regions)
        if home_rx.search(rl) or WORLDWIDE_RE.search(rl):
            return "remote-worldwide", f"remote, open to: {', '.join(regions[:4])}{tz_note}"
        return "remote-restricted", f"remote, limited to: {', '.join(regions[:4])}"
    if remote:
        explicit = bool(WORLDWIDE_RE.search(loc))
        if explicit and not hard:
            extra = " (note: says no visa sponsorship - confirm contractor hiring)" if sponsor_neg else ""
            return "remote-worldwide", f"location says '{loc}'{extra}{tz_note}"
        if hard:
            return "remote-restricted", f"restricted: '{hard[0][:60]}'"
        if not GENERIC_REMOTE_RE.fullmatch(loc.strip()):
            return "remote-restricted", f"remote but tied to '{loc}'"
        if sponsor_neg:
            return "remote-restricted", "says no visa sponsorship / needs existing work authorisation"
        if TEXT_WORLDWIDE_RE.search(text):
            return "remote-worldwide", f"posting says it hires from anywhere{tz_note}"
        return "verify", "remote, but the posting doesn't state which countries - check before applying"
    if allow_reloc and RELO_POS_RE.search(text) and not sponsor_neg:
        return "relocation", f"on-site in '{loc or 'unknown'}'; posting mentions relocation/visa support"
    if sponsor_neg:
        return "onsite-abroad", f"on-site in '{loc or 'unknown'}', no visa sponsorship"
    return "onsite-abroad", f"on-site in '{loc or 'unknown'}', no relocation/visa support mentioned"


def cmd_fetch(args) -> list[dict]:
    roles = split_roles(args.role)
    jobs: list[dict] = []
    if args.jd_dir:
        local = list(load_local_jds(args.jd_dir))
        jobs += local
        print(f"[fetch] local JDs: {len(local)} from {args.jd_dir}")
    specs = [s.strip() for s in args.sources.split(",") if s.strip() and s.strip() != "none"]
    results = run_tasks(build_tasks(specs, roles, args.boards))
    board_ok = board_bad = 0
    bad_names = []
    for kind, label, got, err in results:
        kept = [j for j in got if any(title_relevant(r, j["title"]) for r in roles)]
        if kind == "board":
            if err:
                board_bad += 1
                bad_names.append(f"{label} ({err})")
            else:
                board_ok += 1
        else:
            print(f"[fetch] {label:<28}: " + (f"FAILED ({err})" if err else f"{len(got)} listings, {len(kept)} match role"))
        jobs += kept
    if board_ok or board_bad:
        n_b = sum(len(g) for k, _, g, e in results if k == "board")
        print(f"[fetch] company boards       : {board_ok} responded ({n_b} listings), {board_bad} skipped")
        if bad_names:
            print("        skipped: " + ", ".join(bad_names[:8]) + (" ..." if len(bad_names) > 8 else "")
                  + "\n        (dead slugs are harmless; run `jobfit.py sources` to check, edit boards.txt to prune)")

    if args.location.lower() != "any":
        home_rx = home_regex(args.location)
        keep = {"india", "pasted", "remote-worldwide", "relocation", "verify"}
        if args.no_remote:
            keep -= {"remote-worldwide", "verify"}
        if args.no_relocation:
            keep.discard("relocation")
        if args.strict:
            keep.discard("verify")
        dropped: Counter = Counter()
        kept_jobs = []
        for j in jobs:
            j["eligibility"], j["elig_reason"] = classify(j, home_rx, not args.no_relocation)
            if j["eligibility"] in keep:
                kept_jobs.append(j)
            else:
                dropped[j["eligibility"]] += 1
        kc = Counter(j["eligibility"] for j in kept_jobs)
        print("[fetch] eligibility check    : kept " + ", ".join(f"{n} {k}" for k, n in kc.items())
              if kc else "[fetch] eligibility check    : nothing eligible")
        if dropped:
            print("        dropped: " + ", ".join(f"{n} {k}" for k, n in dropped.items()))
        jobs = kept_jobs
    else:
        for j in jobs:
            j["eligibility"], j["elig_reason"] = "any", "location filter off"

    seen, uniq = set(), []
    for j in jobs:
        k = (j["title"].lower(), j["company"].lower(), j["location"].lower())
        if k not in seen:
            seen.add(k)
            uniq.append(j)
    save_json(Path(args.work) / "jobs.json", uniq)
    print(f"[fetch] {len(uniq)} unique jobs -> {Path(args.work) / 'jobs.json'}")
    if not uniq:
        print("        nothing to score. Run `python jobfit.py sources` to see which sources respond from your\n"
              "        network, add company boards to boards.txt, or paste JDs into a folder and use --jd-dir.")
    return uniq


def cmd_sources(args) -> None:
    """Probe every source so you can see what works from YOUR network."""
    probe = "DevOps Engineer"
    tasks = build_tasks(["remotive", "remoteok", "weworkremotely", "jobicy", "himalayas", "arbeitnow"],
                        [probe], None)
    tasks += build_tasks(["boards"], [probe], args.boards)
    res = run_tasks(tasks)
    print(f"{'source':<30} status")
    for kind, label, got, err in res:
        print(f"{label:<30} " + (f"FAIL  {err}" if err else f"OK    {len(got)} listings"))
    bad = [l for k, l, g, e in res if k == "board" and e]
    empty = [l for k, l, g, e in res if k == "board" and not e and not g]
    if empty:
        print("\nresponded but had 0 listings (slug may be wrong): " + ", ".join(empty))
    if bad and args.prune:
        bp = Path(args.boards)
        bp.with_suffix(".txt.bak").write_text(bp.read_text(encoding="utf-8"), encoding="utf-8")
        kept = [ln for ln in bp.read_text(encoding="utf-8").splitlines() if ln.split("#")[0].strip() not in bad]
        bp.write_text("\n".join(kept) + "\n", encoding="utf-8")
        print(f"\nPruned {len(bad)} dead slug(s) from {bp} (backup: {bp.name}.bak)")
    elif bad:
        print(f"\n{len(bad)} board slug(s) failed. Remove them by re-running: python jobfit.py sources --prune\n  "
              + "\n  ".join(bad))


ATS_URL_RES = [("greenhouse", re.compile(r"(?:job-)?boards(?:\.eu)?\.greenhouse\.io/(?:embed/job_board\?for=)?([A-Za-z0-9_-]+)")),
               ("lever", re.compile(r"jobs(?:\.eu)?\.lever\.co/([A-Za-z0-9_.-]+)")),
               ("ashby", re.compile(r"jobs\.ashbyhq\.com/([A-Za-z0-9_.%-]+)"))]


def cmd_addboard(args) -> None:
    """jobfit.py addboard <careers-page-url>... : detect the ATS + slug, verify it responds, add to boards.txt."""
    bp = Path(args.boards)
    existing = set(load_boards(bp))
    for url in args.urls:
        spec = next((f"{k}:{m.group(1)}" for k, rx in ATS_URL_RES for m in [rx.search(url)] if m), None)
        if not spec:
            print(f"  ?  {url}\n     can't detect Greenhouse/Lever/Ashby in that URL. Open the company's job listing: if its "
                  f"address contains boards.greenhouse.io/<slug>, jobs.lever.co/<slug> or jobs.ashbyhq.com/<slug>, use that.\n"
                  f"     (Workday, Darwinbox, Keka, Zoho etc. aren't supported - paste those JDs into --jd-dir.)")
            continue
        kind, slug = spec.split(":", 1)
        try:
            n = len(list(BOARD_FNS[kind](slug)))
        except Exception as e:  # noqa: BLE001
            print(f"  x  {spec}: doesn't respond ({getattr(e, 'code', e)}) - not added")
            continue
        if spec in existing:
            print(f"  =  {spec}: already in {bp.name} ({n} listings)")
            continue
        with bp.open("a", encoding="utf-8") as f:
            f.write(spec + "\n")
        existing.add(spec)
        print(f"  +  {spec}: OK, {n} listings - added to {bp.name}")


# --------------------------------------------------------------------------
# stage 3: score
# --------------------------------------------------------------------------
def jd_skill_weights(text: str) -> dict[str, float]:
    """Weight each skill the JD mentions: more mentions -> heavier; nice-to-have sections -> half."""
    req: Counter = Counter()
    nice: Counter = Counter()
    in_nice = False
    for line in text.splitlines():
        s = line.strip()
        if not s:
            continue
        if len(s) < 70 and not BULLET_RE.match(s) and (s.endswith(":") or s.istitle() or s.isupper() or len(s.split()) <= 5):
            in_nice = bool(NICE_RE.search(s))
        for canon, rx in SKILL_RE.items():
            n = len(rx.findall(s))
            if n:
                (nice if in_nice else req)[canon] += n
    out = {}
    for canon in set(req) | set(nice):
        total = req[canon] + nice[canon]
        base = 1 + 0.5 * (min(total, 3) - 1)
        out[canon] = round(base * (1.0 if req[canon] > 0 else 0.5), 3)
    return out


def build_idf(docs: list[list[str]]) -> dict[str, float]:
    n = len(docs)
    df: Counter = Counter()
    for d in docs:
        df.update(set(d))
    return {t: math.log((1 + n) / (1 + c)) + 1 for t, c in df.items()}


def vec(toks: list[str], idf: dict[str, float]) -> dict[str, float]:
    tf = Counter(toks)
    return {t: (1 + math.log(c)) * idf.get(t, 1.0) for t, c in tf.items()}


def cosine(a: dict[str, float], b: dict[str, float]) -> float:
    if not a or not b:
        return 0.0
    dot = sum(v * b.get(t, 0.0) for t, v in a.items())
    na = math.sqrt(sum(v * v for v in a.values()))
    nb = math.sqrt(sum(v * v for v in b.values()))
    return dot / (na * nb) if na and nb else 0.0


def title_score(roles: list[str], resume_titles: list[str], job_title: str) -> float:
    jt = norm_title_tokens(job_title)
    best = 0.0
    for cand in list(roles) + list(resume_titles):
        ct = norm_title_tokens(cand)
        if jt and ct:
            best = max(best, 2 * len(jt & ct) / (len(jt) + len(ct)))
    return best


def jd_required_years(text: str):
    for m in YEARS_RE.finditer(text):
        ctx = text[max(0, m.start() - 60): m.end() + 60].lower()
        if "experience" in ctx:
            return float(m.group(1))
    return None


JUNIOR_RE = re.compile(r"\b(?:jr\.?|junior|trainee|graduate|fresher|entry[- ]level)\b", re.I)
STAFF_RE = re.compile(r"\b(?:staff|principal|distinguished|fellow|head of|director|vp|vice president)\b", re.I)
SENIOR_RE = re.compile(r"\b(?:senior|sr\.?)\b", re.I)


def level_penalty(title: str, years):
    """Penalise titles clearly below/above your experience level. Returns (points, flag|None)."""
    if not years:
        return 0.0, None
    if JUNIOR_RE.search(title) and years >= 4:
        return 12.0, f"junior-level title; you have {years:g} yrs"
    if STAFF_RE.search(title) and years < 9:
        return 10.0, f"staff/principal-level title (usually 8-10+ yrs); you have {years:g}"
    if SENIOR_RE.search(title) and years < 4:
        return 5.0, f"senior-level title; you have {years:g} yrs"
    return 0.0, None


def combine(skill, cos_s, title, penalty) -> float:
    if skill is None:
        m = 0.7 * cos_s + 0.3 * title
    else:
        m = 0.55 * skill + 0.25 * cos_s + 0.20 * title
    return max(0.0, min(100.0, 100 * m - penalty))


def score_all(resume: dict, jobs: list[dict], role: str, min_match: float, near: float,
              use_level: bool = True) -> list[dict]:
    r_tokens = tokens(resume["text"])
    j_tokens = [tokens(j["description"]) for j in jobs]
    idf = build_idf(j_tokens + [r_tokens])
    rv = vec(r_tokens, idf)
    rskills = set(resume["skills"])
    years = resume.get("years")
    results = []
    for job, jt in zip(jobs, j_tokens):
        weights = jd_skill_weights(job["description"])
        total_w = sum(weights.values())
        got_w = sum(w for s, w in weights.items() if s in rskills)
        skill = (got_w / total_w) if total_w else None
        cos_s = min(1.0, cosine(rv, vec(jt, idf)) / 0.45)
        tscore = title_score(split_roles(role), resume["titles"], job["title"])
        req_years = jd_required_years(job["description"])
        penalty, flags = 0.0, []
        if req_years and years and req_years > years + 0.5:
            penalty = min(15.0, 4.0 * (req_years - years))
            flags.append(f"JD asks ~{req_years:g} yrs experience; your resume shows {years:g}")
        lvl_pen, lvl_flag = level_penalty(job["title"], years) if use_level else (0.0, None)
        if lvl_flag:
            flags.append(lvl_flag)
        penalty = min(25.0, penalty + lvl_pen)
        match = combine(skill, cos_s, tscore, penalty)

        missing = sorted(((s, w) for s, w in weights.items() if s not in rskills), key=lambda x: -x[1])
        missing_out = [{"skill": s, "weight": w, "required": w >= 1.0} for s, w in missing]
        matched = sorted((s for s in weights if s in rskills), key=lambda s: -weights[s])
        surface = [s for s in matched if weights[s] >= 1.5 and resume["skills"].get(s, 0) <= 1]

        potential = match
        top_req = [m["skill"] for m in missing_out if m["required"]][:3]
        if top_req and total_w and match >= min_match - near - 10:  # skip hopeless roles
            got2 = got_w + sum(weights[s] for s in top_req)
            potential = combine(got2 / total_w, cos_s, tscore, penalty)

        jd_tok = set(jt)
        scored_b = []
        for b in resume["bullets"]:
            bs = sum(weights.get(s, 0) for s in count_skills(b)) + 0.1 * len(set(tokens(b)) & jd_tok)
            if bs > 0:
                scored_b.append((bs, b))
        lead = [b for _, b in sorted(scored_b, key=lambda x: -x[0])[:5]]

        label = "match" if match >= min_match else ("near-miss" if match >= min_match - near else "low")
        results.append({
            "id": job["id"], "title": job["title"], "company": job["company"], "location": job["location"],
            "url": job["url"], "source": job["source"], "where": job.get("eligibility", ""),
            "where_reason": job.get("elig_reason", ""), "match": round(match, 1),
            "potential": round(potential, 1), "label": label,
            "components": {"skills": None if skill is None else round(skill, 3), "text": round(cos_s, 3),
                           "title": round(tscore, 3), "exp_penalty": round(penalty, 1)},
            "matched_skills": matched, "missing": missing_out, "surface": surface,
            "lead_with": lead, "flags": flags, "potential_adds": top_req,
        })
    results.sort(key=lambda r: -r["match"])
    for i, r in enumerate(results, 1):
        r["rank"] = i
    return results


WHERE_LABEL = {"india": "IN", "remote-worldwide": "WORLD", "relocation": "RELO", "verify": "VERIFY",
               "pasted": "PASTE", "any": "-", "": "-"}


def print_table(results: list[dict], top: int, min_match: float):
    print(f"\n  #  match  if-added  where   fit    job")
    print("  " + "-" * 84)
    for r in results[:top]:
        pot = f"{r['potential']:>5.0f}%" if r["potential"] > r["match"] + 0.5 else "     -"
        tag = {"match": "MATCH", "near-miss": "near", "low": "low"}[r["label"]]
        who = f"{r['title']} @ {r['company'] or r['source']}"
        print(f"{r['rank']:>3}  {r['match']:>4.0f}%  {pot}   {WHERE_LABEL.get(r['where'], '-'):<7} {tag:<6} {who[:46]}")
    n_ok = sum(1 for r in results if r["match"] >= min_match)
    print(f"\n  {n_ok} job(s) at or above {min_match:g}%.  where: IN = India, WORLD = remote & open to India, "
          f"RELO = relocation/visa mentioned,\n  VERIFY = remote but countries not stated.  "
          f"'if-added' = score if you genuinely have the top missing required skills.")
    links = [r for r in results if r["match"] >= min_match and r["url"]][:top]
    if links:
        print("\n  Apply links (>= threshold):")
        for r in links:
            print(f"   {r['rank']:>3}. {r['title']} @ {r['company'] or r['source']}  [{WHERE_LABEL.get(r['where'], '-')}]\n       {r['url']}")


def cmd_score(args) -> list[dict]:
    resume = load_json(Path(args.work) / "resume.json", "parse")
    jobs = load_json(Path(args.work) / "jobs.json", "fetch")
    if not jobs:
        die("jobs.json is empty - nothing to score.")
    results = score_all(resume, jobs, args.role, args.min_match, args.near, not args.no_level_penalty)
    save_json(Path(args.work) / "matches.json", results)
    print_table(results, args.top, args.min_match)
    return results


# --------------------------------------------------------------------------
# stage 4: report (+ Excel export + optional local LLM notes)
# --------------------------------------------------------------------------
def ollama_notes(model: str, resume_text: str, job: dict):
    prompt = (
        "You are a careful resume reviewer. Compare the RESUME to the JOB.\n"
        "RULES: only suggest edits that rephrase, reorder or emphasise facts ALREADY in the resume. "
        "Never invent skills, tools, employers, titles or metrics. If something is missing, say it is a gap.\n"
        'Reply as JSON: {"summary": "<2 sentences>", "edits": ["<specific edit>", "..."], '
        '"gaps": ["<genuine gaps>"]}\n\n'
        f"RESUME:\n{resume_text[:6000]}\n\nJOB: {job['title']} @ {job['company']}\n{job['description'][:5000]}"
    )
    body = json.dumps({"model": model, "prompt": prompt, "stream": False, "format": "json"}).encode()
    req = urllib.request.Request("http://localhost:11434/api/generate", data=body,
                                 headers={"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=240) as r:
            return json.loads(json.loads(r.read().decode())["response"])
    except Exception as e:  # noqa: BLE001
        print(f"[report] LLM notes failed for '{job['title']}': {e}")
        return None


def export_excel(results: list[dict], out_path: Path) -> None:
    try:
        import openpyxl
        from openpyxl.styles import Font
        from openpyxl.worksheet.datavalidation import DataValidation
    except ImportError:
        print("[report] openpyxl not installed -> skipping Excel export (run: pip install openpyxl)")
        return

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Job Matches"

    headers = [
        "Rank", "Match %", "If-Added %", "Fit", "Where", "Status", "Applied Date",
        "Title", "Company", "Location", "Direct Apply", "Notes",
        "Matched Skills", "Missing (Required)", "Eligibility Note"
    ]
    ws.append(headers)
    bold = Font(bold=True)
    for cell in ws[1]:
        cell.font = bold

    for r in results:
        req_missing = ", ".join(m["skill"] for m in r["missing"] if m["required"])
        url = r.get("url") or ""
        ws.append([
            r["rank"],
            r["match"],
            r["potential"] if r["potential"] > r["match"] + 0.5 else "",
            r["label"],
            WHERE_LABEL.get(r["where"], "-"),
            "To Apply",
            "",
            r["title"],
            r["company"] or r["source"],
            r["location"],
            url,
            "",
            ", ".join(r["matched_skills"]),
            req_missing,
            r.get("where_reason", ""),
        ])
        if url:
            link_cell = ws.cell(row=ws.max_row, column=11)
            link_cell.hyperlink = url
            link_cell.style = "Hyperlink"

    if results:
        dv = DataValidation(
            type="list",
            formula1='"To Apply,Applied,Interview,Rejected,Offer,Skipped"',
            allow_blank=True,
        )
        ws.add_data_validation(dv)
        dv.add(f"F2:F{len(results) + 1}")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    widths = {
        "A": 7, "B": 10, "C": 12, "D": 11, "E": 9, "F": 13, "G": 14,
        "H": 36, "I": 22, "J": 22, "K": 45, "L": 28, "M": 35, "N": 30, "O": 35,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    print(f"[report] wrote {out_path}  ({len(results)} job(s))")


def cmd_report(args) -> None:
    resume = load_json(Path(args.work) / "resume.json", "parse")
    results = load_json(Path(args.work) / "matches.json", "score")
    jobs_by_id = {j["id"]: j for j in load_json(Path(args.work) / "jobs.json", "fetch")}
    pick = {int(x) for x in args.pick.split(",")} if args.pick else None
    floor = args.min_match - args.near
    pool = [r for r in results if r["match"] >= floor]
    detail = [r for r in results if (r["rank"] in pick)] if pick else [r for r in pool]

    lines = [f"# jobfit report - target role: {args.role}", "",
             f"Resume: `{resume['source_file']}` | skills detected: {len(resume['skills'])} | "
             f"experience: {resume['years'] if resume['years'] else 'unknown'} yrs", "",
             f"**{sum(1 for r in results if r['match'] >= args.min_match)} match(es) >= {args.min_match:g}%**, "
             f"{sum(1 for r in results if floor <= r['match'] < args.min_match)} near-miss(es) "
             f"({floor:g}-{args.min_match:g}%).", "",
             "> Suggestions only select, reorder or re-emphasise things already in your resume. "
             "Anything under **Add only if true** must be backed by real experience.", "",
             "## Candidates", "", "| # | Match | If-added | Where | Job (apply link) | Fit |", "|---|---|---|---|---|---|"]
    for r in results[: args.top]:
        pot = f"{r['potential']:.0f}%" if r["potential"] > r["match"] + 0.5 else "-"
        link = f"[{r['title']}]({r['url']})" if r["url"] else r["title"]
        lines.append(f"| {r['rank']} | {r['match']:.0f}% | {pot} | {WHERE_LABEL.get(r['where'], '-')} | {link} - {r['company'] or r['source']} | {r['label']} |")

    gap = Counter()
    for r in pool:
        for m in r["missing"]:
            gap[m["skill"]] += 1
    if gap and pool:
        lines += ["", f"## Biggest gaps across your {len(pool)} target job(s)", "",
                  "Skills these jobs ask for that your resume doesn't show. Add **only** the ones you really have:", ""]
        for s, n in gap.most_common(10):
            lines.append(f"- **{s}** - asked for by {n} of {len(pool)} jobs")

    lines += ["", "## Per-job resume suggestions"]
    for r in detail:
        lines += ["", f"### {r['rank']}. {r['title']} - {r['company'] or r['source']}  ({r['match']:.0f}%, {r['label']})"]
        lines.append("")
        if r["url"]:
            lines.append(f"- **Apply:** {r['url']}")
        if r.get("where_reason"):
            lines.append(f"- **Can you take it from India?** {r['where']} - {r['where_reason']}")
        lines.append(f"- **Already matching:** {', '.join(r['matched_skills'][:12]) or 'none detected'}")
        if r["surface"]:
            lines.append(f"- **Surface more (you have it, JD stresses it):** {', '.join(r['surface'])} - "
                         f"mention in your summary and in a bullet, not only the skills list")
        if r["lead_with"]:
            lines.append("- **Lead with these existing bullets (reorder, no rewrite needed):**")
            lines += [f"    - {b}" for b in r["lead_with"]]
        req = [m["skill"] for m in r["missing"] if m["required"]]
        nice = [m["skill"] for m in r["missing"] if not m["required"]]
        if req or nice:
            lines.append("- **Add only if true:**")
            if req:
                lines.append(f"    - Asked for prominently: {', '.join(req[:8])}")
            if nice:
                lines.append(f"    - Nice-to-have: {', '.join(nice[:6])}")
            if r["potential"] > r["match"] + 0.5:
                lines.append(f"    - Match would go **{r['match']:.0f}% -> {r['potential']:.0f}%** if you genuinely "
                             f"have: {', '.join(r['potential_adds'])}")
        for f in r["flags"]:
            lines.append(f"- **Heads-up:** {f}")
        if args.llm:
            notes = ollama_notes(args.llm, resume["text"], jobs_by_id[r["id"]])
            if notes:
                lines.append(f"- **Local LLM ({args.llm}):** {notes.get('summary', '')}")
                lines += [f"    - {e}" for e in notes.get("edits", [])[:6]]
                if notes.get("gaps"):
                    lines.append(f"    - Gaps it sees: {', '.join(map(str, notes['gaps'][:6]))}")
    out = Path(args.out) / "report.md"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"[report] wrote {out}  ({len(detail)} job(s) with detailed suggestions)")
    export_excel(results, Path(args.out) / "matches.xlsx")


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------
def build_parser() -> argparse.ArgumentParser:
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--work", default="work", help="stage JSON folder (default: ./work)")
    common.add_argument("--out", default="jobfit_out", help="report folder (default: ./jobfit_out)")
    common.add_argument("--role", default="", help='target designation, e.g. "DevOps Engineer"')
    common.add_argument("--min-match", type=float, default=60.0, help="match threshold in %% (default 60)")
    common.add_argument("--near", type=float, default=15.0, help="also show near-misses within N points (default 15)")
    common.add_argument("--years", type=float, default=None, help="your years of experience (overrides detection)")
    common.add_argument("--top", type=int, default=25, help="rows in the candidate table")
    common.add_argument("--pick", default="", help="report detail only for these ranks, e.g. 1,3,5")
    common.add_argument("--llm", default="", help="local Ollama model for extra notes, e.g. qwen2.5:7b")
    common.add_argument("--sources", default="boards,remotive,remoteok,weworkremotely,jobicy,himalayas",
                        help="comma list: boards, remotive, remoteok, weworkremotely, jobicy, himalayas, arbeitnow, greenhouse:<slug>, lever:<slug>, ashby:<slug>, none")
    common.add_argument("--jd-dir", default="", help="folder of .txt/.md job descriptions you pasted in")
    common.add_argument("--location", default="india",
                        help='keep jobs in this country/city (default: india; use "any" for no filter)')
    common.add_argument("--no-remote", action="store_true", help="drop remote jobs outside your country")
    common.add_argument("--no-relocation", action="store_true", help="drop on-site jobs abroad that offer relocation/visa")
    common.add_argument("--strict", action="store_true", help="drop remote jobs whose eligible countries aren't stated")
    common.add_argument("--boards", default=str(Path(__file__).with_name("boards.txt")),
                        help="file of company career boards, one per line: greenhouse:slug | lever:slug | ashby:slug")
    common.add_argument("--no-level-penalty", action="store_true",
                        help="don't penalise junior/staff/principal titles vs your years of experience")
    common.add_argument("--prune", action="store_true", help="with 'sources': remove dead board slugs from boards.txt")
    common.add_argument("--no-preflight", action="store_true")

    ap = argparse.ArgumentParser(prog="jobfit", description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    for name in ("run", "parse"):
        p = sub.add_parser(name, parents=[common])
        p.add_argument("resume", help="resume file (.pdf, .docx, .txt, .md)")
    for name in ("fetch", "score", "report", "sources"):
        sub.add_parser(name, parents=[common])
    p = sub.add_parser("addboard", parents=[common], help="add company career boards by URL")
    p.add_argument("urls", nargs="+", help="careers/job URLs, e.g. https://jobs.lever.co/meesho")
    return ap


def main():
    args = build_parser().parse_args()
    if args.cmd in ("run", "fetch", "score", "report") and not args.role:
        die('--role is required, e.g. --role "DevOps Engineer"')
    if not args.no_preflight:
        preflight(args)
    if args.cmd == "sources":
        cmd_sources(args)
    elif args.cmd == "addboard":
        cmd_addboard(args)
    elif args.cmd == "parse":
        cmd_parse(args)
    elif args.cmd == "fetch":
        cmd_fetch(args)
    elif args.cmd == "score":
        cmd_score(args)
    elif args.cmd == "report":
        cmd_report(args)
    else:
        cmd_parse(args)
        if not cmd_fetch(args):
            sys.exit(1)
        cmd_score(args)
        cmd_report(args)


if __name__ == "__main__":
    main()
