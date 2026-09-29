import os
import io
import json
import math
import time
import warnings
import urllib.parse
import re
import requests
import toml
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor, as_completed
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pypdf import PdfReader
from google import genai
from google.genai import types
from google.genai.errors import ServerError, ClientError
from serpapi import GoogleSearch

warnings.filterwarnings("ignore", category=UserWarning)

# --- 1. Load Background Secrets ---
def load_all_secrets():
    secrets_dict = {}
    if os.path.exists(".streamlit/secrets.toml"):
        try:
            with open(".streamlit/secrets.toml", "r", encoding="utf-8") as f:
                secrets_dict = toml.load(f)
        except Exception:
            pass
    return secrets_dict

file_secrets = load_all_secrets()
GEMINI_API_KEY = file_secrets.get("GEMINI_API_KEY", "").strip()
SERPAPI_KEY = file_secrets.get("SERPAPI_KEY", "").strip()
RAPIDAPI_KEY = file_secrets.get("RAPIDAPI_KEY", "").strip()

client = genai.Client(api_key=GEMINI_API_KEY)
MODEL_CASCADE = ["gemini-2.5-flash", "gemini-3.5-flash-lite"]

SYSTEM_PROMPT = """
You are a principal technical resume strategist and senior engineering leader.
Your task: Read the candidate's uploaded resume and dynamically tailor it to align with the Target Job Description (JD) while preserving the exact layout structure of the original resume.

HUMAN-ENGINEERING TONE PRINCIPLES:
- Write authentic, realistic technical sentences using active engineering verbs (e.g., "Architected", "Implemented", "Configured", "Migrated", "Optimized", "Refactored", "Automated", "Resolved", "Standardized").
- Avoid generic AI buzzwords and corporate fluff (e.g., NO "spearheaded cutting-edge synergies", "leveraged holistic paradigms", "game-changing solutions").
- Keep technical details grounded in concrete tools, schemas, queries, pipelines, partitioning, and metrics.

STRICT TAILORING CONSTRAINTS:
1. Professional Summary: Provide an updated professional summary matching the point count of the provided resume.
2. Technical Skill Set: Provide an updated technical skill set grouped into domain clusters matching the JD.
3. Experience:
   - Update ONLY the first three projects from the candidate's provided experience history.
   - Do NOT include external project details.
   - For EACH of the first three projects, provide EXACTLY 15 bullet points.
   - Each bullet point must be at least 1.5 lines long.
   - Include the technical environment within the points without subheadings.
   - Align points directly with the target job description and company background.
4. Output MUST be ONLY a valid JSON object matching the requested schema.
"""

# --- 2. Fast API Generation Wrapper ---

def generate_with_retry_and_fallback(prompt: str, system_instruction: str = None, json_mode: bool = True) -> str:
    config_args = {"temperature": 0.2}
    if system_instruction:
        config_args["system_instruction"] = system_instruction
    if json_mode:
        config_args["response_mime_type"] = "application/json"
    
    config = types.GenerateContentConfig(**config_args)
    last_err = None

    for model_name in MODEL_CASCADE:
        for attempt in range(2):
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=prompt,
                    config=config
                )
                if response and response.text:
                    return response.text
            except (ServerError, ClientError, Exception) as e:
                last_err = e
                time.sleep(1.2 * (attempt + 1))
                continue
    
    raise RuntimeError(f"API generation failed across models. Details: {last_err}")

# --- 3. Dynamic Resume Parser (DOCX & PDF) ---

def extract_text_from_file(uploaded_file) -> str:
    text = ""
    if uploaded_file.name.endswith(".docx"):
        doc = Document(uploaded_file)
        for p in doc.paragraphs:
            if p.text:
                text += p.text + "\n"
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    text += cell.text + " "
                text += "\n"
    elif uploaded_file.name.endswith(".pdf"):
        reader = PdfReader(uploaded_file)
        for page in reader.pages:
            t = page.extract_text()
            if t:
                text += t + "\n"
    return text.strip()

def parse_resume_to_master_data(raw_text: str) -> dict:
    parse_prompt = f"""
    Extract candidate information from this resume into the exact JSON format below:
    
    RESUME TEXT:
    \"\"\"{raw_text}\"\"\"

    Output JSON schema:
    {{
      "basics": {{
        "name": "Candidate Full Name",
        "email": "Email address",
        "phone": "Phone number",
        "linkedin": "LinkedIn URL or empty"
      }},
      "summary_point_count": 5,
      "certifications": ["Certification 1", "Certification 2"],
      "skills": ["Skill 1", "Skill 2"],
      "experience": [
        {{
          "company": "Company Name",
          "role": "Job Title",
          "location": "City, State or Location",
          "period": "Start Date – End Date",
          "environment": ["Tool 1", "Tool 2"],
          "bullet_pool": ["Bullet point 1", "Bullet point 2"]
        }}
      ]
    }}
    """
    raw_response = generate_with_retry_and_fallback(prompt=parse_prompt, json_mode=True)
    return json.loads(raw_response)

# --- 4. Direct Multi-Board Ingestion Engine ---

def classify_job_type(title: str, desc: str, raw_extensions: list = None) -> str:
    ext_str = " ".join(raw_extensions) if raw_extensions else ""
    full_text = f"{title} {desc} {ext_str}".lower()
    contract_signals = [
        "contract", "c2c", "corp to corp", "corp-to-corp", "c2h", 
        "contract-to-hire", "contract to hire", "1099", "temporary", 
        "w2 contract", "freelance", "contractor", "temp"
    ]
    if any(sig in full_text for sig in contract_signals):
        return "Contract / C2C"
    return "Full-Time"

def clean_apply_url(raw_url: str, title: str, platform: str, location: str) -> str:
    encoded_q = urllib.parse.quote(f"{title} {location}")
    if not raw_url or raw_url == "#":
        if "dice" in platform.lower():
            return f"https://www.dice.com/jobs?q={encoded_q}"
        elif "linkedin" in platform.lower():
            return f"https://www.linkedin.com/jobs/search?keywords={encoded_q}"
        elif "indeed" in platform.lower():
            return f"https://www.indeed.com/jobs?q={encoded_q}"
        elif "ziprecruiter" in platform.lower():
            return f"https://www.ziprecruiter.com/candidate/search?search={encoded_q}"
        elif "robert half" in platform.lower():
            return f"https://www.roberthalf.com/us/en/jobs?keywords={encoded_q}"
        elif "monster" in platform.lower():
            return f"https://www.monster.com/jobs/search?q={encoded_q}"
        elif "glassdoor" in platform.lower():
            return f"https://www.glassdoor.com/Job/jobs.htm?sc.keyword={encoded_q}"
        return f"https://www.google.com/search?q={encoded_q}+jobs"

    if "indeed.com" in raw_url:
        jk_match = re.search(r"jk=([a-zA-Z0-9]+)", raw_url)
        if jk_match:
            return f"https://www.indeed.com/viewjob?jk={jk_match.group(1)}"
        return raw_url.split("&")[0]

    return raw_url

# Worker 1: Direct Dice API
def fetch_dice_direct(role: str, location: str, target_filter: str):
    jobs = []
    try:
        url = "https://job-search-api.cloud.dice.com/v1/jobs/search"
        params = {
            "q": role,
            "location": location,
            "pageSize": "40",
            "countryCode": "US"
        }
        if target_filter == "CONTRACT":
            params["filters.employmentType"] = "CONTRACTS"
        elif target_filter == "FULLTIME":
            params["filters.employmentType"] = "FULLTIME"

        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
            "Accept": "application/json"
        }
        resp = requests.get(url, headers=headers, params=params, timeout=4)
        if resp.status_code == 200:
            for item in resp.json().get("data", []):
                title = item.get("title", "Untitled Role")
                comp = item.get("company", "Company on Dice")
                loc_val = item.get("location", location)
                job_id = item.get("id", "")
                snippet = item.get("summary", "")
                posted = item.get("postedDate", "Recently posted")
                apply_link = f"https://www.dice.com/job-detail/{job_id}" if job_id else f"https://www.dice.com/jobs?q={urllib.parse.quote(role)}"
                jtype = "Contract / C2C" if target_filter == "CONTRACT" else ("Full-Time" if target_filter == "FULLTIME" else classify_job_type(title, snippet))

                jobs.append({
                    "title": title,
                    "company": comp,
                    "location": loc_val,
                    "job_type": jtype,
                    "posted_date": posted,
                    "source_engine": "Dice Direct API",
                    "platform": "Dice",
                    "job_url": apply_link,
                    "description": snippet if snippet else f"Opening for {title} at {comp}."
                })
    except Exception:
        pass
    return jobs

# Worker 2: Direct LinkedIn Public Guest Feed
def fetch_linkedin_direct(role: str, location: str, target_filter: str):
    jobs = []
    try:
        q = urllib.parse.quote(role)
        loc = urllib.parse.quote(location)
        jt = "&f_JT=C" if target_filter == "CONTRACT" else ("&f_JT=F" if target_filter == "FULLTIME" else "")
        url = f"https://www.linkedin.com/jobs-guest/jobs/api/seeMoreJobPostings/search?keywords={q}&location={loc}&f_TPR=r86400{jt}&start=0"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
        }
        resp = requests.get(url, headers=headers, timeout=4)
        if resp.status_code == 200:
            soup = BeautifulSoup(resp.text, "html.parser")
            for card in soup.find_all("li")[:25]:
                t_el = card.find("h3", class_="base-search-card__title")
                c_el = card.find("h4", class_="base-search-card__subtitle")
                l_el = card.find("span", class_="job-search-card__location")
                a_el = card.find("a", class_="base-card__full-link")
                tm_el = card.find("time")

                if t_el and c_el:
                    title = t_el.get_text(strip=True)
                    comp = c_el.get_text(strip=True)
                    loc_val = l_el.get_text(strip=True) if l_el else location
                    raw_link = a_el["href"].split("?")[0] if a_el and "href" in a_el.attrs else "#"
                    valid_link = clean_apply_url(raw_link, title, "LinkedIn", loc_val)
                    posted = tm_el.get_text(strip=True) if tm_el else "Past 24 hours"
                    jtype = "Contract / C2C" if target_filter == "CONTRACT" else ("Full-Time" if target_filter == "FULLTIME" else classify_job_type(title, ""))

                    jobs.append({
                        "title": title,
                        "company": comp,
                        "location": loc_val,
                        "job_type": jtype,
                        "posted_date": posted,
                        "source_engine": "LinkedIn Direct",
                        "platform": "LinkedIn",
                        "job_url": valid_link,
                        "description": f"Role: {title} at {comp} ({loc_val}). Category: {role}."
                    })
    except Exception:
        pass
    return jobs

# Worker 3: Direct Monster Public API
def fetch_monster_direct(role: str, location: str, target_filter: str):
    jobs = []
    try:
        url = "https://services.monster.com/search/v1/jobs/search"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
            "Accept": "application/json",
            "Content-Type": "application/json"
        }
        payload = {
            "query": role,
            "locations": [{"address": location}],
            "jobPostingTime": "P1D",
            "pageSize": 25
        }
        if target_filter == "CONTRACT":
            payload["employmentType"] = ["CONTRACTOR", "TEMPORARY"]

        resp = requests.post(url, headers=headers, json=payload, timeout=4)
        if resp.status_code == 200:
            for item in resp.json().get("jobResults", []):
                title = item.get("title", "Untitled Role")
                comp = item.get("company", {}).get("name", "Company on Monster")
                job_id = item.get("jobId", "")
                snippet = item.get("description", "")
                posted = item.get("datePosted", "Recently posted")
                apply_link = f"https://www.monster.com/job-openings/{job_id}" if job_id else f"https://www.monster.com/jobs/search?q={urllib.parse.quote(role)}"
                jtype = "Contract / C2C" if target_filter == "CONTRACT" else ("Full-Time" if target_filter == "FULLTIME" else classify_job_type(title, snippet))

                jobs.append({
                    "title": title,
                    "company": comp,
                    "location": location,
                    "job_type": jtype,
                    "posted_date": posted,
                    "source_engine": "Monster Direct API",
                    "platform": "Monster",
                    "job_url": apply_link,
                    "description": snippet if snippet else f"Monster listing for {title} at {comp}."
                })
    except Exception:
        pass
    return jobs

# Worker 4: Board-Specific SerpApi Indexer (Indeed, Glassdoor, Robert Half, ZipRecruiter)
def fetch_google_jobs_board_specific(role: str, location: str, target_filter: str, target_board: str, api_key: str):
    jobs = []
    if not api_key:
        return jobs

    contract_kw = " contract" if target_filter == "CONTRACT" else " full time"
    q_str = f"{role} {target_board} {contract_kw} {location}"

    params = {
        "engine": "google_jobs",
        "q": q_str,
        "api_key": api_key
    }

    try:
        search = GoogleSearch(params)
        res = search.get_dict()
        for j in res.get("jobs_results", [])[:20]:
            apply_opts = j.get("apply_options", [])
            primary_site = apply_opts[0].get("title", target_board) if apply_opts else target_board
            raw_apply_link = apply_opts[0].get("link", "#") if apply_opts else "#"
            desc = j.get("description", "")
            title = j.get("title", "Untitled Role")
            comp = j.get("company_name", "Hiring Employer")
            exts = j.get("detected_extensions", {})
            posted_at = str(exts.get("posted_at", "Recently posted")) if isinstance(exts, dict) else "Recently posted"

            valid_link = clean_apply_url(raw_apply_link, title, primary_site, location)
            jtype = "Contract / C2C" if target_filter == "CONTRACT" else ("Full-Time" if target_filter == "FULLTIME" else classify_job_type(title, desc))

            jobs.append({
                "title": title,
                "company": comp,
                "location": location,
                "job_type": jtype,
                "posted_date": posted_at,
                "source_engine": f"{target_board} Portal",
                "platform": target_board,
                "job_url": valid_link,
                "description": desc if desc else f"Opportunity: {title} at {comp}."
            })
    except Exception:
        pass
    return jobs

# Worker 5: JSearch Multi-Feed Aggregator
def fetch_jsearch_multi(role: str, location: str, target_filter: str, api_key: str):
    jobs = []
    if not api_key:
        return jobs
    try:
        url = "https://jsearch.p.rapidapi.com/search"
        querystring = {
            "query": f"{role} in {location}",
            "page": "1",
            "num_pages": "3"
        }
        if target_filter == "CONTRACT":
            querystring["employment_types"] = "CONTRACTOR"
        elif target_filter == "FULLTIME":
            querystring["employment_types"] = "FULLTIME"

        headers = {
            "X-RapidAPI-Key": api_key,
            "X-RapidAPI-Host": "jsearch.p.rapidapi.com"
        }
        resp = requests.get(url, headers=headers, params=querystring, timeout=5)
        if resp.status_code == 200:
            for item in resp.json().get("data", []):
                title = item.get("job_title", "Untitled Role")
                comp = item.get("employer_name", "Hiring Company")
                desc = item.get("job_description", "")
                publisher = item.get("job_publisher", "Career Board")
                raw_type = str(item.get("job_employment_type", "")).upper()
                raw_link = item.get("job_apply_link", "#")
                loc_val = f"{item.get('job_city', '')}, {item.get('job_state', '')}".strip(", ") or location
                valid_link = clean_apply_url(raw_link, title, publisher, loc_val)
                jtype = "Contract / C2C" if "CONTRACTOR" in raw_type else ("Full-Time" if "FULLTIME" in raw_type else classify_job_type(title, desc))

                jobs.append({
                    "title": title,
                    "company": comp,
                    "location": loc_val,
                    "job_type": jtype,
                    "posted_date": "Recently posted",
                    "source_engine": "JSearch Multi-Portal Feed",
                    "platform": publisher,
                    "job_url": valid_link,
                    "description": desc if desc else f"Role: {title} at {comp}."
                })
    except Exception:
        pass
    return jobs

# High-Speed Cached Aggregator across All Career Boards
@st.cache_data(ttl=900, show_spinner=False)
def aggregate_high_volume_live_jobs(role_query: str, location: str, target_filter: str = "ALL", limit_count: int = 150):
    all_jobs = []
    target_boards = ["Indeed", "Robert Half", "Glassdoor", "ZipRecruiter"]

    with ThreadPoolExecutor(max_workers=8) as executor:
        futures = [
            executor.submit(fetch_dice_direct, role_query, location, target_filter),
            executor.submit(fetch_linkedin_direct, role_query, location, target_filter),
            executor.submit(fetch_monster_direct, role_query, location, target_filter),
            executor.submit(fetch_jsearch_multi, role_query, location, target_filter, RAPIDAPI_KEY)
        ]
        for board in target_boards:
            futures.append(executor.submit(fetch_google_jobs_board_specific, role_query, location, target_filter, board, SERPAPI_KEY))

        for f in as_completed(futures):
            try:
                res = f.result()
                if res:
                    all_jobs.extend(res)
            except Exception:
                continue

    unique_jobs = []
    seen = set()
    for job in all_jobs:
        sig = f"{job['title'].strip().lower()}_{job['company'].strip().lower()}"
        if sig not in seen:
            seen.add(sig)
            unique_jobs.append(job)

    return unique_jobs[:limit_count]

# --- 5. ATS Resume & Outreach Generation ---

def tailor_resume(master_data: dict, jd_text: str) -> dict:
    point_count = master_data.get("summary_point_count", 5)
    prompt = f"""
    TARGET JOB DESCRIPTION:
    \"\"\"{jd_text}\"\"\"

    CANDIDATE PROFILE DATA (READ DYNAMICALLY FROM UPLOADED RESUME):
    {json.dumps(master_data, indent=2)}

    Output JSON schema:
    {{
      "summary_points": [
        "Updated summary point matching exact count ({point_count} points)..."
      ],
      "skills_categorized": {{
        "Core Platforms & Architecture": ["Skill 1", "Skill 2"],
        "Programming & Distributed Engines": ["Skill 1", "Skill 2"],
        "Data Engineering, BI & Analytics": ["Skill 1", "Skill 2"],
        "Data Modeling, Warehousing & Pipelines": ["Skill 1", "Skill 2"],
        "DevOps, CI/CD & Governance": ["Skill 1", "Skill 2"],
        "Domain & Industry Expertise": ["Domain 1", "Domain 2"]
      }},
      "experience": [
        {{
          "company": "Company 1",
          "role": "Role 1",
          "location": "Location 1",
          "period": "Period 1",
          "tailored_bullets": [
            "Bullet 1 (at least 1.5 lines long, includes technical environment inline, no subheadings)...",
            "Bullet 2...", "Bullet 3...", "Bullet 4...", "Bullet 5...",
            "Bullet 6...", "Bullet 7...", "Bullet 8...", "Bullet 9...",
            "Bullet 10...", "Bullet 11...", "Bullet 12...", "Bullet 13...",
            "Bullet 14...", "Bullet 15..."
          ]
        }},
        {{
          "company": "Company 2",
          "role": "Role 2",
          "location": "Location 2",
          "period": "Period 2",
          "tailored_bullets": [
            "15 bullets matching the exact criteria..."
          ]
        }},
        {{
          "company": "Company 3",
          "role": "Role 3",
          "location": "Location 3",
          "period": "Period 3",
          "tailored_bullets": [
            "15 bullets matching the exact criteria..."
          ]
        }}
      ]
    }}
    """
    raw_response = generate_with_retry_and_fallback(
        prompt=prompt,
        system_instruction=SYSTEM_PROMPT,
        json_mode=True
    )
    return json.loads(raw_response)

def generate_cold_outreach(basics: dict, role_title: str, company: str, jd_snippet: str) -> dict:
    prompt = f"""
    You are an expert executive headhunter and cold outreach strategist.
    Generate two high-conversion, professional cold outreach pitches for a senior technical candidate.

    CANDIDATE:
    Name: {basics.get('name', 'Candidate')}
    Email: {basics.get('email', '')}
    Phone: {basics.get('phone', '')}
    LinkedIn: {basics.get('linkedin', '')}

    TARGET ROLE & COMPANY:
    Role: {role_title}
    Company: {company}
    JD Context: {jd_snippet}

    Output JSON schema:
    {{
      "email_subject": "Catchy, professional subject line (under 10 words)",
      "email_body": "3-paragraph concise, confident email pitch highlighting technical fit and measurable delivery without buzzwords.",
      "linkedin_pitch": "Sub-300 character punchy connection note for LinkedIn recruiters."
    }}
    """
    raw = generate_with_retry_and_fallback(prompt=prompt, json_mode=True)
    return json.loads(raw)

def create_word_document(basics: dict, certs: list, tailored: dict) -> io.BytesIO:
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)

    name_p = doc.add_paragraph()
    name_run = name_p.add_run(basics.get("name", "Candidate"))
    name_run.bold = True
    name_run.font.size = Pt(18)
    name_run.font.name = "Calibri"
    name_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    name_p.paragraph_format.space_after = Pt(2)

    contact_parts = [basics.get("email", ""), basics.get("phone", ""), basics.get("linkedin", "")]
    contact_p = doc.add_paragraph(" | ".join([p for p in contact_parts if p]))
    contact_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    contact_p.paragraph_format.space_after = Pt(12)

    def add_section_header(title):
        h = doc.add_paragraph()
        run = h.add_run(title)
        run.bold = True
        run.font.size = Pt(12)
        run.font.name = "Calibri"
        run.font.color.rgb = RGBColor(0, 51, 102)
        h.paragraph_format.space_before = Pt(8)
        h.paragraph_format.space_after = Pt(4)

    add_section_header("PROFESSIONAL SUMMARY")
    for pt in tailored.get("summary_points", []):
        p = doc.add_paragraph(style='List Bullet')
        p.add_run(pt)
        p.paragraph_format.space_after = Pt(2)

    if certs:
        add_section_header("CERTIFICATIONS")
        for cert in certs:
            p = doc.add_paragraph(style='List Bullet')
            p.add_run(cert)
            p.paragraph_format.space_after = Pt(2)

    add_section_header("TECHNICAL SKILLS")
    for category, skill_list in tailored.get("skills_categorized", {}).items():
        p = doc.add_paragraph()
        c_run = p.add_run(f"{category}: ")
        c_run.bold = True
        p.add_run(", ".join(skill_list))
        p.paragraph_format.space_after = Pt(2)

    add_section_header("PROFESSIONAL EXPERIENCE")
    for job in tailored.get("experience", [])[:3]:
        header_p = doc.add_paragraph()
        r_run = header_p.add_run(f"{job.get('role')} | {job.get('company')}, {job.get('location')}")
        r_run.bold = True
        period_p = header_p.add_run(f"  ({job.get('period')})")
        period_p.italic = True
        header_p.paragraph_format.space_before = Pt(6)
        header_p.paragraph_format.space_after = Pt(2)

        for bullet in job.get("tailored_bullets", [])[:15]:
            bp = doc.add_paragraph(style='List Bullet')
            bp.add_run(bullet)
            bp.paragraph_format.space_after = Pt(2)

    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

# --- 6. UI Dashboard & Screen Routing ---

st.set_page_config(page_title="ASK Career Portal | Career Intelligence & ATS Hub", layout="wide", page_icon="*")

if "target_jd" not in st.session_state:
    st.session_state["target_jd"] = ""
if "apply_url" not in st.session_state:
    st.session_state["apply_url"] = ""
if "current_role_name" not in st.session_state:
    st.session_state["current_role_name"] = "Senior Engineer"
if "current_company_name" not in st.session_state:
    st.session_state["current_company_name"] = "Hiring Team"
if "parsed_resume_data" not in st.session_state:
    st.session_state["parsed_resume_data"] = None
if "generated_docx" not in st.session_state:
    st.session_state["generated_docx"] = None
if "generated_filename" not in st.session_state:
    st.session_state["generated_filename"] = "Tailored_Resume.docx"
if "selected_main_tab" not in st.session_state:
    st.session_state["selected_main_tab"] = "💼 Live Market Jobs"
if "auto_execute_tailor" not in st.session_state:
    st.session_state["auto_execute_tailor"] = False

nav_options = ["💼 Live Market Jobs", "📄 ATS Resume Builder", "✉️ Cold Outreach Generator"]
current_tab_index = nav_options.index(st.session_state["selected_main_tab"])

st.markdown("""
<div style="padding: 10px 0px 20px 0px;">
    <h1 style="margin: 0; font-size: 2.2rem; color: #1E3A8A;">* Ask Career portal</h1>
    <p style="color: #64748B; font-size: 1.05rem;">Next-Gen Market Job Aggregation across Dice, LinkedIn, Indeed, Monster, Glassdoor & Robert Half</p>
</div>
""", unsafe_allow_html=True)

selected_nav = st.radio(
    "Navigation Bar",
    nav_options,
    index=current_tab_index,
    horizontal=True,
    label_visibility="collapsed"
)

if selected_nav != st.session_state["selected_main_tab"]:
    st.session_state["selected_main_tab"] = selected_nav
    st.rerun()

def render_job_feed(job_list, search_query_title, key_prefix):
    total_display_items = len(job_list)
    if total_display_items == 0:
        st.info("No postings returned. Try broadening the search terms or location.")
        return

    platform_counts = {}
    for j in job_list:
        p = j.get("platform", "Other")
        platform_counts[p] = platform_counts.get(p, 0) + 1

    pill_html = " ".join([f"<span style='background:#E2E8F0; padding:3px 8px; border-radius:12px; margin-right:6px; font-size:0.85rem;'><b>{k}</b>: {v}</span>" for k, v in platform_counts.items()])
    st.markdown(f"<div style='margin-bottom: 12px;'>{pill_html}</div>", unsafe_allow_html=True)

    pg_key = f"{key_prefix}_current_page"
    sz_key = f"{key_prefix}_page_size"
    if pg_key not in st.session_state:
        st.session_state[pg_key] = 1
    if sz_key not in st.session_state:
        st.session_state[sz_key] = 25

    per_page = st.session_state[sz_key]
    total_pages = max(1, math.ceil(total_display_items / per_page))
    if st.session_state[pg_key] > total_pages:
        st.session_state[pg_key] = total_pages

    curr_p = st.session_state[pg_key]
    start_idx = (curr_p - 1) * per_page
    end_idx = min(start_idx + per_page, total_display_items)
    paginated_jobs = job_list[start_idx:end_idx]

    st.markdown(f"### Found **{total_display_items} Live Opportunities** for **'{search_query_title}'** (Showing {start_idx + 1}–{end_idx})")
    
    for idx, job in enumerate(paginated_jobs):
        global_idx = start_idx + idx
        title = job["title"]
        comp = job["company"]
        loc = job["location"]
        jtype = job["job_type"]
        posted_date = job.get("posted_date", "Recent")
        platform = job["platform"]
        url = job["job_url"]
        desc = job["description"]

        with st.expander(f"📌 **{title}** — {comp} | 📍 {loc} | 🏷️ **[{jtype}]** | 🕒 **{posted_date}** (via {platform})"):
            st.write(desc if desc else "No description text provided.")
            st.markdown("---")
            
            btn_col1, btn_col2, btn_col3 = st.columns(3)
            with btn_col1:
                st.link_button(
                    f"🚀 Open {platform} to Apply",
                    url,
                    use_container_width=True
                )
            with btn_col2:
                if st.button(f"⚡ Tailor Resume", key=f"{key_prefix}_ats_{global_idx}", type="primary", use_container_width=True):
                    st.session_state["target_jd"] = desc if desc else f"Role: {title}\nCompany: {comp}\nLocation: {loc}"
                    st.session_state["apply_url"] = url
                    st.session_state["current_role_name"] = title
                    st.session_state["current_company_name"] = comp
                    st.session_state["selected_main_tab"] = "📄 ATS Resume Builder"
                    st.session_state["auto_execute_tailor"] = True
                    st.rerun()
            with btn_col3:
                if st.button(f"✉️ Draft Outreach", key=f"{key_prefix}_out_{global_idx}", use_container_width=True):
                    st.session_state["target_jd"] = desc if desc else f"Role: {title}\nCompany: {comp}\nLocation: {loc}"
                    st.session_state["current_role_name"] = title
                    st.session_state["current_company_name"] = comp
                    st.session_state["selected_main_tab"] = "✉️ Cold Outreach Generator"
                    st.rerun()

    # Pagination controls
    st.markdown("---")
    pg_col1, pg_col2, pg_col3, pg_col4 = st.columns([1.5, 2, 2, 2.5])

    with pg_col1:
        if st.button("⬅️ Previous", key=f"{key_prefix}_prev", disabled=(curr_p <= 1), use_container_width=True):
            st.session_state[pg_key] = max(1, curr_p - 1)
            st.rerun()

    with pg_col2:
        st.markdown(f"<div style='text-align: center; padding-top: 6px;'><b>Page {curr_p} of {total_pages}</b></div>", unsafe_allow_html=True)

    with pg_col3:
        if st.button("Next ➡️", key=f"{key_prefix}_next", disabled=(curr_p >= total_pages), use_container_width=True):
            st.session_state[pg_key] = min(total_pages, curr_p + 1)
            st.rerun()

    with pg_col4:
        selected_size = st.selectbox(
            "Jobs per page:",
            options=[10, 25, 50, 75, 100],
            index=[10, 25, 50, 75, 100].index(st.session_state[sz_key]),
            key=f"{key_prefix}_size_select"
        )
        if selected_size != st.session_state[sz_key]:
            st.session_state[sz_key] = selected_size
            st.session_state[pg_key] = 1
            st.rerun()

# ==============================================================================
# TAB 1: LIVE MARKET JOBS
# ==============================================================================
if st.session_state["selected_main_tab"] == "💼 Live Market Jobs":
    sub_fte, sub_contract = st.tabs(["🏢 Full Time (FTE)", "📄 Contract / C2C"])

    with sub_fte:
        col_f1, col_f2, col_f3 = st.columns([3, 2, 1.5])
        with col_f1:
            fte_role_input = st.text_input("Role / Skill Target", placeholder="e.g., Data Engineer, Python Developer, BI Analyst...", key="fte_role_in")
        with col_f2:
            fte_loc_input = st.text_input("Location", "United States", key="fte_loc_in")
        with col_f3:
            fte_limit = st.selectbox("Market Batch Size", [50, 100, 150, 200], index=1, key="fte_limit_select")

        if st.button("🔎 Search Full Time Jobs", type="primary", key="fte_search_btn"):
            if not fte_role_input.strip():
                st.warning("Please enter a job title or keyword to search.")
            else:
                with st.spinner(f"Ingesting live full-time openings for '{fte_role_input}' across portals..."):
                    jobs_fte = aggregate_high_volume_live_jobs(fte_role_input, fte_loc_input, target_filter="FULLTIME", limit_count=fte_limit)
                    st.session_state["fte_results"] = jobs_fte
                    st.session_state["fte_search_title"] = fte_role_input

        if "fte_results" in st.session_state:
            render_job_feed(st.session_state["fte_results"], st.session_state.get("fte_search_title", "Full Time"), "fte")

    with sub_contract:
        col_c1, col_c2, col_c3 = st.columns([3, 2, 1.5])
        with col_c1:
            contract_role_input = st.text_input("Contract Role / Tech Stack", placeholder="e.g., Data Engineer, Power BI, SQL DBA, Snowflake Architect...", key="contract_role_in")
        with col_c2:
            contract_loc_input = st.text_input("Location", "United States", key="contract_loc_in")
        with col_c3:
            contract_limit = st.selectbox("Market Batch Size", [50, 100, 150, 200], index=1, key="contract_limit_select")

        if st.button("🔎 Search Contract Jobs", type="primary", key="contract_search_btn"):
            if not contract_role_input.strip():
                st.warning("Please enter a job title or keyword to search.")
            else:
                with st.spinner(f"Ingesting live contract/C2C openings for '{contract_role_input}' across portals..."):
                    jobs_contract = aggregate_high_volume_live_jobs(contract_role_input, contract_loc_input, target_filter="CONTRACT", limit_count=contract_limit)
                    st.session_state["contract_results"] = jobs_contract
                    st.session_state["contract_search_title"] = contract_role_input

        if "contract_results" in st.session_state:
            render_job_feed(st.session_state["contract_results"], st.session_state.get("contract_search_title", "Contract / C2C"), "contract")

# ==============================================================================
# TAB 2: ATS RESUME BUILDER
# ==============================================================================
elif st.session_state["selected_main_tab"] == "📄 ATS Resume Builder":
    st.subheader("📄 Automated ATS Re-Architecting Engine")

    col_up, col_info = st.columns([2, 1])
    with col_up:
        uploaded_doc = st.file_uploader(
            "📤 Upload Candidate Resume (.docx or .pdf)",
            type=["docx", "pdf"]
        )
        if uploaded_doc is not None:
            if "last_uploaded_name" not in st.session_state or st.session_state["last_uploaded_name"] != uploaded_doc.name:
                with st.spinner("Parsing profile structure, technical stacks, and work history..."):
                    raw_text = extract_text_from_file(uploaded_doc)
                    if raw_text:
                        try:
                            parsed_profile = parse_resume_to_master_data(raw_text)
                            st.session_state["parsed_resume_data"] = parsed_profile
                            st.session_state["last_uploaded_name"] = uploaded_doc.name
                            st.success(f"Loaded {len(parsed_profile.get('experience', []))} project roles for {parsed_profile.get('basics', {}).get('name', 'Candidate')}!")
                        except Exception as e:
                            st.error(f"Parsing error: {e}")

    with col_info:
        if st.session_state["parsed_resume_data"]:
            cand_name = st.session_state['parsed_resume_data']['basics'].get('name', 'Candidate')
            st.info(f"✅ Active Candidate: **{cand_name}**")
        else:
            st.info("ℹ️ Upload a resume on the left or system will use **master_profile.json**.")

    jd_input = st.text_area(
        "Target Job Description (JD)",
        value=st.session_state["target_jd"],
        height=220,
        placeholder="Paste full Job Description here..."
    )

    active_profile = None
    if st.session_state["parsed_resume_data"]:
        active_profile = st.session_state["parsed_resume_data"]
    elif os.path.exists("master_profile.json"):
        with open("master_profile.json", "r", encoding="utf-8") as f:
            active_profile = json.load(f)

    manual_tailor_btn = st.button("🚀 Re-Architect Resume & Generate DOCX", type="primary")
    should_tailor = manual_tailor_btn or st.session_state.get("auto_execute_tailor", False)

    if should_tailor:
        st.session_state["auto_execute_tailor"] = False
        if not jd_input.strip():
            st.error("Please paste or transfer a target Job Description first.")
        elif not active_profile:
            st.error("No active profile data found. Please upload a resume.")
        else:
            with st.spinner("Generating summary points, grouped competencies, and 15 project achievements per role..."):
                try:
                    tailored_data = tailor_resume(active_profile, jd_input.strip())
                    docx_stream = create_word_document(
                        basics=active_profile.get("basics", {}),
                        certs=active_profile.get("certifications", []),
                        tailored=tailored_data
                    )
                    st.session_state["generated_docx"] = docx_stream
                    c_name = active_profile.get('basics', {}).get('name', 'Candidate').replace(' ', '_')
                    st.session_state["generated_filename"] = f"{c_name}_Tailored.docx"
                    st.success("Resume tailored successfully with human-engineered phrasing!")
                except Exception as e:
                    st.error(f"Tailoring failed: {e}")

    if st.session_state["generated_docx"]:
        st.markdown("---")
        st.subheader("🎯 Download & Application Actions")
        
        act_col1, act_col2 = st.columns(2)
        with act_col1:
            st.download_button(
                label="📥 Download Tailored Resume (.docx)",
                data=st.session_state["generated_docx"],
                file_name=st.session_state["generated_filename"],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                type="primary"
            )
        with act_col2:
            apply_link = st.session_state["apply_url"]
            if apply_link and apply_link != "#":
                st.link_button(
                    "🚀 Open Official Job Application",
                    apply_link,
                    use_container_width=True
                )
            else:
                st.button("🔗 Direct Apply Link (Open Source Board)", disabled=True, use_container_width=True)

# ==============================================================================
# TAB 3: COLD OUTREACH GENERATOR
# ==============================================================================
elif st.session_state["selected_main_tab"] == "✉️ Cold Outreach Generator":
    st.subheader("✉️ Recruiter Pitch & Cold Outreach Engine")
    st.caption("Draft authentic pitches aimed directly at hiring managers and recruiters.")

    active_profile = st.session_state["parsed_resume_data"]
    if not active_profile and os.path.exists("master_profile.json"):
        with open("master_profile.json", "r", encoding="utf-8") as f:
            active_profile = json.load(f)

    col_t1, col_t2 = st.columns(2)
    with col_t1:
        target_role = st.text_input("Target Job Title", value=st.session_state.get("current_role_name", "Senior Engineer"))
    with col_t2:
        target_company = st.text_input("Target Company / Agency", value=st.session_state.get("current_company_name", "Hiring Team"))

    outreach_jd = st.text_area("Job Summary / Requirements Context", value=st.session_state.get("target_jd", ""), height=150)

    if st.button("⚡ Generate Recruiter Outreach Pitch", type="primary"):
        if not active_profile:
            st.warning("Please upload a candidate resume in the ATS tab or provide master_profile.json to personalize the pitch.")
        else:
            with st.spinner("Drafting custom email pitch and LinkedIn connection note..."):
                try:
                    outreach_data = generate_cold_outreach(
                        basics=active_profile.get("basics", {}),
                        role_title=target_role,
                        company=target_company,
                        jd_snippet=outreach_jd[:800]
                    )
                    st.session_state["outreach_res"] = outreach_data
                    st.success("Pitches generated successfully!")
                except Exception as e:
                    st.error(f"Outreach generation error: {e}")

    if "outreach_res" in st.session_state:
        res = st.session_state["outreach_res"]
        st.markdown("---")
        
        col_em, col_li = st.columns([3, 2])
        with col_em:
            st.markdown("#### 📧 Cold Outreach Email to Recruiter / Hiring Lead")
            st.text_input("Subject Line", value=res.get("email_subject", ""), key="out_sub")
            st.text_area("Email Pitch Body", value=res.get("email_body", ""), height=260, key="out_body")

        with col_li:
            st.markdown("#### 💬 LinkedIn Connection Note (< 300 chars)")
            st.text_area("LinkedIn Message", value=res.get("linkedin_pitch", ""), height=180, key="out_li")
            st.caption("Tip: Send this directly when connecting with recruiters on LinkedIn for this role.")