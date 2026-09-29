import os
import json
import warnings
from google import genai
from google.genai import types
from jinja2 import Template

# Suppress SDK informational warning logs
warnings.filterwarnings("ignore", category=UserWarning)

client = genai.Client(api_key="AQ.Ab8RN6KAlvmEPy6qclRUPOM4USe6KGLN6WocQlwU0FsXfpWfRg")

SYSTEM_PROMPT = """
You are an expert technical resume engineer and ATS optimization specialist.
Your task is to tailor the candidate's master profile to align with the provided Target Job Description (JD).

STRICT GENERATION RULES:
1. Professional Summary: Generate EXACTLY 12 distinct, impactful bullet points highlighting the candidate's experience, architecture expertise, data engineering skills, and domain knowledge relevant to the JD.
2. Technical Skill Set: Group relevant skills into categorized groups matching the JD.
3. Experience:
   - For EACH project/role present in the master profile, generate EXACTLY 15 detailed bullet points.
   - Expand and tailor bullets from the candidate's bullet pool to emphasize technologies, performance metrics, and architectures required by the JD without fabricating untrue qualifications.
   - Retain the exact Environment list for each role.
4. Output MUST be ONLY a valid JSON object matching the schema.
"""

RESUME_TEMPLATE = """# {{ basics.name }}
{{ basics.email }} | {{ basics.phone }} | [LinkedIn]({{ basics.linkedin }})

---

### Professional Summary
{% for point in tailored.summary_points %}
* {{ point }}
{% endfor %}

---

### Certifications
{% for cert in certifications %}
* {{ cert }}
{% endfor %}

---

### Technical Skill Set
{% for category, skill_list in tailored.skills_categorized.items() %}
**{{ category }}:** {{ skill_list | join(', ') }}
{% endfor %}

---

### Professional Experience
{% for job in tailored.experience %}
**{{ job.role }}** | **{{ job.company }}** — {{ job.location }} _({{ job.period }})_

**Environment:** {{ job.environment | join(', ') }}

**Responsibilities:**
{% for bullet in job.tailored_bullets %}
* {{ bullet }}
{% endfor %}

---
{% endfor %}
"""

def tailor_resume(master_data: dict, jd_text: str) -> dict:
    prompt = f"""
    TARGET JOB DESCRIPTION:
    \"\"\"{jd_text}\"\"\"

    CANDIDATE MASTER PROFILE:
    {json.dumps(master_data, indent=2)}

    Output JSON schema:
    {{
      "summary_points": [
        "Point 1...",
        "Point 2...",
        "Point 3...",
        "Point 4...",
        "Point 5...",
        "Point 6...",
        "Point 7...",
        "Point 8...",
        "Point 9...",
        "Point 10...",
        "Point 11...",
        "Point 12..."
      ],
      "skills_categorized": {{
        "Azure Databricks & Governance": ["Skill 1", "Skill 2"],
        "Cloud Ecosystem": ["Skill 1", "Skill 2"],
        "Languages & Distributed Engines": ["Skill 1", "Skill 2"],
        "ETL/ELT & Data Modeling": ["Skill 1", "Skill 2"],
        "DevOps, Security & CI/CD": ["Skill 1", "Skill 2"],
        "Domains": ["Domain 1", "Domain 2"]
      }},
      "experience": [
        {{
          "company": "...",
          "role": "...",
          "location": "...",
          "period": "...",
          "environment": ["Tool1", "Tool2"],
          "tailored_bullets": [
            "Bullet 1",
            "Bullet 2",
            "Bullet 3",
            "Bullet 4",
            "Bullet 5",
            "Bullet 6",
            "Bullet 7",
            "Bullet 8",
            "Bullet 9",
            "Bullet 10",
            "Bullet 11",
            "Bullet 12",
            "Bullet 13",
            "Bullet 14",
            "Bullet 15"
          ]
        }}
      ]
    }}
    """
    response = client.models.generate_content(
        model="gemini-3.6-flash",
        contents=prompt,
        config=types.GenerateContentConfig(
            system_instruction=SYSTEM_PROMPT,
            response_mime_type="application/json",
            temperature=0.2
        )
    )
    return json.loads(response.text)

def main():
    with open("master_profile.json", "r", encoding="utf-8") as f:
        master_data = json.load(f)

    if not os.path.exists("jd.txt") or os.stat("jd.txt").st_size == 0:
        print("Error: 'jd.txt' is missing or empty. Please create jd.txt and paste a job description into it.")
        return

    with open("jd.txt", "r", encoding="utf-8") as f:
        target_jd = f.read().strip()

    print("Analyzing JD from jd.txt and tailoring resume (12 summary points, categorized skills, 15 bullets per project)...")
    tailored_data = tailor_resume(master_data, target_jd)

    template = Template(RESUME_TEMPLATE)
    rendered = template.render(
        basics=master_data["basics"],
        certifications=master_data.get("certifications", []),
        tailored=tailored_data
    )

    output_path = "tailored_resume.md"
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(rendered)

    print(f"Success! Comprehensive tailored resume generated at: {output_path}")

if __name__ == "__main__":
    main()