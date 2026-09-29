import toml
import requests
from serpapi import GoogleSearch

try:
    with open(".streamlit/secrets.toml", "r") as f:
        secrets = toml.load(f)
        print("SUCCESS: Loaded .streamlit/secrets.toml")
except Exception as e:
    print(f"ERROR reading secrets: {e}")
    exit()

serp_key = secrets.get("SERPAPI_KEY", "").strip()
rapid_key = secrets.get("RAPIDAPI_KEY", "").strip()

print(f"SERPAPI_KEY status: {'Present' if serp_key else 'Missing'}")
print(f"RAPIDAPI_KEY status: {'Present' if rapid_key else 'Missing'}")
print("-" * 50)

# 1. Test SerpApi
if serp_key:
    print("Testing SerpApi...")
    try:
        search = GoogleSearch({
            "engine": "google_jobs",
            "q": "Data Engineer United States",
            "api_key": serp_key
        })
        res = search.get_dict()
        if "jobs_results" in res:
            print(f"-> SerpApi Success: Found {len(res['jobs_results'])} jobs.")
        elif "error" in res:
            print(f"-> SerpApi Server Error: {res['error']}")
        else:
            print(f"-> SerpApi Response keys: {list(res.keys())}")
    except Exception as e:
        print(f"-> SerpApi Exception: {e}")

# 2. Test JSearch
if rapid_key:
    print("\nTesting JSearch RapidAPI...")
    try:
        url = "https://jsearch.p.rapidapi.com/search"
        headers = {
            "X-RapidAPI-Key": rapid_key,
            "X-RapidAPI-Host": "jsearch.p.rapidapi.com"
        }
        params = {"query": "Data Engineer in USA", "page": "1", "num_pages": "1"}
        resp = requests.get(url, headers=headers, params=params, timeout=10)
        if resp.status_code == 200:
            count = len(resp.json().get("data", []))
            print(f"-> JSearch Success: Found {count} jobs.")
        else:
            print(f"-> JSearch Error ({resp.status_code}): {resp.text}")
    except Exception as e:
        print(f"-> JSearch Exception: {e}")