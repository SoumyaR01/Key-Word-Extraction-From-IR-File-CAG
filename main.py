import os
import re
import json
import pandas as pd
from tqdm import tqdm
from dotenv import load_dotenv
from docx import Document
from langchain_groq import ChatGroq
import logging
from logging.handlers import RotatingFileHandler

LOG_FILE = r"C:/Users/Soumy/OneDrive/Desktop/IR/script.log"
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

console_handler = logging.StreamHandler()
console_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

file_handler = RotatingFileHandler(LOG_FILE, maxBytes=5*1024*1024, backupCount=3)
file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

logger.handlers = []
logger.addHandler(console_handler)
logger.addHandler(file_handler)

logger.info("Script started at %s", pd.Timestamp.now(tz='Asia/Kolkata').strftime('%Y-%m-%d %H:%M:%S %Z'))

load_dotenv()

GROQ_API_KEY = os.getenv("GROQ_API_KEY")
if not GROQ_API_KEY:
    logger.error("GROQ_API_KEY not found in .env file. Please set it.")
    raise ValueError("GROQ_API_KEY environment variable is not set.")

# ====== CONFIGURATION ======
MY_DOCS_FOLDER = r"C:/Users/Soumy/OneDrive/Desktop/IR/input"
RESULTS_FILE = r"C:/Users/Soumy/OneDrive/Desktop/IR/results.xlsx"
MODEL = "qwen/qwen3-32b"
#MODEL = "DeepSeek-R1-Distill-Llama-70B"
MAX_TEXT_LENGTH = 5000

if not os.path.exists(MY_DOCS_FOLDER):
    logger.error(f"Directory {MY_DOCS_FOLDER} does not exist.")
    raise FileNotFoundError(f"Directory {MY_DOCS_FOLDER} does not exist.")

# Initialize ChatGroq model
try:
    llm = ChatGroq(model=MODEL)
    logger.info(f"Initialized ChatGroq model: {MODEL}")
except Exception as e:
    logger.error(f"Failed to initialize ChatGroq model: {str(e)}")
    raise

# ---------- UTILITIES ----------
def extract_text_from_docx(doc_path):
    try:
        doc = Document(doc_path)
        lines = [para.text.strip() for para in doc.paragraphs if para.text.strip()]
        extracted_text = "\n".join(lines)
        if len(extracted_text) > MAX_TEXT_LENGTH:
            extracted_text = extracted_text[:MAX_TEXT_LENGTH]
            logger.warning(f"Text from {doc_path} truncated to {MAX_TEXT_LENGTH} characters")
        return extracted_text
    except Exception as e:
        logger.error(f"Error reading {doc_path}: {str(e)}")
        return ""

def clean_location(location):
    if not location:
        return ""
    if "," in location:
        return location.split(",")[-1].strip()
    location = re.sub(r"\b(highway|road|street|main|lane)\b.*", "", location, flags=re.IGNORECASE).strip()
    return location

def analyze_ir_content(doc_text):
    if not doc_text:
        return "Unknown", "Unknown", "Unknown Department", "Unknown", "Unknown"

    prompt = (
"""You are an IR analyst. I will provide you the content of an IR file. Your task is:

1. Identify the following details from the IR file:
   - State = Name of the Indian state (return 'Unknown' if not found).
   - Location = Clean town/city/taluka name (avoid full addresses, return 'Unknown' if not found).
   - Department = Only one overall department, always ending with 'Department' (return 'Unknown Department' if not found).
   - Audit Conducted Year = The **Date of Audit** (must be present in the Scope of Audit section, return 'Unviable' if not found, formats 'DD-MM-YYYY' or 'YYYY-YYYY' when possible).
   - Financial Year = The **Period of Audit / Reporting Period** (must be present in the Headings or Scope of Audit section, return 'Unviable' if not found, formats 'DD-MM-YYYY' or 'YYYY-YYYY' when possible).


Return only a valid JSON object with exactly these 5 keys:

{
  "state": "...",
  "location": "...",
  "department": "...",
  "audit_conducted_year": "...",
  "financial_year": "..."
}
"""

    )

    try:
        response = llm.invoke([
            {"role": "system", "content": prompt},
            {"role": "user", "content": doc_text}
        ])
        content = response.content.strip()
        logger.info("Received raw response from Groq API: %s", content)

        try:
            parsed = json.loads(content)
            state = parsed.get("state", "Unknown").strip()
            location = clean_location(parsed.get("location", "Unknown").strip())
            department = parsed.get("department", "Unknown Department").strip()
            audit_year = parsed.get("audit_conducted_year", "Unknown").strip()
            financial_year = parsed.get("financial_year", "Unknown").strip()
            return state, location, department, audit_year, financial_year
        except json.JSONDecodeError:
            json_match = re.search(r'\{.*\}', content, re.DOTALL)
            if json_match:
                parsed = json.loads(json_match.group())
                state = parsed.get("state", "Unknown").strip()
                location = clean_location(parsed.get("location", "Unknown").strip())
                department = parsed.get("department", "Unknown Department").strip()
                audit_year = parsed.get("audit_conducted_year", "Unknown").strip()
                financial_year = parsed.get("financial_year", "Unknown").strip()
                return state, location, department, audit_year, financial_year
            return "Unknown", "Unknown", "Unknown Department", "Unknown", "Unknown"
    except Exception as e:
        logger.error(f"Groq API error: {str(e)}")
        return "Unknown", "Unknown", "Unknown Department", "Unknown", "Unknown"

def main_process():
    rows = []
    docx_files = [f for f in os.listdir(MY_DOCS_FOLDER) if f.lower().endswith('.docx')]
    if not docx_files:
        logger.warning(f"No .docx files found in {MY_DOCS_FOLDER}")
        pd.DataFrame([{
            "Filename": "",
            "State": "",
            "Location": "",
            "Department": "",
            "Audit Conducted Year": "",
            "Financial Year": ""
        }]).to_excel(RESULTS_FILE, index=False, engine='openpyxl')
        return

    logger.info(f"Found {len(docx_files)} .docx files in {MY_DOCS_FOLDER}")

    for root, _, files in os.walk(MY_DOCS_FOLDER):
        for fname in tqdm(files, desc=f"Processing files"):
            if not fname.lower().endswith(".docx"):
                continue

            fpath = os.path.join(root, fname)
            text = extract_text_from_docx(fpath)
            state, location, department, audit_year, financial_year = analyze_ir_content(text)

            rows.append({
                "Filename": fname,
                "State": state,
                "Location": location,
                "Department": department,
                "Audit Conducted Year": audit_year,
                "Financial Year": financial_year
            })

            df = pd.DataFrame(rows)
            df.to_excel(RESULTS_FILE, index=False, engine='openpyxl')

    if not rows:
        pd.DataFrame(columns=["Filename", "State", "Location", "Department", "Audit Conducted Year", "Financial Year"]).to_excel(
            RESULTS_FILE, index=False, engine='openpyxl'
        )

    logger.info(f"✅ Processing complete. Results saved in {RESULTS_FILE}")

# ---------- MAIN ----------
if __name__ == "__main__":
    try:
        if not os.path.exists(RESULTS_FILE):
            pd.DataFrame(columns=["Filename", "State", "Location", "Department", "Audit Conducted Year", "Financial Year"]).to_excel(
                RESULTS_FILE, index=False, engine='openpyxl'
            )
            logger.info(f"Created empty Excel file: {RESULTS_FILE}")

        main_process()
    except Exception as e:
        logger.error(f"Script failed: {str(e)}")
        pd.DataFrame([{
            "Filename": "",
            "State": "",
            "Location": "",
            "Department": "",
            "Audit Conducted Year": "",
            "Financial Year": ""
        }]).to_excel(RESULTS_FILE, index=False, engine='openpyxl')
