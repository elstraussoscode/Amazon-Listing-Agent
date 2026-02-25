#!/usr/bin/env python3
"""Test script - runs the content optimization pipeline against test data"""

import pandas as pd
import os
import logging
import sys
from typing import List
from pydantic import BaseModel, Field
from openai import OpenAI
from dotenv import load_dotenv

load_dotenv()

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class MainContent(BaseModel):
    model_config = {"extra": "forbid"}
    artikelname: str = Field(description="Produkttitel 170-200 BYTES")
    produktbeschreibung: str = Field(description="Produktbeschreibung 1700-2000 BYTES")
    bullet_points: List[str] = Field(min_length=5, max_length=5, description="5 Bullets je 170-200 BYTES. Format 'HOOK: Beschreibung.' NUR Hook in CAPS!")

class KeywordContent(BaseModel):
    model_config = {"extra": "forbid"}
    suchbegriffe: str = Field(description="210-249 BYTES. NUR komma-getrennte Keywords!")

class VerificationResult(BaseModel):
    model_config = {"extra": "forbid"}
    approved: bool = Field(description="True wenn alles korrekt")
    artikelname: str = Field(description="Korrigierter Titel oder Original")
    bullet_1: str = Field(description="Korrigierter Bullet 1 oder Original")
    bullet_2: str = Field(description="Korrigierter Bullet 2 oder Original")
    bullet_3: str = Field(description="Korrigierter Bullet 3 oder Original")
    bullet_4: str = Field(description="Korrigierter Bullet 4 oder Original")
    bullet_5: str = Field(description="Korrigierter Bullet 5 oder Original")
    produktbeschreibung: str = Field(description="Korrigierte Beschreibung oder Original")
    suchbegriffe: str = Field(description="Korrigierte Keywords oder Original")
    issues_found: str = Field(description="Gefundene Probleme oder 'Keine'")

def get_byte_length(text: str) -> int:
    return len(text.encode('utf-8'))

# Import prompts from app_production
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from app_production import MAIN_CONTENT_PROMPT, KEYWORD_PROMPT, VERIFICATION_PROMPT

api_key = os.getenv("OPENAI_API_KEY")
if not api_key:
    print("ERROR: Set OPENAI_API_KEY in .env or environment")
    sys.exit(1)

client = OpenAI(api_key=api_key)
MODEL = "gpt-5.1"
LANG = "Italiano - Scrivi in italiano con le convenzioni Amazon italiane"

test_file = "data/test/Top10 Italien Produkte.xlsx"
print(f"\nLoading test data from: {test_file}")

df = pd.read_excel(test_file)
print(f"Loaded {len(df)} products")
print(f"Columns: {list(df.columns)}")
print(f"\nFirst product preview:")
for col in df.columns:
    val = df.iloc[0][col]
    if pd.notna(val):
        print(f"  {col}: {str(val)[:100]}")

product_data_str = "\n".join([f"- {k}: {v}" for k, v in df.iloc[0].to_dict().items() if pd.notna(v)])

print("\n" + "="*70)
print("CALL 1: Main Content (Title, Description, Bullets)")
print("="*70)

main_prompt = MAIN_CONTENT_PROMPT.replace("{{product_data}}", product_data_str).replace("{{language}}", LANG)

response1 = client.chat.completions.create(
    model=MODEL,
    messages=[
        {"role": "system", "content": f"Amazon SEO-Experte. KRITISCH: Schreibe 100% in der Zielsprache! OUTPUT: {LANG}"},
        {"role": "user", "content": main_prompt}
    ],
    response_format={
        "type": "json_schema",
        "json_schema": {
            "name": "main_content",
            "schema": MainContent.model_json_schema(),
            "strict": True
        }
    },
    max_completion_tokens=4000
)

main_content = MainContent.model_validate_json(response1.choices[0].message.content)

print(f"\nTitel: {main_content.artikelname}")
print(f"  ({get_byte_length(main_content.artikelname)} bytes)")
print()
for i, bp in enumerate(main_content.bullet_points, 1):
    print(f"Bullet {i}: {bp}")
    print(f"  ({get_byte_length(bp)} bytes)")
print()
print(f"Beschreibung: {main_content.produktbeschreibung[:200]}...")
print(f"  ({get_byte_length(main_content.produktbeschreibung)} bytes)")

print("\n" + "="*70)
print("CALL 2: Keywords")
print("="*70)

kw_prompt = KEYWORD_PROMPT.replace("{{language}}", LANG)
kw_prompt = kw_prompt.replace("{{title}}", main_content.artikelname)
for i in range(5):
    bp = main_content.bullet_points[i] if i < len(main_content.bullet_points) else ""
    kw_prompt = kw_prompt.replace(f"{{{{bullet{i+1}}}}}", bp)
kw_prompt = kw_prompt.replace("{{description}}", main_content.produktbeschreibung)
kw_prompt = kw_prompt.replace("{{poe_data}}", "")

response2 = client.chat.completions.create(
    model=MODEL,
    messages=[
        {"role": "system", "content": f"Keyword-Spezialist. OUTPUT LANGUAGE: {LANG}. NUR komma-getrennte Keywords! KEINE Satze!"},
        {"role": "user", "content": kw_prompt}
    ],
    response_format={
        "type": "json_schema",
        "json_schema": {
            "name": "keyword_content",
            "schema": KeywordContent.model_json_schema(),
            "strict": True
        }
    },
    max_completion_tokens=500
)

keyword_content = KeywordContent.model_validate_json(response2.choices[0].message.content)

print(f"\nKeywords: {keyword_content.suchbegriffe}")
print(f"  ({get_byte_length(keyword_content.suchbegriffe)} bytes)")

print("\n" + "="*70)
print("CALL 3: Verification")
print("="*70)

verify_prompt = VERIFICATION_PROMPT.replace("{{language}}", LANG)
verify_prompt = verify_prompt.replace("{{title}}", main_content.artikelname)
for i in range(5):
    bp = main_content.bullet_points[i] if i < len(main_content.bullet_points) else ""
    verify_prompt = verify_prompt.replace(f"{{{{bullet{i+1}}}}}", bp)
verify_prompt = verify_prompt.replace("{{description}}", main_content.produktbeschreibung)
verify_prompt = verify_prompt.replace("{{keywords}}", keyword_content.suchbegriffe)

response3 = client.chat.completions.create(
    model=MODEL,
    messages=[
        {"role": "system", "content": f"Strenger Qualitaetspruefer. Zielsprache: {LANG}"},
        {"role": "user", "content": verify_prompt}
    ],
    response_format={
        "type": "json_schema",
        "json_schema": {
            "name": "verification_result",
            "schema": VerificationResult.model_json_schema(),
            "strict": True
        }
    },
    max_completion_tokens=4000
)

verification = VerificationResult.model_validate_json(response3.choices[0].message.content)

print(f"\nApproved: {verification.approved}")
print(f"Issues: {verification.issues_found}")

final_title = verification.artikelname
final_bullets = [verification.bullet_1, verification.bullet_2, verification.bullet_3, verification.bullet_4, verification.bullet_5]
final_desc = verification.produktbeschreibung
final_kw = verification.suchbegriffe

print("\n" + "="*70)
print("FINAL RESULTS + VALIDATION")
print("="*70)

errors = []

title_bytes = get_byte_length(final_title)
print(f"\nTitel ({title_bytes} bytes): {final_title}")
if title_bytes < 170: errors.append(f"Title too short: {title_bytes} < 170")
if title_bytes > 200: errors.append(f"Title too long: {title_bytes} > 200")
if "." in final_title and not any(x in final_title for x in ["0.", "1.", "2.", "3.", "4.", "5.", "6.", "7.", "8.", "9."]):
    errors.append("Title contains period")

for i, bp in enumerate(final_bullets, 1):
    bp_bytes = get_byte_length(bp)
    print(f"\nBullet {i} ({bp_bytes} bytes): {bp}")
    if bp_bytes < 170: errors.append(f"Bullet {i} too short: {bp_bytes} < 170")
    if bp_bytes > 200: errors.append(f"Bullet {i} too long: {bp_bytes} > 200")
    if ":" not in bp[:60]: errors.append(f"Bullet {i} missing HOOK: pattern")
    hook = bp.split(":")[0] if ":" in bp else ""
    if hook and hook != hook.upper(): errors.append(f"Bullet {i} hook not in CAPS: '{hook}'")
    body = bp.split(":", 1)[1].strip() if ":" in bp else bp
    caps_words = [w for w in body.split() if w == w.upper() and len(w) > 2 and not w.replace("-", "").replace(",", "").replace(".", "").isdigit()]
    if caps_words: errors.append(f"Bullet {i} has extra CAPS in body: {caps_words}")

desc_bytes = get_byte_length(final_desc)
print(f"\nDescription ({desc_bytes} bytes): {final_desc[:200]}...")
if desc_bytes < 1700: errors.append(f"Description too short: {desc_bytes} < 1700")
if desc_bytes > 2000: errors.append(f"Description too long: {desc_bytes} > 2000")

kw_bytes = get_byte_length(final_kw)
print(f"\nKeywords ({kw_bytes} bytes): {final_kw}")
if kw_bytes < 210: errors.append(f"Keywords too short: {kw_bytes} < 210")
if kw_bytes > 249: errors.append(f"Keywords too long: {kw_bytes} > 249")
if "." in final_kw: errors.append("Keywords contain periods (sentence-like)")
kw_parts = [k.strip() for k in final_kw.split(",")]
long_kws = [k for k in kw_parts if len(k.split()) > 3]
if long_kws: errors.append(f"Keywords with >3 words: {long_kws}")

german_words = ["spülmaschinenfest", "auslaufsicher", "Trinkflasche", "Edelstahl", "BPA-frei", "hitzebeständig", "Vorratsdose", "Käsereibe", "Pizzaschere"]
all_text = final_title + " " + " ".join(final_bullets) + " " + final_desc + " " + final_kw
for word in german_words:
    if word.lower() in all_text.lower():
        errors.append(f"German word found in Italian output: '{word}'")

print("\n" + "="*70)
if errors:
    print(f"VALIDATION: {len(errors)} ISSUES FOUND")
    for e in errors:
        print(f"  - {e}")
else:
    print("VALIDATION: ALL CHECKS PASSED")
print("="*70)
