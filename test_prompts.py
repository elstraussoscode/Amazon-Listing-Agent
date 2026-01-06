#!/usr/bin/env python3
"""
Test script for the COSMO prompts - loads prompts from app_production.py
"""
import pandas as pd
from openai import OpenAI
import os
import json
from pydantic import BaseModel, Field
from typing import List
from dotenv import load_dotenv
import re

# Load API key from .env
load_dotenv()
client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))

# Pydantic models (same as production)
class MainContent(BaseModel):
    """Main content: Title, Description, Bullets"""
    model_config = {"extra": "forbid"}
    artikelname: str = Field(description="Produkttitel 150-175 Zeichen. Format: MARKE Produkt Farbe, Eigenschaften. 100% nicht 100 Prozent!")
    produktbeschreibung: str = Field(description="Produktbeschreibung, 1500-1750 Zeichen!")
    bullet_points: List[str] = Field(min_length=5, max_length=5, description="5 Bullets im Format 'HOOK: Beschreibung.' NUR Hook in CAPS!")

class KeywordContent(BaseModel):
    """Keywords only"""
    model_config = {"extra": "forbid"}
    suchbegriffe: str = Field(description="NUR komma-getrennte Keywords! KEINE Sätze!")

def get_byte_length(text: str) -> int:
    return len(text.encode('utf-8'))

# Load test data
df = pd.read_excel('data/test/Top10 Italien Produkte.xlsx')
print(f"Loaded {len(df)} products\n")

# Test with first product
row = df.iloc[0]
product_data_str = "\n".join([f"- {k}: {v}" for k, v in row.to_dict().items() if pd.notna(v)])

print("="*80)
print("PRODUCT DATA:")
print("="*80)
print(product_data_str[:1000])
print("\n")

# Language instruction for Italian
lang_instruction = "Italienisch (IT) - Schreibe in perfektem Italienisch. Amazon.it Konventionen. Verwende italienische Keywords. BPA-frei → senza BPA"

# MAIN CONTENT PROMPT
main_prompt = f"""Erstelle Amazon-optimierten Content für dieses Produkt.

🔤 AUSGABESPRACHE: {lang_instruction}

⛔⛔⛔ ABSOLUT KRITISCH - SPRACHE ⛔⛔⛔
DER GESAMTE OUTPUT MUSS ZU 100% IN DER ZIELSPRACHE SEIN!

EINZIGE AUSNAHMEN (dürfen original bleiben):
- Markennamen: "Dreamfarm", "EMSA", "Philips", "Nike"
- Produktserien: "Aveo", "Scizza"
- Technische Bezeichnungen: "USB-C", "LED"

ALLES ANDERE MUSS ÜBERSETZT WERDEN:
❌ VERBOTEN: "Pizzaschere", "Trinkflasche", "spülmaschinenfest"
✅ ITALIENISCH: "forbici per pizza", "borraccia", "lavabile in lavastoviglie"

📦 PRODUKTDATEN:
{product_data_str}

═══════════════════════════════════════════════════════════════════
📝 TITEL-REGELN:
═══════════════════════════════════════════════════════════════════
1. STRUKTUR: Marke + Produktname + Farbe + Eigenschaften
2. KEINE Sätze! KEINE Punkte am Ende!
3. Komma-getrennt
4. "100%" nicht "100 Prozent"
5. Farbe direkt nach Produktname
6. ZIEL: 170-200 Bytes
7. 100% IN ZIELSPRACHE (außer Markennamen)!

═══════════════════════════════════════════════════════════════════
📝 BULLET POINTS REGELN:
═══════════════════════════════════════════════════════════════════
FORMAT: HOOK IN CAPS: Beschreibender Text in normaler Schreibweise.

1. Hook = 2-4 Wörter in GROSSBUCHSTABEN
2. Nach Hook ein DOPPELPUNKT
3. Danach normaler Text (keine weiteren CAPS!)
4. ZIEL: 170-200 Bytes pro Bullet
5. 100% IN ZIELSPRACHE!

✅ KORREKT: "TAGLIO PRECISO: Le lame in acciaio inossidabile tagliano ogni pizza con precisione."
❌ FALSCH: "TAGLIO PRECISO: Le FORBICI in ACCIAIO..." (zu viele CAPS!)
❌ FALSCH: "Le lame affilate tagliano..." (kein Hook!)

═══════════════════════════════════════════════════════════════════
📝 BESCHREIBUNG:
═══════════════════════════════════════════════════════════════════
- Fließtext, keine Bullets
- ZIEL: 1700-2000 Bytes
- 100% IN ZIELSPRACHE!
"""

print("="*80)
print("CALL 1: Main Content (Title, Bullets, Description)")
print("="*80)

response1 = client.chat.completions.create(
    model="gpt-5.1",
    messages=[
        {"role": "system", "content": f"Amazon SEO-Experte. KRITISCH: Schreibe 100% in ITALIENISCH! Keine deutschen Wörter außer Markennamen! OUTPUT: {lang_instruction}"},
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

print("\n📌 TITEL:")
print(f"   {main_content.artikelname}")
print(f"   [{get_byte_length(main_content.artikelname)} bytes]")

print("\n📌 BULLETS:")
for i, bp in enumerate(main_content.bullet_points, 1):
    print(f"   {i}. {bp}")
    print(f"      [{get_byte_length(bp)} bytes]")

print("\n📌 BESCHREIBUNG (first 300 chars):")
print(f"   {main_content.produktbeschreibung[:300]}...")
print(f"   [{get_byte_length(main_content.produktbeschreibung)} bytes total]")

# KEYWORDS PROMPT
kw_prompt = f"""Generiere Backend-Keywords für dieses Amazon-Produkt.

🎯 OUTPUT LANGUAGE: {lang_instruction}

═══════════════════════════════════════════════════════════════════
📌 BEREITS ERSTELLTER CONTENT (Keywords NICHT wiederholen!):
═══════════════════════════════════════════════════════════════════

TITEL: {main_content.artikelname}

BULLET 1: {main_content.bullet_points[0]}
BULLET 2: {main_content.bullet_points[1]}
BULLET 3: {main_content.bullet_points[2]}
BULLET 4: {main_content.bullet_points[3]}
BULLET 5: {main_content.bullet_points[4]}

BESCHREIBUNG: {main_content.produktbeschreibung[:500]}...

═══════════════════════════════════════════════════════════════════
⚠️ KEYWORD-REGELN (EXTREM STRIKT!):
═══════════════════════════════════════════════════════════════════

FORMAT: keyword1, keyword2, keyword3, keyword4, ...

1. NUR einzelne Wörter oder kurze 2-3 Wort-Kombinationen!
2. Getrennt durch KOMMAS!
3. KEINE ganzen Sätze!
4. KEINE Beschreibungen!
5. Alles kleingeschrieben (außer Markennamen)!
6. ZIEL: 210-249 Bytes

✅ KORREKT: "forbici pizza, taglia pizza, utensile cucina, accessorio pizzeria, lama acciaio"
❌ FALSCH: "Le forbici per pizza Dreamfarm sono perfette per tagliare ogni tipo di pizza."
"""

print("\n" + "="*80)
print("CALL 2: Keywords")
print("="*80)

response2 = client.chat.completions.create(
    model="gpt-5.1",
    messages=[
        {"role": "system", "content": f"Keyword-Spezialist. OUTPUT LANGUAGE: {lang_instruction}. NUR komma-getrennte Keywords! KEINE Sätze! KEINE Beschreibungen!"},
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

print("\n📌 KEYWORDS:")
print(f"   {keyword_content.suchbegriffe}")
print(f"   [{get_byte_length(keyword_content.suchbegriffe)} bytes]")

# Validation Summary
print("\n" + "="*80)
print("VALIDATION SUMMARY")
print("="*80)

issues = []

# Check title
title_bytes = get_byte_length(main_content.artikelname)
if title_bytes < 170:
    issues.append(f"❌ Titel zu kurz: {title_bytes} bytes (min 170)")
elif title_bytes > 200:
    issues.append(f"❌ Titel zu lang: {title_bytes} bytes (max 200)")
else:
    print(f"✅ Titel OK: {title_bytes} bytes")

if main_content.artikelname.endswith('.'):
    issues.append("❌ Titel endet mit Punkt!")

# Check bullets
for i, bp in enumerate(main_content.bullet_points, 1):
    bp_bytes = get_byte_length(bp)
    if bp_bytes < 170:
        issues.append(f"❌ Bullet {i} zu kurz: {bp_bytes} bytes (min 170)")
    elif bp_bytes > 200:
        issues.append(f"❌ Bullet {i} zu lang: {bp_bytes} bytes (max 200)")
    else:
        print(f"✅ Bullet {i} OK: {bp_bytes} bytes")
    
    # Check for HOOK: format
    if ':' not in bp[:50]:
        issues.append(f"❌ Bullet {i} hat kein HOOK: Format!")

# Check description
desc_bytes = get_byte_length(main_content.produktbeschreibung)
if desc_bytes < 1700:
    issues.append(f"❌ Beschreibung zu kurz: {desc_bytes} bytes (min 1700)")
elif desc_bytes > 2000:
    issues.append(f"❌ Beschreibung zu lang: {desc_bytes} bytes (max 2000)")
else:
    print(f"✅ Beschreibung OK: {desc_bytes} bytes")

# Check keywords
kw_bytes = get_byte_length(keyword_content.suchbegriffe)
if kw_bytes < 210:
    issues.append(f"❌ Keywords zu kurz: {kw_bytes} bytes (min 210)")
elif kw_bytes > 249:
    issues.append(f"❌ Keywords zu lang: {kw_bytes} bytes (max 249)")
else:
    print(f"✅ Keywords OK: {kw_bytes} bytes")

# Check if keywords are sentences
if '. ' in keyword_content.suchbegriffe:
    issues.append("❌ Keywords enthalten Sätze (Punkt + Leerzeichen gefunden)!")

if issues:
    print("\n⚠️ ISSUES FOUND:")
    for issue in issues:
        print(f"   {issue}")
else:
    print("\n✅ ALL VALIDATIONS PASSED!")

# ========================================
# CALL 3: Verifikation
# ========================================
print("\n" + "="*80)
print("CALL 3: Verifikation")
print("="*80)

from pydantic import BaseModel, Field

class VerificationResult(BaseModel):
    """Ergebnis der Qualitätsprüfung"""
    model_config = {"extra": "forbid"}
    
    approved: bool = Field(description="True wenn alles korrekt ist")
    artikelname: str = Field(description="Korrigierter Titel oder Original")
    bullet_1: str = Field(description="Korrigierter Bullet 1 oder Original")
    bullet_2: str = Field(description="Korrigierter Bullet 2 oder Original")
    bullet_3: str = Field(description="Korrigierter Bullet 3 oder Original")
    bullet_4: str = Field(description="Korrigierter Bullet 4 oder Original")
    bullet_5: str = Field(description="Korrigierter Bullet 5 oder Original")
    produktbeschreibung: str = Field(description="Korrigierte Beschreibung oder Original")
    suchbegriffe: str = Field(description="Korrigierte Keywords oder Original")
    issues_found: str = Field(description="Gefundene Probleme oder 'Keine'")

verify_prompt = f"""Du bist ein strenger Qualitätsprüfer für Amazon-Listings.

🔤 ZIELSPRACHE: {lang_instruction}

Prüfe den folgenden Content und korrigiere NUR wenn wirklich nötig.
Wenn alles korrekt ist, gib den Original-Content zurück mit approved=True.

TITEL: {main_content.artikelname}

BULLET 1: {main_content.bullet_points[0]}
BULLET 2: {main_content.bullet_points[1]}
BULLET 3: {main_content.bullet_points[2]}
BULLET 4: {main_content.bullet_points[3]}
BULLET 5: {main_content.bullet_points[4]}

BESCHREIBUNG: {main_content.produktbeschreibung[:500]}...

KEYWORDS: {keyword_content.suchbegriffe}

PRÜFKRITERIEN:
1. SPRACHE: Ist ALLES in der Zielsprache (außer Markennamen)?
2. TITEL: Keine Sätze, keine Punkte am Ende, "100%" nicht "100 Prozent"
3. BULLETS: HOOK in CAPS, Doppelpunkt, keine weiteren CAPS im Text
4. KEYWORDS: Komma-getrennt, keine Sätze

Korrigiere NUR echte Fehler! Wenn alles OK: approved=True
"""

response3 = client.chat.completions.create(
    model="gpt-5.1",
    messages=[
        {"role": "system", "content": f"Strenger Qualitätsprüfer. Korrigiere NUR bei echten Fehlern! Zielsprache: {lang_instruction}"},
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

if verification.approved:
    print("\n✅ VERIFIKATION BESTANDEN - Keine Änderungen nötig!")
else:
    print(f"\n⚠️ KORREKTUREN VORGENOMMEN:")
    print(f"   Issues: {verification.issues_found}")
    
    # Zeige Unterschiede
    if verification.artikelname != main_content.artikelname:
        print(f"\n   TITEL korrigiert:")
        print(f"   Vorher: {main_content.artikelname}")
        print(f"   Nachher: {verification.artikelname}")

