"""
Amazon Listing Agent - Template Filler (PARKED)
Automatisches Befüllen von Amazon-Templates mit KI-optimierten Inhalten
"""

import streamlit as st
import pandas as pd
import openpyxl
from openai import OpenAI
import io
from typing import Dict, List, Optional, Annotated
from pydantic import StringConstraints
import logging
from pydantic import BaseModel, Field
import tempfile
import os
import re
import unicodedata
from dynamic_template_analyzer import analyze_template, TemplateFormat

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Constrained string type for bullet points (max 160 chars)
BulletPoint = Annotated[str, StringConstraints(max_length=160)]

# Pydantic Models
class ProductContent(BaseModel):
    """AI-generated Amazon listing content"""
    model_config = {"extra": "forbid"}
    
    artikelname: str = Field(max_length=170, description="Produkttitel, MAXIMAL 170 Zeichen!")
    bullet_points: List[BulletPoint] = Field(min_length=5, max_length=5, description="5 Sätze, je MAXIMAL 160 Zeichen!")
    suchbegriffe: str = Field(max_length=210, description="Suchbegriffe, MAXIMAL 210 Zeichen!")

# Default Prompt
DEFAULT_PROMPT = """Erstelle einen optimierten Amazon-Listing für folgendes Produkt:

{{product_data}}

{{poe_data}}

🔤 AUSGABESPRACHE: {{language}}

🎯 DENKE WIE EIN KUNDE! Was will der Käufer WIRKLICH wissen?

🔍 ANALYSIERE DIE PRODUKTDATEN - Finde ALLE EIGENSCHAFTEN:
- Material? (z.B. Edelstahl, Tritan, BPA-frei, Kunststoff, Glas)
- Größe/Maße/Kapazität? (MUSS im Titel wenn vorhanden!)
- Form? (z.B. rechteckig, quadratisch, rund)
- Komponenten? (z.B. mit Deckel, mit Griffen, inkl. Sieb)
- Farbe?
- Besondere Features? (z.B. auslaufsicher, spülmaschinenfest)

📌 TITEL-STRUKTUR (EXAKT DIESE REIHENFOLGE!):
MARKE SERIE PRODUKTART KOMPONENTEN, EIGENSCHAFTEN (USP, Material, Größe, Farbe, Form) Synonym

REGELN FÜR TITEL:
- KEINE Bindestriche (-), NUR Kommata!
- Komponenten IMMER direkt nach Produktbezeichnung (Long-Tail-Keywords!)
- Material MUSS enthalten sein wenn bekannt
- Größe/Maße MUSS enthalten sein wenn bekannt
- Form MUSS enthalten sein wenn bekannt (rechteckig, rund, etc.)

📌 BULLET POINTS - VOLLSTÄNDIGE SÄTZE:
Jeder Bullet Point MUSS ein vollständiger, abgeschlossener Satz sein!

⚠️ EXAKTE ZEICHEN-LIMITS (STRIKT EINHALTEN!):
- Titel: 140-170 ZEICHEN (= 150-200 Bytes wegen Umlauten)
- Bullet Points: Je 130-160 ZEICHEN (= 150-200 Bytes) - NICHT LÄNGER!
- Keywords: 170-210 ZEICHEN (= 200-250 Bytes)

WICHTIG: Sachliche Produktinfos! ERFINDE NICHTS!"""

# Page Config
st.set_page_config(
    page_title="Amazon Template Filler (PARKED)",
    page_icon="📝",
    layout="wide"
)

# Initialize Session State
if 'api_key' not in st.session_state:
    st.session_state.api_key = ""
if 'prompt_template' not in st.session_state:
    st.session_state.prompt_template = DEFAULT_PROMPT

# Title
st.title("📝 Amazon Template Filler (PARKED)")
st.warning("⚠️ Diese Funktion ist vorübergehend geparkt. Bitte nutze die Content-Optimierung App.")

# Sidebar - Configuration
with st.sidebar:
    st.header("⚙️ Konfiguration")
    
    api_key_input = st.text_input(
        "OpenAI API Key",
        value=st.session_state.api_key,
        type="password"
    )
    
    if st.button("💾 API Key speichern"):
        st.session_state.api_key = api_key_input
        st.success("✅ API Key gespeichert!")

# Helper functions
def get_byte_length(text: str) -> int:
    return len(text.encode('utf-8'))

def ensure_minimum_length_with_ai(text: str, min_bytes: int, field_name: str, client_instance, product_context: str = "") -> str:
    current_bytes = get_byte_length(text)
    if current_bytes >= min_bytes:
        return text
    
    expand_prompt = f"""ERWEITERE diesen Text auf MINDESTENS {min_bytes} BYTES.
Aktuell: {current_bytes} Bytes - zu kurz!
Text: {text}
Antworte NUR mit dem erweiterten Text:"""
    
    try:
        resp = client_instance.chat.completions.create(
            model="gpt-5.1",
            messages=[
                {"role": "system", "content": f"Erweitere den Text auf mindestens {min_bytes} Bytes."},
                {"role": "user", "content": expand_prompt}
            ],
            max_completion_tokens=500
        )
        return resp.choices[0].message.content.strip()
    except:
        return text

def ensure_length_with_ai(text: str, max_bytes: int, field_name: str, client_instance, min_bytes: int = 0) -> str:
    current_bytes = get_byte_length(text)
    if current_bytes <= max_bytes:
        return text
    
    target_bytes = int(max_bytes * 0.95)
    target_chars = int(target_bytes / 1.1)
    
    shorten_prompt = f"""KÜRZE diesen Text auf {target_chars} ZEICHEN.
Text: {text}
Antworte NUR mit dem gekürzten Text:"""
    
    try:
        resp = client_instance.chat.completions.create(
            model="gpt-5.1",
            messages=[
                {"role": "system", "content": f"Kürze auf {target_chars} Zeichen."},
                {"role": "user", "content": shorten_prompt}
            ],
            max_completion_tokens=target_chars + 50
        )
        return resp.choices[0].message.content.strip()
    except:
        return text[:int(max_bytes * 0.9)]

# Main Content
st.info("Diese App wurde geparkt. Der Template Filler wird später wieder aktiviert.")
st.markdown("""
### Was diese App tut:
1. Lädt Produktdaten aus Excel
2. Lädt ein Amazon-Template (.xlsm)
3. Erkennt automatisch das Template-Format
4. Generiert KI-optimierte Inhalte
5. Befüllt das Template automatisch

### Status: 🔴 PARKED
""")

