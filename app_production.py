"""
Amazon Listing Agent - Content Optimization
KI-optimierte Amazon-Listings basierend auf COSMO & RUFUS
"""

import streamlit as st
import pandas as pd
from openai import OpenAI
import io
from typing import List
import logging
from pydantic import BaseModel, Field

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Pydantic Model for structured output
# NO max_length constraints - let ensure_optimal_length_with_ai handle length enforcement
# This prevents Pydantic from truncating text mid-sentence!
class CosmoOptimizedContent(BaseModel):
    """COSMO/RUFUS optimized Amazon listing content"""
    model_config = {"extra": "forbid"}
    
    artikelname: str = Field(description="Produkttitel, 150-175 Zeichen, KEINE Sätze, KEINE Punkte! Nur Keywords mit Kommata!")
    produktbeschreibung: str = Field(description="Produktbeschreibung, 1500-1750 Zeichen!")
    bullet_points: List[str] = Field(min_length=5, max_length=5, description="5 VOLLSTÄNDIGE Sätze, je 150-175 Zeichen!")
    suchbegriffe: str = Field(description="Komma-getrennte Keywords die NICHT im Titel/Bullets stehen, 180-220 Zeichen!")

# COSMO Prompt
COSMO_PROMPT = """Erstelle ein vollständig COSMO & RUFUS optimiertes Amazon-Listing für folgendes Produkt.

Produktdaten:
{{product_data}}

{{poe_data}}

🔤 AUSGABESPRACHE: {{language}}
⚠️ WICHTIG: Schreibe den GESAMTEN Output (Titel, Bullets, Beschreibung, Keywords) in dieser Sprache!

🎯 DENKE WIE EIN KUNDE! Was will der Käufer WIRKLICH wissen?

🔍 ANALYSIERE DIE PRODUKTDATEN - Finde ALLE EIGENSCHAFTEN:
- Material? (z.B. Edelstahl, Tritan, BPA-frei, Kunststoff, Glas)
- Größe/Maße/Kapazität? (MUSS im Titel wenn vorhanden!)
- Form? (z.B. rechteckig, quadratisch, rund)
- Komponenten? (z.B. mit Deckel, mit Griffen, inkl. Sieb)
- Farbe?
- Besondere Features? (z.B. auslaufsicher, spülmaschinenfest)

✅ COSMO-/RUFUS - 15 SEMANTISCHE BEZIEHUNGSTYPEN (NUR INTERN - NICHT IM OUTPUT!):
1. is - Was ist das Produkt?
2. has_property - Eigenschaften (Farbe, Material, Größe, Form...)
3. has_component - Teile, Zubehör, Bestandteile
4. used_for - Verwendungszweck, Funktion
5. used_in - Situation, Umgebung, Ort der Nutzung
6. used_by - Zielgruppe, Nutzertyp
7. used_with - Kombinationsprodukte
8. made_of - Material, Rohstoffe
9. has_quality - Qualitätsmerkmale
10. has_brand - Marke
11. has_style - Stil, Design
12. targets_audience - Zielgruppe
13. associated_with - Themen
14. has_certification - Zertifikate
15. enables_activity - Ermöglichte Aktivitäten

📌 TITEL-STRUKTUR (EXAKT DIESE REIHENFOLGE!):
MARKE SERIE PRODUKTART KOMPONENTEN, EIGENSCHAFTEN (USP, Material, Größe, Farbe, Form) Synonym

PRIORITÄT IM TITEL (in dieser Reihenfolge!):
1. is - Was ist das? (Produktart)
2. has_component - Komponenten DIREKT nach Produktart! (z.B. "mit Deckel und Griffen")
3. has_property - Eigenschaften (Material, Größe, Farbe, Form)

BEISPIELE:
✅ GUT: "EMSA Clip & Close Vorratsdosen mit Deckel und Griffen, 3er Set rechteckig 1L 2L 3L, Kunststoff BPA-frei spülmaschinenfest, Frischhaltedosen"
✅ GUT: "Vtopmart Aufbewahrungsbox mit Deckel luftdicht, 24er Set rechteckig, BPA-frei Kunststoff, für Mehl Zucker Müsli, Vorratsbehälter"
❌ SCHLECHT: "EMSA – Premium Frischhalte-System für optimale Lebensmittel-Aufbewahrung" (= kein USP, keine Komponenten)

REGELN FÜR TITEL:
- KEINE vollständigen Sätze! KEINE Punkte am Ende!
- Titel = Auflistung von Features/Keywords, getrennt durch Kommata
- Trenne Sinnabschnitte mit Kommata, KEINE Bindestriche als Trennzeichen!
- ABER: Behalte technische Bezeichnungen wie "18/10 Edelstahl", "BPA-frei"
- Komponenten IMMER direkt nach Produktbezeichnung (Long-Tail-Keywords!)
- Material MUSS enthalten sein wenn bekannt
- Größe/Maße MUSS enthalten sein wenn bekannt
- Am Ende: Synonym für Produktgattung

TITEL-BEISPIEL:
✅ GUT: "KÜCHENPROFI Käsereibe mit Kurbel und Trommel, 18/10 Edelstahl, 20 cm, für Parmesan und Hartkäse, spülmaschinenfest, Trommelreibe"
❌ SCHLECHT: "KÜCHENPROFI Käsereibe mit Kurbel. Die robuste Reibe ist langlebig." (= Satz mit Punkt!)

📌 BULLET POINTS - VOLLSTÄNDIGE SÄTZE:
Jeder Bullet Point MUSS ein vollständiger, abgeschlossener Satz sein!
NIEMALS mitten im Satz abbrechen!

1. HAUPTVORTEIL: Der größte Nutzen für den Kunden
2. MATERIAL/QUALITÄT: Material und warum es gut ist
3. ANWENDUNG: Wo und wie wird es benutzt
4. BESONDERHEIT: Was unterscheidet es von anderen Produkten
5. LIEFERUMFANG: Was ist enthalten

📌 PRODUKTBESCHREIBUNG (1500-1700 ZEICHEN):
Ausführliche Beschreibung die ALLE 15 Beziehungstypen inhaltlich abdeckt.
KEINE technischen Begriffe wie "is", "has_property" etc. verwenden!

⚠️ AMAZON ZEICHEN-LIMITS (85-90% AUSNUTZEN!):
- Titel: 150-175 ZEICHEN (Amazon max: 200) → NUTZE VOLL AUS!
- Bullet Points: Je 150-175 ZEICHEN (Amazon max: 200) → NUTZE VOLL AUS!
- Beschreibung: 1500-1750 ZEICHEN (Amazon max: 2000)
- Keywords: 180-220 ZEICHEN (Amazon max: 249 Bytes)

WICHTIG: Nutze die verfügbare Länge MAXIMAL aus! 
Ein kurzer Titel verschenkt SEO-Potenzial!
Umlaute (ä,ö,ü,ß) zählen als 2 Bytes.

🔑 KEYWORDS/SUCHBEGRIFFE - BACKEND SEARCH TERMS (210-249 BYTES!):
FORMAT: Komma-getrennte Liste von Keywords
BEISPIEL: "parmesan reibe, käsehobel, reibemaschine, küchengerät manuell, hartkäse raspel"

WICHTIG - NUR KEYWORDS DIE NICHT IM TEXT STEHEN:
- KEINE Begriffe die bereits im Titel oder Bullets vorkommen!
- Amazon indexiert den sichtbaren Text automatisch
- Backend-Keywords sind für ZUSÄTZLICHE Suchbegriffe!

WAS GEHÖRT REIN:
- Synonyme (z.B. "Käsereibe" im Titel → "Käsehobel, Reibemaschine" in Keywords)
- Schreibvarianten (z.B. "Kaesereibe" ohne Umlaut)
- Long-Tail-Keywords (z.B. "manuell ohne strom", "küchenhelfer hand")
- Relevante Anwendungsfälle die NICHT im Text stehen
- Wenn POE-Daten vorhanden: Nutze Top-Suchbegriffe als Inspiration!

WAS GEHÖRT NICHT REIN:
- Begriffe die schon im Titel/Bullets stehen (Verschwendung!)
- Komplementäre Produkte
- Marken von Wettbewerbern

🚫 VERBOTEN IM OUTPUT:
- NIEMALS "is", "has_property", "used_for" etc. im Text!
- KEINE Verweise auf COSMO, RUFUS oder "Beziehungstypen"
- KEINE Bindestriche im Titel!

✍️ STILISTISCHE REGELN - SCHREIBE WIE EIN MENSCH:
- KEINE künstlichen Wortkombinationen
- KEINE übertriebenen Adjektivketten
- KEINE Marketing-Floskeln ("perfekt für", "ideal für")
- Schreibe SACHLICH und DIREKT
- KONKRETE Fakten statt vage Beschreibungen
- ERFINDE NICHTS was nicht in den Produktdaten steht!
- Behalte technische Bezeichnungen EXAKT: "18/10 Edelstahl", "BPA-frei", "0,5L"
- Titel muss LESBAR sein - nicht nur Keywords aneinanderreihen!

BEISPIEL GUTER TITEL:
"KÜCHENPROFI Käsereibe mit Kurbel und Trommel aus 18/10 Edelstahl, 20 cm, für Parmesan und Hartkäse, spülmaschinenfest"
NICHT: "KÜCHENPROFI Käsereibe Kurbel Trommel 18 10 Edelstahl 20 cm Parmesan Hartkäse Spülmaschinenfest"

WICHTIG: Sachliche Produktinfos! ERFINDE NICHTS!
"""

# Page Config
st.set_page_config(
    page_title="Amazon Content Optimizer",
    page_icon="✍️",
    layout="wide"
)

# Initialize Session State
if 'api_key' not in st.session_state:
    st.session_state.api_key = ""
if 'cosmo_prompt_template' not in st.session_state:
    st.session_state.cosmo_prompt_template = COSMO_PROMPT

# Title
st.title("✍️ Amazon Content Optimizer")
st.markdown("**COSMO & RUFUS optimierte Listings mit KI**")

# Sidebar - Configuration
with st.sidebar:
    st.header("⚙️ Konfiguration")
    
    api_key_input = st.text_input(
        "OpenAI API Key",
        value=st.session_state.api_key,
        type="password",
        help="Ihr OpenAI API Key für GPT-5.1"
    )
    
    if st.button("💾 API Key speichern"):
        st.session_state.api_key = api_key_input
        st.success("✅ API Key gespeichert!")
    
    st.markdown("---")
    if st.session_state.get('api_key', '').strip():
        st.success("✅ API Key konfiguriert")
    else:
        st.error("❌ API Key fehlt")
    
    with st.expander("✏️ Prompt bearbeiten"):
        cosmo_prompt = st.text_area(
            "COSMO Prompt",
            value=st.session_state.cosmo_prompt_template,
            height=300
        )
        if st.button("💾 Prompt speichern"):
            st.session_state.cosmo_prompt_template = cosmo_prompt
            st.success("✅ Prompt gespeichert!")

# Helper functions
def get_byte_length(text: str) -> int:
    return len(text.encode('utf-8'))

def ensure_optimal_length_with_ai(text: str, min_bytes: int, max_bytes: int, field_name: str, client_instance, product_context: str = "", max_retries: int = 5) -> str:
    """
    Ensure text is within optimal byte range (min_bytes to max_bytes).
    Uses LLM to adjust text - NO truncation fallback.
    Retries until text is in correct range or max_retries reached.
    """
    current_text = text
    current_bytes = get_byte_length(current_text)
    
    # Already in optimal range
    if min_bytes <= current_bytes <= max_bytes:
        logger.info(f"✅ {field_name}: {current_bytes} bytes (OK)")
        return current_text
    
    target_bytes = (min_bytes + max_bytes) // 2
    
    for attempt in range(max_retries):
        current_bytes = get_byte_length(current_text)
        
        # Check if now in range
        if min_bytes <= current_bytes <= max_bytes:
            logger.info(f"✅ {field_name}: {current_bytes} bytes (nach {attempt} Anpassungen)")
            return current_text
        
        # Determine action
        if current_bytes < min_bytes:
            action = "ERWEITERE"
            direction = "zu kurz"
            needed = min_bytes - current_bytes
            target_chars = int((target_bytes) / 1.12)
        else:
            action = "KÜRZE"
            direction = "zu lang"
            needed = current_bytes - max_bytes
            target_chars = int((target_bytes) / 1.12)
        
        logger.info(f"⚠️ {field_name} {direction} ({current_bytes} bytes, Ziel: {min_bytes}-{max_bytes}). Versuch {attempt + 1}...")
        
        prompt = f"""{action} diesen Text auf EXAKT {min_bytes}-{max_bytes} BYTES.

AKTUELL: {current_bytes} Bytes ({direction} um {needed} Bytes)
ZIEL: {min_bytes}-{max_bytes} Bytes (ca. {target_chars}-{target_chars+15} Zeichen)

PRODUKTKONTEXT:
{product_context[:1000] if product_context else "Nicht verfügbar"}

STRENGE REGELN:
1. Der Text MUSS zwischen {min_bytes} und {max_bytes} Bytes liegen!
2. Umlaute (ä,ö,ü,ß) = 2 Bytes, Sonderzeichen beachten!
3. VOLLSTÄNDIGE SÄTZE - niemals mitten im Satz abbrechen!
4. Behalte technische Bezeichnungen: "18/10 Edelstahl", "BPA-frei", "0,5L"
5. Der Text muss SINN ergeben und lesbar sein!
6. Behalte die wichtigsten Produktinfos

TEXT ZUM ANPASSEN:
{current_text}

Antworte NUR mit dem angepassten Text:"""
        
        try:
            resp = client_instance.chat.completions.create(
                model="gpt-5.1",
                messages=[
                    {"role": "system", "content": f"Du passt Texte auf exakt {min_bytes}-{max_bytes} Bytes an. Ziel: ~{target_bytes} Bytes. Vollständige, sinnvolle Sätze!"},
                    {"role": "user", "content": prompt}
                ],
                max_completion_tokens=int(max_bytes / 0.9) + 100
            )
            new_text = resp.choices[0].message.content.strip()
            
            # Remove quotes if AI added them
            if new_text.startswith('"') and new_text.endswith('"'):
                new_text = new_text[1:-1]
            if new_text.startswith("'") and new_text.endswith("'"):
                new_text = new_text[1:-1]
            # Remove markdown code blocks if present
            if new_text.startswith('```') and new_text.endswith('```'):
                new_text = new_text[3:-3].strip()
            
            new_bytes = get_byte_length(new_text)
            logger.info(f"  → {current_bytes} → {new_bytes} bytes")
            
            # Update for next iteration
            current_text = new_text
            
        except Exception as e:
            logger.error(f"Error adjusting text (attempt {attempt + 1}): {e}")
            # Continue with current text
    
    # Final check after all retries
    final_bytes = get_byte_length(current_text)
    if min_bytes <= final_bytes <= max_bytes:
        logger.info(f"✅ {field_name}: {final_bytes} bytes (nach {max_retries} Anpassungen)")
    else:
        logger.warning(f"⚠️ {field_name}: {final_bytes} bytes - konnte nicht in Zielbereich {min_bytes}-{max_bytes} gebracht werden!")
    
    return current_text


# Main Content
st.markdown("""
Erstellt **hoch-optimierte Amazon-Listings** basierend auf den 15 semantischen Beziehungstypen von COSMO & RUFUS.

**Output:** Eine Excel-Datei mit:
- Produkt-ID & Alter Titel
- ✅ **Neuer Titel** (Marke, Produktart, Komponenten, Eigenschaften, Synonym)
- ✅ **Neue Beschreibung** (15 Beziehungstypen abgedeckt)
- ✅ **5 Bullet Points** (Verkaufsstark & nutzenorientiert)
- ✅ **Keywords** (Exklusiv - nicht in anderen Feldern enthalten!)
""")

st.markdown("---")

# Language Selection
st.subheader("🔤 Ausgabesprache")
language_options = {
    "Deutsch": "Deutsch - Schreibe alles auf Deutsch mit deutschen Amazon-Konventionen",
    "English (UK)": "English (UK) - Write in British English with UK Amazon conventions",
    "English (US)": "English (US) - Write in American English with US Amazon conventions",
    "Français": "Français - Écrivez en français avec les conventions Amazon françaises",
    "Italiano": "Italiano - Scrivi in italiano con le convenzioni Amazon italiane",
    "Español": "Español - Escribe en español con las convenciones de Amazon España",
    "Nederlands": "Nederlands - Schrijf in het Nederlands met Nederlandse Amazon-conventies",
    "Polski": "Polski - Pisz po polsku z polskimi konwencjami Amazon",
    "Svenska": "Svenska - Skriv på svenska med svenska Amazon-konventioner"
}
selected_language = st.selectbox(
    "Wähle die Ausgabesprache für alle generierten Inhalte:",
    options=list(language_options.keys()),
    index=0,
    help="Die KI wird alle Titel, Beschreibungen, Bullets und Keywords in dieser Sprache erstellen."
)

st.markdown("---")

# POE Data Upload
st.subheader("📊 POE-Daten (Product Opportunity Explorer) - Optional")
st.info("💡 **Tipp:** Lade eine CSV aus dem Amazon POE hoch, um die Keywords mit echten Suchvolumen-Daten zu optimieren!")

poe_file = st.file_uploader(
    "POE-Export CSV hochladen (optional)",
    type=["csv"],
    key="poe_upload",
    help="Die CSV-Datei aus dem Amazon Product Opportunity Explorer mit Suchbegriffen und Suchvolumen."
)

poe_keywords = []
if poe_file:
    try:
        poe_file.seek(0)
        lines = poe_file.read().decode('utf-8').split('\n')
        
        # Find the actual CSV header row (must contain comma AND start with Suchbegriff/Search)
        header_row_idx = 0
        for i, line in enumerate(lines):
            line_lower = line.lower().strip()
            # Must start with "suchbegriff" or contain "suchbegriff," (with comma = CSV header)
            if line_lower.startswith('suchbegriff,') or line_lower.startswith('"suchbegriff"') or \
               'suchbegriff,' in line_lower or 'search term,' in line_lower or 'keyword,' in line_lower:
                header_row_idx = i
                break
        
        logger.info(f"POE Header found at row {header_row_idx}: {lines[header_row_idx][:50]}...")
        
        poe_file.seek(0)
        poe_df = pd.read_csv(poe_file, encoding='utf-8', skiprows=header_row_idx, on_bad_lines='skip')
        
        search_col = None
        volume_col = None
        for col in poe_df.columns:
            col_lower = str(col).lower()
            # Search term column
            if search_col is None and ('suchbegriff' in col_lower or 'search' in col_lower or 'keyword' in col_lower):
                search_col = col
            # Volume column - must contain "volumen" or "volume" but NOT "wachstum" (growth)
            if volume_col is None and ('volumen' in col_lower or 'volume' in col_lower) and 'wachstum' not in col_lower and 'growth' not in col_lower:
                volume_col = col
        
        if search_col:
            poe_data_list = []
            for _, row in poe_df.head(20).iterrows():
                term = str(row[search_col]).strip().strip('"')
                if term and term != 'nan' and len(term) > 2:
                    volume = row[volume_col] if volume_col else "N/A"
                    if isinstance(volume, str):
                        volume = volume.strip('"').replace(',', '')
                    # Format volume with thousands separator
                    try:
                        volume_int = int(float(volume))
                        volume_formatted = f"{volume_int:,}".replace(',', '.')
                    except:
                        volume_formatted = str(volume)
                    poe_data_list.append({"Suchbegriff": term, "Suchvolumen": volume_formatted})
                    poe_keywords.append(term)
            
            st.success(f"✅ {len(poe_keywords)} Suchbegriffe aus POE geladen!")
            
            with st.expander("📋 Top POE-Suchbegriffe", expanded=True):
                # Display as table
                import pandas as pd
                poe_display_df = pd.DataFrame(poe_data_list[:15])
                st.dataframe(poe_display_df, use_container_width=True, hide_index=True)
        else:
            st.warning(f"⚠️ Konnte keine Suchbegriff-Spalte finden. Gefundene Spalten: {list(poe_df.columns)[:5]}")
    except Exception as e:
        st.error(f"❌ Fehler beim Laden der POE-Daten: {e}")
        logger.error(f"POE parse error: {e}", exc_info=True)

st.markdown("---")

# Product Data Upload
st.subheader("📁 Produktdaten")
opt_file = st.file_uploader("Produktdaten Excel hochladen", type=["xlsx", "xls"], key="opt_upload")

if opt_file:
    try:
        df_test = pd.read_excel(opt_file, header=None, nrows=10)
        header_row = 0
        best_score = 0
        for row_idx in range(min(5, len(df_test))):
            row = df_test.iloc[row_idx]
            score = sum(1 for val in row if pd.notna(val) and isinstance(val, str) and len(val) > 2 and not val.replace('.', '').replace('-', '').isdigit())
            if score > best_score:
                best_score = score
                header_row = row_idx
        
        df_opt = pd.read_excel(opt_file, header=header_row)
        st.success(f"✅ {len(df_opt)} Produkte geladen")
        
        with st.expander("📋 Daten-Vorschau"):
            st.dataframe(df_opt.head(), use_container_width=True)
        
        cols = df_opt.columns.tolist()
        
        # Auto-detect ID column
        id_col_idx = 0
        for i, col in enumerate(cols):
            col_lower = col.lower()
            if any(x in col_lower for x in ['sku', 'asin', 'ean', 'gtin', 'artikel-nr', 'artikelnummer', 'product_id']):
                id_col_idx = i
                break
        
        # Auto-detect Title column
        title_col_idx = 0
        title_keywords = ['title', 'titel', 'artikelname', 'item_name', 'produktname', 'product name', 'name']
        for i, col in enumerate(cols):
            col_lower = col.lower()
            if any(x in col_lower for x in title_keywords):
                title_col_idx = i + 1
                break
        
        col1, col2 = st.columns(2)
        with col1:
            id_col = st.selectbox("Spalte für Produkt-ID (SKU/EAN)", cols, index=id_col_idx)
        with col2:
            title_col = st.selectbox("Spalte für Alten Titel (Optional)", ["-"] + cols, index=title_col_idx)
        
        num_products_opt = st.slider("Anzahl Produkte", 1, len(df_opt), min(5, len(df_opt)), key="opt_slider")
        
        if st.button("🚀 Optimierung Starten", type="primary", key="opt_start", use_container_width=True):
            if not st.session_state.get('api_key', '').strip():
                st.error("❌ Bitte API Key in der Seitenleiste eingeben")
            else:
                results = []
                progress_bar = st.progress(0)
                status = st.empty()
                
                client = OpenAI(api_key=st.session_state.api_key)
                
                for idx in range(num_products_opt):
                    status.text(f"✍️ Optimiere Produkt {idx + 1}/{num_products_opt}...")
                    row = df_opt.iloc[idx]
                    
                    product_data_str = "\n".join([f"- {k}: {v}" for k, v in row.to_dict().items() if pd.notna(v)])
                    
                    poe_data_str = ""
                    if poe_keywords:
                        poe_data_str = f"""
📊 POE-DATEN (Top-Suchbegriffe mit hohem Suchvolumen):
{', '.join(poe_keywords[:15])}

NUTZE diese Suchbegriffe als Inspiration für:
- Keywords die NICHT im Titel/Bullets stehen
- Long-Tail Varianten
- Synonyme und verwandte Begriffe
"""
                    
                    prompt = st.session_state.cosmo_prompt_template
                    prompt = prompt.replace("{{product_data}}", product_data_str)
                    prompt = prompt.replace("{{poe_data}}", poe_data_str)
                    lang_instruction = language_options[selected_language]
                    prompt = prompt.replace("{{language}}", lang_instruction)
                    
                    try:
                        response = client.chat.completions.create(
                            model="gpt-5.1",
                            messages=[
                                {"role": "system", "content": f"Amazon SEO-Experte für COSMO & RUFUS. OUTPUT LANGUAGE: {lang_instruction}. Schreibe VOLLSTÄNDIGE Sätze!"},
                                {"role": "user", "content": prompt}
                            ],
                            response_format={
                                "type": "json_schema",
                                "json_schema": {
                                    "name": "cosmo_content",
                                    "schema": CosmoOptimizedContent.model_json_schema(),
                                    "strict": True
                                }
                            },
                            max_completion_tokens=4000  # Enough for full content generation
                        )
                        
                        content = CosmoOptimizedContent.model_validate_json(response.choices[0].message.content)
                        
                        # Amazon Limits (85-90% ausnutzen!):
                        # - Title: 200 chars max → Ziel: 170-190 chars (~185-210 bytes für DE)
                        # - Bullets: 255 chars max (empf. 200) → Ziel: 170-190 chars
                        # - Description: 2000 chars max → Ziel: 1700-1900 chars
                        # - Keywords: 249 bytes max → Ziel: 210-245 bytes
                        
                        # TITEL: 170-200 Bytes (85-100% von 200 chars)
                        titel_bytes = get_byte_length(content.artikelname)
                        if titel_bytes < 170 or titel_bytes > 200:
                            content.artikelname = ensure_optimal_length_with_ai(
                                content.artikelname, 170, 200, "Titel", client, product_data_str)
                        
                        # BULLET POINTS: 170-200 Bytes each (85-100% von 200 chars)
                        new_bullets = []
                        for i, bp in enumerate(content.bullet_points):
                            bp_bytes = get_byte_length(bp)
                            if bp_bytes < 170 or bp_bytes > 200:
                                bp = ensure_optimal_length_with_ai(
                                    bp, 170, 200, f"Bullet {i+1}", client, product_data_str)
                            new_bullets.append(bp)
                        content.bullet_points = new_bullets
                        
                        # BESCHREIBUNG: 1700-2000 Bytes (85-100% von 2000 chars)
                        desc_bytes = get_byte_length(content.produktbeschreibung)
                        if desc_bytes < 1700 or desc_bytes > 2000:
                            content.produktbeschreibung = ensure_optimal_length_with_ai(
                                content.produktbeschreibung, 1700, 2000, "Beschreibung", client, product_data_str)
                        
                        # KEYWORDS: 210-249 Bytes (85-100% von 249 bytes)
                        kw_bytes = get_byte_length(content.suchbegriffe)
                        if kw_bytes < 210 or kw_bytes > 249:
                            content.suchbegriffe = ensure_optimal_length_with_ai(
                                content.suchbegriffe, 210, 249, "Keywords", client, product_data_str)
                        
                        result_row = {
                            "Identifier": row[id_col],
                            "Old Title": row[title_col] if title_col != "-" else "",
                            "New Title": content.artikelname,
                            "New Description": content.produktbeschreibung,
                            "New Keyword": content.suchbegriffe
                        }
                        for i, bp in enumerate(content.bullet_points, 1):
                            result_row[f"New Bullet {i}"] = bp
                        
                        results.append(result_row)
                        
                        with st.expander(f"✅ {row[id_col]}: {content.artikelname[:60]}..."):
                            st.write("**Titel:**", content.artikelname)
                            st.write(f"*({get_byte_length(content.artikelname)} bytes)*")
                            st.write("**Bullets:**")
                            for i, bp in enumerate(content.bullet_points, 1):
                                st.write(f"• {bp} *({get_byte_length(bp)} bytes)*")
                            st.write("**Keywords:**", content.suchbegriffe)
                            st.write(f"*({get_byte_length(content.suchbegriffe)} bytes)*")
                            st.caption("**Beschreibung:** " + content.produktbeschreibung[:100] + "...")
                    
                    except Exception as e:
                        st.error(f"Fehler bei Produkt {idx}: {e}")
                        logger.error(f"Error: {e}", exc_info=True)
                    
                    progress_bar.progress((idx + 1) / num_products_opt)
                
                status.text("✅ Fertig!")
                
                if results:
                    df_result = pd.DataFrame(results)
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_result.to_excel(writer, index=False, sheet_name="Optimized Content")
                        ws = writer.sheets["Optimized Content"]
                        for column_cells in ws.columns:
                            length = max(len(str(cell.value) or "") for cell in column_cells)
                            ws.column_dimensions[column_cells[0].column_letter].width = min(length + 2, 50)
                    
                    output.seek(0)
                    
                    st.download_button(
                        "📥 Optimierte Daten herunterladen (XLSX)",
                        data=output.getvalue(),
                        file_name=f"cosmo_optimized_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True
                    )
                
    except Exception as e:
        st.error(f"Fehler beim Lesen der Datei: {e}")
        logger.error(f"File read error: {e}", exc_info=True)
