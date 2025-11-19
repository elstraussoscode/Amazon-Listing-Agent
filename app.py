"""
Amazon Listing Agent - Streamlit App
KI-gestützte Optimierung nach COSMO und RUFUS Richtlinien
"""

import streamlit as st
import pandas as pd
import openpyxl
from openai import OpenAI
import io
from typing import Dict, List, Optional
import json

# Page configuration
st.set_page_config(
    page_title="Amazon Listing Agent",
    page_icon="🛒",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Default COSMO/RUFUS optimized prompt
DEFAULT_PROMPT = """Du bist ein Experte für Amazon SEO und die neuen COSMO/RUFUS KI-Systeme von Amazon.

Erstelle optimierte Amazon-Listing-Inhalte basierend auf den 15 Beziehungstypen von COSMO:

**DIE 15 COSMO-BEZIEHUNGSTYPEN:**
1. **is** - Was ist das? (Art des Produkts)
2. **has_property** - Eigenschaften (Farbe, Material, Größe, Form)
3. **has_component** - Teile/Komponenten (Deckel, Filter, Zubehör)
4. **used_for** - Wofür wird es verwendet? (Backen, Reinigung, Meal-Prep)
5. **used_in** - Situation/Umgebung (Camping, Weihnachten, Büro, Outdoor)
6. **used_by** - Wer verwendet es? (Kinder, Senioren, Hobbyköche, Profis)
7. **used_with** - Mit welchen Produkten zusammen? (Komplementärprodukte)
8. **made_of** - Material (Edelstahl, Glas, BPA-frei)
9. **has_quality** - Qualitative Merkmale (langlebig, energiesparend, leicht zu reinigen)
10. **has_brand** - Marke und deren Bedeutung
11. **has_style** - Stil (skandinavisch, modern, retro, rustikal)
12. **targets_audience** - Zielgruppe (Berufspendler, Fitnessbegeisterte, Familien)
13. **associated_with** - Übergeordnete Themen (Nachhaltigkeit, Gesundheit, Minimalismus)
14. **has_certification** - Zertifikate (TÜV, BPA-frei, Öko-Test, CE, GS)
15. **enables_activity** - Ermöglichte Aktivität (Meal-Prep, Organisation, gesund Kochen)

**PRODUKTDATEN:**
{product_data}

**AUFGABE:**
Erstelle für dieses Produkt:

1. **Optimierten Artikelnamen** (max. 200 Zeichen):
   - Enthält Marke, Produkttyp, wichtigste USPs
   - Natürlich lesbar, keine Keyword-Stuffing
   - Deckt mehrere COSMO-Dimensionen ab (is, has_property, used_for)

2. **5 Aufzählungspunkte** (je max. 250 Zeichen):
   - Jeder Punkt adressiert 2-3 der 15 COSMO-Beziehungstypen
   - Beginne mit dem Nutzen/Vorteil
   - Beantworte die W-Fragen: Wer? Was? Wann? Wo? Wie? Warum?
   - Integriere Keywords organisch
   - Nutze konkrete Anwendungsfälle und Szenarien

3. **Suchbegriffe** (5 Begriffe, durch Komma getrennt):
   - Long-Tail-Keywords
   - Verwendungsszenarien (z.B. "camping kochgeschirr set")
   - Zielgruppen-bezogene Begriffe (z.B. "geschenk für hobbyköche")
   - Problemlösungen (z.B. "plastikfrei lebensmittel aufbewahren")
   - Synonyme und Varianten

**FORMAT:**
Gib die Antwort als JSON zurück:
{
  "artikelname": "...",
  "bullet_points": ["1. ...", "2. ...", "3. ...", "4. ...", "5. ..."],
  "suchbegriffe": "begriff1, begriff2, begriff3, begriff4, begriff5"
}

Achte darauf, dass die Inhalte RUFUS-optimiert sind - sie sollten natürliche Antworten auf Kundenfragen sein!"""

# Initialize session state
if 'api_key' not in st.session_state:
    st.session_state.api_key = ""
if 'prompt_template' not in st.session_state:
    st.session_state.prompt_template = DEFAULT_PROMPT
if 'generated_data' not in st.session_state:
    st.session_state.generated_data = None

def load_excel_sheet(file, sheet_name=None):
    """Load Excel file and return dataframe"""
    try:
        if sheet_name:
            df = pd.read_excel(file, sheet_name=sheet_name)
        else:
            df = pd.read_excel(file)
        return df
    except Exception as e:
        st.error(f"Fehler beim Laden der Datei: {str(e)}")
        return None

def find_vorlage_sheet(file):
    """Find the Vorlage sheet in template file"""
    try:
        xl = pd.ExcelFile(file)
        vorlage_sheets = [s for s in xl.sheet_names if 'vorlage' in s.lower() 
                         and not s.lower().startswith('änderungen')]
        if vorlage_sheets:
            return vorlage_sheets[0]
        return None
    except Exception as e:
        st.error(f"Fehler beim Finden der Vorlage: {str(e)}")
        return None

def find_template_columns(template_df):
    """Find key columns in template (Artikelname, Aufzählungspunkt, Suchbegriffe)"""
    columns = {}
    
    # Search in first 10 rows for headers
    for row_idx in range(min(10, len(template_df))):
        for col_idx, cell_value in enumerate(template_df.iloc[row_idx]):
            if pd.notna(cell_value):
                cell_str = str(cell_value).lower()
                
                if 'artikelname' in cell_str and 'title' not in columns:
                    columns['title'] = col_idx
                    columns['title_row'] = row_idx
                
                if 'aufzählungspunkt' in cell_str:
                    if 'bullet_points' not in columns:
                        columns['bullet_points'] = []
                    columns['bullet_points'].append(col_idx)
                    columns['bp_row'] = row_idx
                
                if 'suchbegriffe' in cell_str and 'search_terms' not in columns:
                    # Take first search terms column
                    if 'search_terms' not in columns:
                        columns['search_terms'] = col_idx
                        columns['search_row'] = row_idx
    
    return columns

def clean_gpt5_response(response_text: str) -> str:
    """Clean GPT-5-mini response to get valid JSON"""
    logger.info(f"Raw response (first 500 chars): {response_text[:500]}")
    
    # Remove reasoning field if present
    response_text = re.sub(r'"Reasoning":\s*\[.*?\],?\s*', '', response_text, flags=re.DOTALL)
    
    # Extract JSON from markdown code blocks
    json_match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', response_text, re.DOTALL)
    if json_match:
        response_text = json_match.group(1)
        logger.info("Extracted JSON from markdown block")
    
    # Remove leading/trailing whitespace and markdown markers
    response_text = response_text.strip()
    if response_text.startswith('```'):
        lines = response_text.split('\n')
        response_text = '\n'.join(lines[1:-1] if lines[-1].strip() == '```' else lines[1:])
        response_text = response_text.strip()
    
    logger.info(f"Cleaned response (first 500 chars): {response_text[:500]}")
    return response_text

def generate_content_with_openai(product_data: Dict, api_key: str, prompt_template: str) -> Optional[Dict]:
    """Generate optimized content using OpenAI"""
    try:
        client = OpenAI(api_key=api_key)
        
        # Format product data nicely
        product_info = "\n".join([f"- {k}: {v}" for k, v in product_data.items() if pd.notna(v)])
        
        prompt = prompt_template.format(product_data=product_info)
        
        logger.info("Sending request to GPT-5-mini...")
        
        response = client.chat.completions.create(
            model="gpt-5-mini",  # GPT-5-mini: Latest, fast, cost-efficient!
            messages=[
                {"role": "system", "content": "Du bist ein Amazon SEO-Experte, spezialisiert auf COSMO und RUFUS Optimierung. Antworte IMMER nur mit gültigem JSON im folgenden Format ohne zusätzlichen Text, Markdown oder Reasoning-Felder:\n{\"artikelname\": \"...\", \"bullet_points\": [\"...\", \"...\", \"...\", \"...\", \"...\"], \"suchbegriffe\": \"..., ..., ..., ..., ...\"}"},
                {"role": "user", "content": prompt}
            ]
            # Note: gpt-5-mini only supports temperature=1 (default)
        )
        
        content = response.choices[0].message.content
        logger.info(f"Received response from API (length: {len(content)})")
        
        # Clean the response (handles GPT-5-mini quirks)
        cleaned_content = clean_gpt5_response(content)
        
        # Try to parse JSON
        try:
            result = json.loads(cleaned_content)
            logger.info("Successfully parsed JSON response")
            return result
            
        except json.JSONDecodeError as je:
            logger.error(f"JSON parsing failed: {str(je)}")
            logger.error(f"Failed content: {cleaned_content[:500]}")
            
            # Last resort: try to fix common issues
            try:
                # Remove trailing commas
                fixed = re.sub(r',\s*}', '}', cleaned_content)
                fixed = re.sub(r',\s*]', ']', fixed)
                result = json.loads(fixed)
                logger.info("Successfully parsed after fixing trailing commas")
                return result
            except:
                st.error(f"JSON Parse Error: {str(je)}\n\nResponse:\n{cleaned_content[:300]}")
                return None
        
    except Exception as e:
        logger.error(f"Error in generate_content_with_openai: {str(e)}", exc_info=True)
        st.error(f"Fehler bei der Content-Generierung: {str(e)}")
        return None

def create_output_dataframe(products_df, generated_contents):
    """Create output dataframe with generated content"""
    output_data = []
    
    for idx, (_, product) in enumerate(products_df.iterrows()):
        if idx < len(generated_contents) and generated_contents[idx]:
            content = generated_contents[idx]
            row = {
                **product.to_dict(),
                'Artikelname_Optimiert': content.get('artikelname', ''),
                'Bullet_Point_1': content.get('bullet_points', [''])[0] if len(content.get('bullet_points', [])) > 0 else '',
                'Bullet_Point_2': content.get('bullet_points', [''])[1] if len(content.get('bullet_points', [])) > 1 else '',
                'Bullet_Point_3': content.get('bullet_points', [''])[2] if len(content.get('bullet_points', [])) > 2 else '',
                'Bullet_Point_4': content.get('bullet_points', [''])[3] if len(content.get('bullet_points', [])) > 3 else '',
                'Bullet_Point_5': content.get('bullet_points', [''])[4] if len(content.get('bullet_points', [])) > 4 else '',
                'Suchbegriffe_Optimiert': content.get('suchbegriffe', '')
            }
            output_data.append(row)
    
    return pd.DataFrame(output_data)

# Main UI
st.title("🛒 Amazon Listing Agent")
st.markdown("**KI-gestützte Content-Optimierung nach COSMO & RUFUS**")

# Create tabs
tab1, tab2 = st.tabs(["📝 Content Generator", "⚙️ Konfiguration"])

with tab1:
    st.header("Content Generator")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("1️⃣ Produktdaten hochladen")
        products_file = st.file_uploader(
            "Excel-Datei mit Produktdaten (1 Produkt pro Zeile)",
            type=['xlsx', 'xls', 'xlsm'],
            key='products'
        )
        
        if products_file:
            products_df = load_excel_sheet(products_file)
            if products_df is not None:
                st.success(f"✅ {len(products_df)} Produkte geladen")
                st.dataframe(products_df.head(), use_container_width=True)
    
    with col2:
        st.subheader("2️⃣ Template hochladen")
        template_file = st.file_uploader(
            "Amazon Upload-Template",
            type=['xlsx', 'xls', 'xlsm'],
            key='template'
        )
        
        if template_file:
            vorlage_sheet = find_vorlage_sheet(template_file)
            if vorlage_sheet:
                st.success(f"✅ Vorlage-Sheet gefunden: {vorlage_sheet}")
                template_df = load_excel_sheet(template_file, vorlage_sheet)
                if template_df is not None:
                    columns = find_template_columns(template_df)
                    st.info(f"🔍 Erkannte Spalten:\n- Artikelname: Spalte {columns.get('title', 'nicht gefunden')}\n" + 
                           f"- Bullet Points: {len(columns.get('bullet_points', []))} Spalten\n" +
                           f"- Suchbegriffe: Spalte {columns.get('search_terms', 'nicht gefunden')}")
    
    st.markdown("---")
    
    # Generation section
    if products_file and st.session_state.get('api_key', '').strip():
        st.subheader("3️⃣ Content generieren")
        
        if st.button("🚀 Content für alle Produkte generieren", type="primary", use_container_width=True):
            if not st.session_state.get('api_key', '').strip():
                st.error("❌ Bitte OpenAI API Key in der Konfiguration eingeben!")
            else:
                products_df = load_excel_sheet(products_file)
                
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                generated_contents = []
                
                for idx, (_, product) in enumerate(products_df.iterrows()):
                    status_text.text(f"Generiere Content für Produkt {idx + 1} von {len(products_df)}...")
                    
                    product_data = product.to_dict()
                    content = generate_content_with_openai(
                        product_data,
                        st.session_state.api_key,
                        st.session_state.prompt_template
                    )
                    generated_contents.append(content)
                    
                    progress_bar.progress((idx + 1) / len(products_df))
                
                status_text.text("✅ Generierung abgeschlossen!")
                
                # Store in session state
                st.session_state.generated_data = create_output_dataframe(products_df, generated_contents)
        
        # Display generated content
        if st.session_state.generated_data is not None:
            st.markdown("---")
            st.subheader("📊 Generierte Inhalte")
            
            st.dataframe(st.session_state.generated_data, use_container_width=True)
            
            # Download buttons
            col1, col2 = st.columns(2)
            
            with col1:
                # Excel download
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    st.session_state.generated_data.to_excel(writer, index=False, sheet_name='Optimierte Produkte')
                output.seek(0)
                
                st.download_button(
                    label="📥 Als Excel herunterladen",
                    data=output,
                    file_name="amazon_listings_optimiert.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            with col2:
                # CSV download
                csv = st.session_state.generated_data.to_csv(index=False)
                st.download_button(
                    label="📥 Als CSV herunterladen",
                    data=csv,
                    file_name="amazon_listings_optimiert.csv",
                    mime="text/csv",
                    use_container_width=True
                )
    elif not st.session_state.get('api_key', '').strip():
        st.warning("⚠️ Bitte OpenAI API Key in der Konfiguration eingeben!")

with tab2:
    st.header("⚙️ Konfiguration")
    
    # API Key section
    st.subheader("🔑 OpenAI API Key")
    st.markdown("Ihr API Key wird nur in dieser Session gespeichert und nicht dauerhaft.")
    
    api_key_input = st.text_input(
        "OpenAI API Key",
        type="password",
        value=st.session_state.api_key,
        help="Geben Sie Ihren OpenAI API Key ein. Dieser wird sicher behandelt und nur für die aktuelle Session gespeichert."
    )
    
    if st.button("💾 API Key speichern"):
        st.session_state.api_key = api_key_input
        st.success("✅ API Key gespeichert!")
    
    if st.session_state.api_key:
        st.success("✅ API Key ist gesetzt")
    else:
        st.error("❌ Kein API Key gesetzt")
    
    st.markdown("---")
    
    # Prompt template section
    st.subheader("📝 Prompt Template")
    st.markdown("Passen Sie den Prompt an Ihre spezifischen Anforderungen an.")
    
    prompt_input = st.text_area(
        "Prompt Template",
        value=st.session_state.prompt_template,
        height=400,
        help="Verwenden Sie {product_data} als Platzhalter für die Produktinformationen."
    )
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("💾 Prompt speichern", use_container_width=True):
            st.session_state.prompt_template = prompt_input
            st.success("✅ Prompt gespeichert!")
    
    with col2:
        if st.button("🔄 Prompt zurücksetzen", use_container_width=True):
            st.session_state.prompt_template = DEFAULT_PROMPT
            st.success("✅ Prompt auf Standard zurückgesetzt!")
            st.rerun()
    
    # Information box
    st.markdown("---")
    st.info("""
    **📚 COSMO & RUFUS Optimierung**
    
    Diese App optimiert Ihre Amazon-Listings nach den neuesten KI-Richtlinien:
    
    - **COSMO**: Amazons Relevanz-Engine mit 15 Beziehungstypen
    - **RUFUS**: Amazons KI-Shopping-Assistent
    
    Die generierten Inhalte decken alle 15 COSMO-Dimensionen ab und sind für
    RUFUS-Dialoge optimiert, sodass Ihre Produkte in KI-gestützten Kaufberatungen
    empfohlen werden können.
    """)
    
    with st.expander("📖 Die 15 COSMO-Beziehungstypen"):
        st.markdown("""
        1. **is** - Was ist das?
        2. **has_property** - Eigenschaften
        3. **has_component** - Teile/Komponenten
        4. **used_for** - Wofür verwendet?
        5. **used_in** - Situation/Umgebung
        6. **used_by** - Wer verwendet es?
        7. **used_with** - Komplementärprodukte
        8. **made_of** - Material
        9. **has_quality** - Qualitätsmerkmale
        10. **has_brand** - Marke
        11. **has_style** - Stil
        12. **targets_audience** - Zielgruppe
        13. **associated_with** - Übergeordnete Themen
        14. **has_certification** - Zertifikate
        15. **enables_activity** - Ermöglichte Aktivität
        """)

# Sidebar
with st.sidebar:
    st.image("https://via.placeholder.com/200x80/FF9900/FFFFFF?text=Amazon", use_container_width=True)
    st.markdown("### Amazon Listing Agent")
    st.markdown("Optimiert für **COSMO** & **RUFUS**")
    
    st.markdown("---")
    
    st.markdown("### 📊 Status")
    if st.session_state.get('api_key', '').strip():
        st.success("✅ API Key gesetzt")
    else:
        st.error("❌ Kein API Key")
    
    if st.session_state.generated_data is not None:
        st.success(f"✅ {len(st.session_state.generated_data)} Produkte generiert")
    else:
        st.info("ℹ️ Noch keine Daten generiert")
    
    st.markdown("---")
    
    st.markdown("### 🚀 Quick Start")
    st.markdown("""
    1. Gehe zu **Konfiguration**
    2. OpenAI API Key eingeben
    3. Zurück zu **Content Generator**
    4. Produktdaten hochladen
    5. Template hochladen (optional)
    6. Auf **Generieren** klicken
    """)
    
    st.markdown("---")
    st.markdown("**Version 1.0** | Made with ❤️ for Amazon Sellers")

