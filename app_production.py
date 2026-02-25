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
from concurrent.futures import ThreadPoolExecutor, as_completed, TimeoutError
import threading
import time

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Pydantic Models for structured output - 2 SEPARATE CALLS!
# NO max_length constraints - let ensure_optimal_length_with_ai handle length enforcement

# CALL 1: Hauptinhalt (Titel, Beschreibung, Bullets)
class MainContent(BaseModel):
    """Hauptinhalt des Amazon-Listings: Titel, Beschreibung, Bullets"""
    model_config = {"extra": "forbid"}
    
    artikelname: str = Field(description="Produkttitel 170-200 BYTES (~185 Bytes, ca. 30 Wörter). Format: MARKE Produkt Farbe, Eigenschaften. 100% nicht 100 Prozent!")
    produktbeschreibung: str = Field(description="Produktbeschreibung 1700-2000 BYTES (~1850 Bytes, ca. 280 Wörter)!")
    bullet_points: List[str] = Field(min_length=5, max_length=5, description="5 Bullets je 170-200 BYTES (~185 Bytes, ca. 30 Wörter). Format 'HOOK: Beschreibung.' NUR Hook in CAPS!")

# CALL 2: Keywords (bekommt Titel, Beschreibung, Bullets als Kontext)
class KeywordContent(BaseModel):
    """Backend-Keywords für Amazon - komma-getrennte Liste"""
    model_config = {"extra": "forbid"}
    
    suchbegriffe: str = Field(description="210-249 BYTES (~230 Bytes, ca. 15-20 Keywords). NUR komma-getrennte Keywords! KEINE Sätze!")

# CALL 3: Verifikation (prüft und korrigiert nur bei Bedarf)
class VerificationResult(BaseModel):
    """Ergebnis der Qualitätsprüfung"""
    model_config = {"extra": "forbid"}
    
    approved: bool = Field(description="True wenn alles korrekt ist, False wenn Korrekturen nötig waren")
    artikelname: str = Field(description="Korrigierter Titel oder Original wenn approved=True")
    bullet_1: str = Field(description="Korrigierter Bullet 1 oder Original")
    bullet_2: str = Field(description="Korrigierter Bullet 2 oder Original")
    bullet_3: str = Field(description="Korrigierter Bullet 3 oder Original")
    bullet_4: str = Field(description="Korrigierter Bullet 4 oder Original")
    bullet_5: str = Field(description="Korrigierter Bullet 5 oder Original")
    produktbeschreibung: str = Field(description="Korrigierte Beschreibung oder Original")
    suchbegriffe: str = Field(description="Korrigierte Keywords oder Original")
    issues_found: str = Field(description="Liste der gefundenen und korrigierten Probleme, oder 'Keine' wenn approved=True")

# CALL 1 PROMPT: Hauptinhalt (Titel, Beschreibung, Bullets) - OHNE Keywords!
MAIN_CONTENT_PROMPT = """Erstelle Titel, Beschreibung und Bullet Points für folgendes Amazon-Produkt.
WICHTIG: Erstelle KEINE Keywords - diese werden separat generiert!

Produktdaten:
{{product_data}}

🔤 AUSGABESPRACHE: {{language}}

⛔⛔⛔ ABSOLUT KRITISCH - SPRACHE ⛔⛔⛔
DER GESAMTE OUTPUT MUSS ZU 100% IN DER ZIELSPRACHE SEIN!

EINZIGE AUSNAHMEN (dürfen original bleiben):
- Markennamen: "Dreamfarm", "EMSA", "Philips", "Nike"
- Produktserien: "Aveo", "Clip & Close"
- Technische Bezeichnungen: "USB-C", "LED", "LCD"

ALLES ANDERE MUSS ÜBERSETZT WERDEN:
❌ VERBOTEN: "Pizzaschere", "Trinkflasche", "spülmaschinenfest", "auslaufsicher"
✅ STATTDESSEN (Italienisch): "forbici per pizza", "borraccia", "lavabile in lavastoviglie", "a prova di perdite"
✅ STATTDESSEN (Englisch): "pizza scissors", "water bottle", "dishwasher safe", "leak-proof"
✅ STATTDESSEN (Französisch): "ciseaux à pizza", "gourde", "lavable au lave-vaisselle", "étanche"
✅ STATTDESSEN (Spanisch): "tijeras para pizza", "botella", "apto lavavajillas", "hermético"

🔄 PFLICHT-ÜBERSETZUNGEN:
| Deutsch | Englisch | Italienisch | Französisch | Spanisch |
|---------|----------|-------------|-------------|----------|
| Trinkflasche | water bottle | borraccia | gourde | botella |
| Vorratsdose | storage container | contenitore | boîte de conservation | recipiente |
| Pizzaschere | pizza scissors | forbici per pizza | ciseaux à pizza | tijeras para pizza |
| Käsereibe | cheese grater | grattugia | râpe à fromage | rallador |
| spülmaschinenfest | dishwasher safe | lavabile in lavastoviglie | lave-vaisselle | apto lavavajillas |
| BPA-frei | BPA free | senza BPA | sans BPA | libre de BPA |
| Edelstahl | stainless steel | acciaio inox | acier inoxydable | acero inoxidable |
| auslaufsicher | leak-proof | a prova di perdite | étanche | hermético |
| hitzebeständig | heat resistant | resistente al calore | résistant à la chaleur | resistente al calor |

⚠️ PRÜFE JEDEN SATZ: Enthält er deutsche Wörter? → ÜBERSETZEN!

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

BEISPIELE AUS VERSCHIEDENEN INDUSTRIEN:
✅ Küche: "EMSA Vorratsdosen mit Deckel und Griffen, 3er Set 1L 2L 3L, BPA-frei spülmaschinenfest, Frischhaltedosen"
✅ Sport: "Nike Air Zoom Laufschuhe Herren Schwarz, Größe 43, atmungsaktives Mesh, gedämpfte Sohle, Joggingschuhe"
✅ Elektronik: "Anker PowerCore Powerbank 20000mAh, USB-C Schnellladung, 2 Ports, LED-Anzeige, kompakt für Reisen"
✅ Baby: "Philips Avent Babyflasche mit Sauger, 260ml, Anti-Kolik-Ventil, BPA-frei, spülmaschinenfest, Erstausstattung"
✅ Beauty: "Revlon Haartrockner mit Diffusor, 2200W, Ionentechnologie, 3 Heizstufen, Kaltluft, Föhn Salon"
❌ SCHLECHT: "Premium Aufbewahrungs-System für optimale Organisation" (= keine konkreten Eigenschaften)

⛔ TITEL DARF KEINEN PUNKT ENTHALTEN!
⛔ TITEL DARF KEIN VOLLSTÄNDIGER SATZ SEIN!
⛔ TITEL DARF KEINE BINDESTRICHE ALS TRENNZEICHEN ENTHALTEN!

REGELN FÜR TITEL:
- KEIN Punkt (.) im gesamten Titel! NIEMALS!
- KEINE vollständigen Sätze! Titel = Auflistung von Features/Keywords!
- Trenne Sinnabschnitte mit Kommata, KEINE Bindestriche als Trennzeichen!
- ABER: Behalte technische Bezeichnungen wie "18/10 Edelstahl", "BPA-frei"
- Komponenten IMMER direkt nach Produktbezeichnung (Long-Tail-Keywords!)
- Material MUSS enthalten sein wenn bekannt
- Größe/Maße MUSS enthalten sein wenn bekannt
- Am Ende: Synonym für Produktgattung

SPEZIELLE TITEL-REGELN:
- Prozent als Symbol: "100%" NICHT "100 Prozent"
- Farbe DIREKT nach Produktname, OHNE "Farbe" davor!
  ❌ "Lunch Bag Farbe blau Schiefer" → ✅ "Lunch Bag Schiefer"
  ❌ "Flasche Farbe olive grün" → ✅ "Flasche Olive"
- Bei Farbvarianten: Farbe kommt direkt nach Produktbezeichnung!

TITEL-BEISPIELE (VERSCHIEDENE INDUSTRIEN):
✅ Küche: "WMF Profi Bratpfanne Schwarz 28cm, antihaftbeschichtet, induktionsgeeignet, Edelstahlgriff"
✅ Sport: "Adidas Rucksack Classic Schwarz 25L, wasserabweisend, gepolsterte Gurte, Laptopfach 15 Zoll"
✅ Elektronik: "Samsung Galaxy Buds Pro Weiß, Bluetooth 5.0, aktive Geräuschunterdrückung, IPX7 wasserdicht"
✅ Baby: "MAM Easy Start Babyflasche 260ml Rosa, Anti-Kolik, selbststerilisierend, ab 0 Monate"
❌ SCHLECHT: "Das Produkt ist ein tolles Gerät für jeden Tag." (= Satz mit Punkt!)
❌ SCHLECHT: "Flasche 850 ml, 100 Prozent auslaufsicher" (= "Prozent" statt "%"!)
❌ SCHLECHT: "Kühltasche - isolierend - wasserabweisend" (= Bindestriche als Trennzeichen!)

📌 BULLET POINTS - EXAKTES FORMAT:

⛔⛔⛔ KRITISCH - JEDER BULLET MUSS EXAKT DIESES FORMAT HABEN:
HOOK IN CAPS: Beschreibender Text in normaler Schreibweise mit einem vollständigen Satz.

SCHEMA: [2-4 WÖRTER IN GROSSBUCHSTABEN]: [Normaler Satz mit ca. 25-35 Wörtern gesamt.]

BEISPIELE KORREKTES FORMAT (VERSCHIEDENE INDUSTRIEN):
✅ Küche: "SCHARFE KLINGEN: Die Edelstahlklingen schneiden mühelos durch alle Zutaten und bleiben langlebig scharf."
✅ Sport: "ATMUNGSAKTIVES MESH: Das leichte Obermaterial sorgt für optimale Belüftung bei intensivem Training und bleibt angenehm auf der Haut."
✅ Elektronik: "SCHNELLLADUNG: Mit USB-C Power Delivery lädt das Gerät bis zu 3x schneller als herkömmliche Ladegeräte und spart wertvolle Zeit."
✅ Baby: "ANTI-KOLIK-VENTIL: Das spezielle Ventil reduziert Luftschlucken und beugt Bauchschmerzen wirksam vor, damit das Baby entspannt trinkt."
✅ Beauty: "IONENTECHNOLOGIE: Negative Ionen versiegeln die Haarstruktur und sorgen für glänzendes, frizzfreies Haar nach jedem Föhnen."

BEISPIELE FALSCHES FORMAT:
❌ "VIELSEITIGES MATERIAL: Die BOXEN aus EDELSTAHL sind robust." (EDELSTAHL/BOXEN darf NICHT caps sein im Text nach dem Doppelpunkt!)
❌ "Das atmungsaktive Material sorgt für Komfort." (KEINE Hook vorhanden! Muss mit CAPS-HOOK: beginnen!)
❌ "PERFEKT FÜR SPORT: IDEAL für TRAINING und YOGA." (IDEAL/TRAINING/YOGA dürfen NICHT caps sein nach dem Doppelpunkt!)
❌ "EXTRA WEITE ÖFFNUNG: Die Öffnung erleichtert das Packen und Entnehmen von Dosen, Flaschen, Snacks und Obst. Der Roll-up Verschl" (ABGEBROCHEN! Satz muss vollständig sein!)

REGELN:
1. JEDER Bullet beginnt mit einem HOOK in GROSSBUCHSTABEN (2-4 Wörter)
2. Nach dem Hook kommt ein DOPPELPUNKT und ein Leerzeichen
3. Dann folgt der beschreibende Text in NORMALER Schreibweise
4. Im beschreibenden Text nach dem Doppelpunkt KEINE EINZIGEN weiteren CAPS-Wörter! Nur normale Groß-/Kleinschreibung!
5. Vollständige Sätze, NIEMALS mitten im Wort oder Satz abbrechen!
6. Jeder Bullet = ca. 25-35 Wörter gesamt (Hook + beschreibender Text)

DIE 5 BULLETS MÜSSEN ABDECKEN:
1. HAUPTVORTEIL: Der größte Nutzen für den Kunden
2. MATERIAL/QUALITÄT: Material und warum es gut ist
3. ANWENDUNG: Wo und wie wird es benutzt
4. BESONDERHEIT: Was unterscheidet es von anderen Produkten
5. LIEFERUMFANG/DETAILS: Was ist enthalten, technische Details

📌 PRODUKTBESCHREIBUNG (1500-1700 ZEICHEN):
Ausführliche Beschreibung die ALLE 15 Beziehungstypen inhaltlich abdeckt.
KEINE technischen Begriffe wie "is", "has_property" etc. verwenden!

⚠️ EXAKTE BYTE-LIMITS (KRITISCH!):

📏 ZIEL-LÄNGEN (in BYTES, nicht Zeichen!):
- TITEL: 170-200 BYTES (Ziel: ~185 Bytes)
- BULLET POINTS: Je 170-200 BYTES (Ziel: ~185 Bytes pro Bullet)
- BESCHREIBUNG: 1700-2000 BYTES (Ziel: ~1850 Bytes)

⚠️ BYTE-ZÄHLUNG: Umlaute (ä,ö,ü,ß) = 2 Bytes, normale Buchstaben = 1 Byte

💡 TIPP FÜR RICHTIGE LÄNGE:
- Titel: ca. 25-35 Wörter
- Bullet: ca. 25-35 Wörter pro Bullet
- Beschreibung: ca. 250-300 Wörter

WICHTIG: Nutze die Länge MAXIMAL aus! Kurze Texte verschwenden SEO-Potenzial!

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

WICHTIG: Sachliche Produktinfos! ERFINDE NICHTS! Titel muss LESBAR sein - nicht nur Keywords aneinanderreihen!
"""

# CALL 2 PROMPT: Keywords - SEPARATER FOKUSSIERTER CALL
KEYWORD_PROMPT = """Erstelle Backend-Suchbegriffe (Keywords) für dieses Amazon-Produkt.

🔤 AUSGABESPRACHE: {{language}}

📊 BEREITS ERSTELLTER INHALT (DIESE BEGRIFFE NICHT WIEDERHOLEN!):
---
TITEL: {{title}}
---
BULLET 1: {{bullet1}}
BULLET 2: {{bullet2}}
BULLET 3: {{bullet3}}
BULLET 4: {{bullet4}}
BULLET 5: {{bullet5}}
---
BESCHREIBUNG: {{description}}
---

{{poe_data}}

🔑 DEINE AUFGABE: BACKEND-KEYWORDS ERSTELLEN (210-249 BYTES)

⛔⛔⛔ ABSOLUT KRITISCH - FORMAT ⛔⛔⛔

Dein Output MUSS so aussehen:
wort1, wort2, begriff drei, wort4, wort5, ...

WENN DEIN OUTPUT EINEN PUNKT (.) ENTHÄLT, IST ER FALSCH!
WENN DEIN OUTPUT EINEN GANZEN SATZ ENTHÄLT, IST ER FALSCH!
WENN EIN KEYWORD MEHR ALS 3 WÖRTER HAT, IST ES FALSCH!

Jedes Keyword = 1-3 Wörter. KEINE Artikel (der/die/das/ein/eine). KEINE Verben im Infinitiv als Satzanfang.

✅ RICHTIG:
küchenhelfer, küchenzubehör, kochwerkzeug, haushalt gadget, edelstahl reibe, parmesan hobel

✅ RICHTIG (Sport):
fitness zubehör, gym equipment, workout gear, trainingsband, resistance band

✅ RICHTIG (Elektronik):
ladegerät usb, portable charger, schnellladung, power bank klein, reise adapter

✅ RICHTIG (Baby):
babyzubehör, säuglingsausstattung, babycare, erstausstattung, neugeborene

❌ FALSCH - SÄTZE:
"Die Käsereibe ist ideal für Parmesan. Robustes Material." → MUSS SEIN: "parmesan hobel, robuste reibe, hartkäse raspel"

❌ FALSCH - ZU LANG:
"Perfektes Ladegerät für unterwegs mit schneller Aufladung" → MUSS SEIN: "ladegerät unterwegs, schnellladung, reise charger"

❌ FALSCH - MIT PUNKTEN:
"Küchenhelfer. Gadget. Haushalt." → MUSS SEIN: "küchenhelfer, gadget, haushalt"

STRENGE REGELN:
1. NUR 1-3 Wörter pro Keyword!
2. Getrennt durch KOMMAS!
3. KEINE Punkte im gesamten Output!
4. KEINE ganzen Sätze!
5. KEINE Beschreibungen oder Erklärungen!
6. Alles kleingeschrieben (außer Markennamen)!
7. KEINE Begriffe die bereits im Titel, Bullets oder Beschreibung stehen!

WAS GEHÖRT REIN:
- Synonyme (z.B. Trinkflasche im Titel → "wasserflasche, flasche, bottle")
- Schreibvarianten ohne Umlaut (z.B. "käsereibe" → "kaesereibe")
- Long-Tail 2-3 Wort Kombinationen (z.B. "manuell ohne strom", "küchenhelfer hand")
- POE-Suchbegriffe die NICHT im Text stehen

WAS GEHÖRT NICHT REIN:
- Begriffe die schon im Titel/Bullets/Beschreibung stehen!
- Ganze Sätze oder Beschreibungen!
- Komplementäre Produkte
- Wettbewerber-Marken

📏 EXAKTES BYTE-ZIEL: 210-249 BYTES (Ziel: ~230 Bytes)
💡 TIPP: ca. 15-20 Keywords/Phrasen erreichen diese Länge
Umlaute (ä,ö,ü) = 2 Bytes!
"""

# CALL 3 PROMPT: Verifikation - prüft und korrigiert nur bei Bedarf
VERIFICATION_PROMPT = """Du bist ein strenger Qualitätsprüfer für Amazon-Listings.

🔤 ZIELSPRACHE: {{language}}

Prüfe den folgenden Content und korrigiere NUR wenn wirklich nötig.
Wenn alles korrekt ist, gib den Original-Content zurück mit approved=True.

═══════════════════════════════════════════════════════════════════
📋 ZU PRÜFENDER CONTENT:
═══════════════════════════════════════════════════════════════════

TITEL: {{title}}

BULLET 1: {{bullet1}}
BULLET 2: {{bullet2}}
BULLET 3: {{bullet3}}
BULLET 4: {{bullet4}}
BULLET 5: {{bullet5}}

BESCHREIBUNG: {{description}}

KEYWORDS: {{keywords}}

═══════════════════════════════════════════════════════════════════
✅ PRÜFKRITERIEN (JEDES IST PFLICHT!):
═══════════════════════════════════════════════════════════════════

1. SPRACHE: Ist ALLES in der Zielsprache (außer Markennamen, Seriennamen, technische Begriffe)?
   ❌ FEHLER: Deutsche Wörter in italienischem/englischem/etc. Output (z.B. "BPA-frei" in EN statt "BPA free")
   ❌ FEHLER: "spülmaschinenfest" in EN statt "dishwasher safe"

2. TITEL-FORMAT:
   ❌ FEHLER: Enthält einen Punkt (.)
   ❌ FEHLER: Ist ein vollständiger Satz
   ❌ FEHLER: "100 Prozent" statt "100%"
   ❌ FEHLER: "Farbe blau" statt Farbe direkt nach Produktname
   ❌ FEHLER: Bindestriche als Trennzeichen zwischen Abschnitten (nur Kommata erlaubt!)
   → Wenn Titel einen Punkt enthält, ENTFERNE ihn und kürze wenn nötig!

3. BULLET-FORMAT (JEDER einzelne Bullet prüfen!):
   ❌ FEHLER: Kein HOOK am Anfang (muss 2-4 CAPS-Wörter + Doppelpunkt sein!)
   ❌ FEHLER: Kein Doppelpunkt nach HOOK
   ❌ FEHLER: Weitere CAPS-Wörter im beschreibenden Text nach dem Doppelpunkt
   ❌ FEHLER: Unvollständiger Satz (abgebrochen, letztes Wort unvollständig)
   → Wenn Hook fehlt: Füge einen passenden HOOK hinzu!
   → Wenn CAPS im Text: Wandle in normale Schreibweise um!
   → Wenn Satz abgebrochen: Vervollständige den Satz!

4. KEYWORDS:
   ❌ FEHLER: Enthält Punkte (.) → ENTFERNE alle Punkte!
   ❌ FEHLER: Ganze Sätze statt komma-getrennte Wörter → Zerlege in einzelne Keywords!
   ❌ FEHLER: Keywords mit mehr als 3 Wörtern → Kürze auf 1-3 Wörter!
   ❌ FEHLER: Beschreibende Phrasen mit Verben → Reduziere auf Substantive/Adjektive!
   → Keywords müssen IMMER so aussehen: "wort1, wort2, begriff drei, wort4"

═══════════════════════════════════════════════════════════════════
⚠️ WICHTIG:
═══════════════════════════════════════════════════════════════════

- Korrigiere NUR echte Fehler gegen die obigen Regeln!
- Ändere NICHT den inhaltlichen Kern, nur Format wenn nötig!
- Behalte die Länge der Texte möglichst bei! Nicht unnötig kürzen oder verlängern!
- Wenn alles OK ist: approved=True und Original-Content EXAKT zurückgeben!
- Bei Korrekturen: approved=False und korrigierten Content + issues_found
"""

# Page Config
st.set_page_config(
    page_title="Amazon Content Optimizer",
    page_icon="✍️",
    layout="wide"
)

# Known GPT models as fallback if API listing fails
KNOWN_GPT_MODELS = ["gpt-5.1", "gpt-5", "gpt-4.1", "gpt-4.1-mini", "gpt-4.1-nano", "gpt-4o", "gpt-4o-mini", "gpt-4-turbo"]

# Initialize Session State
if 'api_key' not in st.session_state:
    st.session_state.api_key = ""
if 'available_models' not in st.session_state:
    st.session_state.available_models = KNOWN_GPT_MODELS
if 'selected_model' not in st.session_state:
    st.session_state.selected_model = "gpt-5.1"
if 'main_content_prompt' not in st.session_state:
    st.session_state.main_content_prompt = MAIN_CONTENT_PROMPT
if 'keyword_prompt' not in st.session_state:
    st.session_state.keyword_prompt = KEYWORD_PROMPT
if 'verification_prompt' not in st.session_state:
    st.session_state.verification_prompt = VERIFICATION_PROMPT

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
        help="Ihr OpenAI API Key"
    )
    
    if st.button("💾 API Key speichern"):
        st.session_state.api_key = api_key_input
        if api_key_input.strip():
            try:
                temp_client = OpenAI(api_key=api_key_input)
                models_response = temp_client.models.list()
                gpt_models = sorted(
                    [m.id for m in models_response.data if m.id.startswith("gpt-") and "realtime" not in m.id and "audio" not in m.id],
                    key=lambda x: getattr(next((m for m in models_response.data if m.id == x), None), 'created', 0),
                    reverse=True
                )
                if gpt_models:
                    st.session_state.available_models = gpt_models
                    logger.info(f"Loaded {len(gpt_models)} GPT models from API")
            except Exception as e:
                logger.warning(f"Could not load models from API: {e}")
                st.session_state.available_models = KNOWN_GPT_MODELS
        st.success("✅ API Key gespeichert!")
    
    st.markdown("---")
    if st.session_state.get('api_key', '').strip():
        st.success("✅ API Key konfiguriert")
    else:
        st.error("❌ API Key fehlt")
    
    available = st.session_state.available_models
    default_idx = 0
    if st.session_state.selected_model in available:
        default_idx = available.index(st.session_state.selected_model)
    
    selected = st.selectbox(
        "KI-Modell",
        options=available,
        index=default_idx,
        help="Wähle das OpenAI-Modell für die Content-Generierung"
    )
    st.session_state.selected_model = selected
    
    with st.expander("✏️ Prompts bearbeiten"):
        st.caption("**Call 1:** Titel, Beschreibung, Bullets")
        main_prompt = st.text_area(
            "Hauptinhalt-Prompt",
            value=st.session_state.main_content_prompt,
            height=200,
            key="main_prompt_edit"
        )
        st.caption("**Call 2:** Keywords (bekommt Output von Call 1)")
        kw_prompt = st.text_area(
            "Keyword-Prompt",
            value=st.session_state.keyword_prompt,
            height=200,
            key="kw_prompt_edit"
        )
        if st.button("💾 Prompts speichern"):
            st.session_state.main_content_prompt = main_prompt
            st.session_state.keyword_prompt = kw_prompt
            st.success("✅ Prompts gespeichert!")

# Helper functions
def get_byte_length(text: str) -> int:
    return len(text.encode('utf-8'))

def ensure_optimal_length_with_ai(text: str, min_bytes: int, max_bytes: int, field_name: str, client_instance, product_context: str = "", max_retries: int = 5, model_name: str = "gpt-5.1", language: str = "") -> str:
    """
    Ensure text is within optimal byte range using word-count guidance.
    Computes avg bytes/word from actual text, then tells the LLM exactly
    how many words to add/remove — something it can actually follow.
    Always operates in the specified output language.
    """
    current_text = text
    current_bytes = get_byte_length(current_text)
    
    if min_bytes <= current_bytes <= max_bytes:
        logger.info(f"✅ {field_name}: {current_bytes} bytes (OK)")
        return current_text
    
    # Aim at 70% up the range — safety margin below max
    target_bytes = min_bytes + int((max_bytes - min_bytes) * 0.7)
    
    lang_hint = f"CRITICAL: Output MUST be 100% in {language}. Do NOT change the language!" if language else ""
    
    field_lower = field_name.lower()
    if "titel" in field_lower or "title" in field_lower:
        field_type_hint = "Amazon product title. NO period at the end! NO full sentence! Comma-separated feature listing only."
    elif "bullet" in field_lower:
        field_type_hint = "Amazon bullet point. Format: HOOK IN CAPS: Descriptive text. Complete sentence, NEVER cut off mid-sentence!"
    elif "keyword" in field_lower:
        field_type_hint = "Amazon backend keywords. ONLY comma-separated words/phrases (1-3 words). NO sentences! NO periods!"
    else:
        field_type_hint = "Amazon product description. Complete, readable sentences."
    
    for attempt in range(max_retries):
        current_bytes = get_byte_length(current_text)
        
        if min_bytes <= current_bytes <= max_bytes:
            logger.info(f"✅ {field_name}: {current_bytes} bytes (nach {attempt} Anpassungen)")
            return current_text
        
        words = current_text.split()
        word_count = len(words)
        avg_bytes_per_word = current_bytes / word_count if word_count > 0 else 6
        target_word_count = round(target_bytes / avg_bytes_per_word)
        word_diff = target_word_count - word_count
        
        if current_bytes < min_bytes:
            action = "EXPAND"
            word_instruction = f"Add approximately {abs(word_diff)} words with relevant product details. Target: {target_word_count} words total."
        else:
            action = "SHORTEN"
            word_instruction = f"Remove approximately {abs(word_diff)} words. Target: {target_word_count} words total."
        
        logger.info(f"⚠️ {field_name} ({current_bytes}B, {word_count}W → target: ~{target_word_count}W). Attempt {attempt + 1}/{max_retries}...")
        
        prompt = f"""{action} this text. It currently has {word_count} words.
{word_instruction}

FIELD TYPE: {field_type_hint}

{lang_hint}

PRODUCT CONTEXT:
{product_context[:1000] if product_context else "N/A"}

RULES:
1. Write EXACTLY {target_word_count} words (±2 words)!
2. COMPLETE SENTENCES - never cut off mid-sentence or mid-word!
3. Keep technical terms exactly as-is: "18/10 Edelstahl", "BPA-frei", "0,5L"
4. Text must make sense and be readable!
5. Keep the most important product information!
6. Do NOT change the language of the text!

TEXT TO ADJUST ({word_count} words):
{current_text}

Respond ONLY with the adjusted text, WITHOUT quotation marks:"""
        
        try:
            resp = client_instance.chat.completions.create(
                model=model_name,
                messages=[
                    {"role": "system", "content": f"You adjust texts to an exact word count. Target: {target_word_count} words. {field_type_hint} {lang_hint}"},
                    {"role": "user", "content": prompt}
                ],
                max_completion_tokens=int(max_bytes / 0.9) + 100
            )
            new_text = resp.choices[0].message.content.strip()
            
            if new_text.startswith('"') and new_text.endswith('"'):
                new_text = new_text[1:-1]
            if new_text.startswith("'") and new_text.endswith("'"):
                new_text = new_text[1:-1]
            if new_text.startswith('```') and new_text.endswith('```'):
                new_text = new_text[3:-3].strip()
            
            new_bytes = get_byte_length(new_text)
            new_words = len(new_text.split())
            logger.info(f"  → {current_bytes}B/{word_count}W → {new_bytes}B/{new_words}W")
            
            current_text = new_text
            
        except Exception as e:
            logger.error(f"Error adjusting text (attempt {attempt + 1}): {e}")
            time.sleep(min(2 ** attempt, 10))
    
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
        
        # Parallelization slider
        parallel_workers = st.slider("Parallele Verarbeitung", 1, 5, 3, key="parallel_slider", 
                                     help="Anzahl der Produkte die gleichzeitig verarbeitet werden")
        
        if st.button("🚀 Optimierung Starten", type="primary", key="opt_start", use_container_width=True):
            if not st.session_state.get('api_key', '').strip():
                st.error("❌ Bitte API Key in der Seitenleiste eingeben")
            else:
                results = []
                results_lock = threading.Lock()
                progress_bar = st.progress(0)
                status = st.empty()
                completed_count = [0]  # Mutable for thread access
                
                client = OpenAI(api_key=st.session_state.api_key)
                
                # Store prompts and model locally for thread safety
                main_prompt_template = st.session_state.main_content_prompt
                keyword_prompt_template = st.session_state.keyword_prompt
                verification_prompt_template = st.session_state.verification_prompt
                lang_instruction = language_options[selected_language]
                active_model = st.session_state.selected_model
                
                def process_product(idx, row, id_col_name, title_col_name, poe_kws):
                    """Process a single product with Call 1 → Call 2 sequentially"""
                    product_data_str = "\n".join([f"- {k}: {v}" for k, v in row.to_dict().items() if pd.notna(v)])
                    
                    try:
                        # ========================================
                        # CALL 1: Hauptinhalt (Titel, Beschreibung, Bullets)
                        # ========================================
                        logger.info(f"Produkt {idx + 1}: Call 1 - Hauptinhalt...")
                        
                        main_prompt = main_prompt_template
                        main_prompt = main_prompt.replace("{{product_data}}", product_data_str)
                        main_prompt = main_prompt.replace("{{language}}", lang_instruction)
                        
                        response1 = client.chat.completions.create(
                            model=active_model,
                            messages=[
                                {"role": "system", "content": f"Amazon SEO-Experte. KRITISCH: Schreibe 100% in der Zielsprache! Keine deutschen Wörter außer Markennamen! OUTPUT: {lang_instruction}"},
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
                        
                        # Längenanpassung für Hauptinhalt
                        titel_bytes = get_byte_length(main_content.artikelname)
                        if titel_bytes < 170 or titel_bytes > 200:
                            main_content.artikelname = ensure_optimal_length_with_ai(
                                main_content.artikelname, 170, 200, f"Titel P{idx+1}", client, product_data_str, model_name=active_model, language=lang_instruction)
                        
                        new_bullets = []
                        for i, bp in enumerate(main_content.bullet_points):
                            bp_bytes = get_byte_length(bp)
                            if bp_bytes < 170 or bp_bytes > 200:
                                bp = ensure_optimal_length_with_ai(
                                    bp, 170, 200, f"Bullet {i+1} P{idx+1}", client, product_data_str, model_name=active_model, language=lang_instruction)
                            new_bullets.append(bp)
                        main_content.bullet_points = new_bullets
                        
                        desc_bytes = get_byte_length(main_content.produktbeschreibung)
                        if desc_bytes < 1700 or desc_bytes > 2000:
                            main_content.produktbeschreibung = ensure_optimal_length_with_ai(
                                main_content.produktbeschreibung, 1700, 2000, f"Beschreibung P{idx+1}", client, product_data_str, model_name=active_model, language=lang_instruction)
                        
                        # ========================================
                        # CALL 2: Keywords (bekommt Output von Call 1)
                        # ========================================
                        logger.info(f"Produkt {idx + 1}: Call 2 - Keywords...")
                        
                        poe_data_str = ""
                        if poe_kws:
                            poe_data_str = f"""
📊 POE-DATEN (Top-Suchbegriffe - nutze als Inspiration):
{', '.join(poe_kws[:15])}
"""
                        
                        kw_prompt = keyword_prompt_template
                        kw_prompt = kw_prompt.replace("{{language}}", lang_instruction)
                        kw_prompt = kw_prompt.replace("{{title}}", main_content.artikelname)
                        kw_prompt = kw_prompt.replace("{{bullet1}}", main_content.bullet_points[0] if len(main_content.bullet_points) > 0 else "")
                        kw_prompt = kw_prompt.replace("{{bullet2}}", main_content.bullet_points[1] if len(main_content.bullet_points) > 1 else "")
                        kw_prompt = kw_prompt.replace("{{bullet3}}", main_content.bullet_points[2] if len(main_content.bullet_points) > 2 else "")
                        kw_prompt = kw_prompt.replace("{{bullet4}}", main_content.bullet_points[3] if len(main_content.bullet_points) > 3 else "")
                        kw_prompt = kw_prompt.replace("{{bullet5}}", main_content.bullet_points[4] if len(main_content.bullet_points) > 4 else "")
                        kw_prompt = kw_prompt.replace("{{description}}", main_content.produktbeschreibung)
                        kw_prompt = kw_prompt.replace("{{poe_data}}", poe_data_str)
                        
                        response2 = client.chat.completions.create(
                            model=active_model,
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
                        
                        # Längenanpassung für Keywords
                        kw_bytes = get_byte_length(keyword_content.suchbegriffe)
                        if kw_bytes < 210 or kw_bytes > 249:
                            keyword_content.suchbegriffe = ensure_optimal_length_with_ai(
                                keyword_content.suchbegriffe, 210, 249, f"Keywords P{idx+1}", client, product_data_str, model_name=active_model, language=lang_instruction)
                        
                        # ========================================
                        # CALL 3: Verifikation (prüft und korrigiert nur bei Bedarf)
                        # ========================================
                        logger.info(f"Produkt {idx + 1}: Call 3 - Verifikation...")
                        
                        verify_prompt = verification_prompt_template
                        verify_prompt = verify_prompt.replace("{{language}}", lang_instruction)
                        verify_prompt = verify_prompt.replace("{{title}}", main_content.artikelname)
                        verify_prompt = verify_prompt.replace("{{bullet1}}", main_content.bullet_points[0] if len(main_content.bullet_points) > 0 else "")
                        verify_prompt = verify_prompt.replace("{{bullet2}}", main_content.bullet_points[1] if len(main_content.bullet_points) > 1 else "")
                        verify_prompt = verify_prompt.replace("{{bullet3}}", main_content.bullet_points[2] if len(main_content.bullet_points) > 2 else "")
                        verify_prompt = verify_prompt.replace("{{bullet4}}", main_content.bullet_points[3] if len(main_content.bullet_points) > 3 else "")
                        verify_prompt = verify_prompt.replace("{{bullet5}}", main_content.bullet_points[4] if len(main_content.bullet_points) > 4 else "")
                        verify_prompt = verify_prompt.replace("{{description}}", main_content.produktbeschreibung)
                        verify_prompt = verify_prompt.replace("{{keywords}}", keyword_content.suchbegriffe)
                        
                        response3 = client.chat.completions.create(
                            model=active_model,
                            messages=[
                                {"role": "system", "content": f"Strenger Qualitätsprüfer für Amazon-Listings. Prüfe und korrigiere NUR bei echten Fehlern! Zielsprache: {lang_instruction}"},
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
                        
                        # Übernehme verifizierte/korrigierte Werte
                        final_title = verification.artikelname
                        final_bullets = [
                            verification.bullet_1,
                            verification.bullet_2,
                            verification.bullet_3,
                            verification.bullet_4,
                            verification.bullet_5
                        ]
                        final_description = verification.produktbeschreibung
                        final_keywords = verification.suchbegriffe
                        
                        if verification.approved:
                            logger.info(f"Produkt {idx + 1}: ✅ Verifikation bestanden")
                        else:
                            logger.info(f"Produkt {idx + 1}: ⚠️ Korrekturen: {verification.issues_found}")
                            
                            # ========================================
                            # Längenkorrektur für geänderte Texte nach Verifikation
                            # ========================================
                            
                            # Titel prüfen (falls geändert)
                            if final_title != main_content.artikelname:
                                title_bytes = get_byte_length(final_title)
                                if title_bytes < 170 or title_bytes > 200:
                                    final_title = ensure_optimal_length_with_ai(
                                        final_title, 170, 200, f"Titel P{idx+1} post-verify", client, product_data_str, model_name=active_model, language=lang_instruction)
                            
                            # Bullets prüfen (falls geändert)
                            original_bullets = main_content.bullet_points
                            for i in range(5):
                                if i < len(original_bullets) and final_bullets[i] != original_bullets[i]:
                                    bp_bytes = get_byte_length(final_bullets[i])
                                    if bp_bytes < 170 or bp_bytes > 200:
                                        final_bullets[i] = ensure_optimal_length_with_ai(
                                            final_bullets[i], 170, 200, f"Bullet {i+1} P{idx+1} post-verify", client, product_data_str, model_name=active_model, language=lang_instruction)
                            
                            # Beschreibung prüfen (falls geändert)
                            if final_description != main_content.produktbeschreibung:
                                desc_bytes = get_byte_length(final_description)
                                if desc_bytes < 1700 or desc_bytes > 2000:
                                    final_description = ensure_optimal_length_with_ai(
                                        final_description, 1700, 2000, f"Beschreibung P{idx+1} post-verify", client, product_data_str, model_name=active_model, language=lang_instruction)
                            
                            # Keywords prüfen (falls geändert)
                            if final_keywords != keyword_content.suchbegriffe:
                                kw_bytes = get_byte_length(final_keywords)
                                if kw_bytes < 210 or kw_bytes > 249:
                                    final_keywords = ensure_optimal_length_with_ai(
                                        final_keywords, 210, 249, f"Keywords P{idx+1} post-verify", client, product_data_str, model_name=active_model, language=lang_instruction)
                        
                        # ========================================
                        # Ergebnis zusammenführen
                        # ========================================
                        result_row = {
                            "idx": idx,  # For sorting later
                            "Identifier": row[id_col_name],
                            "Old Title": row[title_col_name] if title_col_name != "-" else "",
                            "New Title": final_title,
                            "New Description": final_description,
                            "New Keyword": final_keywords
                        }
                        for i, bp in enumerate(final_bullets, 1):
                            result_row[f"New Bullet {i}"] = bp
                        
                        logger.info(f"✅ Produkt {idx + 1} fertig!")
                        return {
                            "success": True, 
                            "data": result_row, 
                            "final_title": final_title,
                            "final_bullets": final_bullets,
                            "final_description": final_description,
                            "final_keywords": final_keywords,
                            "verified": verification.approved,
                            "issues": verification.issues_found
                        }
                    
                    except Exception as e:
                        logger.error(f"Fehler bei Produkt {idx + 1}: {e}", exc_info=True)
                        return {"success": False, "error": str(e), "idx": idx}
                
                # Process products in parallel
                status.text(f"🚀 Starte parallele Verarbeitung ({parallel_workers} gleichzeitig)...")
                error_count = [0]
                
                with ThreadPoolExecutor(max_workers=parallel_workers) as executor:
                    futures = {}
                    for idx in range(num_products_opt):
                        row = df_opt.iloc[idx]
                        future = executor.submit(process_product, idx, row, id_col, title_col, poe_keywords)
                        futures[future] = idx
                    
                    for future in as_completed(futures):
                        idx = futures[future]
                        
                        try:
                            result = future.result(timeout=180)
                        except TimeoutError:
                            logger.error(f"Produkt {idx + 1}: Timeout nach 180 Sekunden")
                            result = {"success": False, "error": "Timeout (180s)", "idx": idx}
                        except Exception as e:
                            logger.error(f"Produkt {idx + 1}: Unerwarteter Fehler: {e}", exc_info=True)
                            result = {"success": False, "error": str(e), "idx": idx}
                        
                        with results_lock:
                            completed_count[0] += 1
                            status.text(f"✅ {completed_count[0]}/{num_products_opt} Produkte fertig...")
                            progress_bar.progress(completed_count[0] / num_products_opt)
                        
                        if result["success"]:
                            with results_lock:
                                results.append(result["data"])
                            
                            final_title = result["final_title"]
                            final_bullets = result["final_bullets"]
                            final_keywords = result["final_keywords"]
                            final_description = result["final_description"]
                            verified = result["verified"]
                            issues = result["issues"]
                            
                            status_icon = "✅" if verified else "🔧"
                            with st.expander(f"{status_icon} Produkt {idx + 1}: {final_title[:50]}..."):
                                if not verified:
                                    st.warning(f"⚠️ Korrekturen vorgenommen: {issues}")
                                st.write("**Titel:**", final_title)
                                st.write(f"*({get_byte_length(final_title)} bytes)*")
                                st.write("**Bullets:**")
                                for i, bp in enumerate(final_bullets, 1):
                                    st.write(f"• {bp} *({get_byte_length(bp)} bytes)*")
                                st.write("**Keywords:**", final_keywords)
                                st.write(f"*({get_byte_length(final_keywords)} bytes)*")
                                st.caption("**Beschreibung:** " + final_description[:100] + "...")
                        else:
                            with results_lock:
                                error_count[0] += 1
                            st.error(f"❌ Fehler bei Produkt {idx + 1}: {result['error']}")
                
                # Sort results by original index
                results.sort(key=lambda x: x["idx"])
                # Remove idx from final output
                for r in results:
                    del r["idx"]
                
                if error_count[0] > 0:
                    status.text(f"✅ Fertig! {len(results)} erfolgreich, {error_count[0]} fehlgeschlagen")
                    st.warning(f"⚠️ {error_count[0]} von {num_products_opt} Produkten konnten nicht verarbeitet werden. Teilergebnisse sind verfügbar.")
                else:
                    status.text(f"✅ Fertig! {len(results)} Produkte erfolgreich verarbeitet")
                
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
