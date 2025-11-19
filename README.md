# Amazon Listing Agent

Automatisches Befüllen von beliebigen Amazon-Templates mit KI-optimierten Inhalten.

## 🚀 Quick Start

```bash
cd "/Users/florianstrauss/Desktop/Amazon Listing Agent"
source venv/bin/activate
streamlit run app_production.py
```

**Browser öffnet automatisch:** http://localhost:8501

## 📋 Verwendung

1. **API Key eingeben** (Sidebar)
2. **Produktdaten hochladen** (Excel)
3. **Template hochladen** (beliebiges Amazon .xlsm)
4. **"Starten" klicken**
5. **Download** - Fertig!

## ✅ Features

- ✅ **Beliebige Templates** - XML & Flat File Format
- ✅ **Automatische Erkennung** - Keine Konfiguration nötig
- ✅ **AI-Content** - GPT-5-mini mit COSMO/RUFUS Optimierung
- ✅ **Struktur-Erhaltung** - Formeln, Formatierung, Makros bleiben erhalten
- ✅ **Dynamische Pflichtfelder** - Füllt automatisch ALLE erforderlichen Felder pro Produkttyp

## 📁 Test-Dateien

**Produktdaten:**
- `data/input/Amazon Products_example.xlsx`

**Templates:**
- `data/BOTTLE_KITCHEN_KNIFE_SAUTE_FRY_PAN.xlsm` (XML Format)
- `data/template/DRINKING_CUP_BOTTLE_SEASONING_MILL_CUTTING_BOARD_FOOD_STORAGE_CONTAINER.xlsm` (Flat File)

## 🔧 Technische Details

### Hauptdateien:
- `app_production.py` - Streamlit App
- `dynamic_template_analyzer.py` - Template-Analyse
- `requirements.txt` - Dependencies

### Unterstützte Formate:
- **XML Format** - Seller Central (Header Row 4, Data Row 7)
- **Flat File Format** - Seller Central (Header Row 2, Data Row 4)

## 💰 Kosten

- **Template-Analyse:** Kostenlos (keine API-Calls)
- **Content-Generierung:** ~€0.001-0.003 pro Produkt (GPT-5-mini)
- **10 Produkte:** ~€0.01-0.03

## 🐛 Troubleshooting

**API Key Error:**
- Prüfen Sie ob der Key korrekt ist
- Guthaben vorhanden?

**Template-Analyse fehlgeschlagen:**
- Ist die Datei .xlsm oder .xlsx?
- Hat die Datei ein "Vorlage" Sheet?

**Logs ansehen:**
```bash
# Terminal wo App gestartet wurde zeigt detaillierte Logs
```

## 📊 Was wird befüllt?

**AI-Generiert:**
- Artikelname (COSMO/RUFUS optimiert)
- 5 Bullet Points (ohne Nummerierung)
- 5 Suchbegriffe

**Dynamisch pro Produkttyp:**
- **ALLE Pflichtfelder** werden automatisch erkannt und gefüllt
- Jeder Produkttyp (bottle, cup, knife, etc.) hat eigene Anforderungen
- Die App liest diese aus dem "AttributePTDMAP" Sheet und befüllt automatisch

**Von Produktdaten (wenn verfügbar):**
- SKU, Marke, EAN, Modellnummer
- Material, Farbe, Größe, Gewicht, Volumen
- Abmessungen (Länge, Breite, Höhe)
- Hersteller, Herkunftsland
- Und 40+ weitere Felder basierend auf Template-Anforderungen

**Intelligente Defaults:**
- Condition: "new_new"
- Fulfillment: "DEFAULT"
- Batterien: "false" (wenn nicht anders angegeben)
- Maßeinheiten: Automatisch (kg, ml, cm, etc.)

## ⚡ Performance

- Template-Analyse: ~1 Sekunde
- Content pro Produkt: ~3 Sekunden
- **10 Produkte:** ~30-40 Sekunden

## 🎯 Nächste Schritte

1. API Key von OpenAI holen
2. App starten
3. Mit Test-Dateien testen
4. Eigene Produktdaten verwenden

**Fertig!** 🎉
