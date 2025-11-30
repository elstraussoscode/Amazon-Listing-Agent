#!/usr/bin/env python3
"""Test script to analyze why generated content exceeds byte limits"""

import os
from openai import OpenAI
from pydantic import BaseModel, Field
from typing import List

# Initialize client
client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))

class CosmoOptimizedContent(BaseModel):
    """COSMO/RUFUS optimized Amazon listing content"""
    model_config = {"extra": "forbid"}
    
    artikelname: str = Field(description="Optimierter Produkttitel: 170-200 BYTES. Marke + Produkt + USP + differenzierende Eigenschaften")
    produktbeschreibung: str = Field(description="Detaillierte Produktbeschreibung: 1700-2000 BYTES. Alle 15 COSMO-Beziehungstypen abdecken")
    bullet_points: List[str] = Field(description="5 Aufzählungspunkte: Je 170-200 BYTES. Verkaufsstark, nutzenorientiert, sauberes Deutsch", min_length=5, max_length=5)
    suchbegriffe: str = Field(description="Keywords: 225-250 BYTES (90%!). Synonyme, Long-Tail-Keywords, Kundeneigenschaften. KEINE komplementären Produkte!")


def get_byte_length(text: str) -> int:
    return len(text.encode('utf-8'))


# Simple test prompt
TEST_PROMPT = """Erstelle ein Amazon-Listing für folgendes Produkt:

Produktname: Edelstahl Trinkflasche
Material: 18/8 Edelstahl, doppelwandig
Kapazität: 750ml
Features: BPA-frei, vakuumisoliert, 24h kalt/12h warm
Farbe: Schwarz Matt

⚠️ KRITISCHE LÄNGEN-REGELN IN BYTES:
1. TITEL: MAXIMAL 200 BYTES (nicht mehr!)
2. BULLET POINTS: Je MAXIMAL 200 BYTES
3. BESCHREIBUNG: MAXIMAL 2000 BYTES  
4. KEYWORDS: MAXIMAL 250 BYTES

HINWEIS: Deutsche Umlaute (ä,ö,ü,ß) zählen als 2 Bytes!
"""

print("=" * 60)
print("Testing GPT-5.1 content generation with byte limits...")
print("=" * 60)

try:
    response = client.chat.completions.create(
        model="gpt-5.1",
        messages=[
            {"role": "system", "content": "Du bist ein Amazon SEO Experte. Halte dich STRIKT an die Byte-Limits!"},
            {"role": "user", "content": TEST_PROMPT}
        ],
        response_format={
            "type": "json_schema",
            "json_schema": {
                "name": "cosmo_content",
                "schema": CosmoOptimizedContent.model_json_schema(),
                "strict": True
            }
        }
    )
    
    content = CosmoOptimizedContent.model_validate_json(response.choices[0].message.content)
    
    print("\n📊 RESULTS:")
    print("-" * 60)
    
    # Title
    title_bytes = get_byte_length(content.artikelname)
    title_status = "✅" if title_bytes <= 200 else "❌ TOO LONG"
    print(f"\n📝 TITEL ({title_bytes}/200 bytes) {title_status}")
    print(f"   {content.artikelname}")
    
    # Bullets
    print(f"\n📌 BULLET POINTS:")
    for i, bp in enumerate(content.bullet_points, 1):
        bp_bytes = get_byte_length(bp)
        bp_status = "✅" if bp_bytes <= 200 else "❌ TOO LONG"
        print(f"   {i}. ({bp_bytes}/200 bytes) {bp_status}")
        print(f"      {bp[:100]}..." if len(bp) > 100 else f"      {bp}")
    
    # Description
    desc_bytes = get_byte_length(content.produktbeschreibung)
    desc_status = "✅" if desc_bytes <= 2000 else "❌ TOO LONG"
    print(f"\n📄 BESCHREIBUNG ({desc_bytes}/2000 bytes) {desc_status}")
    print(f"   {content.produktbeschreibung[:200]}...")
    
    # Keywords
    kw_bytes = get_byte_length(content.suchbegriffe)
    kw_status = "✅" if kw_bytes <= 250 else "❌ TOO LONG"
    print(f"\n🔑 KEYWORDS ({kw_bytes}/250 bytes) {kw_status}")
    print(f"   {content.suchbegriffe}")
    
    print("\n" + "=" * 60)
    print("SUMMARY:")
    print(f"  Title: {title_bytes}/200 bytes {'(OK)' if title_bytes <= 200 else '(EXCEEDED by ' + str(title_bytes-200) + ')'}")
    for i, bp in enumerate(content.bullet_points, 1):
        bp_bytes = get_byte_length(bp)
        print(f"  Bullet {i}: {bp_bytes}/200 bytes {'(OK)' if bp_bytes <= 200 else '(EXCEEDED by ' + str(bp_bytes-200) + ')'}")
    print(f"  Description: {desc_bytes}/2000 bytes {'(OK)' if desc_bytes <= 2000 else '(EXCEEDED by ' + str(desc_bytes-2000) + ')'}")
    print(f"  Keywords: {kw_bytes}/250 bytes {'(OK)' if kw_bytes <= 250 else '(EXCEEDED by ' + str(kw_bytes-250) + ')'}")
    print("=" * 60)

except Exception as e:
    print(f"Error: {e}")

