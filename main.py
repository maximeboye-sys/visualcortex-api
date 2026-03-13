"""
Visual Cortex — PPTX Generator API v5 (High-End B2B Edition)
Approche : Copie stricte de l'XML + Prompting "Consulting B2B" + Injection textuelle robuste (fallbacks)
"""

import os
import io
import json
import zipfile
import time
import copy
import re
from collections import defaultdict

import anthropic
from fastapi import FastAPI, File, Form, UploadFile, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pptx import Presentation
from pptx.util import Pt
from pptx.oxml.ns import qn
import uvicorn

app = FastAPI(title="Visual Cortex API", version="5.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ─────────────────────────────────────────────
# CONFIG & QUOTAS (inchangé)
# ─────────────────────────────────────────────

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
PRO_SECRET_TOKEN  = os.environ.get("PRO_SECRET_TOKEN", "change-me-in-railway")
FREE_QUOTA_PER_IP = int(os.environ.get("FREE_QUOTA_PER_IP", "50"))

_usage: dict = defaultdict(list)
DAY_SECONDS = 86400

def _get_ip(request: Request) -> str:
    forwarded = request.headers.get("x-forwarded-for")
    return forwarded.split(",")[0].strip() if forwarded else request.client.host

def _is_pro(authorization: str | None) -> bool:
    if not authorization: return False
    return authorization.replace("Bearer ", "").strip() == PRO_SECRET_TOKEN

def _check_and_increment_quota(ip: str) -> tuple[int, int]:
    now = time.time()
    _usage[ip] = [t for t in _usage[ip] if now - t < DAY_SECONDS]
    used = len(_usage[ip])
    if used >= FREE_QUOTA_PER_IP:
        raise HTTPException(status_code=429, detail={"error": "quota_exceeded"})
    _usage[ip].append(now)
    return used + 1, FREE_QUOTA_PER_IP


# ─────────────────────────────────────────────
# 1. ANALYSE DU TEMPLATE (inchangé)
# ─────────────────────────────────────────────

def analyze_template_slides(prs: Presentation) -> list[dict]:
    slide_profiles = []
    for i, slide in enumerate(prs.slides):
        profile = {
            "index": i, "layout_name": slide.slide_layout.name,
            "text_placeholders": [], "has_image": False,
            "has_table": False, "has_chart": False,
            "text_content": [], "guessed_type": "content",
        }
        for shape in slide.shapes:
            if shape.has_text_frame:
                ph_idx = shape.placeholder_format.idx if shape.is_placeholder else None
                texts = [r.text.strip() for p in shape.text_frame.paragraphs for r in p.runs if r.text.strip()]
                if texts:
                    profile["text_placeholders"].append({"ph_idx": ph_idx, "texts": texts})
                    profile["text_content"].extend(texts)
            if shape.shape_type == 13: profile["has_image"] = True
            if shape.has_table: profile["has_table"] = True
            if shape.has_chart: profile["has_chart"] = True

        layout = slide.slide_layout.name.lower()
        full_text = " ".join(profile["text_content"]).lower()
        if i == 0 or "couverture" in layout or "cover" in layout: profile["guessed_type"] = "cover"
        elif i == len(prs.slides) - 1 or "merci" in full_text or "conclusion" in full_text: profile["guessed_type"] = "conclusion"
        elif "section" in layout or "divider" in layout: profile["guessed_type"] = "section"
        elif profile["has_table"] or profile["has_chart"]: profile["guessed_type"] = "data"
        else: profile["guessed_type"] = "content"
        slide_profiles.append(profile)
    return slide_profiles

def extract_brand_identity(pptx_bytes: bytes) -> dict:
    prs = Presentation(io.BytesIO(pptx_bytes))
    fonts, colors, slide_texts = set(), set(), []
    slide_profiles = analyze_template_slides(prs)
    for slide in prs.slides:
        slide_content = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if run.font.name: fonts.add(run.font.name)
                        if run.text.strip(): slide_content.append(run.text.strip())
        if slide_content: slide_texts.append(" | ".join(slide_content[:6]))

    return {
        "fonts": list(fonts)[:5],
        "layouts": list(dict.fromkeys([p["layout_name"] for p in slide_profiles])),
        "slide_count": len(prs.slides),
        "slide_profiles": slide_profiles,
        "sample_texts": slide_texts,
    }


# ─────────────────────────────────────────────
# 2. INJECTION TEXTUELLE ROBUSTE (Amélioré)
# ─────────────────────────────────────────────

def duplicate_slide_xml(prs: Presentation, source_index: int) -> any:
    source_slide = prs.slides[source_index]
    new_slide = prs.slides.add_slide(source_slide.slide_layout)
    sp_tree = new_slide.shapes._spTree
    for child in list(sp_tree): sp_tree.remove(child)
    source_sp_tree = source_slide.shapes._spTree
    for child in source_sp_tree: sp_tree.append(copy.deepcopy(child))
    return new_slide

def safe_inject_text(slide, title_text: str, content_lines: list[str]):
    """
    Tente d'abord les placeholders. Si échec, cherche les zones de texte libres.
    Préserve le style du premier caractère trouvé dans la zone cible.
    """
    title_injected = False
    body_injected = False

    # 1. Tentative d'injection propre via Placeholders
    for shape in slide.placeholders:
        if shape.placeholder_format.idx == 0 and title_text:
            _apply_text_to_shape(shape, [title_text])
            title_injected = True
        elif shape.placeholder_format.idx in [1, 2] and content_lines:
            _apply_text_to_shape(shape, content_lines)
            body_injected = True

    # 2. Fallback (Secours) si les placeholders n'ont pas fonctionné
    if not title_injected or not body_injected:
        text_shapes = [s for s in slide.shapes if s.has_text_frame]
        # Trie par taille de police ou position pour deviner le titre du corps
        if text_shapes:
            # On suppose que la première forme est le titre, la deuxième le corps
            if not title_injected and title_text:
                _apply_text_to_shape(text_shapes[0], [title_text])
            if not body_injected and content_lines and len(text_shapes) > 1:
                _apply_text_to_shape(text_shapes[1], content_lines)

def _apply_text_to_shape(shape, lines: list[str]):
    tf = shape.text_frame
    if not tf.paragraphs: return
    
    # Sauvegarde du style d'origine pour le réappliquer
    ref_run_xml = None
    if tf.paragraphs[0].runs:
        ref_run_xml = copy.deepcopy(tf.paragraphs[0].runs[0]._r)

    tf.clear()
    for i, line in enumerate(lines):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        run = para.add_run()
        run.text = line
        if ref_run_xml is not None:
            try:
                rpr_tag = qn("a:rPr")
                rpr = ref_run_xml.find(rpr_tag)
                if rpr is not None:
                    existing_rpr = run._r.find(rpr_tag)
                    if existing_rpr is not None: run._r.remove(existing_rpr)
                    run._r.insert(0, copy.deepcopy(rpr))
            except: pass


# ─────────────────────────────────────────────
# 3. GÉNÉRATION CLAUDE (PROMPT CONSULTING B2B)
# ─────────────────────────────────────────────

def generate_content_with_claude(prompt: str, brand: dict, nb_slides: int) -> dict:
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    profiles = brand.get("slide_profiles", [])
    
    template_structure = "\n".join([
        f"  - Slide {p['index']+1} ({p['guessed_type']}) : layout '{p['layout_name']}' | Textes type : {' | '.join(p['text_content'][:2])}"
        for p in profiles
    ])

    system = """Tu es un Partner dans un grand cabinet de conseil en stratégie (ex: McKinsey, BCG). 
Ton rôle est de structurer et rédiger des présentations B2B (Slide Decks) à fort impact.
Règles d'or du copywriting B2B :
1. Pyramide de Minto : Commence toujours par la conclusion/l'idée maîtresse, puis développe les arguments.
2. MECE (Mutually Exclusive, Collectively Exhaustive) : Les points d'une slide ne doivent pas se chevaucher.
3. Titres d'action (Action Titles) : Le titre de la slide DOIT être une phrase complète qui résume le message clé de la slide (ex: "L'automatisation réduit les coûts de 30% d'ici 2025" au lieu de "Introduction").
4. Concision : Supprime les mots inutiles. Sois percutant, orienté bénéfices et résultats.
5. Calibrage : Adapte-toi parfaitement aux types de slides proposés dans le template.
Réponds UNIQUEMENT avec un JSON valide."""

    user = f"""Génère un deck stratégique basé sur la demande suivante : "{prompt}".
La présentation doit contenir exactement {nb_slides} slides.

ANALYSE DU TEMPLATE CLIENT (Pour calquer la structure) :
{template_structure}

Format JSON strict requis :
{{
  "title": "Nom du fichier (court)",
  "narrative": "L'Action Title global de la présentation",
  "slides": [
    {{
      "index": 1,
      "type": "cover",
      "template_slide_index": 0,
      "title": "Titre principal accrocheur (Orienté valeur)",
      "subtitle": "Sous-titre explicatif ou chiffre choc",
      "notes": "Script pour l'orateur."
    }},
    {{
      "index": 2,
      "type": "content",
      "template_slide_index": 1,
      "title": "Action Title: [Phrase complète résumant l'insight de la slide]",
      "body": [
        "Insight fort 1 (verbe d'action, data si possible)",
        "Insight fort 2 (impact business)",
        "Insight fort 3 (preuve ou next step)"
      ]
    }}
  ]
}}

IMPORTANT : "template_slide_index" doit correspondre au meilleur index de slide du template à copier (entre 0 et {len(profiles)-1})."""

    msg = client.messages.create(
        model="claude-3-5-sonnet-20241022", # Utilise la dernière version de Sonnet !
        max_tokens=8000,
        system=system,
        messages=[{"role": "user", "content": user}],
    )

    raw = msg.content[0].text.strip()
    if "```" in raw:
        parts = raw.split("```")
        raw = parts[1] if len(parts) > 1 else parts[0]
        if raw.startswith("json"): raw = raw[4:]
    return json.loads(raw.strip())


# ─────────────────────────────────────────────
# 4. CONSTRUCTION DU PPTX (Simplifié et Robuste)
# ─────────────────────────────────────────────

def build_pptx_from_template(pptx_bytes: bytes, content: dict) -> bytes:
    prs = Presentation(io.BytesIO(pptx_bytes))
    slides_data = content.get("slides", [])
    nb_orig = len(list(prs.slides))

    type_to_indices = defaultdict(list)
    for p in analyze_template_slides(prs):
        type_to_indices[p["guessed_type"]].append(p["index"])

    def best_template_index(slide_type: str, suggested_index: int) -> int:
        if suggested_index is not None and 0 <= suggested_index < nb_orig: return suggested_index
        candidates = type_to_indices.get(slide_type, [])
        if candidates: return candidates[0]
        if slide_type == "cover": return 0
        if slide_type == "conclusion": return nb_orig - 1
        return 1 if nb_orig > 1 else 0

    # Purge des slides originales
    sldIdLst = prs.slides._sldIdLst
    for ref in list(sldIdLst): sldIdLst.remove(ref)

    # Re-création
    for sd in slides_data:
        slide_type = sd.get("type", "content")
        src_idx = best_template_index(slide_type, sd.get("template_slide_index"))
        
        try:
            new_slide = duplicate_slide_xml(prs, src_idx)
        except Exception:
            layout = prs.slide_layouts[min(src_idx, len(prs.slide_layouts)-1)]
            new_slide = prs.slides.add_slide(layout)

        title_text = sd.get("title", "")
        body_lines = sd.get("body", [])
        if not body_lines and sd.get("subtitle"):
            body_lines = [sd.get("subtitle")]

        # Utilisation de notre nouvelle fonction d'injection robuste
        safe_inject_text(new_slide, title_text, body_lines)

        if sd.get("notes"):
            try: new_slide.notes_slide.notes_text_frame.text = sd.get("notes")
            except Exception: pass

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out.read()


# ─────────────────────────────────────────────
# 5. ROUTES (Inchangées, gardées minimales)
# ─────────────────────────────────────────────

@app.post("/generate")
async def generate_presentation(
    request: Request,
    template: UploadFile = File(...),
    prompt: str = Form(...),
    nb_slides: int = Form(default=8)
):
    template_bytes = await template.read()
    brand = extract_brand_identity(template_bytes)
    content = generate_content_with_claude(prompt, brand, nb_slides)
    pptx_bytes = build_pptx_from_template(template_bytes, content)
    
    filename = "presentation_b2b.pptx"
    return StreamingResponse(
        io.BytesIO(pptx_bytes),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=int(os.environ.get("PORT", 8000)), reload=False)
