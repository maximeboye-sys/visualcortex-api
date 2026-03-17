"""
Visual Cortex — PPTX Generator API v6 (Shielded Edition)
Approche : Injection "bulldozer" + Bouclier anti-crash avec gestion CORS absolue.
"""

import os
import io
import json
import zipfile
import time
import copy
import re
import traceback
from collections import defaultdict

import anthropic
from fastapi import FastAPI, File, Form, UploadFile, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, StreamingResponse
from pptx import Presentation
from pptx.util import Pt
from pptx.oxml.ns import qn
import uvicorn

app = FastAPI(title="Visual Cortex API", version="6.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ─────────────────────────────────────────────
# 0. BOUCLIER ANTI-CRASH GLOBAL (Évite le "Failed to fetch")
# ─────────────────────────────────────────────
@app.exception_handler(Exception)
async def universal_exception_handler(request: Request, exc: Exception):
    """
    Capture TOUTES les erreurs Python critiques pour empêcher la perte des en-têtes CORS.
    Renvoie toujours un JSON propre au front-end au lieu de couper la connexion.
    """
    traceback.print_exc() # Affiche l'erreur complète dans les logs Railway
    return JSONResponse(
        status_code=500,
        content={"detail": {"message": f"Erreur critique du serveur : {str(exc)}"}},
        headers={"Access-Control-Allow-Origin": "*"}
    )


# ─────────────────────────────────────────────
# CONFIG & QUOTAS
# ─────────────────────────────────────────────

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
PRO_SECRET_TOKEN  = os.environ.get("PRO_SECRET_TOKEN", "change-me-in-railway")
FREE_QUOTA_PER_IP = int(os.environ.get("FREE_QUOTA_PER_IP", "3"))

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
        raise HTTPException(
            status_code=429,
            detail={"message": f"Limite gratuite atteinte ({FREE_QUOTA_PER_IP} présentations/jour)."}
        )
    _usage[ip].append(now)
    return used + 1, FREE_QUOTA_PER_IP

# ─────────────────────────────────────────────
# 1. ANALYSE APPROFONDIE DU TEMPLATE
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
            if getattr(shape, "has_text_frame", False):
                ph_idx = None
                if getattr(shape, "is_placeholder", False):
                    try: ph_idx = shape.placeholder_format.idx
                    except Exception: pass
                texts = [r.text.strip() for p in shape.text_frame.paragraphs for r in p.runs if r.text.strip()]
                if texts:
                    profile["text_placeholders"].append({"ph_idx": ph_idx, "texts": texts})
                    profile["text_content"].extend(texts)
            if shape.shape_type == 13: profile["has_image"] = True
            if getattr(shape, "has_table", False): profile["has_table"] = True
            if getattr(shape, "has_chart", False): profile["has_chart"] = True

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
            if getattr(shape, "has_text_frame", False):
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if run.font.name: fonts.add(run.font.name)
                        try:
                            if run.font.color and run.font.color.type: colors.add(str(run.font.color.rgb))
                        except Exception: pass
                        if run.text.strip(): slide_content.append(run.text.strip())
        if slide_content: slide_texts.append(" | ".join(slide_content[:6]))

    return {
        "fonts": list(fonts)[:5], "colors": list(colors)[:10],
        "theme_colors": _extract_theme_colors(pptx_bytes),
        "layouts": list(dict.fromkeys([p["layout_name"] for p in slide_profiles])),
        "slide_count": len(prs.slides), "slide_profiles": slide_profiles,
        "sample_texts": slide_texts,
    }

def _extract_theme_colors(pptx_bytes: bytes) -> list:
    try:
        with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
            theme_files = [f for f in z.namelist() if "theme/theme" in f]
            if theme_files:
                xml = z.read(theme_files[0]).decode("utf-8")
                return list(dict.fromkeys(re.findall(r'val="([0-9A-Fa-f]{6})"', xml)))[:8]
    except Exception: return []


# ─────────────────────────────────────────────
# 2. COPIE FIDÈLE & INJECTION ULTRA-SÉCURISÉE
# ─────────────────────────────────────────────

def duplicate_slide_xml(prs: Presentation, source_index: int) -> any:
    source_slide = prs.slides[source_index]
    new_slide = prs.slides.add_slide(source_slide.slide_layout)
    sp_tree = new_slide.shapes._spTree
    for child in list(sp_tree): sp_tree.remove(child)
    for child in source_slide.shapes._spTree: sp_tree.append(copy.deepcopy(child))
    return new_slide

def safe_inject_text(slide, title_text: str, content_lines: list) -> None:
    title_injected, body_injected = False, False

    def _apply_text(shape, lines: list) -> bool:
        if not getattr(shape, "has_text_frame", False): return False
        tf = shape.text_frame
        if not tf.paragraphs: return False

        ref_run_xml = copy.deepcopy(tf.paragraphs[0].runs[0]._r) if tf.paragraphs[0].runs else None
        tf.clear()
        
        for i, line in enumerate(lines):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            run = para.add_run()
            run.text = str(line)
            if ref_run_xml is not None:
                try:
                    rpr_tag = qn("a:rPr")
                    rpr = ref_run_xml.find(rpr_tag)
                    if rpr is not None:
                        if run._r.find(rpr_tag) is not None: run._r.remove(run._r.find(rpr_tag))
                        run._r.insert(0, copy.deepcopy(rpr))
                except Exception: pass
        return True

    # 1. Via Placeholders
    for shape in slide.placeholders:
        try:
            idx = shape.placeholder_format.idx
            if idx == 0 and title_text and not title_injected: title_injected = _apply_text(shape, [title_text])
            elif idx in [1, 2] and content_lines and not body_injected: body_injected = _apply_text(shape, content_lines)
        except Exception: pass

    # 2. Le Bulldozer (récupération des zones libres)
    if not title_injected or not body_injected:
        text_shapes = [s for s in slide.shapes if getattr(s, "has_text_frame", False) and not getattr(s, "is_placeholder", False)]
        if not title_injected and title_text and text_shapes: title_injected = _apply_text(text_shapes[0], [title_text])
        if not body_injected and content_lines and len(text_shapes) > 1: body_injected = _apply_text(text_shapes[1], content_lines)
        elif not body_injected and content_lines and text_shapes and not title_injected: _apply_text(text_shapes[0], content_lines)


# ─────────────────────────────────────────────
# 3. GÉNÉRATION CLAUDE (PROMPT BLINDÉ)
# ─────────────────────────────────────────────

def generate_content_with_claude(prompt: str, brand: dict, nb_slides: int) -> dict:
    if not ANTHROPIC_API_KEY: raise ValueError("Clé API Claude manquante.")
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    
    profiles = brand.get("slide_profiles", [])
    template_structure = "\n".join([
        f"  - Slide {p['index']+1} ({p['guessed_type']}) : textes : {' | '.join(p['text_content'][:3])}"
        for p in profiles
    ])

    user = f"""Demande : {prompt}
Structure du template original :
{template_structure}

Crée {nb_slides} slides concises (max 3 points par slide).
FORMAT JSON STRICT (pas de markdown) :
{{
  "title": "Titre présentation",
  "narrative": "Idée globale",
  "slides": [
    {{
      "index": 1,
      "type": "cover",
      "template_slide_index": 0,
      "title": "Titre",
      "subtitle": "Sous-titre",
      "body": []
    }},
    {{
      "index": 2,
      "type": "content",
      "template_slide_index": 1,
      "title": "Titre court",
      "subtitle": "",
      "body": ["Point 1", "Point 2"]
    }}
  ]
}}
IMPORTANT : "template_slide_index" doit être un entier entre 0 et {len(profiles)-1}."""

    msg = client.messages.create(
        model="claude-3-5-sonnet-20241022", max_tokens=6000,
        system="Tu es expert en B2B. Réponds uniquement en JSON valide.",
        messages=[{"role": "user", "content": user}],
    )
    
    raw = msg.content[0].text.strip()
    if "```" in raw:
        parts = raw.split("```")
        raw = parts[1] if len(parts) > 1 else parts[0]
        if raw.startswith("json"): raw = raw[4:]
    return json.loads(raw.strip())


# ─────────────────────────────────────────────
# 4. CONSTRUCTION PPTX
# ─────────────────────────────────────────────

def build_pptx_from_template(pptx_bytes: bytes, content: dict) -> bytes:
    prs = Presentation(io.BytesIO(pptx_bytes))
    slides_data = content.get("slides", [])
    nb_orig = len(list(prs.slides))
    if nb_orig == 0: raise ValueError("Template vide.")

    type_to_indices = defaultdict(list)
    for p in analyze_template_slides(prs): type_to_indices[p["guessed_type"]].append(p["index"])

    def best_template_index(slide_type: str, suggested: int) -> int:
        if suggested is not None and 0 <= suggested < nb_orig: return suggested
        if type_to_indices.get(slide_type): return type_to_indices[slide_type][0]
        return 0 if slide_type == "cover" else (nb_orig - 1 if slide_type == "conclusion" else (1 if nb_orig > 1 else 0))

    sldIdLst = prs.slides._sldIdLst
    for ref in list(sldIdLst): sldIdLst.remove(ref)

    for sd in slides_data:
        src_idx = best_template_index(sd.get("type", "content"), sd.get("template_slide_index"))
        try: new_slide = duplicate_slide_xml(prs, src_idx)
        except Exception: new_slide = prs.slides.add_slide(prs.slide_layouts[min(src_idx, len(prs.slide_layouts)-1)])

        title_text = str(sd.get("title", ""))
        body_lines = sd.get("body", [])
        if not isinstance(body_lines, list): body_lines = [str(body_lines)]
        if not body_lines and sd.get("subtitle"): body_lines = [str(sd.get("subtitle"))]

        safe_inject_text(new_slide, title_text, body_lines)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out.read()


# ─────────────────────────────────────────────
# 5. ROUTES
# ─────────────────────────────────────────────

@app.get("/")
def root(): return {"status": "ok", "version": "6.0.0"}

@app.post("/analyze-template")
async def analyze_template(file: UploadFile = File(...)):
    pptx_bytes = await file.read()
    brand = extract_brand_identity(pptx_bytes)
    return {"success": True, "brand": brand, "slide_count": brand["slide_count"]}

@app.post("/generate-preview")
async def generate_preview(request: Request, template: UploadFile = File(...), prompt: str = Form(...), nb_slides: int = Form(default=8), authorization: str = Form(default=None)):
    pro = _is_pro(authorization)
    quota_info = {"plan": "pro"} if pro else {"used": _check_and_increment_quota(_get_ip(request))[0], "plan": "free"}
    
    brand = extract_brand_identity(await template.read())
    content = generate_content_with_claude(prompt, brand, nb_slides)
    
    return {
        "success": True, "title": content.get("title", "Présentation"),
        "slides": [{"title": s.get("title", "")} for s in content.get("slides", [])],
        "quota": quota_info
    }

@app.post("/generate")
async def generate_presentation(request: Request, template: UploadFile = File(...), prompt: str = Form(...), nb_slides: int = Form(default=8), authorization: str = Form(default=None)):
    if not _is_pro(authorization): _check_and_increment_quota(_get_ip(request))
    
    template_bytes = await template.read()
    content = generate_content_with_claude(prompt, extract_brand_identity(template_bytes), nb_slides)
    pptx_bytes = build_pptx_from_template(template_bytes, content)
    
    return StreamingResponse(io.BytesIO(pptx_bytes), media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation", headers={"Content-Disposition": "attachment; filename=presentation.pptx"})

if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=int(os.environ.get("PORT", 8000)), reload=False)
