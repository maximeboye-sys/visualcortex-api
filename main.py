"""
Visual Cortex — PPTX Generator API v3
"""

import os
import io
import json
import zipfile
import time
from collections import defaultdict

import anthropic
from fastapi import FastAPI, File, Form, UploadFile, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pptx import Presentation
from pptx.util import Pt
import uvicorn

app = FastAPI(title="Visual Cortex API", version="3.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
PRO_SECRET_TOKEN  = os.environ.get("PRO_SECRET_TOKEN", "change-me-in-railway")
FREE_QUOTA_PER_IP = int(os.environ.get("FREE_QUOTA_PER_IP", "3"))

_usage: dict = defaultdict(list)
DAY_SECONDS = 86400


# ─────────────────────────────────────────────
# QUOTA
# ─────────────────────────────────────────────

def _get_ip(request: Request) -> str:
    forwarded = request.headers.get("x-forwarded-for")
    return forwarded.split(",")[0].strip() if forwarded else request.client.host

def _is_pro(authorization: str | None) -> bool:
    if not authorization:
        return False
    return authorization.replace("Bearer ", "").strip() == PRO_SECRET_TOKEN

def _check_and_increment_quota(ip: str) -> tuple[int, int]:
    now = time.time()
    _usage[ip] = [t for t in _usage[ip] if now - t < DAY_SECONDS]
    used = len(_usage[ip])
    if used >= FREE_QUOTA_PER_IP:
        raise HTTPException(
            status_code=429,
            detail={
                "error": "quota_exceeded",
                "message": f"Limite gratuite atteinte ({FREE_QUOTA_PER_IP} présentations/jour). Passez en Pro pour un accès illimité.",
                "used": used,
                "max": FREE_QUOTA_PER_IP,
            }
        )
    _usage[ip].append(now)
    return used + 1, FREE_QUOTA_PER_IP


# ─────────────────────────────────────────────
# 1. EXTRACTION CHARTE
# ─────────────────────────────────────────────

def extract_brand_identity(pptx_bytes: bytes) -> dict:
    prs = Presentation(io.BytesIO(pptx_bytes))
    fonts, colors, layouts, slide_texts = set(), set(), [], []

    for slide in prs.slides:
        layouts.append(slide.slide_layout.name)
        slide_content = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if run.font.name:
                            fonts.add(run.font.name)
                        try:
                            if run.font.color and run.font.color.type:
                                colors.add(str(run.font.color.rgb))
                        except Exception:
                            pass
                        if run.text.strip():
                            slide_content.append(run.text.strip())
        if slide_content:
            slide_texts.append(" | ".join(slide_content[:5]))

    return {
        "fonts": list(fonts)[:5],
        "colors": list(colors)[:10],
        "theme_colors": _extract_theme_colors(pptx_bytes),
        "layouts": list(dict.fromkeys(layouts)),
        "slide_count": len(prs.slides),
        "sample_texts": slide_texts,
    }


def _extract_theme_colors(pptx_bytes: bytes) -> list:
    import re
    try:
        with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
            theme_files = [f for f in z.namelist() if "theme/theme" in f]
            if theme_files:
                xml = z.read(theme_files[0]).decode("utf-8")
                found = re.findall(r'val="([0-9A-Fa-f]{6})"', xml)
                return list(dict.fromkeys(found))[:8]
    except Exception:
        pass
    return []


# ─────────────────────────────────────────────
# 2. GÉNÉRATION CONTENU CLAUDE
# ─────────────────────────────────────────────

def generate_content_with_claude(prompt: str, brand: dict, nb_slides: int) -> dict:
    if not ANTHROPIC_API_KEY:
        raise HTTPException(500, "Clé API Claude non configurée sur le serveur.")

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    system = """Tu es un expert en communication d'entreprise et création de présentations PowerPoint B2B.
Tu génères du contenu structuré, professionnel et percutant en respectant le ton éditorial fourni.
Réponds UNIQUEMENT avec un JSON valide, sans texte avant ou après, sans backticks."""

    user = f"""Génère une présentation PowerPoint professionnelle.

DEMANDE : {prompt}

CHARTE DE L'ENTREPRISE :
- Polices : {', '.join(brand.get('fonts', ['Arial']))}
- Couleurs thème : {', '.join(brand.get('theme_colors', [])[:5])}
- Layouts disponibles : {', '.join(brand.get('layouts', []))}
- Exemples de textes : {' | '.join(brand.get('sample_texts', [])[:3])}

CONTRAINTES :
- Exactement {nb_slides} slides
- Respecte le ton et le wording des exemples
- Slide 1 = couverture, slide {nb_slides} = conclusion
- Contenu concis, orienté action, langage B2B
- Le champ "title" de la présentation doit toujours être une chaîne non vide

FORMAT JSON attendu :
{{
  "title": "Titre de la présentation",
  "slides": [
    {{
      "index": 1,
      "type": "cover",
      "title": "Titre",
      "subtitle": "Accroche",
      "notes": "Notes présentateur"
    }},
    {{
      "index": 2,
      "type": "content",
      "title": "Titre slide",
      "body": ["Point 1", "Point 2", "Point 3"],
      "notes": "Notes présentateur"
    }}
  ]
}}

Types disponibles : cover | content | section | conclusion"""

    msg = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4000,
        system=system,
        messages=[{"role": "user", "content": user}],
    )

    raw = msg.content[0].text.strip()
    if "```" in raw:
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    return json.loads(raw.strip())


# ─────────────────────────────────────────────
# 3. CONSTRUCTION PPTX
# ─────────────────────────────────────────────

def build_pptx_from_template(pptx_bytes: bytes, content: dict) -> bytes:
    prs = Presentation(io.BytesIO(pptx_bytes))
    slides_data = content.get("slides", [])

    layout_map = {}
    for layout in prs.slide_layouts:
        n = layout.name.lower()
        if "couverture" in n or "cover" in n:
            layout_map["cover"] = layout
        elif "titre" in n or "title" in n:
            layout_map["section"] = layout
        elif "texte" in n or "content" in n or "bullet" in n:
            layout_map["content"] = layout

    available = list(prs.slide_layouts)
    layout_map.setdefault("cover", available[0])
    layout_map.setdefault("content", available[min(1, len(available) - 1)])
    layout_map.setdefault("section", layout_map["content"])
    layout_map["conclusion"] = layout_map["cover"]

    sldIdLst = prs.slides._sldIdLst
    for ref in list(sldIdLst):
        sldIdLst.remove(ref)

    for sd in slides_data:
        layout = layout_map.get(sd.get("type", "content"), layout_map["content"])

        try:
            slide = prs.slides.add_slide(layout)
        except Exception as e:
            raise HTTPException(500, f"Impossible d'ajouter une slide : {e}")

        try:
            if slide.shapes.title:
                slide.shapes.title.text = sd.get("title") or ""
        except Exception:
            pass

        body = sd.get("body") or []
        subtitle = sd.get("subtitle") or ""

        for ph in slide.placeholders:
            try:
                if ph.placeholder_format.idx == 1:
                    tf = ph.text_frame
                    tf.clear()
                    if body:
                        tf.text = body[0]
                        for pt in body[1:]:
                            p = tf.add_paragraph()
                            p.text = pt
                    elif subtitle:
                        tf.text = subtitle
                    break
            except Exception:
                pass

        try:
            notes = sd.get("notes") or ""
            if notes:
                slide.notes_slide.notes_text_frame.text = notes
        except Exception:
            pass

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out.read()


# ─────────────────────────────────────────────
# 4. ROUTES
# ─────────────────────────────────────────────

@app.get("/")
def root():
    return {"status": "ok", "service": "Visual Cortex API 🚀"}


@app.get("/health")
def health():
    return {"status": "healthy"}


@app.post("/analyze-template")
async def analyze_template(file: UploadFile = File(...)):
    """Analyse un template PPTX. Gratuit, sans quota."""
    if not file.filename.endswith(".pptx"):
        raise HTTPException(400, "Le fichier doit être un .pptx")
    try:
        brand = extract_brand_identity(await file.read())
    except Exception as e:
        raise HTTPException(500, f"Erreur analyse template : {e}")
    return {
        "success": True,
        "brand": brand,
        "message": f"{brand['slide_count']} slides • {len(brand['fonts'])} polices détectées"
    }


@app.post("/generate-preview")
async def generate_preview(
    request: Request,
    template: UploadFile = File(...),
    prompt: str = Form(...),
    nb_slides: int = Form(default=8),
    authorization: str = Form(default=None),
):
    """Génère le plan sans créer le fichier."""
    if not template.filename.endswith(".pptx"):
        raise HTTPException(400, "Le template doit être un .pptx")
    if nb_slides < 3 or nb_slides > 30:
        raise HTTPException(400, "Entre 3 et 30 slides")

    pro = _is_pro(authorization)
    quota_info = {"plan": "pro"} if pro else {}
    if not pro:
        used, max_q = _check_and_increment_quota(_get_ip(request))
        quota_info = {"used": used, "max": max_q, "plan": "free"}

    brand = extract_brand_identity(await template.read())
    content = generate_content_with_claude(prompt, brand, nb_slides)

    return {
        "success": True,
        "title": content.get("title") or "Présentation",
        "slides": [
            {
                "index": s.get("index"),
                "type": s.get("type"),
                "title": s.get("title") or "",
                "summary": s.get("subtitle") or (s.get("body", [""])[0] if s.get("body") else ""),
            }
            for s in content.get("slides", [])
        ],
        "brand": {
            "fonts": brand["fonts"],
            "theme_colors": brand["theme_colors"],
            "layouts": brand["layouts"],
        },
        "quota": quota_info,
    }


@app.post("/generate")
async def generate_presentation(
    request: Request,
    template: UploadFile = File(...),
    prompt: str = Form(...),
    nb_slides: int = Form(default=8),
    authorization: str = Form(default=None),
):
    """Génère et retourne le .pptx chartée."""
    if not template.filename.endswith(".pptx"):
        raise HTTPException(400, "Le template doit être un .pptx")
    if nb_slides < 3 or nb_slides > 30:
        raise HTTPException(400, "Entre 3 et 30 slides")

    pro = _is_pro(authorization)
    if not pro:
        _check_and_increment_quota(_get_ip(request))

    template_bytes = await template.read()
    brand = extract_brand_identity(template_bytes)

    try:
        content = generate_content_with_claude(prompt, brand, nb_slides)
    except json.JSONDecodeError as e:
        raise HTTPException(500, f"Erreur parsing contenu IA : {e}")
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, f"Erreur Claude API : {e}")

    try:
        pptx_bytes = build_pptx_from_template(template_bytes, content)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, f"Erreur génération PPTX : {e}")

    # Nom de fichier sécurisé
    raw_name = (content.get("title") or "presentation")[:40].replace(" ", "_")
    filename = "".join(c for c in raw_name if c.isalnum() or c in "_-.") + ".pptx"

    return StreamingResponse(
        io.BytesIO(pptx_bytes),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("main:app", host="0.0.0.0", port=port, reload=False)
