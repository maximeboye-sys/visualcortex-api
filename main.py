"""
Visual Cortex — PPTX Generator API v4
Approche : duplication des slides template + remplacement du texte in-place.
La charte (photos, couleurs, positions) est préservée intégralement.
"""

import os
import io
import json
import zipfile
import time
import copy
from collections import defaultdict

import anthropic
from fastapi import FastAPI, File, Form, UploadFile, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pptx import Presentation
from pptx.oxml.ns import qn
from lxml import etree
import uvicorn

app = FastAPI(title="Visual Cortex API", version="4.0.0")

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
PRO_SECRET_TOKEN  = os.environ.get("PRO_SECRET_TOKEN", "change-me")
FREE_QUOTA_PER_IP = int(os.environ.get("FREE_QUOTA_PER_IP", "3"))

_usage: dict = defaultdict(list)
DAY_SECONDS = 86400


# ─────────────────────────────────────────────
# QUOTA
# ─────────────────────────────────────────────

def _get_ip(request: Request) -> str:
    fwd = request.headers.get("x-forwarded-for")
    return fwd.split(",")[0].strip() if fwd else request.client.host

def _is_pro(authorization: str | None) -> bool:
    if not authorization:
        return False
    return authorization.replace("Bearer ", "").strip() == PRO_SECRET_TOKEN

def _check_quota(ip: str) -> tuple[int, int]:
    now = time.time()
    _usage[ip] = [t for t in _usage[ip] if now - t < DAY_SECONDS]
    used = len(_usage[ip])
    if used >= FREE_QUOTA_PER_IP:
        raise HTTPException(429, detail={
            "error": "quota_exceeded",
            "message": f"Limite gratuite atteinte ({FREE_QUOTA_PER_IP}/jour). Passez en Pro.",
            "used": used, "max": FREE_QUOTA_PER_IP,
        })
    _usage[ip].append(now)
    return used + 1, FREE_QUOTA_PER_IP


# ─────────────────────────────────────────────
# 1. ANALYSE TEMPLATE
# ─────────────────────────────────────────────

def extract_brand_identity(pptx_bytes: bytes) -> dict:
    prs = Presentation(io.BytesIO(pptx_bytes))
    fonts, colors, layouts, slide_texts = set(), set(), [], []

    for slide in prs.slides:
        layouts.append(slide.slide_layout.name)
        texts = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                t = shape.text_frame.text.strip()
                if t:
                    texts.append(t[:80])
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if run.font.name:
                            fonts.add(run.font.name)
                        try:
                            if run.font.color and run.font.color.type:
                                colors.add(str(run.font.color.rgb))
                        except Exception:
                            pass
        if texts:
            slide_texts.append(" | ".join(texts[:3]))

    return {
        "fonts": list(fonts)[:5],
        "colors": list(colors)[:8],
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


def _analyze_template_slides(pptx_bytes: bytes) -> list:
    """
    Analyse les slides du template et retourne leur structure :
    index, layout, shapes avec leur texte.
    Permet de choisir la bonne slide template pour chaque type de contenu.
    """
    prs = Presentation(io.BytesIO(pptx_bytes))
    slides_info = []
    for i, slide in enumerate(prs.slides):
        shapes_info = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                shapes_info.append({
                    "name": shape.name,
                    "text": shape.text_frame.text.strip()[:100],
                })
        slides_info.append({
            "index": i,
            "layout": slide.slide_layout.name,
            "shapes": shapes_info,
        })
    return slides_info


# ─────────────────────────────────────────────
# 2. GÉNÉRATION CONTENU CLAUDE
# ─────────────────────────────────────────────

def generate_content_with_claude(
    prompt: str,
    brand: dict,
    nb_slides: int,
    template_slides_info: list,
) -> dict:
    if not ANTHROPIC_API_KEY:
        raise HTTPException(500, "Clé API Claude non configurée.")

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    # Décrit les slides template disponibles pour que Claude choisisse
    template_desc = ""
    for s in template_slides_info:
        template_desc += f"\n- Slide {s['index']} ({s['layout']}) : {', '.join([sh['text'][:40] for sh in s['shapes'] if sh['text']])}"

    system = """Tu es un expert en communication d'entreprise B2B et création de présentations PowerPoint.
Tu génères du contenu structuré, professionnel et percutant.
Réponds UNIQUEMENT avec un JSON valide, sans texte avant ou après, sans backticks."""

    user = f"""Génère une présentation PowerPoint professionnelle.

DEMANDE : {prompt}

CHARTE DE L'ENTREPRISE :
- Polices : {', '.join(brand.get('fonts', ['Arial']))}
- Couleurs thème : {', '.join(brand.get('theme_colors', [])[:5])}
- Exemples de textes : {' | '.join(brand.get('sample_texts', [])[:3])}

SLIDES DISPONIBLES DANS LE TEMPLATE :
{template_desc}

INSTRUCTIONS :
- Génère exactement {nb_slides} slides
- Pour chaque slide, choisis l'index de slide template le plus adapté (template_slide_index)
  * Couverture → utilise la slide 0 (couverture avec photo)
  * Conclusion/Merci → utilise la dernière slide du template
  * Contenu → utilise les slides du milieu
- Slide 1 = couverture, slide {nb_slides} = conclusion
- Contenu concis, langage B2B professionnel

FORMAT JSON :
{{
  "title": "Titre de la présentation",
  "slides": [
    {{
      "index": 1,
      "type": "cover",
      "template_slide_index": 0,
      "title": "Titre principal",
      "subtitle": "Sous-titre accroche",
      "body": [],
      "notes": "Notes présentateur"
    }},
    {{
      "index": 2,
      "type": "content",
      "template_slide_index": 2,
      "title": "Titre de la slide",
      "subtitle": "",
      "body": ["Point clé 1", "Point clé 2", "Point clé 3"],
      "notes": "Notes présentateur"
    }}
  ]
}}"""

    msg = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4000,
        system=system,
        messages=[{"role": "user", "content": user}],
    )

    raw = msg.content[0].text.strip()
    if "```" in raw:
        parts = raw.split("```")
        raw = parts[1] if len(parts) > 1 else parts[0]
        if raw.startswith("json"):
            raw = raw[4:]
    return json.loads(raw.strip())


# ─────────────────────────────────────────────
# 3. CONSTRUCTION PPTX PAR DUPLICATION
# ─────────────────────────────────────────────

def _duplicate_slide(prs: Presentation, slide_index: int) -> object:
    """
    Duplique une slide existante du template et l'ajoute à la fin.
    Préserve intégralement la charte : images, formes, couleurs, positions.
    """
    template_slide = prs.slides[slide_index]
    
    # Copie profonde du XML de la slide
    xml_copy = copy.deepcopy(template_slide._element)
    
    # Ajoute la nouvelle slide dans la présentation
    prs.slides._sldIdLst.append(
        prs.slides._sldIdLst[-1].__class__()
    )
    
    # Utilise le même layout que la slide source
    slide_layout = template_slide.slide_layout
    new_slide = prs.slides.add_slide(slide_layout)
    
    # Remplace le XML par la copie
    new_slide._element.getparent().replace(new_slide._element, xml_copy)
    
    # Met à jour l'ID de la slide
    sp_tree = prs.slides._sldIdLst
    last_id = max(int(el.get('id', 256)) for el in sp_tree) 
    xml_copy.set('id', str(last_id))
    
    return new_slide


def _set_text_in_shape(shape, new_text: str):
    """
    Remplace le texte d'une forme en préservant le formatage du premier run.
    """
    if not shape.has_text_frame:
        return
    tf = shape.text_frame
    
    # Préserve le format du premier run disponible
    first_run_rpr = None
    for para in tf.paragraphs:
        for run in para.runs:
            rpr = run._r.find(qn('a:rPr'))
            if rpr is not None:
                first_run_rpr = copy.deepcopy(rpr)
            break
        if first_run_rpr is not None:
            break

    # Vide tous les paragraphes sauf le premier
    for para in tf.paragraphs[1:]:
        p_elem = para._p
        p_elem.getparent().remove(p_elem)

    # Réécrit le premier paragraphe
    first_para = tf.paragraphs[0]
    # Supprime tous les runs
    for r in first_para._p.findall(qn('a:r')):
        first_para._p.remove(r)
    for br in first_para._p.findall(qn('a:br')):
        first_para._p.remove(br)

    lines = new_text.split('\n') if '\n' in new_text else [new_text]

    for i, line in enumerate(lines):
        if i == 0:
            # Premier run dans le premier paragraphe
            r_elem = etree.SubElement(first_para._p, qn('a:r'))
            if first_run_rpr is not None:
                r_elem.insert(0, copy.deepcopy(first_run_rpr))
            t_elem = etree.SubElement(r_elem, qn('a:t'))
            t_elem.text = line
        else:
            # Nouveau paragraphe pour chaque ligne
            new_p = etree.SubElement(tf._txBody, qn('a:p'))
            r_elem = etree.SubElement(new_p, qn('a:r'))
            if first_run_rpr is not None:
                r_elem.insert(0, copy.deepcopy(first_run_rpr))
            t_elem = etree.SubElement(r_elem, qn('a:t'))
            t_elem.text = line


def _find_main_title_shape(slide):
    """Trouve la forme titre principale d'une slide."""
    # Cherche d'abord un placeholder titre (idx=0)
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 0 and ph.has_text_frame:
            return ph
    # Sinon cherche par nom
    for shape in slide.shapes:
        if shape.has_text_frame:
            name = shape.name.lower()
            if 'title' in name or 'titre' in name:
                return shape
    return None


def _find_content_shapes(slide):
    """Retourne les formes de contenu (hors titre, hors numéro de page, hors pied de page)."""
    excluded_keywords = ['titre', 'title', 'numéro', 'number', 'pied', 'footer', 'slide number']
    shapes = []
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        name = shape.name.lower()
        if any(kw in name for kw in excluded_keywords):
            continue
        text = shape.text_frame.text.strip()
        if not text or text.isdigit():
            continue
        shapes.append(shape)
    return shapes


def build_pptx_from_template(pptx_bytes: bytes, content: dict) -> bytes:
    """
    Construit le PPTX final en dupliquant les slides template
    et en injectant le contenu généré par Claude.
    La charte graphique est préservée intégralement.
    """
    prs = Presentation(io.BytesIO(pptx_bytes))
    slides_data = content.get("slides", [])
    nb_template = len(prs.slides)

    if nb_template == 0:
        raise HTTPException(500, "Template vide.")

    # Indices des slides template disponibles
    cover_idx = 0
    conclusion_idx = nb_template - 1
    content_indices = list(range(1, nb_template - 1)) if nb_template > 2 else [0]

    # Collecte les slides originales avant modification
    original_slide_elements = [copy.deepcopy(slide._element) for slide in prs.slides]
    original_layouts = [slide.slide_layout for slide in prs.slides]

    # Supprime toutes les slides existantes
    sldIdLst = prs.slides._sldIdLst
    for ref in list(sldIdLst):
        sldIdLst.remove(ref)

    # Reconstruit slide par slide
    for sd in slides_data:
        slide_type = sd.get("type", "content")
        requested_idx = sd.get("template_slide_index", None)

        # Choisit l'index template
        if requested_idx is not None and 0 <= requested_idx < nb_template:
            tpl_idx = requested_idx
        elif slide_type == "cover":
            tpl_idx = cover_idx
        elif slide_type == "conclusion":
            tpl_idx = conclusion_idx
        else:
            # Alterne entre les slides de contenu
            ci = (sd.get("index", 1) - 2) % max(len(content_indices), 1)
            tpl_idx = content_indices[ci]

        # Ajoute une slide avec le bon layout
        layout = original_layouts[tpl_idx]
        new_slide = prs.slides.add_slide(layout)

        # Remplace le XML par une copie de la slide template
        tpl_xml = copy.deepcopy(original_slide_elements[tpl_idx])
        new_slide._element.getparent().replace(new_slide._element, tpl_xml)
        # Relie le bon layout
        tpl_xml.set(qn('r:id') if False else 'r:id', '')  # Railway gère les rels

        # Trouve le titre et injecte
        title_shape = _find_main_title_shape(new_slide)
        title_text = sd.get("title", "")
        if title_shape and title_text:
            try:
                _set_text_in_shape(title_shape, title_text)
            except Exception:
                pass

        # Trouve les zones de contenu et injecte
        body = sd.get("body", [])
        subtitle = sd.get("subtitle", "")
        content_text = "\n".join(body) if body else subtitle

        content_shapes = _find_content_shapes(new_slide)
        # Exclut la forme titre déjà remplie
        if title_shape:
            content_shapes = [s for s in content_shapes if s.name != title_shape.name]

        if content_text and content_shapes:
            try:
                _set_text_in_shape(content_shapes[0], content_text)
            except Exception:
                pass

        # Notes
        try:
            notes = sd.get("notes", "")
            if notes:
                new_slide.notes_slide.notes_text_frame.text = notes
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
    return {"status": "ok", "service": "Visual Cortex API 🚀 v4"}

@app.get("/health")
def health():
    return {"status": "healthy"}


@app.post("/analyze-template")
async def analyze_template(file: UploadFile = File(...)):
    if not file.filename.endswith(".pptx"):
        raise HTTPException(400, "Le fichier doit être un .pptx")
    try:
        data = await file.read()
        brand = extract_brand_identity(data)
    except Exception as e:
        raise HTTPException(500, f"Erreur analyse : {e}")
    return {"success": True, "brand": brand,
            "message": f"{brand['slide_count']} slides • {len(brand['fonts'])} polices"}


@app.post("/generate-preview")
async def generate_preview(
    request: Request,
    template: UploadFile = File(...),
    prompt: str = Form(...),
    nb_slides: int = Form(default=8),
    authorization: str = Form(default=None),
):
    if not template.filename.endswith(".pptx"):
        raise HTTPException(400, "Le template doit être un .pptx")
    if nb_slides < 3 or nb_slides > 30:
        raise HTTPException(400, "Entre 3 et 30 slides")

    pro = _is_pro(authorization)
    quota_info = {"plan": "pro"} if pro else {}
    if not pro:
        used, max_q = _check_quota(_get_ip(request))
        quota_info = {"used": used, "max": max_q, "plan": "free"}

    data = await template.read()
    brand = extract_brand_identity(data)
    tpl_info = _analyze_template_slides(data)
    content = generate_content_with_claude(prompt, brand, nb_slides, tpl_info)

    return {
        "success": True,
        "title": content.get("title"),
        "slides": [
            {
                "index": s.get("index"),
                "type": s.get("type"),
                "title": s.get("title"),
                "summary": s.get("subtitle") or (s.get("body", [""])[0] if s.get("body") else ""),
            }
            for s in content.get("slides", [])
        ],
        "brand": {"fonts": brand["fonts"], "theme_colors": brand["theme_colors"]},
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
    if not template.filename.endswith(".pptx"):
        raise HTTPException(400, "Le template doit être un .pptx")
    if nb_slides < 3 or nb_slides > 30:
        raise HTTPException(400, "Entre 3 et 30 slides")

    pro = _is_pro(authorization)
    if not pro:
        _check_quota(_get_ip(request))

    data = await template.read()
    brand = extract_brand_identity(data)
    tpl_info = _analyze_template_slides(data)

    try:
        content = generate_content_with_claude(prompt, brand, nb_slides, tpl_info)
    except json.JSONDecodeError as e:
        raise HTTPException(500, f"Erreur parsing IA : {e}")
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, f"Erreur Claude API : {e}")

    try:
        pptx_bytes = build_pptx_from_template(data, content)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, f"Erreur génération PPTX : {e}")

    raw_name = content.get("title", "presentation")[:40].replace(" ", "_")
    filename = "".join(c for c in raw_name if c.isalnum() or c in "_-.") + ".pptx"

    return StreamingResponse(
        io.BytesIO(pptx_bytes),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("main:app", host="0.0.0.0", port=port, reload=False)
