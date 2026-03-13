"""
Visual Cortex — PPTX Generator API v4
Approche : copie fidèle des slides du template + remplacement intelligent du texte.
Résultat : présentation 100% chartée, visuellement identique au template original.
"""

import os
import io
import json
import zipfile
import time
import copy
import re
from collections import defaultdict
from lxml import etree

import anthropic
from fastapi import FastAPI, File, Form, UploadFile, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pptx import Presentation
from pptx.util import Pt
from pptx.oxml.ns import qn
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
PRO_SECRET_TOKEN  = os.environ.get("PRO_SECRET_TOKEN", "change-me-in-railway")
FREE_QUOTA_PER_IP = int(os.environ.get("FREE_QUOTA_PER_IP", "50"))

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
# 1. ANALYSE APPROFONDIE DU TEMPLATE
# ─────────────────────────────────────────────

def analyze_template_slides(prs: Presentation) -> list[dict]:
    """
    Analyse chaque slide du template pour comprendre sa structure et son rôle.
    Retourne une liste de descripteurs de slides utilisables pour la génération.
    """
    slide_profiles = []

    for i, slide in enumerate(prs.slides):
        profile = {
            "index": i,
            "layout_name": slide.slide_layout.name,
            "text_placeholders": [],
            "has_image": False,
            "has_table": False,
            "has_chart": False,
            "text_content": [],
            "guessed_type": "content",
        }

        for shape in slide.shapes:
            if shape.has_text_frame:
                ph_idx = None
                if shape.is_placeholder:
                    try:
                        ph_idx = shape.placeholder_format.idx
                    except Exception:
                        pass
                texts = [r.text.strip() for p in shape.text_frame.paragraphs for r in p.runs if r.text.strip()]
                if texts:
                    profile["text_placeholders"].append({"ph_idx": ph_idx, "texts": texts})
                    profile["text_content"].extend(texts)

            if shape.shape_type == 13:  # MSO_SHAPE_TYPE.PICTURE
                profile["has_image"] = True
            if shape.has_table:
                profile["has_table"] = True
            if shape.has_chart:
                profile["has_chart"] = True

        # Devine le type de slide
        layout = slide.slide_layout.name.lower()
        full_text = " ".join(profile["text_content"]).lower()
        if i == 0 or "couverture" in layout or "cover" in layout:
            profile["guessed_type"] = "cover"
        elif i == len(prs.slides) - 1 or "merci" in full_text or "conclusion" in full_text:
            profile["guessed_type"] = "conclusion"
        elif "section" in layout or "divider" in layout:
            profile["guessed_type"] = "section"
        elif profile["has_table"] or profile["has_chart"]:
            profile["guessed_type"] = "data"
        else:
            profile["guessed_type"] = "content"

        slide_profiles.append(profile)

    return slide_profiles


def extract_brand_identity(pptx_bytes: bytes) -> dict:
    """Extrait la charte graphique complète du template."""
    prs = Presentation(io.BytesIO(pptx_bytes))
    fonts, colors, slide_texts = set(), set(), []

    slide_profiles = analyze_template_slides(prs)

    for slide in prs.slides:
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
            slide_texts.append(" | ".join(slide_content[:6]))

    return {
        "fonts": list(fonts)[:5],
        "colors": list(colors)[:10],
        "theme_colors": _extract_theme_colors(pptx_bytes),
        "layouts": list(dict.fromkeys([p["layout_name"] for p in slide_profiles])),
        "slide_count": len(prs.slides),
        "slide_profiles": slide_profiles,
        "sample_texts": slide_texts,
    }


def _extract_theme_colors(pptx_bytes: bytes) -> list:
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
# 2. COPIE FIDÈLE D'UNE SLIDE (cœur du système)
# ─────────────────────────────────────────────

def duplicate_slide_xml(prs: Presentation, source_index: int) -> any:
    """
    Duplique une slide du template en copiant fidèlement tout son XML.
    Préserve : formes, couleurs, images, positions, polices, styles.
    Retourne la nouvelle slide ajoutée à la présentation.
    """
    source_slide = prs.slides[source_index]

    # Copie profonde de l'XML de la slide source
    xml_copy = copy.deepcopy(source_slide._element)

    # Ajoute la nouvelle slide en utilisant le même layout
    new_slide = prs.slides.add_slide(source_slide.slide_layout)

    # Remplace l'arbre XML de la nouvelle slide par la copie
    # (on garde uniquement les relations, on remplace le contenu)
    sp_tree = new_slide.shapes._spTree
    # Supprime tous les éléments existants sauf les relations de layout
    for child in list(sp_tree):
        sp_tree.remove(child)

    # Copie tous les éléments de la slide source
    source_sp_tree = source_slide.shapes._spTree
    for child in source_sp_tree:
        sp_tree.append(copy.deepcopy(child))

    return new_slide


def replace_text_in_slide(slide, replacements: dict[str, str]) -> None:
    """
    Remplace du texte dans une slide en préservant TOUT le formatage.
    replacements = {"ancien texte": "nouveau texte"}
    """
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                for old_text, new_text in replacements.items():
                    if old_text and old_text in run.text:
                        run.text = run.text.replace(old_text, new_text or "")


def set_placeholder_text(slide, ph_idx: int, lines: list[str]) -> bool:
    """
    Remplace le contenu d'un placeholder par une liste de lignes.
    Préserve le style du premier run trouvé.
    Retourne True si le placeholder a été trouvé.
    """
    for shape in slide.placeholders:
        try:
            if shape.placeholder_format.idx != ph_idx:
                continue

            tf = shape.text_frame
            if not tf.paragraphs:
                return True

            # Récupère le style du premier run pour le préserver
            first_para = tf.paragraphs[0]
            ref_run_xml = None
            if first_para.runs:
                ref_run_xml = copy.deepcopy(first_para.runs[0]._r)

            # Vide le text frame
            tf.clear()

            # Réécrit les lignes en préservant le style
            for i, line in enumerate(lines):
                if i == 0:
                    para = tf.paragraphs[0]
                else:
                    para = tf.add_paragraph()

                run = para.add_run()
                run.text = line

                # Applique le style de référence si disponible
                if ref_run_xml is not None:
                    try:
                        rpr_tag = qn("a:rPr")
                        rpr = ref_run_xml.find(rpr_tag)
                        if rpr is not None:
                            existing_rpr = run._r.find(rpr_tag)
                            if existing_rpr is not None:
                                run._r.remove(existing_rpr)
                            run._r.insert(0, copy.deepcopy(rpr))
                    except Exception:
                        pass

            return True
        except Exception:
            continue
    return False


# ─────────────────────────────────────────────
# 3. GÉNÉRATION CONTENU CLAUDE
# ─────────────────────────────────────────────

def generate_content_with_claude(prompt: str, brand: dict, nb_slides: int) -> dict:
    if not ANTHROPIC_API_KEY:
        raise HTTPException(500, "Clé API Claude non configurée sur le serveur.")

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    # Prépare le contexte des slides du template
    profiles = brand.get("slide_profiles", [])
    template_structure = "\n".join([
        f"  - Slide {p['index']+1} ({p['guessed_type']}) : layout '{p['layout_name']}' | textes : {' | '.join(p['text_content'][:3])}"
        for p in profiles
    ])

    system = """Tu es un expert senior en communication d'entreprise, storytelling B2B et design de présentations.
Tu crées des présentations PowerPoint professionnelles, percutantes et visuellement cohérentes.
Tu analyses la structure et le ton d'un template existant pour en reproduire l'esprit tout en adaptant le contenu.
Réponds UNIQUEMENT avec un JSON valide, sans texte avant ou après, sans backticks markdown."""

    user = f"""Génère une présentation PowerPoint B2B professionnelle et complète.

DEMANDE DU CLIENT : {prompt}

ANALYSE DU TEMPLATE D'ENTREPRISE :
Structure des slides originales :
{template_structure}

Exemples de wording utilisé dans ce template :
{chr(10).join(brand.get('sample_texts', [])[:4])}

Polices de la charte : {', '.join(brand.get('fonts', ['Arial']))}
Couleurs du thème : {', '.join(brand.get('theme_colors', [])[:6])}

INSTRUCTIONS IMPORTANTES :
1. Génère exactement {nb_slides} slides
2. Slide 1 = couverture avec titre accrocheur et sous-titre percutant
3. Slide {nb_slides} = conclusion avec call-to-action ou message fort
4. Respecte scrupuleusement le ton et le wording des exemples fournis
5. Chaque slide doit avoir un objectif narratif clair (pas juste une liste de bullets)
6. Pour les slides "content" : max 4 points, formulations concises et impactantes
7. Pour les slides "section" : titre court et accrocheur, sous-titre optionnel
8. Assure une progression narrative cohérente (problème → solution → bénéfices → action)
9. Le champ "title" de la présentation doit être une string non vide

MAPPING avec le template (utilise les types existants) :
Types disponibles : {', '.join(list(dict.fromkeys([p['guessed_type'] for p in profiles])))}

FORMAT JSON STRICT :
{{
  "title": "Titre de la présentation",
  "narrative": "Résumé en 1 phrase du fil conducteur de la présentation",
  "slides": [
    {{
      "index": 1,
      "type": "cover",
      "template_slide_index": 0,
      "title": "Titre principal accrocheur",
      "subtitle": "Sous-titre ou accroche forte",
      "notes": "Note présentateur : contexte et intention de cette slide"
    }},
    {{
      "index": 2,
      "type": "content",
      "template_slide_index": 1,
      "title": "Titre de la slide",
      "body": [
        "Point clé 1 formulé de façon percutante",
        "Point clé 2 avec chiffre ou fait concret si pertinent",
        "Point clé 3 orienté bénéfice"
      ],
      "notes": "Note présentateur : comment présenter cette slide"
    }}
  ]
}}

IMPORTANT : "template_slide_index" doit être un entier entre 0 et {len(profiles)-1} indiquant quelle slide du template copier visuellement."""

    msg = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=6000,
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
# 4. CONSTRUCTION PPTX FIDÈLE AU TEMPLATE
# ─────────────────────────────────────────────

def build_pptx_from_template(pptx_bytes: bytes, content: dict) -> bytes:
    """
    Construit le PPTX final en :
    1. Copiant les slides du template (préserve tout le visuel)
    2. Remplaçant uniquement le texte avec le contenu généré
    """
    prs = Presentation(io.BytesIO(pptx_bytes))
    slides_data = content.get("slides", [])
    original_slides = list(prs.slides)
    nb_orig = len(original_slides)

    if nb_orig == 0:
        raise HTTPException(500, "Le template ne contient aucune slide.")

    # Identifie les indices de slides template par type
    type_to_indices: dict[str, list[int]] = defaultdict(list)
    profiles = analyze_template_slides(prs)
    for p in profiles:
        type_to_indices[p["guessed_type"]].append(p["index"])

    def best_template_index(slide_type: str, suggested_index: int | None) -> int:
        """Choisit le meilleur index de slide template à copier."""
        # Si Claude a suggéré un index valide, l'utilise en priorité
        if suggested_index is not None and 0 <= suggested_index < nb_orig:
            return suggested_index
        # Sinon choisit par type
        candidates = type_to_indices.get(slide_type, [])
        if candidates:
            return candidates[0]
        # Fallback : cover = première slide, conclusion = dernière
        if slide_type == "cover":
            return 0
        if slide_type == "conclusion":
            return nb_orig - 1
        # Pour le contenu, alterne entre les slides du milieu
        mid = list(range(1, nb_orig - 1)) if nb_orig > 2 else list(range(nb_orig))
        return mid[0] if mid else 0

    # Supprime les slides existantes de la présentation
    sldIdLst = prs.slides._sldIdLst
    for ref in list(sldIdLst):
        sldIdLst.remove(ref)

    # Reconstruit slide par slide
    for sd in slides_data:
        slide_type = sd.get("type") or "content"
        suggested_idx = sd.get("template_slide_index")
        src_idx = best_template_index(slide_type, suggested_idx)

        # Copie fidèle de la slide template
        try:
            new_slide = duplicate_slide_xml(prs, src_idx)
        except Exception as e:
            # Fallback : crée depuis le layout si la copie échoue
            layout = prs.slide_layouts[min(src_idx, len(prs.slide_layouts)-1)]
            new_slide = prs.slides.add_slide(layout)

        title_text = sd.get("title") or ""
        body_lines = sd.get("body") or []
        subtitle_text = sd.get("subtitle") or ""
        notes_text = sd.get("notes") or ""

        # Injecte le titre (placeholder idx=0)
        if title_text:
            injected = set_placeholder_text(new_slide, 0, [title_text])
            if not injected:
                # Fallback via shapes.title
                try:
                    if new_slide.shapes.title:
                        new_slide.shapes.title.text = title_text
                except Exception:
                    pass

        # Injecte le corps (placeholder idx=1)
        content_lines = body_lines if body_lines else ([subtitle_text] if subtitle_text else [])
        if content_lines:
            set_placeholder_text(new_slide, 1, content_lines)

        # Injecte les notes présentateur
        if notes_text:
            try:
                new_slide.notes_slide.notes_text_frame.text = notes_text
            except Exception:
                pass

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out.read()


# ─────────────────────────────────────────────
# 5. ROUTES
# ─────────────────────────────────────────────

@app.get("/")
def root():
    return {"status": "ok", "service": "Visual Cortex API 🚀", "version": "4.0.0"}


@app.get("/health")
def health():
    return {"status": "healthy"}


@app.post("/analyze-template")
async def analyze_template(file: UploadFile = File(...)):
    """Analyse un template PPTX. Gratuit, sans quota."""
    if not file.filename.endswith(".pptx"):
        raise HTTPException(400, "Le fichier doit être un .pptx")
    try:
        pptx_bytes = await file.read()
        brand = extract_brand_identity(pptx_bytes)
        prs = Presentation(io.BytesIO(pptx_bytes))
        profiles = analyze_template_slides(prs)
    except Exception as e:
        raise HTTPException(500, f"Erreur analyse template : {e}")

    return {
        "success": True,
        "brand": brand,
        "slide_profiles": [
            {
                "index": p["index"],
                "type": p["guessed_type"],
                "layout": p["layout_name"],
                "has_image": p["has_image"],
                "has_table": p["has_table"],
            }
            for p in profiles
        ],
        "message": f"{brand['slide_count']} slides analysées • {len(brand['fonts'])} polices • {len(brand['theme_colors'])} couleurs thème"
    }


@app.post("/generate-preview")
async def generate_preview(
    request: Request,
    template: UploadFile = File(...),
    prompt: str = Form(...),
    nb_slides: int = Form(default=8),
    authorization: str = Form(default=None),
):
    """Génère le plan narratif sans créer le fichier."""
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
        "narrative": content.get("narrative") or "",
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
    """Génère et retourne le .pptx chartée et fidèle au template."""
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
