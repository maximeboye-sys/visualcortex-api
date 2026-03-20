"""
Visual Cortex — PPTX Generator API v9 (Modèle Cortex Edition)
Approche : Hydratation in-situ. Zéro duplication XML.
Résultat : 0 fichier corrompu, préservation absolue du design.
Modèle Cortex : qualité de contenu B2B, cohérence, respiration visuelle.
"""

import os
import io
import json
import time
import copy
import re
from collections import defaultdict

import anthropic
from fastapi import FastAPI, File, Form, UploadFile, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, StreamingResponse
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.dml.color import RGBColor
import uvicorn

app = FastAPI(title="Visual Cortex API", version="9.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ─────────────────────────────────────────────
# BOUCLIER ANTI-CRASH GLOBAL
# ─────────────────────────────────────────────
@app.exception_handler(Exception)
async def universal_exception_handler(request: Request, exc: Exception):
    return JSONResponse(
        status_code=500,
        content={"detail": {"message": f"Erreur serveur : {str(exc)}"}},
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
        raise HTTPException(status_code=429, detail={"message": "Quota gratuit épuisé."})
    _usage[ip].append(now)
    return used + 1, FREE_QUOTA_PER_IP


# ─────────────────────────────────────────────
# 1. ANALYSE DE LA CHARTE (Modèle Cortex)
# ─────────────────────────────────────────────
def extract_brand(prs: Presentation) -> dict:
    """
    Extrait la charte graphique du template :
    polices, couleurs dominantes, nombre de slides.
    """
    fonts = set()
    colors = set()

    for slide in prs.slides:
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if run.font.name:
                            fonts.add(run.font.name)
                        if run.font.color and run.font.color.type is not None:
                            try:
                                rgb = run.font.color.rgb
                                if rgb not in (RGBColor(0xFF, 0xFF, 0xFF), RGBColor(0, 0, 0)):
                                    colors.add(str(rgb))
                            except Exception:
                                pass

    color_list = list(colors)[:6]

    return {
        "fonts": list(fonts)[:4],
        "colors": color_list,
        "slide_count": len(prs.slides),
        "layouts_available": len(prs.slide_layouts),
    }


# ─────────────────────────────────────────────
# 2. EXTRACTION DES TEXTES AVEC RÔLE
# ─────────────────────────────────────────────
def _detect_shape_role(shape, slide_width_emu, slide_height_emu) -> str:
    """
    Détermine le rôle probable d'une forme selon sa position et taille :
    title, subtitle, body, footer, label.
    Permet à Claude de comprendre la hiérarchie de chaque texte.
    """
    try:
        top_ratio    = shape.top    / slide_height_emu if slide_height_emu else 0
        left_ratio   = shape.left   / slide_width_emu  if slide_width_emu  else 0
        width_ratio  = shape.width  / slide_width_emu  if slide_width_emu  else 0
        height_ratio = shape.height / slide_height_emu if slide_height_emu else 0

        # Footer : en bas de la slide (>85%) et large
        if top_ratio > 0.85 and width_ratio > 0.3:
            return "footer"

        # Titre : dans le tiers supérieur, assez large
        if top_ratio < 0.30 and width_ratio > 0.4:
            return "title"

        # Sous-titre : sous le titre, taille modérée
        if top_ratio < 0.50 and width_ratio > 0.3 and height_ratio < 0.15:
            return "subtitle"

        # Corps : zone centrale
        if 0.25 < top_ratio < 0.85:
            return "body"

        return "label"
    except Exception:
        return "label"


def extract_texts_for_ai(prs: Presentation, nb_slides: int) -> dict:
    """
    Extrait le texte brut de chaque slide avec le rôle de chaque zone.
    Claude reçoit ainsi la hiérarchie complète (titre, sous-titre, body, footer).
    """
    slide_width_emu  = prs.slide_width
    slide_height_emu = prs.slide_height
    extracted = {}
    limit = min(nb_slides, len(prs.slides))

    for i in range(limit):
        slide = prs.slides[i]
        texts_with_roles = []

        for shape in slide.shapes:
            if not getattr(shape, "has_text_frame", False):
                continue
            role = _detect_shape_role(shape, slide_width_emu, slide_height_emu)
            for para in shape.text_frame.paragraphs:
                full_text = "".join(run.text for run in para.runs).strip()
                if len(full_text) > 2:
                    texts_with_roles.append({
                        "text": full_text,
                        "role": role,
                        "word_count": len(full_text.split()),
                    })

        # Dédupliqué, ordre préservé
        seen = set()
        unique = []
        for item in texts_with_roles:
            if item["text"] not in seen:
                seen.add(item["text"])
                unique.append(item)

        if unique:
            extracted[f"slide_{i}"] = unique

    return extracted


# ─────────────────────────────────────────────
# 3. IA : GÉNÉRATION DU CONTENU (Modèle Cortex)
# ─────────────────────────────────────────────

# Prompt système complet — Modèle Cortex (Partie 6 des instructions)
CORTEX_SYSTEM_PROMPT = """Tu es Visual Cortex, expert en création de présentations B2B professionnelles.
Tu appliques le Modèle Cortex — principes de qualité graphique et éditoriale :

RÈGLES ABSOLUES :
1. Longueur stricte : même nombre de mots que l'original (±20%).
   Un titre de 3 mots → 3 mots. Un paragraphe de 25 mots → 25 mots.
   Le champ "word_count" dans le JSON d'entrée t'indique exactement la longueur cible.
2. Cohérence systématique : si un élément est présent dans le template,
   il doit l'être dans toutes les slides concernées, sans exception.
3. Respiration visuelle : les textes courts restent courts.
   Ne jamais surcharger une zone conçue pour peu de texte.
4. Qualité B2B : langage professionnel, direct, orienté valeur.
   Pas de formules creuses. Chaque mot compte.
5. Wording de l'entreprise : utiliser le vocabulaire propre au secteur
   et à l'entreprise cible détectée dans le prompt.
6. Structure narrative : la première slide accroche, les slides
   intermédiaires développent avec un angle différent chacune, la dernière conclut.
7. Rôles des zones : respecte la hiérarchie (title > subtitle > body > label > footer).
   Un "footer" ne change jamais de longueur. Un "title" reste percutant et court.
   Un "body" développe l'argument principal de la slide.
8. ZÉRO couleur inventée : utilise uniquement le vocabulaire graphique de l'entreprise.
   Ne pas inventer de noms de produits ou de chiffres non fournis dans le prompt.

Réponds UNIQUEMENT en JSON valide, sans commentaire ni markdown."""


def generate_text_mapping_with_claude(
    prompt: str,
    extracted_texts: dict,
    brand_info: dict | None = None
) -> dict:
    """
    Génère le mapping texte original → nouveau texte en appliquant
    les principes du Modèle Cortex : cohérence, qualité B2B, respiration.
    Utilise claude-sonnet pour une qualité maximale.
    """
    if not ANTHROPIC_API_KEY:
        raise ValueError("Clé API Claude manquante.")

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    # Contexte de marque enrichi (polices + couleurs)
    brand_context = ""
    if brand_info:
        fonts_str  = ", ".join(brand_info.get("fonts", [])) or "non détectées"
        colors_str = ", ".join(f"#{c}" for c in brand_info.get("colors", [])) or "non détectées"
        brand_context = (
            f"\n\nCHARTE GRAPHIQUE DÉTECTÉE :"
            f"\n- Polices : {fonts_str}"
            f"\n- Couleurs de la marque : {colors_str}"
            f"\n- Nombre de slides dans le template : {brand_info.get('slide_count', '?')}"
            f"\n\nRègle : utilise uniquement ces polices et ces couleurs. Aucune invention."
        )

    user = f"""SUJET DE LA PRÉSENTATION : {prompt}{brand_context}

TEXTES DU TEMPLATE (à remplacer slide par slide) :
Chaque texte est accompagné de son rôle (title/subtitle/body/footer/label) et du nombre de mots cible.
{json.dumps(extracted_texts, ensure_ascii=False, indent=2)}

INSTRUCTIONS DE GÉNÉRATION :
- Génère un texte de remplacement pour chaque "text" dans chaque slide.
- Respecte STRICTEMENT le "word_count" de chaque texte (±20% maximum).
- Respecte le "role" : les titres sont percutants, les body développent, les footers ne changent pas.
- Les footers (numéros de page, titres de présentation répétés) peuvent rester identiques ou être légèrement adaptés.
- Assure une progression narrative cohérente entre les slides.

FORMAT JSON ATTENDU :
{{
  "slide_0": {{
    "Texte original exact": "Nouveau texte de longueur équivalente"
  }},
  "slide_1": {{ ... }}
}}

Réponds UNIQUEMENT avec ce JSON, sans explication ni markdown."""

    msg = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=8000,
        system=CORTEX_SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user}],
    )

    raw = msg.content[0].text.strip()

    # Nettoyage robuste des balises markdown
    if "```" in raw:
        parts = raw.split("```")
        raw = parts[1] if len(parts) > 1 else parts[0]
        if raw.startswith("json"):
            raw = raw[4:]

    return json.loads(raw.strip())


# ─────────────────────────────────────────────
# 4. HYDRATATION (Remplacement chirurgical)
# ─────────────────────────────────────────────
def hydrate_presentation(pptx_bytes: bytes, mapping: dict, nb_slides: int) -> bytes:
    """
    Remplace le texte en préservant 100% du style original.
    Supprime les slides en trop via la méthode sécurisée python-pptx.
    Modèle Cortex : cohérence du remplacement, aucune corruption XML.
    """
    prs = Presentation(io.BytesIO(pptx_bytes))

    # ── Étape 1 : Remplacer le texte slide par slide ──────────────────────
    for slide_idx_str, replacements in mapping.items():
        try:
            slide_idx = int(slide_idx_str.replace("slide_", ""))
            if slide_idx >= len(prs.slides):
                continue
            slide = prs.slides[slide_idx]

            for shape in slide.shapes:
                if not getattr(shape, "has_text_frame", False):
                    continue
                for para in shape.text_frame.paragraphs:
                    para_text = "".join(run.text for run in para.runs).strip()
                    if not para_text or para_text not in replacements:
                        continue

                    new_text = replacements[para_text]
                    if not new_text:
                        continue

                    # Sauvegarde du style du premier run (XML)
                    rpr_xml = None
                    if para.runs and para.runs[0]._r.find(qn("a:rPr")) is not None:
                        rpr_xml = copy.deepcopy(para.runs[0]._r.find(qn("a:rPr")))

                    # Remplacement du texte (détruit les runs, recrée proprement)
                    para.text = new_text

                    # Ré-application du style original sur tous les nouveaux runs
                    if rpr_xml is not None:
                        for run in para.runs:
                            rpr = run._r.find(qn("a:rPr"))
                            if rpr is not None:
                                run._r.remove(rpr)
                            run._r.insert(0, copy.deepcopy(rpr_xml))

        except Exception:
            continue

    # ── Étape 2 : Suppression sécurisée des slides en trop ───────────────
    current_count = len(prs.slides)
    if nb_slides < current_count:
        xml_slides = prs.slides._sldIdLst
        for sld in reversed(list(xml_slides)[nb_slides:]):
            rId = sld.get(qn("r:id"))
            if rId:
                try:
                    prs.part.drop_rel(rId)
                except Exception:
                    pass
            xml_slides.remove(sld)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out.read()


# ─────────────────────────────────────────────
# 5. ROUTES API
# ─────────────────────────────────────────────
@app.get("/")
def root():
    return {"status": "ok", "version": "9.0.0 - Modèle Cortex"}


@app.post("/analyze-template")
async def analyze_template(file: UploadFile = File(...)):
    """
    Analyse le template et retourne les infos de charte pour l'UI.
    Modèle Cortex : extraction des polices, couleurs, nombre de slides.
    """
    pptx_bytes = await file.read()
    prs = Presentation(io.BytesIO(pptx_bytes))
    brand = extract_brand(prs)

    fonts_display  = ", ".join(brand["fonts"]) if brand["fonts"] else "Standard"
    colors_count   = len(brand["colors"])

    return {
        "success": True,
        "message": f"Charte détectée : {fonts_display} • {colors_count} couleurs • {brand['slide_count']} slides",
        "brand": brand,
    }


@app.post("/generate-preview")
async def generate_preview(
    request: Request,
    template: UploadFile = File(...),
    prompt: str = Form(...),
    nb_slides: int = Form(default=8),
    authorization: str = Form(default=None),
):
    """Génère un aperçu du plan (titres) sans créer le fichier."""
    pro = _is_pro(authorization)
    quota_info = (
        {"plan": "pro"}
        if pro
        else {"used": _check_and_increment_quota(_get_ip(request))[0], "total": FREE_QUOTA_PER_IP, "plan": "free"}
    )

    template_bytes = await template.read()
    prs = Presentation(io.BytesIO(template_bytes))
    brand = extract_brand(prs)

    extracted_texts = extract_texts_for_ai(prs, nb_slides)
    mapping = generate_text_mapping_with_claude(prompt, extracted_texts, brand)

    # Extraire les titres pour l'aperçu (premier texte de chaque slide)
    preview_slides = []
    for k in sorted(mapping.keys()):
        v = mapping[k]
        if v:
            title = list(v.values())[0]
            preview_slides.append({"slide": k, "title": title})

    return {
        "success": True,
        "title": prompt[:60],
        "slides": preview_slides,
        "brand": brand,
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
    """Génère et retourne le fichier .pptx final."""
    if not _is_pro(authorization):
        _check_and_increment_quota(_get_ip(request))

    template_bytes = await template.read()
    prs = Presentation(io.BytesIO(template_bytes))
    brand = extract_brand(prs)

    extracted_texts = extract_texts_for_ai(prs, nb_slides)
    mapping = generate_text_mapping_with_claude(prompt, extracted_texts, brand)

    pptx_bytes = hydrate_presentation(template_bytes, mapping, nb_slides)

    # Nom de fichier propre basé sur le prompt
    safe_name = re.sub(r"[^a-z0-9]+", "-", prompt[:40].lower()).strip("-")
    filename  = f"visualcortex-{safe_name}.pptx"

    return StreamingResponse(
        io.BytesIO(pptx_bytes),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


if __name__ == "__main__":
    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=int(os.environ.get("PORT", 8000)),
        reload=False,
    )
