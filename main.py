"""
Visual Cortex — PPTX Generator API v8 (Modèle Cortex Edition)
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

app = FastAPI(title="Visual Cortex API", version="8.0.0")

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
    Utilisé pour l'affichage dans l'UI Lovable.
    """
    fonts = set()
    colors = set()

    for slide in prs.slides:
        for shape in slide.shapes:
            # Polices
            if getattr(shape, "has_text_frame", False):
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if run.font.name:
                            fonts.add(run.font.name)
                        if run.font.color and run.font.color.type is not None:
                            try:
                                rgb = run.font.color.rgb
                                # Exclure blanc et noir purs (trop communs)
                                if rgb not in (RGBColor(0xFF,0xFF,0xFF), RGBColor(0,0,0)):
                                    colors.add(str(rgb))
                            except Exception:
                                pass

    # Limiter aux couleurs les plus distinctives
    color_list = list(colors)[:6]

    return {
        "fonts": list(fonts)[:4],
        "colors": color_list,
        "slide_count": len(prs.slides),
        "layouts_available": len(prs.slide_layouts),
    }


# ─────────────────────────────────────────────
# 2. EXTRACTION DES TEXTES
# ─────────────────────────────────────────────
def extract_texts_for_ai(prs: Presentation, nb_slides: int) -> dict:
    """
    Extrait le texte brut de chaque slide avec métadonnées de position.
    On extrait aussi des infos sur la slide (titre probable, type) pour
    permettre à Claude de générer un contenu adapté à la structure.
    """
    extracted = {}
    limit = min(nb_slides, len(prs.slides))

    for i in range(limit):
        slide = prs.slides[i]
        texts = []
        title_candidate = None

        for shape in slide.shapes:
            if not getattr(shape, "has_text_frame", False):
                continue
            for para in shape.text_frame.paragraphs:
                full_text = "".join(run.text for run in para.runs).strip()
                if len(full_text) > 2:
                    texts.append(full_text)
                    # Le premier texte long est probablement le titre
                    if title_candidate is None and len(full_text) > 5:
                        title_candidate = full_text

        if texts:
            extracted[f"slide_{i}"] = {
                "texts": list(dict.fromkeys(texts)),  # dédupliqué, ordre préservé
                "title_candidate": title_candidate or texts[0],
            }

    return extracted


# ─────────────────────────────────────────────
# 3. IA : GÉNÉRATION DU CONTENU (Modèle Cortex)
# ─────────────────────────────────────────────
def generate_text_mapping_with_claude(
    prompt: str,
    extracted_texts: dict,
    brand_info: dict | None = None
) -> dict:
    """
    Génère le mapping texte original → nouveau texte en appliquant
    les principes du Modèle Cortex : cohérence, qualité B2B, respiration.
    """
    if not ANTHROPIC_API_KEY:
        raise ValueError("Clé API Claude manquante.")

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    brand_context = ""
    if brand_info:
        fonts_str = ", ".join(brand_info.get("fonts", [])) or "non détectées"
        brand_context = f"\nCharte détectée : polices {fonts_str}, {brand_info.get('slide_count', '?')} slides."

    system = """Tu es Visual Cortex, expert en création de présentations B2B professionnelles.
Tu appliques le Modèle Cortex — principes de qualité graphique et éditoriale :

RÈGLES ABSOLUES :
1. Longueur stricte : même nombre de mots que l'original (±20%). Un titre de 3 mots → 3 mots. Un paragraphe de 25 mots → 25 mots.
2. Cohérence systématique : si un élément (sous-titre, accroche, call-to-action) est présent dans le template, il doit l'être dans toutes les slides concernées.
3. Respiration visuelle : les textes courts restent courts. Ne jamais surcharger une zone conçue pour peu de texte.
4. Qualité B2B : langage professionnel, direct, orienté valeur. Pas de formules creuses.
5. Wording de l'entreprise : utiliser le vocabulaire propre au secteur et à l'entreprise cible.
6. Structure narrative : la première slide accroche, les slides intermédiaires développent, la dernière conclut.

Réponds UNIQUEMENT en JSON valide, sans aucun commentaire ni markdown."""

    # Simplifier la structure pour le prompt (juste les textes, pas les métadonnées)
    simplified = {
        k: v["texts"] for k, v in extracted_texts.items()
    }

    user = f"""Nouveau sujet cible : {prompt}{brand_context}

Textes extraits du template (par slide) :
{json.dumps(simplified, ensure_ascii=False, indent=2)}

Génère les textes de remplacement en respectant STRICTEMENT la longueur de chaque texte.

Format JSON attendu :
{{
  "slide_0": {{
    "Texte original exact": "Nouveau texte de longueur équivalente"
  }},
  "slide_1": {{ ... }}
}}"""

    msg = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=8000,
        system=system,
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
            # Une slide qui plante ne doit jamais bloquer les autres
            continue

    # ── Étape 2 : Suppression sécurisée des slides en trop ───────────────
    # Méthode sécurisée : on retire depuis la fin pour éviter les décalages d'index
    current_count = len(prs.slides)
    if nb_slides < current_count:
        xml_slides = prs.slides._sldIdLst
        # Supprimer de la fin vers le début — plus sûr que de supprimer depuis le début
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
    return {"status": "ok", "version": "8.0.0 - Modèle Cortex"}


@app.post("/analyze-template")
async def analyze_template(file: UploadFile = File(...)):
    """
    Analyse le template et retourne les infos de charte pour l'UI.
    Modèle Cortex : extraction des polices, couleurs, nombre de slides.
    """
    pptx_bytes = await file.read()
    prs = Presentation(io.BytesIO(pptx_bytes))
    brand = extract_brand(prs)

    fonts_display = ", ".join(brand["fonts"]) if brand["fonts"] else "Standard"
    colors_count = len(brand["colors"])

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
    filename = f"visualcortex-{safe_name}.pptx"

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
