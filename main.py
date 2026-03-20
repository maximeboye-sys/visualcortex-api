"""
Visual Cortex — PPTX Generator API v10 (Architecture 3 Phases)
═══════════════════════════════════════════════════════════════
Phase 1 — COMPRÉHENSION : analyse profonde du template, classification
           de chaque slide par type, construction d'une bibliothèque de layouts.
Phase 2 — PLANIFICATION : Claude décide la structure narrative et quel
           layout utiliser pour chaque slide avant de générer quoi que ce soit.
Phase 3 — GÉNÉRATION + HYDRATATION : contenu adapté au layout choisi,
           injection chirurgicale dans le template, zéro corruption XML.

Modèle : claude-sonnet-4-6 (qualité maximale, configurable via CLAUDE_MODEL)
"""

import os
import io
import json
import time
import copy
import re
import logging
from collections import defaultdict
from enum import Enum
from typing import Optional

import anthropic
from fastapi import FastAPI, File, Form, UploadFile, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, StreamingResponse
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.oxml.ns import qn
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
import uvicorn

# ─────────────────────────────────────────────
# LOGGING
# ─────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger("visual-cortex")

# ─────────────────────────────────────────────
# APP
# ─────────────────────────────────────────────
app = FastAPI(title="Visual Cortex API", version="10.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.exception_handler(Exception)
async def universal_exception_handler(request: Request, exc: Exception):
    log.error(f"Unhandled exception: {exc}", exc_info=True)
    return JSONResponse(
        status_code=500,
        content={"detail": {"message": f"Erreur serveur : {str(exc)}"}},
        headers={"Access-Control-Allow-Origin": "*"},
    )


# ─────────────────────────────────────────────
# CONFIG & QUOTAS
# ─────────────────────────────────────────────
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
PRO_SECRET_TOKEN  = os.environ.get("PRO_SECRET_TOKEN", "change-me-in-railway")
FREE_QUOTA_PER_IP = int(os.environ.get("FREE_QUOTA_PER_IP", "3"))
CLAUDE_MODEL      = os.environ.get("CLAUDE_MODEL", "claude-sonnet-4-6")

_usage: dict = defaultdict(list)
DAY_SECONDS = 86400


def _get_ip(request: Request) -> str:
    fwd = request.headers.get("x-forwarded-for")
    return fwd.split(",")[0].strip() if fwd else request.client.host


def _is_pro(authorization: Optional[str]) -> bool:
    if not authorization:
        return False
    return authorization.replace("Bearer ", "").strip() == PRO_SECRET_TOKEN


def _check_and_increment_quota(ip: str) -> tuple:
    now = time.time()
    _usage[ip] = [t for t in _usage[ip] if now - t < DAY_SECONDS]
    used = len(_usage[ip])
    if used >= FREE_QUOTA_PER_IP:
        raise HTTPException(
            status_code=429,
            detail={
                "message": "Quota gratuit épuisé. Passez en Pro pour un accès illimité."
            },
        )
    _usage[ip].append(now)
    return used + 1, FREE_QUOTA_PER_IP


# ═══════════════════════════════════════════════════════════════
# PHASE 1 — COMPRÉHENSION PROFONDE DU TEMPLATE
# ═══════════════════════════════════════════════════════════════

SLIDE_TYPES = {
    "cover":      "Slide de couverture (titre + sous-titre)",
    "section":    "Slide de séparation de section",
    "two_col":    "Deux colonnes de contenu",
    "kpi":        "Chiffres clés / statistiques",
    "quote":      "Citation ou accroche forte",
    "timeline":   "Chronologie / processus étapes",
    "list":       "Liste à items ou bullet points",
    "image_text": "Image + texte côte à côte",
    "full_text":  "Contenu textuel dense",
    "closing":    "Slide de conclusion / remerciement / CTA",
    "unknown":    "Non classifiable",
}


def _emu_to_in(emu: int) -> float:
    return emu / 914400.0


def _classify_slide(slide, idx: int, total: int, w: int, h: int) -> str:
    """
    Classifie une slide par type via plusieurs heuristiques combinées :
    position dans la présentation, nombre de shapes, distribution spatiale,
    présence d'images, longueur des textes.
    """
    shapes      = list(slide.shapes)
    text_shapes = [s for s in shapes if getattr(s, "has_text_frame", False)]
    img_shapes  = [
        s for s in shapes
        if s.shape_type in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.LINKED_PICTURE)
    ]

    texts = []
    for s in text_shapes:
        for para in s.text_frame.paragraphs:
            t = "".join(r.text for r in para.runs).strip()
            if len(t) > 3:
                texts.append({"text": t, "shape": s})

    total_chars = sum(len(t["text"]) for t in texts)
    n           = len(texts)

    # Position dans la présentation
    if idx == 0:
        return "cover"
    if idx == total - 1:
        return "closing"

    # Image + texte
    if img_shapes and n >= 2:
        return "image_text"

    # Section : très peu de texte
    if n <= 2 and total_chars < 80:
        return "section"

    # KPI : nombreux textes courts distribués horizontalement
    if n >= 4:
        short = [t for t in texts if len(t["text"]) < 25]
        if len(short) >= 3:
            lefts = sorted(
                [t["shape"].left for t in texts if hasattr(t["shape"], "left")]
            )
            if len(lefts) >= 3:
                spread = _emu_to_in(lefts[-1]) - _emu_to_in(lefts[0])
                if spread > 4.0:
                    return "kpi"

    # Timeline : shapes étalés verticalement
    if n >= 3:
        tops = sorted(
            [t["shape"].top for t in texts if hasattr(t["shape"], "top")]
        )
        if len(tops) >= 3:
            v_spread = _emu_to_in(tops[-1]) - _emu_to_in(tops[0])
            if v_spread > 2.5 and total_chars < 500:
                return "timeline"

    # Deux colonnes : shapes distribués gauche / droite
    if n >= 4:
        half = w / 2
        left_c  = sum(1 for t in texts if hasattr(t["shape"], "left") and t["shape"].left < half)
        right_c = sum(1 for t in texts if hasattr(t["shape"], "left") and t["shape"].left >= half)
        if left_c >= 2 and right_c >= 2:
            return "two_col"

    # Citation : 1-2 textes, dont un long
    long_texts = [t for t in texts if len(t["text"]) > 60]
    if len(long_texts) == 1 and n <= 3:
        return "quote"

    # Liste : textes de longueur similaire
    if n >= 3:
        lengths = [len(t["text"]) for t in texts]
        avg = sum(lengths) / len(lengths)
        variance = sum((l - avg) ** 2 for l in lengths) / len(lengths)
        if variance < 600 and avg < 120:
            return "list"

    if total_chars > 200:
        return "full_text"

    return "unknown"


def _shape_role(shape, w: int, h: int) -> str:
    """
    Détermine le rôle d'une shape via le type placeholder PPTX natif
    (le plus fiable), puis par heuristique géométrique en fallback.
    """
    # Type natif PPTX (vérifier is_placeholder avant d'accéder à placeholder_format)
    try:
        if shape.is_placeholder:
            ph = shape.placeholder_format
            ph_map = {
                0: "title", 1: "body", 2: "subtitle",
                3: "date",  4: "footer", 5: "page_number",
                13: "title", 15: "subtitle",
            }
            return ph_map.get(ph.idx, "placeholder")
    except Exception:
        pass

    if not getattr(shape, "has_text_frame", False):
        return "decoration"

    try:
        top_r   = shape.top    / h
        left_r  = shape.left   / w
        width_r = shape.width  / w
        height_r = shape.height / h

        if top_r > 0.87 and width_r > 0.15:
            return "footer"
        if top_r > 0.87 and width_r < 0.10:
            return "page_number"
        if top_r < 0.28 and width_r > 0.45:
            return "title"
        if 0.20 < top_r < 0.50 and width_r > 0.35 and height_r < 0.15:
            return "subtitle"
        if 0.25 < top_r < 0.75 and width_r < 0.40 and height_r < 0.12:
            return "label"
        if 0.25 < top_r < 0.87:
            return "body"
    except Exception:
        pass

    return "text"


def extract_brand(prs: Presentation) -> dict:
    """
    Extraction complète de la charte graphique du template :
    polices, couleurs texte, couleurs de fond, dimensions, layouts disponibles.
    """
    fonts  = set()
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
                                if rgb not in (
                                    RGBColor(0xFF, 0xFF, 0xFF),
                                    RGBColor(0x00, 0x00, 0x00),
                                ):
                                    colors.add(str(rgb))
                            except Exception:
                                pass

    layout_names = [lay.name for lay in prs.slide_layouts]
    w, h = prs.slide_width, prs.slide_height

    return {
        "fonts":           list(fonts)[:5],
        "colors":          list(colors)[:8],
        "slide_count":     len(prs.slides),
        "layouts":         layout_names,
        "slide_width_in":  round(_emu_to_in(w), 2),
        "slide_height_in": round(_emu_to_in(h), 2),
        "aspect_ratio": (
            "16:9" if abs(w / h - 16 / 9) < 0.05
            else "4:3" if abs(w / h - 4 / 3) < 0.05
            else "custom"
        ),
    }


def build_layout_library(prs: Presentation) -> list:
    """
    Construit la bibliothèque de layouts du template.
    Pour chaque slide : type détecté, zones texte avec rôle et word_count.
    C'est la carte que Claude utilise en Phase 2 pour planifier.
    """
    library = []
    total   = len(prs.slides)
    w, h    = prs.slide_width, prs.slide_height

    for idx, slide in enumerate(prs.slides):
        slide_type = _classify_slide(slide, idx, total, w, h)
        zones = []
        seen  = set()

        for shape in slide.shapes:
            if not getattr(shape, "has_text_frame", False):
                continue
            role = _shape_role(shape, w, h)
            if role == "decoration":
                continue

            for para in shape.text_frame.paragraphs:
                text = "".join(r.text for r in para.runs).strip()
                if len(text) < 2 or text in seen:
                    continue
                seen.add(text)
                zones.append({
                    "original_text": text,
                    "role":          role,
                    "word_count":    len(text.split()),
                    "char_count":    len(text),
                })

        if zones:
            library.append({
                "slide_index": idx,
                "slide_type":  slide_type,
                "description": SLIDE_TYPES.get(slide_type, ""),
                "position": (
                    "cover"   if idx == 0 else
                    "closing" if idx == total - 1 else
                    f"{idx + 1}/{total}"
                ),
                "zones":       zones,
                "total_words": sum(z["word_count"] for z in zones),
            })

    return library


# ═══════════════════════════════════════════════════════════════
# PHASE 2 — PLANIFICATION NARRATIVE
# ═══════════════════════════════════════════════════════════════

PLANNER_SYSTEM = """Tu es Visual Cortex Planner, architecte narratif de présentations B2B professionnelles.

TON RÔLE : concevoir la structure narrative COMPLÈTE avant toute génération de contenu.

PRINCIPES DU PLANNER :
- Chaque slide doit avoir un angle narratif UNIQUE et une contribution claire à l'histoire.
- La slide 1 (cover) capte l'attention : titre fort + sous-titre qui donne envie de lire.
- Les slides intermédiaires progressent logiquement selon des séquences éprouvées :
  contexte → enjeux → solution → différenciation → preuves → bénéfices → ROI → next steps
- La dernière slide (closing) laisse une impression forte : CTA clair ou message mémorable.
- Choisir le type de slide le plus adapté au contenu prévu :
  données chiffrées → kpi | étapes chronologiques → timeline |
  comparaison → two_col | élément fort à mettre en valeur → quote |
  liste d'arguments → list | photo/visuel → image_text
- Éviter de répéter le même type consécutivement (sauf list/full_text si justifié).
- Le nombre de slides dans le plan doit être EXACTEMENT égal à nb_slides.

Réponds UNIQUEMENT en JSON valide, sans markdown ni commentaire."""

PLANNER_USER = """SUJET : {prompt}
NOMBRE DE SLIDES : {nb_slides}

LAYOUTS DISPONIBLES DANS LE TEMPLATE :
{library_json}

CHARTE :
- Polices : {fonts}
- Couleurs : {colors}
- Format : {aspect_ratio} ({w}" × {h}")

Génère le plan narratif. Le tableau "slides" doit contenir EXACTEMENT {nb_slides} entrées.

FORMAT ATTENDU :
{{
  "presentation_title": "Titre accrocheur",
  "narrative_arc": "Logique narrative en 1-2 phrases",
  "slides": [
    {{
      "plan_index": 0,
      "template_slide_index": 0,
      "slide_type": "cover",
      "narrative_angle": "Ce que cette slide accomplit dans l'histoire",
      "key_message": "Le message principal en 1 phrase courte",
      "content_hints": "Indications concrètes sur le contenu idéal"
    }}
  ]
}}"""


def plan_presentation(prompt: str, nb_slides: int, library: list, brand: dict) -> dict:
    """
    Phase 2 : Claude planifie la structure narrative complète.
    Retourne un plan JSON avec, pour chaque slide, son rôle et le layout à utiliser.
    """
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    # Version allégée de la library pour le planner (sans les textes originaux)
    lib_light = [
        {
            "slide_index": s["slide_index"],
            "slide_type":  s["slide_type"],
            "description": s["description"],
            "position":    s["position"],
            "zone_roles":  [z["role"] for z in s["zones"]],
            "total_words": s["total_words"],
        }
        for s in library
    ]

    user = PLANNER_USER.format(
        prompt       = prompt,
        nb_slides    = nb_slides,
        library_json = json.dumps(lib_light, ensure_ascii=False, indent=2),
        fonts        = ", ".join(brand.get("fonts", [])) or "Standard",
        colors       = ", ".join(f"#{c}" for c in brand.get("colors", [])) or "non détectées",
        aspect_ratio = brand.get("aspect_ratio", "16:9"),
        w            = brand.get("slide_width_in", 13.33),
        h            = brand.get("slide_height_in", 7.5),
    )

    msg = client.messages.create(
        model      = CLAUDE_MODEL,
        max_tokens = 4000,
        system     = PLANNER_SYSTEM,
        messages   = [{"role": "user", "content": user}],
    )

    plan = json.loads(_clean_json(msg.content[0].text.strip()))
    log.info(
        f"Plan : {len(plan.get('slides', []))} slides — {plan.get('narrative_arc', '')[:80]}"
    )
    return plan


def _pad_plan(plan: dict, nb_slides: int, library: list) -> dict:
    """
    S'assure que le plan contient exactement nb_slides entrées.
    Complète si Claude en a retourné moins, tronque si plus.
    """
    slides = plan.get("slides", [])
    while len(slides) < nb_slides:
        fallback_idx = min(len(library) - 2, max(1, len(slides) - 1))
        slides.append({
            "plan_index":           len(slides),
            "template_slide_index": library[fallback_idx]["slide_index"] if library else 1,
            "slide_type":           "list",
            "narrative_angle":      "Développement complémentaire",
            "key_message":          "Argument additionnel",
            "content_hints":        "",
        })
    plan["slides"] = slides[:nb_slides]
    return plan


# ═══════════════════════════════════════════════════════════════
# PHASE 3 — GÉNÉRATION DU CONTENU
# ═══════════════════════════════════════════════════════════════

CORTEX_SYSTEM = """Tu es Visual Cortex, expert en création de présentations B2B professionnelles.
Tu appliques le Modèle Cortex — principes de qualité graphique et éditoriale :

RÈGLES ABSOLUES :
1. Longueur stricte : respecte le "word_count" de chaque zone (±20% maximum).
   Un titre de 3 mots → 3 mots. Un body de 20 mots → 18 à 24 mots.
2. Cohérence systématique : footer identique (ou quasi) sur toutes les slides de contenu.
3. Respiration visuelle : les zones courtes restent courtes. Jamais de surcharge.
4. Qualité B2B : langage professionnel, direct, orienté valeur. Zéro formule creuse.
5. Wording sectoriel : adapter au vocabulaire du secteur et de l'entreprise.
6. Progression narrative : chaque slide avance l'histoire selon le "narrative_angle" fourni.
7. Rôles respectés :
   - title       → percutant, court, mémorable (respecter word_count)
   - subtitle    → précise et complète le titre
   - body        → développe l'argument, structuré, lisible
   - label       → court, factuel, 2-4 mots de préférence
   - footer      → conserver quasi identique au template (nom entreprise, date, etc.)
   - page_number → NE PAS MODIFIER
   - kpi_value   → chiffre ou métrique impactante, format court
   - kpi_label   → 2-4 mots maximum
   - quote       → citation forte, ton affirmatif
8. Zéro invention de chiffres ou données non fournis dans le prompt.

Réponds UNIQUEMENT en JSON valide, sans commentaire ni markdown."""

CORTEX_USER = """PRÉSENTATION : {title}
ARC NARRATIF : {arc}
SUJET DÉTAILLÉ : {prompt}

CHARTE :
- Polices : {fonts}
- Couleurs : {colors}

SLIDES À GÉNÉRER ({n} slides) :
{slides_json}

Pour chaque slide, génère un texte de remplacement par zone.
Clés du JSON de sortie = "template_slide_index" de chaque slide du plan.

FORMAT ATTENDU :
{{
  "0": {{"Texte original": "Nouveau texte"}},
  "1": {{"Texte original": "Nouveau texte"}},
  ...
}}"""


def generate_content(prompt: str, plan: dict, library: list, brand: dict) -> dict:
    """
    Phase 3 : génère le contenu de toutes les slides en une seule requête Claude,
    guidée par le plan narratif et les zones du template.
    """
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    lib_by_idx = {s["slide_index"]: s for s in library}

    slides_payload = []
    for sp in plan.get("slides", []):
        tmpl_idx  = sp.get("template_slide_index", 0)
        tmpl_data = lib_by_idx.get(tmpl_idx, {})

        slides_payload.append({
            "template_slide_index": tmpl_idx,
            "slide_type":           sp.get("slide_type"),
            "narrative_angle":      sp.get("narrative_angle"),
            "key_message":          sp.get("key_message"),
            "content_hints":        sp.get("content_hints", ""),
            "zones": [
                {
                    "original_text": z["original_text"],
                    "role":          z["role"],
                    "word_count":    z["word_count"],
                }
                for z in tmpl_data.get("zones", [])
            ],
        })

    user = CORTEX_USER.format(
        title      = plan.get("presentation_title", prompt[:60]),
        arc        = plan.get("narrative_arc", ""),
        prompt     = prompt,
        fonts      = ", ".join(brand.get("fonts", [])) or "Standard",
        colors     = ", ".join(f"#{c}" for c in brand.get("colors", [])) or "non détectées",
        n          = len(slides_payload),
        slides_json = json.dumps(slides_payload, ensure_ascii=False, indent=2),
    )

    msg = client.messages.create(
        model      = CLAUDE_MODEL,
        max_tokens = 8000,
        system     = CORTEX_SYSTEM,
        messages   = [{"role": "user", "content": user}],
    )

    mapping = json.loads(_clean_json(msg.content[0].text.strip()))
    log.info(f"Contenu généré : {len(mapping)} slides mappées.")
    return mapping


# ═══════════════════════════════════════════════════════════════
# HYDRATATION — Injection chirurgicale dans le template
# ═══════════════════════════════════════════════════════════════

def hydrate_presentation(
    pptx_bytes: bytes,
    mapping: dict,
    plan_slides: list,
    nb_slides: int,
) -> bytes:
    """
    Injecte les textes générés dans le template PPTX en préservant
    intégralement le style XML (polices, couleurs, tailles).

    Gestion du nombre de slides :
    - Trop peu dans le template → duplique des slides de contenu
    - Trop beaucoup → supprime depuis la fin (méthode sécurisée)
    """
    prs = Presentation(io.BytesIO(pptx_bytes))

    # Duplication si nécessaire
    _ensure_slide_count(prs, plan_slides, nb_slides)

    # Remplacement texte par texte
    for slide_key, replacements in mapping.items():
        try:
            slide_idx = int(str(slide_key).replace("slide_", ""))
            if slide_idx >= len(prs.slides):
                continue
            slide = prs.slides[slide_idx]

            for shape in slide.shapes:
                if not getattr(shape, "has_text_frame", False):
                    continue
                for para in shape.text_frame.paragraphs:
                    para_text = "".join(r.text for r in para.runs).strip()
                    if not para_text or para_text not in replacements:
                        continue

                    new_text = replacements[para_text]
                    if not new_text:
                        continue

                    # Sauvegarder le style du premier run
                    rpr_xml = None
                    if para.runs:
                        rpr_el = para.runs[0]._r.find(qn("a:rPr"))
                        if rpr_el is not None:
                            rpr_xml = copy.deepcopy(rpr_el)

                    para.text = new_text

                    # Ré-appliquer le style
                    if rpr_xml is not None:
                        for run in para.runs:
                            existing = run._r.find(qn("a:rPr"))
                            if existing is not None:
                                run._r.remove(existing)
                            run._r.insert(0, copy.deepcopy(rpr_xml))

        except Exception as e:
            log.warning(f"Hydratation slide {slide_key} ignorée : {e}")
            continue

    # Suppression des slides en trop
    current = len(prs.slides)
    if nb_slides < current:
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


def _ensure_slide_count(prs: Presentation, plan_slides: list, nb_slides: int):
    """
    Duplique des slides de contenu si le plan demande plus de slides
    que le template n'en contient.
    """
    current = len(prs.slides)
    if nb_slides <= current:
        return

    content_types = {"list", "full_text", "two_col", "image_text", "kpi", "unknown"}
    content_indices = [
        s["template_slide_index"]
        for s in plan_slides
        if s.get("slide_type") in content_types
        and s.get("template_slide_index", 0) < current
    ]
    if not content_indices:
        content_indices = [max(1, current // 2)]

    to_add = nb_slides - current
    cycle  = content_indices * (to_add // len(content_indices) + 1)

    for i in range(to_add):
        _duplicate_slide(prs, cycle[i % len(cycle)])

    log.info(f"Slides dupliquées : {current} → {len(prs.slides)}")


def _duplicate_slide(prs: Presentation, src_idx: int):
    """Duplique une slide existante à la fin de la présentation."""
    try:
        src        = prs.slides[src_idx]
        layout     = prs.slide_layouts[-1]
        new_slide  = prs.slides.add_slide(layout)

        for elem in src.shapes._spTree:
            new_slide.shapes._spTree.append(copy.deepcopy(elem))

        for rel in src.part.rels.values():
            if "image" in rel.reltype:
                try:
                    new_slide.part.add_relationship(rel.reltype, rel._target)
                except Exception:
                    pass
    except Exception as e:
        log.warning(f"Duplication slide {src_idx} échouée : {e}")


# ═══════════════════════════════════════════════════════════════
# UTILITAIRES
# ═══════════════════════════════════════════════════════════════

def _clean_json(raw: str) -> str:
    """Supprime les balises markdown éventuelles autour du JSON."""
    if "```" in raw:
        parts = raw.split("```")
        raw = parts[1] if len(parts) > 1 else parts[0]
        if raw.startswith("json"):
            raw = raw[4:]
    return raw.strip()


def _safe_filename(prompt: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", prompt[:40].lower()).strip("-")


# ═══════════════════════════════════════════════════════════════
# PIPELINE COMPLET (3 phases enchaînées)
# ═══════════════════════════════════════════════════════════════

def run_pipeline(pptx_bytes: bytes, prompt: str, nb_slides: int) -> tuple:
    """
    Exécute les 3 phases en séquence.
    Retourne : (bytes du .pptx final, plan narratif, charte extraite)
    """
    prs = Presentation(io.BytesIO(pptx_bytes))

    # Phase 1 — Compréhension
    log.info("Phase 1 : analyse profonde du template...")
    brand   = extract_brand(prs)
    library = build_layout_library(prs)
    log.info(
        f"Template : {len(library)} slides analysées, "
        f"{len(brand['fonts'])} polices, "
        f"{len(brand['colors'])} couleurs, "
        f"format {brand['aspect_ratio']}"
    )

    nb_slides = max(1, min(nb_slides, 30))

    # Phase 2 — Planification
    log.info("Phase 2 : planification narrative...")
    plan = plan_presentation(prompt, nb_slides, library, brand)
    plan = _pad_plan(plan, nb_slides, library)

    # Phase 3 — Génération + Hydratation
    log.info("Phase 3 : génération du contenu...")
    mapping = generate_content(prompt, plan, library, brand)

    log.info("Hydratation du fichier PPTX...")
    final_bytes = hydrate_presentation(
        pptx_bytes,
        mapping,
        plan["slides"],
        nb_slides,
    )

    return final_bytes, plan, brand


# ═══════════════════════════════════════════════════════════════
# ROUTES API
# ═══════════════════════════════════════════════════════════════

@app.get("/")
def root():
    return {
        "status":  "ok",
        "version": "10.0.0 - Architecture 3 Phases",
        "model":   CLAUDE_MODEL,
    }


@app.post("/analyze-template")
async def analyze_template(file: UploadFile = File(...)):
    """
    Analyse le template : charte graphique + classification des slides.
    Permet à l'UI d'afficher les infos détectées avant la génération.
    """
    pptx_bytes = await file.read()
    prs     = Presentation(io.BytesIO(pptx_bytes))
    brand   = extract_brand(prs)
    library = build_layout_library(prs)

    type_counts: dict = defaultdict(int)
    for s in library:
        type_counts[s["slide_type"]] += 1

    fonts_display = ", ".join(brand["fonts"]) if brand["fonts"] else "Standard"

    return {
        "success": True,
        "message": (
            f"Charte détectée : {fonts_display} • "
            f"{len(brand['colors'])} couleurs • "
            f"{brand['slide_count']} slides • "
            f"{brand['aspect_ratio']}"
        ),
        "brand":       brand,
        "slide_types": dict(type_counts),
    }


@app.post("/generate-preview")
async def generate_preview(
    request:       Request,
    template:      UploadFile = File(...),
    prompt:        str        = Form(...),
    nb_slides:     int        = Form(default=8),
    authorization: str        = Form(default=None),
):
    """
    Phase 1 + Phase 2 : retourne le plan narratif sans générer le fichier.
    L'utilisateur peut valider la structure avant de lancer la génération complète.
    """
    pro = _is_pro(authorization)
    quota_info = (
        {"plan": "pro"}
        if pro
        else {
            "used":  _check_and_increment_quota(_get_ip(request))[0],
            "total": FREE_QUOTA_PER_IP,
            "plan":  "free",
        }
    )

    template_bytes = await template.read()
    prs     = Presentation(io.BytesIO(template_bytes))
    brand   = extract_brand(prs)
    library = build_layout_library(prs)
    plan    = plan_presentation(prompt, nb_slides, library, brand)
    plan    = _pad_plan(plan, nb_slides, library)

    return {
        "success":            True,
        "presentation_title": plan.get("presentation_title", prompt[:60]),
        "narrative_arc":      plan.get("narrative_arc", ""),
        "slides": [
            {
                "index":           s.get("plan_index"),
                "type":            s.get("slide_type"),
                "narrative_angle": s.get("narrative_angle"),
                "key_message":     s.get("key_message"),
            }
            for s in plan.get("slides", [])
        ],
        "brand": brand,
        "quota": quota_info,
    }


@app.post("/generate")
async def generate_presentation(
    request:       Request,
    template:      UploadFile = File(...),
    prompt:        str        = Form(...),
    nb_slides:     int        = Form(default=8),
    authorization: str        = Form(default=None),
):
    """Pipeline complet 3 phases → retourne le fichier .pptx final."""
    if not _is_pro(authorization):
        _check_and_increment_quota(_get_ip(request))

    template_bytes = await template.read()
    final_bytes, plan, _brand = run_pipeline(template_bytes, prompt, nb_slides)

    filename = f"visualcortex-{_safe_filename(prompt)}.pptx"

    return StreamingResponse(
        io.BytesIO(final_bytes),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


if __name__ == "__main__":
    uvicorn.run(
        "main:app",
        host   = "0.0.0.0",
        port   = int(os.environ.get("PORT", 8000)),
        reload = False,
    )
