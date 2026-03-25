"""
Visual Cortex — PPTX Generator API v13 (Niveau 1 + Niveau 2)
═════════════════════════════════════════════════════════════════════════════
[Niveau 1 — inchangé depuis v12]
Architecture 3 phases : Compréhension → Planification → Génération/Hydratation

[Niveau 2 — nouveau dans v13]
Architecture 4 phases :
  Phase 1 — Analyse brand (identique Niveau 1)
  Phase 2 — Planification narrative (identique Niveau 1)
  Phase 3 — Génération de code python-pptx par Claude
  Phase 4 — Exécution sécurisée (sandbox exec) + assemblage PPTX

Nouveautés v13 :
  - CLIENT_PROFILES : 5 profils visuels (finance, industrial, institutional, startup, creative)
  - Bibliothèque de helpers h2_* : shapes, textes, KPIs, dividers, numéros décoratifs
  - Sandbox exec() à namespace restreint (aucun import, os, open exposé)
  - Timeout threading 30s + fallback automatique Niveau 1 si exécution échoue
  - Routes GET /profiles et POST /generate-v2

Modèle : claude-sonnet-4-6 (configurable via CLAUDE_MODEL)
"""

import os, io, json, time, copy, re, logging, threading
from collections import defaultdict
from typing import Optional

import anthropic
from fastapi import FastAPI, File, Form, UploadFile, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, StreamingResponse
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.oxml.ns import qn
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
import uvicorn

# ─────────────────────────────────────────────
# LOGGING & APP
# ─────────────────────────────────────────────
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger("visual-cortex")

app = FastAPI(title="Visual Cortex API", version="13.0.0")
app.add_middleware(
    CORSMiddleware, allow_origins=["*"], allow_credentials=False,
    allow_methods=["*"], allow_headers=["*"],
)

@app.exception_handler(Exception)
async def global_error(request: Request, exc: Exception):
    log.error(f"Exception: {exc}", exc_info=True)
    return JSONResponse(
        status_code=500,
        content={"detail": {"message": f"Erreur serveur : {str(exc)}"}},
        headers={"Access-Control-Allow-Origin": "*"},
    )

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
PRO_SECRET_TOKEN  = os.environ.get("PRO_SECRET_TOKEN", "change-me")
FREE_QUOTA_PER_IP = int(os.environ.get("FREE_QUOTA_PER_IP", "3"))
CLAUDE_MODEL      = os.environ.get("CLAUDE_MODEL", "claude-sonnet-4-6")

_usage: dict = defaultdict(list)
DAY_SEC = 86400

def _ip(r: Request) -> str:
    fwd = r.headers.get("x-forwarded-for")
    return fwd.split(",")[0].strip() if fwd else r.client.host

def _is_pro(auth: Optional[str]) -> bool:
    return bool(auth and auth.replace("Bearer ", "").strip() == PRO_SECRET_TOKEN)

def _quota(ip: str) -> tuple:
    now = time.time()
    _usage[ip] = [t for t in _usage[ip] if now - t < DAY_SEC]
    if len(_usage[ip]) >= FREE_QUOTA_PER_IP:
        raise HTTPException(429, {"message": "Quota gratuit épuisé. Passez en Pro."})
    _usage[ip].append(now)
    return len(_usage[ip]), FREE_QUOTA_PER_IP


# ══════════════════════════════════════════════════════════════
# UTILITAIRES FONDAMENTAUX
# ══════════════════════════════════════════════════════════════

def _emu(v: int) -> float:
    return v / 914400.0

def _clean_json(raw: str) -> str:
    if "```" in raw:
        parts = raw.split("```")
        raw = parts[1] if len(parts) > 1 else parts[0]
        if raw.startswith("json"):
            raw = raw[4:]
    return raw.strip()

def _safe_name(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", s[:40].lower()).strip("-")

FOOTER_PLACEHOLDERS = [
    "date - pied de page de votre présentation",
    "pied de page de votre présentation",
    "titre de la présentation",
    "date de la présentation",
    "[date]",
    "[footer]",
    "click to edit",
]

def _is_footer_placeholder(text: str) -> bool:
    return any(p in text.lower() for p in FOOTER_PLACEHOLDERS)


# ══════════════════════════════════════════════════════════════
# TRAVERSÉE RÉCURSIVE DES SHAPES (FIX CRITIQUE GROUP SHAPES)
# ══════════════════════════════════════════════════════════════

def iter_all_shapes(shapes):
    """
    Parcourt TOUS les shapes, y compris ceux imbriqués dans des GROUP shapes.
    FIX CRITIQUE : le bug "contenu fantôme" venait de l'absence de cette traversée.
    """
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_all_shapes(shape.shapes)
        else:
            yield shape


# ══════════════════════════════════════════════════════════════
# PHASE 1 — COMPRÉHENSION PROFONDE DU TEMPLATE
# ══════════════════════════════════════════════════════════════

SLIDE_TYPE_DESC = {
    "cover":      "Couverture — titre fort + sous-titre clair",
    "section":    "Ouverture de chapitre — numéro + titre de section",
    "two_col":    "Deux colonnes — 2 blocs symétriques côte à côte",
    "kpi":        "Chiffres clés — métriques visuelles impactantes",
    "quote":      "Citation forte — accroche ou message clé mis en valeur",
    "timeline":   "Chronologie — étapes ou jalons temporels",
    "list":       "Liste structurée — items avec labels courts",
    "image_text": "Image + texte — visuel fort côte à côte",
    "full_text":  "Contenu développé — argumentaire structuré",
    "closing":    "Conclusion — remerciement ou call-to-action",
    "complex":    "Slide graphique — diagramme ou organigramme",
    "unknown":    "Layout générique",
}

DENSITY_RULES = {
    "cover": {
        "max_title_words":    7,
        "max_subtitle_words": 12,
        "style":              "Titre : accroche forte, verbe d'action ou tension. Sous-titre : contexte ou angle en une phrase.",
        "never":              "Jamais de bullet points. Jamais plus d'un sous-titre.",
        "example_title":      "TotalEnergies sous pression",
        "example_subtitle":   "Anatomie des campagnes ONG depuis 2020",
    },
    "section": {
        "max_title_words":    6,
        "max_label_words":    2,
        "style":              "Numéro de section (01, 02…) + titre court et percutant. Pas de body.",
        "never":              "Jamais de body text sur une slide de section. Jamais plus de 2 zones.",
        "example_title":      "Liens financiers & fiscaux",
        "example_label":      "01",
    },
    "two_col": {
        "max_title_words":     8,
        "max_col_title_words": 4,
        "max_col_body_words":  30,
        "max_items_per_col":   4,
        "style":               "Titre général + 2 colonnes symétriques. Chaque colonne : label court + 2-4 items courts.",
        "never":               "Jamais de paragraphes dans les colonnes. Jamais d'asymétrie.",
        "example":             "Colonne gauche : 'LIENS FINANCIERS' + 4 items de 10 mots max. Colonne droite : 'LIENS RÉGLEMENTAIRES' + 4 items.",
    },
    "kpi": {
        "max_kpi_count":      6,
        "max_value_words":    2,
        "max_label_words":    5,
        "max_sublabel_words": 12,
        "style":              "Chiffre ou métrique très visible + label court + sous-label contextuel.",
        "never":              "Jamais plus de 6 KPIs. Jamais de phrases complètes pour les valeurs.",
        "example":            "'600 M€' / 'contribution exceptionnelle' / 'versée en 2022, loi superprofits'",
    },
    "quote": {
        "max_quote_words":    20,
        "max_author_words":   6,
        "style":              "Citation courte et percutante, ton affirmatif. Idéalement entre guillemets.",
        "never":              "Jamais de bullet points. Jamais plus d'une citation. Jamais > 20 mots.",
        "example":            "'La transition énergétique ne se fera pas sans les majors.' — Analyse 2025",
    },
    "timeline": {
        "max_steps":            6,
        "max_step_title_words": 4,
        "max_step_body_words":  12,
        "style":                "4 à 6 jalons chronologiques. Chaque étape : date + titre court + phrase optionnelle.",
        "never":                "Jamais plus de 6 étapes. Jamais de paragraphes. Jamais sans repère temporel.",
        "example":              "1924 / 'Création CFP' / 'Fondation par décret d'État'",
    },
    "list": {
        "max_items":            5,
        "max_item_title_words": 4,
        "max_item_body_words":  20,
        "style":                "3 à 5 items. Chaque item : titre en gras court + phrase de développement concise.",
        "never":                "Jamais plus de 5 items. Jamais sans titre par item.",
        "example":              "'Lobbying institutionnel' / '~2,3 M€ déclarés/an — contacts réguliers Élysée, Bercy.'",
    },
    "image_text": {
        "max_title_words":  8,
        "max_body_words":   40,
        "max_body_items":   3,
        "style":            "Titre + corps structuré en 2-3 points courts. L'image fait le travail visuel.",
        "never":            "Jamais de mur de texte. Jamais plus de 3 bullet points.",
    },
    "full_text": {
        "max_title_words":  8,
        "max_body_words":   60,
        "max_paragraphs":   3,
        "style":            "Titre + 2-3 paragraphes courts et aérés. Chaque paragraphe = une idée.",
        "never":            "Jamais de corps > 60 mots. Jamais plus de 3 paragraphes.",
    },
    "closing": {
        "max_title_words":    5,
        "max_subtitle_words": 15,
        "style":              "Message mémorable ou 'Merci !' + sous-titre : sources, contact ou CTA.",
        "never":              "Jamais de bullet points. Jamais de corps long. Simple, élégant.",
        "example_title":      "Merci !",
        "example_subtitle":   "Sources : Rapport Annuel 2023 · HATVP · Loi de vigilance 2017",
    },
    "complex": {
        "max_label_words": 4,
        "max_body_words":  15,
        "style":           "Labels ultra-courts pour les éléments du diagramme.",
        "never":           "Jamais de phrases complètes dans un diagramme.",
    },
    "unknown": {
        "max_title_words": 8,
        "max_body_words":  35,
        "style":           "Titre + corps aéré. Respecter la structure du template.",
        "never":           "Jamais de surcharge.",
    },
}

WORD_LIMITS = {
    "title":       8,
    "subtitle":   15,
    "section_num": 2,
    "label":       5,
    "kpi_value":   3,
    "kpi_label":   5,
    "body":       40,
    "footer":     10,
    "page_number": 1,
    "quote":      20,
    "list_item":  18,
    "placeholder": 8,
    "text":       30,
}


def _classify_slide(slide, idx: int, total: int, w: int, h: int) -> str:
    shapes      = list(iter_all_shapes(slide.shapes))
    raw_shapes  = list(slide.shapes)
    img_shapes  = [s for s in shapes if s.shape_type in (13, 11)]
    group_shapes = [s for s in raw_shapes if s.shape_type == MSO_SHAPE_TYPE.GROUP]

    texts = []
    for s in shapes:
        if not getattr(s, "has_text_frame", False):
            continue
        for para in s.text_frame.paragraphs:
            t = "".join(r.text for r in para.runs).strip()
            if len(t) > 3:
                texts.append({"text": t, "shape": s})

    n           = len(texts)
    total_chars = sum(len(t["text"]) for t in texts)

    if idx == 0:
        return "cover"
    if idx == total - 1:
        return "closing"
    if len(group_shapes) >= 5:
        return "complex"
    if img_shapes and n >= 2:
        return "image_text"
    if n <= 3 and total_chars < 100:
        return "section"
    if n >= 4:
        short = [t for t in texts if len(t["text"]) < 25]
        if len(short) >= 3:
            try:
                lefts = sorted([t["shape"].left for t in texts])
                spread = _emu(lefts[-1]) - _emu(lefts[0])
                if spread > 4.0:
                    return "kpi"
            except Exception:
                pass
    if n >= 3 and total_chars < 500:
        try:
            tops = sorted([t["shape"].top for t in texts])
            if _emu(tops[-1]) - _emu(tops[0]) > 2.5:
                return "timeline"
        except Exception:
            pass
    if n >= 4:
        try:
            half = w / 2
            lc = sum(1 for t in texts if t["shape"].left < half)
            rc = sum(1 for t in texts if t["shape"].left >= half)
            if lc >= 2 and rc >= 2:
                return "two_col"
        except Exception:
            pass
    if n <= 3 and any(len(t["text"]) > 60 for t in texts):
        return "quote"
    if n >= 3:
        lengths = [len(t["text"]) for t in texts]
        avg = sum(lengths) / len(lengths)
        var = sum((l - avg) ** 2 for l in lengths) / len(lengths)
        if var < 800 and avg < 150:
            return "list"
    if total_chars > 200:
        return "full_text"
    return "unknown"


def _shape_role(shape, w: int, h: int) -> str:
    """Détermine le rôle via placeholder natif PPTX, puis géométrie."""
    try:
        if shape.is_placeholder:
            ph = shape.placeholder_format
            ph_map = {
                0: "title", 1: "body", 2: "subtitle",
                3: "date", 4: "footer", 5: "page_number",
                13: "title", 15: "subtitle",
            }
            return ph_map.get(ph.idx, "placeholder")
    except Exception:
        pass

    if not getattr(shape, "has_text_frame", False):
        return "decoration"

    try:
        tr = shape.top    / h
        wr = shape.width  / w
        hr = shape.height / h

        if tr > 0.87 and wr > 0.15:
            return "footer"
        if tr > 0.87:
            return "page_number"
        if tr < 0.28 and wr > 0.45:
            return "title"
        if 0.20 < tr < 0.50 and wr > 0.35 and hr < 0.15:
            return "subtitle"
        if 0.25 < tr < 0.75 and wr < 0.40 and hr < 0.12:
            return "label"
        if 0.25 < tr < 0.87:
            return "body"
    except Exception:
        pass

    return "text"


def extract_brand(prs: Presentation) -> dict:
    fonts, colors = set(), set()
    for slide in prs.slides:
        for shape in iter_all_shapes(slide.shapes):
            if not getattr(shape, "has_text_frame", False):
                continue
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

    w, h = prs.slide_width, prs.slide_height
    return {
        "fonts":           list(fonts)[:5],
        "colors":          list(colors)[:8],
        "slide_count":     len(prs.slides),
        "layouts":         [l.name for l in prs.slide_layouts],
        "slide_width_in":  round(_emu(w), 2),
        "slide_height_in": round(_emu(h), 2),
        "aspect_ratio":    (
            "16:9" if abs(w / h - 16 / 9) < 0.05 else
            "4:3"  if abs(w / h - 4 / 3)  < 0.05 else "custom"
        ),
    }


def build_layout_library(prs: Presentation) -> list:
    library = []
    total = len(prs.slides)
    w, h  = prs.slide_width, prs.slide_height

    for idx, slide in enumerate(prs.slides):
        slide_type  = _classify_slide(slide, idx, total, w, h)
        root_groups = sum(1 for s in slide.shapes if s.shape_type == MSO_SHAPE_TYPE.GROUP)
        has_images  = any(s.shape_type in (13, 11) for s in iter_all_shapes(slide.shapes))

        zones = []
        seen  = set()

        for shape in iter_all_shapes(slide.shapes):
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

                is_placeholder_footer = _is_footer_placeholder(text)
                zones.append({
                    "original_text":         text,
                    "role":                  "footer" if is_placeholder_footer else role,
                    "word_count":            len(text.split()),
                    "word_limit":            WORD_LIMITS.get(
                        "footer" if is_placeholder_footer else role, 30
                    ),
                    "is_placeholder_footer": is_placeholder_footer,
                    "char_count":            len(text),
                })

        if zones:
            library.append({
                "slide_index":  idx,
                "slide_type":   slide_type,
                "description":  SLIDE_TYPE_DESC.get(slide_type, ""),
                "position":     (
                    "cover"   if idx == 0 else
                    "closing" if idx == total - 1 else
                    f"{idx+1}/{total}"
                ),
                "root_groups":  root_groups,
                "has_images":   has_images,
                "zones":        zones,
                "total_words":  sum(z["word_count"] for z in zones),
                "visual_score": (
                    has_images * 3 +
                    (1 if slide_type in ("kpi", "timeline", "two_col", "image_text", "quote") else 0) * 2 +
                    (1 if root_groups > 0 else 0) * 1
                ),
            })

    return library


def select_template_slides(library: list, nb_slides: int) -> list:
    if not library:
        return []

    cover   = [s for s in library if s["slide_type"] == "cover"]
    closing = [s for s in library if s["slide_type"] == "closing"]
    middle  = [s for s in library if s["slide_type"] not in ("cover", "closing")]

    middle_sorted = sorted(
        middle,
        key=lambda s: (
            0 if s["slide_type"] == "complex" else 1,
            s["visual_score"],
            -s["total_words"],
        ),
        reverse=True,
    )

    n_cover   = min(len(cover), 1)
    n_closing = min(len(closing), 1)
    n_middle  = nb_slides - n_cover - n_closing

    if n_middle < 0:
        n_middle = 0

    selected_middle = []
    last_type = None
    pool = middle_sorted.copy()

    while len(selected_middle) < n_middle and pool:
        for i, s in enumerate(pool):
            if s["slide_type"] != last_type or i == len(pool) - 1:
                selected_middle.append(s)
                last_type = s["slide_type"]
                pool.pop(i)
                break

    if len(selected_middle) < n_middle and middle_sorted:
        cycle = middle_sorted * 10
        for s in cycle:
            if len(selected_middle) >= n_middle:
                break
            dup = {**s, "duplicate": True}
            selected_middle.append(dup)

    result = (
        cover[:n_cover] +
        sorted(selected_middle, key=lambda s: s["slide_index"]) +
        closing[:n_closing]
    )

    return result[:nb_slides]


# ══════════════════════════════════════════════════════════════
# PHASE 2 — PLANIFICATION NARRATIVE
# ══════════════════════════════════════════════════════════════

PLANNER_SYSTEM = """Tu es Visual Cortex Planner, architecte narratif de présentations B2B professionnelles.

MISSION : concevoir une structure narrative percutante et cohérente, slide par slide.

PRINCIPES NARRATIFS :
- Cover : accroche forte, crée une tension ou une promesse. Titre ≤ 7 mots.
- Slides de contenu : progression logique et variée —
  ne jamais enchaîner deux fois le même type de slide.
  Séquences éprouvées : contexte → enjeux → solution → preuves → bénéfices → ROI → next steps
- Closing : message simple et mémorable. CTA clair ou remerciement élégant.

AFFECTATION DES TYPES AUX CONTENUS :
- données chiffrées, métriques, résultats     → kpi
- étapes, jalons, chronologie, processus       → timeline
- comparaison, dualité, pour/contre, avant/après → two_col
- message fort, conviction, prise de position  → quote
- liste d'acteurs, d'arguments, de features    → list
- concept visuel, cas d'usage, illustration    → image_text
- argumentaire développé, analyse, contexte    → full_text
- séparation de partie                         → section

DENSITÉ PAR TYPE (à indiquer dans visual_hint) :
- kpi       : 4 à 6 chiffres max, valeur courte + label + sous-label
- timeline  : 4 à 6 jalons max, date + titre court + phrase optionnelle
- two_col   : 2 colonnes symétriques, 3-4 items de ≤ 10 mots chacun
- quote     : 1 citation de ≤ 20 mots, source courte
- list      : 3 à 5 items structurés (titre gras + corps ≤ 20 mots)
- section   : numéro (01) + titre ≤ 6 mots, rien d'autre
- cover     : titre ≤ 7 mots + sous-titre ≤ 12 mots
- closing   : titre ≤ 5 mots + sous-titre/sources ≤ 15 mots

RÈGLE ABSOLUE : le tableau "slides" doit contenir EXACTEMENT {nb_slides} entrées.

Réponds UNIQUEMENT en JSON valide, sans markdown."""

PLANNER_USER = """SUJET : {prompt}
NB SLIDES SOUHAITÉ : {nb_slides}

SLIDES DISPONIBLES DANS LE TEMPLATE :
{selection_json}

CHARTE : Polices {fonts} | Couleurs {colors} | Format {ratio}

Génère le plan. Le tableau "slides" doit contenir EXACTEMENT {nb_slides} entrées.

FORMAT :
{{
  "presentation_title": "Titre accrocheur ≤ 7 mots",
  "narrative_arc": "Logique narrative en 1 phrase",
  "footer_text": "Entreprise · Contexte · Année  (≤ 8 mots)",
  "slides": [
    {{
      "plan_index": 0,
      "template_slide_index": 0,
      "slide_type": "cover",
      "narrative_angle": "Ce que cette slide accomplit dans l'histoire (1 phrase)",
      "key_message": "Le message principal ≤ 10 mots",
      "visual_hint": "Contrainte de densité et de style pour cette slide"
    }}
  ]
}}\n"""


def plan_presentation(prompt: str, nb_slides: int, selection: list, brand: dict) -> dict:
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    sel_light = [
        {
            "template_slide_index": s["slide_index"],
            "slide_type":           s["slide_type"],
            "description":          s["description"],
            "zone_roles":           [z["role"] for z in s["zones"]],
            "has_images":           s["has_images"],
            "visual_score":         s["visual_score"],
        }
        for s in selection
    ]

    system = PLANNER_SYSTEM.replace("{nb_slides}", str(nb_slides))

    user = PLANNER_USER.format(
        prompt         = prompt,
        nb_slides      = nb_slides,
        selection_json = json.dumps(sel_light, ensure_ascii=False, indent=2),
        fonts          = ", ".join(brand.get("fonts", [])) or "Standard",
        colors         = ", ".join(f"#{c}" for c in brand.get("colors", [])) or "non détectées",
        ratio          = brand.get("aspect_ratio", "16:9"),
    )

    # max_tokens adaptatif : ~120 tokens/slide suffisent pour le plan JSON
    planner_tokens = max(1200, nb_slides * 130)

    msg = client.messages.create(
        model=CLAUDE_MODEL, max_tokens=planner_tokens,
        system=system,
        messages=[{"role": "user", "content": user}],
    )
    plan = json.loads(_clean_json(msg.content[0].text.strip()))
    log.info(f"Plan: {len(plan.get('slides', []))} slides — {plan.get('narrative_arc', '')[:80]}")
    return plan


# ══════════════════════════════════════════════════════════════
# PHASE 3 — GÉNÉRATION DU CONTENU (visuel-first, Niveau 1)
# ══════════════════════════════════════════════════════════════

CORTEX_SYSTEM = """Tu es Visual Cortex, expert en présentations B2B professionnelles et visuelles.
Philosophie : une slide = une idée. Le texte est une accroche, pas un rapport.

═══════════════════════════════════════════════════
RÈGLES UNIVERSELLES (s'appliquent à toutes les slides)
═══════════════════════════════════════════════════
1. LIMITES PAR RÔLE — ne jamais dépasser :
   title       → ≤ 8 mots   subtitle    → ≤ 12 mots  label       → ≤ 5 mots
   kpi_value   → ≤ 3 mots   kpi_label   → ≤ 5 mots   body        → ≤ 40 mots
   list_item   → ≤ 18 mots  quote       → ≤ 20 mots  footer      → ≤ 8 mots
   page_number → NE PAS MODIFIER

2. COHÉRENCE : footer identique sur toutes les slides de contenu.
3. B2B : vocabulaire du secteur, ton direct, orienté valeur, zéro formule creuse.
4. PROGRESSION : chaque slide fait avancer l'histoire selon son narrative_angle.
5. ZÉRO invention de données, chiffres ou noms non fournis dans le prompt.
6. FOOTERS PLACEHOLDER : zones "is_placeholder_footer: true" → remplacer par le footer_text.

═══════════════════════════════════════════════════
RÈGLES PAR TYPE DE SLIDE (density rules)
═══════════════════════════════════════════════════

[cover] Titre ≤ 7 mots — accroche forte. Sous-titre ≤ 12 mots.
[section] Numéro (01, 02…) + titre ≤ 6 mots. RIEN D'AUTRE.
[kpi] 4 à 6 KPIs MAX. Valeur courte + label + sous-label contextuel.
[timeline] 4 à 6 jalons MAX. Repère temporel + titre ≤ 4 mots + phrase optionnelle.
[two_col] 2 colonnes SYMÉTRIQUES — max 4 items/colonne ≤ 18 mots.
[quote] 1 SEULE citation ≤ 20 mots. Source optionnelle ≤ 6 mots.
[list] 3 à 5 items MAX. Titre ≤ 5 mots + corps ≤ 20 mots. Structure parallèle.
[image_text] Titre ≤ 8 mots + corps 2-3 points ≤ 40 mots total.
[full_text] Titre ≤ 8 mots + 2-3 §§ courts = total body ≤ 60 mots.
[closing] Titre ≤ 5 mots. Sous-titre ≤ 15 mots. Simple, élégant.
[complex] Labels ≤ 4 mots. Titres ≤ 5 mots. Jamais de phrases complètes.

Réponds UNIQUEMENT en JSON valide, sans commentaire ni markdown."""

CORTEX_USER = """PRÉSENTATION : {title}
ARC NARRATIF : {arc}
FOOTER : "{footer}"
SUJET COMPLET : {prompt}
CHARTE : Polices {fonts} | Couleurs {colors}

═══════════════════════
SLIDES À GÉNÉRER — {n} slides
═══════════════════════
{slides_json}

FORMAT DE SORTIE (clés = template_slide_index) :
{{
  "0": {{"Texte original exact": "Nouveau texte généré"}},
  ...
}}"""


def generate_content(prompt: str, plan: dict, selection: list, brand: dict) -> dict:
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    sel_by_idx  = {s["slide_index"]: s for s in selection}
    footer_text = plan.get("footer_text", "")

    slides_payload = []
    for sp in plan.get("slides", []):
        tidx       = sp.get("template_slide_index", 0)
        tmpl       = sel_by_idx.get(tidx, {})
        slide_type = sp.get("slide_type", "unknown")
        density    = DENSITY_RULES.get(slide_type, DENSITY_RULES["unknown"])

        slides_payload.append({
            "template_slide_index": tidx,
            "slide_type":           slide_type,
            "narrative_angle":      sp.get("narrative_angle"),
            "key_message":          sp.get("key_message"),
            "visual_hint":          sp.get("visual_hint", ""),
            "density_rules": {
                "style":  density.get("style", ""),
                "never":  density.get("never", ""),
                "limits": {
                    k: v for k, v in density.items()
                    if k not in ("style", "never", "example", "example_title",
                                 "example_subtitle", "example_label")
                },
            },
            "zones": [
                {
                    "original_text":         z["original_text"],
                    "role":                  z["role"],
                    "word_limit":            z["word_limit"],
                    "char_count":            z.get("char_count", 0),
                    "is_placeholder_footer": z.get("is_placeholder_footer", False),
                }
                for z in tmpl.get("zones", [])
            ],
        })

    user = CORTEX_USER.format(
        title       = plan.get("presentation_title", prompt[:60]),
        arc         = plan.get("narrative_arc", ""),
        footer      = footer_text,
        prompt      = prompt,
        fonts       = ", ".join(brand.get("fonts", [])) or "Standard",
        colors      = ", ".join(f"#{c}" for c in brand.get("colors", [])) or "non détectées",
        n           = len(slides_payload),
        slides_json = json.dumps(slides_payload, ensure_ascii=False, indent=2),
    )

    # max_tokens adaptatif : ~220 tokens/slide (JSON + textes courts)
    content_tokens = max(2000, len(slides_payload) * 220)

    msg = client.messages.create(
        model=CLAUDE_MODEL, max_tokens=content_tokens,
        system=CORTEX_SYSTEM,
        messages=[{"role": "user", "content": user}],
    )

    raw     = _clean_json(msg.content[0].text.strip())
    mapping = json.loads(raw)
    mapping = _validate_and_trim(mapping, slides_payload)

    log.info(f"Contenu généré et validé : {len(mapping)} slides.")
    return mapping


def _validate_and_trim(mapping: dict, slides_payload: list) -> dict:
    zone_limits: dict = {}
    for sp in slides_payload:
        tidx = str(sp["template_slide_index"])
        for z in sp.get("zones", []):
            key = (tidx, z["original_text"])
            zone_limits[key] = {
                "word_limit": z["word_limit"],
                "role":       z["role"],
            }

    validated = {}
    for slide_key, replacements in mapping.items():
        validated[slide_key] = {}
        for orig, new_text in replacements.items():
            if not new_text:
                validated[slide_key][orig] = new_text
                continue

            zone_info = zone_limits.get((str(slide_key), orig), {})
            role      = zone_info.get("role", "text")
            limit     = zone_info.get("word_limit", WORD_LIMITS.get(role, 40))

            if role == "page_number":
                validated[slide_key][orig] = orig
                continue

            words = new_text.split()
            if len(words) > limit:
                trimmed = " ".join(words[:limit])
                for punct in [".", ",", ";", ":", "—", "–"]:
                    last = trimmed.rfind(punct)
                    if last > len(trimmed) * 0.6:
                        trimmed = trimmed[:last + 1].strip()
                        break
                log.debug(f"Trim slide {slide_key} role={role}: {len(words)}→{len(trimmed.split())} mots")
                validated[slide_key][orig] = trimmed
            else:
                validated[slide_key][orig] = new_text

    return validated


# ══════════════════════════════════════════════════════════════
# HYDRATATION NIVEAU 1 — Injection avec traversée récursive
# ══════════════════════════════════════════════════════════════

def _replace_text_in_para(para, replacements: dict):
    para_text = "".join(r.text for r in para.runs).strip()
    if not para_text:
        return

    new_text = replacements.get(para_text)
    if new_text is None:
        normalized = para_text.replace("\u2019", "'").replace("\u2018", "'")
        for k, v in replacements.items():
            k_norm = k.replace("\u2019", "'").replace("\u2018", "'")
            if k_norm == normalized:
                new_text = v
                break

    if not new_text:
        return

    rpr_xml = None
    if para.runs:
        rpr_el = para.runs[0]._r.find(qn("a:rPr"))
        if rpr_el is not None:
            rpr_xml = copy.deepcopy(rpr_el)

    para.text = new_text

    if rpr_xml is not None:
        for run in para.runs:
            ex = run._r.find(qn("a:rPr"))
            if ex is not None:
                run._r.remove(ex)
            run._r.insert(0, copy.deepcopy(rpr_xml))


def _hydrate_slide(slide, replacements: dict):
    for shape in iter_all_shapes(slide.shapes):
        if not getattr(shape, "has_text_frame", False):
            continue
        for para in shape.text_frame.paragraphs:
            _replace_text_in_para(para, replacements)


def hydrate_presentation(
    pptx_bytes: bytes,
    mapping: dict,
    plan_slides: list,
    nb_slides: int,
) -> bytes:
    prs = Presentation(io.BytesIO(pptx_bytes))
    template_indices = [s.get("template_slide_index", 0) for s in plan_slides]
    _reorder_and_hydrate(prs, template_indices, mapping, nb_slides)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out.read()


def _reorder_and_hydrate(prs: Presentation, template_indices: list, mapping: dict, nb_slides: int):
    total_tmpl = len(prs.slides)

    for slide_key, replacements in mapping.items():
        try:
            idx = int(str(slide_key).replace("slide_", ""))
            if idx < total_tmpl:
                _hydrate_slide(prs.slides[idx], replacements)
        except Exception as e:
            log.warning(f"Hydratation slide {slide_key}: {e}")

    xml_slides  = prs.slides._sldIdLst
    all_sld_ids = list(xml_slides)

    new_order = []
    for tidx in template_indices:
        if 0 <= tidx < len(all_sld_ids):
            new_order.append(all_sld_ids[tidx])

    seen_ids   = set()
    final_order = []
    for sld_el in new_order:
        sld_id = sld_el.get("id")
        if sld_id not in seen_ids:
            seen_ids.add(sld_id)
            final_order.append(sld_el)
        else:
            dup = _duplicate_slide_element(prs, sld_el)
            if dup is not None:
                final_order.append(dup)

    final_order = final_order[:nb_slides]

    for sld in list(xml_slides):
        xml_slides.remove(sld)
    for sld in final_order:
        xml_slides.append(sld)

    _cleanup_orphan_slides(prs, final_order)


def _duplicate_slide_element(prs: Presentation, src_sld_el):
    try:
        src_rId  = src_sld_el.get(qn("r:id"))
        src_part = prs.part.related_parts.get(src_rId)
        if src_part is None:
            return None

        blank_layout = prs.slide_layouts[-1]
        new_slide    = prs.slides.add_slide(blank_layout)

        import lxml.etree as etree
        src_xml = copy.deepcopy(src_part._element)
        new_slide._element.getparent().replace(new_slide._element, src_xml)

        return list(prs.slides._sldIdLst)[-1]
    except Exception as e:
        log.warning(f"Duplication slide: {e}")
        return None


def _cleanup_orphan_slides(prs: Presentation, kept_sld_els: list):
    # drop_rel est instable selon les versions python-pptx — la sldIdLst
    # controle les slides affichees, les rels orphelins sont ignores par PowerPoint.
    pass


# ══════════════════════════════════════════════════════════════
# PIPELINE NIVEAU 1
# ══════════════════════════════════════════════════════════════

def run_pipeline(pptx_bytes: bytes, prompt: str, nb_slides: int) -> tuple:
    if not ANTHROPIC_API_KEY:
        raise ValueError("Clé API Claude manquante.")

    prs = Presentation(io.BytesIO(pptx_bytes))
    nb_slides = max(2, min(nb_slides, 30))

    log.info("Phase 1 : analyse du template...")
    brand     = extract_brand(prs)
    library   = build_layout_library(prs)
    selection = select_template_slides(library, nb_slides)
    log.info(f"Template : {len(library)} slides → {len(selection)} sélectionnées pour {nb_slides} demandées")

    log.info("Phase 2 : planification narrative...")
    plan = plan_presentation(prompt, nb_slides, selection, brand)

    plan_slides = plan.get("slides", [])
    while len(plan_slides) < nb_slides:
        fallback = selection[min(len(plan_slides), len(selection) - 2)]
        plan_slides.append({
            "plan_index":           len(plan_slides),
            "template_slide_index": fallback["slide_index"],
            "slide_type":           fallback["slide_type"],
            "narrative_angle":      "Développement complémentaire",
            "key_message":          "Argument additionnel",
            "visual_hint":          "",
        })
    plan["slides"] = plan_slides[:nb_slides]

    log.info("Phase 3 : génération du contenu...")
    mapping = generate_content(prompt, plan, selection, brand)

    log.info("Hydratation PPTX...")
    final_bytes = hydrate_presentation(pptx_bytes, mapping, plan["slides"], nb_slides)

    return final_bytes, plan, brand


# ══════════════════════════════════════════════════════════════
# ROUTES NIVEAU 1
# ══════════════════════════════════════════════════════════════

@app.get("/")
def root():
    return {"status": "ok", "version": "13.0.0", "model": CLAUDE_MODEL,
            "levels": ["L1: /generate", "L2: /generate-v2"]}


@app.post("/analyze-template")
async def analyze_template(file: UploadFile = File(...)):
    pptx_bytes = await file.read()
    prs   = Presentation(io.BytesIO(pptx_bytes))
    brand = extract_brand(prs)
    lib   = build_layout_library(prs)

    type_counts: dict = defaultdict(int)
    for s in lib:
        type_counts[s["slide_type"]] += 1

    fonts_display = ", ".join(brand["fonts"]) if brand["fonts"] else "Standard"

    return {
        "success":     True,
        "message":     (
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
    pro = _is_pro(authorization)
    quota_info = (
        {"plan": "pro"} if pro
        else {"used": _quota(_ip(request))[0], "total": FREE_QUOTA_PER_IP, "plan": "free"}
    )

    nb_slides  = max(2, min(nb_slides, 30))
    pptx_bytes = await template.read()
    prs        = Presentation(io.BytesIO(pptx_bytes))
    brand      = extract_brand(prs)
    lib        = build_layout_library(prs)
    sel        = select_template_slides(lib, nb_slides)
    plan       = plan_presentation(prompt, nb_slides, sel, brand)

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
    if not _is_pro(authorization):
        _quota(_ip(request))

    pptx_bytes = await template.read()
    final_bytes, plan, _ = run_pipeline(pptx_bytes, prompt, nb_slides)

    filename = f"visualcortex-{_safe_name(prompt)}.pptx"
    return StreamingResponse(
        io.BytesIO(final_bytes),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={
            "Content-Disposition": f"attachment; filename={filename}",
            "Content-Length":      str(len(final_bytes)),
        },
    )


# ─────────────────────────────────────────────────────────────
# GÉNÉRATION AVEC PROGRESSION SSE
# Affiche la progression en temps réel côté frontend.
# Protocole : POST multipart, réponse text/event-stream
# Chaque event : { step, label, progress (0-100) }
# Event final "done" : + file_b64 (base64), filename
# ─────────────────────────────────────────────────────────────

_PROGRESS_LABELS = {
    "start":     ("Analyse du template…",           5),
    "planned":   ("Structure narrative créée…",     35),
    "generated": ("Contenu généré…",                70),
    "hydrating": ("Assemblage de la présentation…", 90),
    "done":      ("Prêt !",                         100),
    "error":     ("Erreur…",                         0),
}

def _sse(event: str, data: dict) -> str:
    label, pct = _PROGRESS_LABELS.get(event, (event, 50))
    payload = json.dumps({"step": event, "label": label, "progress": pct, **data},
                         ensure_ascii=False)
    return f"data: {payload}\n\n"


@app.post("/generate-stream")
async def generate_stream(
    request:       Request,
    template:      UploadFile = File(...),
    prompt:        str        = Form(...),
    nb_slides:     int        = Form(default=8),
    authorization: str        = Form(default=None),
):
    """
    Génération avec progression SSE.
    Event final "done" contient le fichier PPTX encodé en base64.
    """
    if not _is_pro(authorization):
        _quota(_ip(request))

    pptx_bytes = await template.read()

    async def _stream():
        import asyncio, base64
        loop = asyncio.get_event_loop()

        try:
            yield _sse("start", {"nb_slides": nb_slides})

            prs       = Presentation(io.BytesIO(pptx_bytes))
            brand     = extract_brand(prs)
            library   = build_layout_library(prs)
            selection = select_template_slides(library, nb_slides)

            # Appel Claude 1 — planning (exécuté dans un thread pour ne pas bloquer)
            plan = await loop.run_in_executor(
                None,
                lambda: plan_presentation(prompt, nb_slides, selection, brand)
            )
            yield _sse("planned", {"title": plan.get("presentation_title", "")})

            # Compléter le plan si insuffisant
            plan_slides = plan.get("slides", [])
            while len(plan_slides) < nb_slides:
                fb = selection[min(len(plan_slides), len(selection) - 2)]
                plan_slides.append({
                    "plan_index":           len(plan_slides),
                    "template_slide_index": fb["slide_index"],
                    "slide_type":           fb["slide_type"],
                    "narrative_angle":      "Développement complémentaire",
                    "key_message":          "Argument additionnel",
                    "visual_hint":          "",
                })
            plan["slides"] = plan_slides[:nb_slides]

            # Appel Claude 2 — génération contenu
            mapping = await loop.run_in_executor(
                None,
                lambda: generate_content(prompt, plan, selection, brand)
            )
            yield _sse("generated", {})

            # Hydratation PPTX
            yield _sse("hydrating", {})
            final_bytes = hydrate_presentation(
                pptx_bytes, mapping, plan["slides"], nb_slides
            )

            # Envoi du fichier en base64 dans l'event final
            b64      = base64.b64encode(final_bytes).decode()
            filename = f"visualcortex-{_safe_name(prompt)}.pptx"
            yield _sse("done", {"file_b64": b64, "filename": filename})

        except Exception as e:
            log.error(f"[stream] Erreur : {e}", exc_info=True)
            yield _sse("error", {"message": str(e)})

    return StreamingResponse(
        _stream(),
        media_type="text/event-stream",
        headers={
            "Cache-Control":               "no-cache",
            "X-Accel-Buffering":           "no",
            "Access-Control-Allow-Origin": "*",
        },
    )


# ══════════════════════════════════════════════════════════════════════════════
#
#  ███╗   ██╗██╗██╗   ██╗███████╗ █████╗ ██╗   ██╗    ██████╗
#  ████╗  ██║██║██║   ██║██╔════╝██╔══██╗██║   ██║    ╚════██╗
#  ██╔██╗ ██║██║██║   ██║█████╗  ███████║██║   ██║     █████╔╝
#  ██║╚██╗██║██║╚██╗ ██╔╝██╔══╝  ██╔══██║██║   ██║    ██╔═══╝
#  ██║ ╚████║██║ ╚████╔╝ ███████╗██║  ██║╚██████╔╝    ███████╗
#  ╚═╝  ╚═══╝╚═╝  ╚═══╝  ╚══════╝╚═╝  ╚═╝ ╚═════╝    ╚══════╝
#
#  GÉNÉRATION CRÉATIVE — Brand Book → Slides nouvelles de A à Z
#
# ══════════════════════════════════════════════════════════════════════════════

# ─────────────────────────────────────────────────────────────
# PROFILS CLIENT
# ─────────────────────────────────────────────────────────────

CLIENT_PROFILES = {
    "finance": {
        "label":       "Finance / Conseil",
        "description": "Premium, minimaliste, données en avant",
        "style_guide": (
            "Backgrounds blancs ou très clairs. Grandes marges. Hiérarchie typographique nette. "
            "Accent doré ou bleu marine. Données chiffrées proéminentes. "
            "Zéro décoration superflue. Lignes fines. Espacement aéré."
        ),
        "layout_prefs": ["kpi", "two_col", "full_text", "quote", "timeline"],
        "bg_dark":      False,
        "accent_style": "thin_lines",
    },
    "industrial": {
        "label":       "Industriel / Technique",
        "description": "Clarté, données, schémas, sobriété",
        "style_guide": (
            "Fond blanc ou gris très clair. Sans-serif sobre. Données et faits en premier. "
            "Couleurs d'accent : bleu, gris, orange sécurité. "
            "Timelines et schémas favorisés. Densité de contenu modérée à haute."
        ),
        "layout_prefs": ["timeline", "kpi", "two_col", "list"],
        "bg_dark":      False,
        "accent_style": "bar_left",
    },
    "institutional": {
        "label":       "Institutionnel / Public",
        "description": "Formalismes respectés, sobre, structuré",
        "style_guide": (
            "Fond blanc. Typographie classique. Structure formelle et lisible. "
            "Éviter les effets visuels marqués. Hiérarchie claire. "
            "Bleu institutionnel, blanc, rouge discret. Sobriété maximale."
        ),
        "layout_prefs": ["full_text", "list", "two_col", "timeline"],
        "bg_dark":      False,
        "accent_style": "bar_left",
    },
    "startup": {
        "label":       "Startup / Tech",
        "description": "Moderne, aéré, accents colorés, iconographie",
        "style_guide": (
            "Fond très clair ou très sombre. Sans-serif bold. Grands espaces blancs. "
            "1-2 couleurs vives en accent. Peu de texte, impact fort. "
            "Icônes, chiffres oversize, layout asymétrique bienvenu."
        ),
        "layout_prefs": ["cover", "quote", "kpi", "image_text", "closing"],
        "bg_dark":      True,
        "accent_style": "bold_color",
    },
    "creative": {
        "label":       "Créatif / Agence",
        "description": "Audacieux, typographie expressive, compositions asymétriques",
        "style_guide": (
            "Fond sombre ou couleur franche. Typographie oversize. Compositions audacieuses. "
            "Couleurs inattendues. Géométrie forte. "
            "Peu de texte — chaque mot compte. Impact visuel prime sur tout."
        ),
        "layout_prefs": ["cover", "quote", "image_text", "section", "closing"],
        "bg_dark":      True,
        "accent_style": "bold_color",
    },
}


# ─────────────────────────────────────────────────────────────
# HELPERS H2_* — Fonctions de génération de shapes
# Ces fonctions sont exposées dans le sandbox exec() du Niveau 2
# ─────────────────────────────────────────────────────────────

def _h2_parse_hex(hex_str: str) -> RGBColor:
    """Parse 'RRGGBB' ou '#RRGGBB' → RGBColor. Fallback bleu corporate si invalide."""
    try:
        h = str(hex_str).lstrip("#").strip()
        if len(h) == 3:
            h = h[0] * 2 + h[1] * 2 + h[2] * 2
        if len(h) != 6:
            return RGBColor(0x1A, 0x3A, 0x6B)
        return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    except Exception:
        return RGBColor(0x1A, 0x3A, 0x6B)


def _h2_blank_slide(prs: Presentation):
    """
    Ajoute une slide entièrement vierge (sans placeholders résiduels).
    Retourne (slide, W, H) en pouces.
    Cherche un layout 'Blank' par nom, sinon prend le dernier layout.
    """
    target = None
    for layout in prs.slide_layouts:
        if "blank" in layout.name.lower() or layout.name.strip() == "":
            target = layout
            break
    if target is None and len(prs.slide_layouts) > 6:
        target = prs.slide_layouts[6]
    if target is None:
        target = prs.slide_layouts[-1]

    slide = prs.slides.add_slide(target)

    # Supprimer les placeholders résiduels du layout
    sp_tree = slide.shapes._spTree
    for ph in list(slide.placeholders):
        try:
            sp_tree.remove(ph._element)
        except Exception:
            pass

    W = prs.slide_width  / 914400.0
    H = prs.slide_height / 914400.0
    return slide, W, H


def _h2_rect(slide, left: float, top: float, width: float, height: float, color: str):
    """
    Ajoute un rectangle coloré (en pouces).
    color = hex string 'RRGGBB' ou brand["primary"] etc.
    """
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        Inches(left), Inches(top), Inches(width), Inches(height),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = _h2_parse_hex(color)
    shape.line.fill.background()  # pas de bordure
    return shape


def _h2_text(slide, text: str,
             left: float, top: float, width: float, height: float,
             font: str, size_pt: float, color: str,
             bold: bool = False, italic: bool = False, align: str = "left"):
    """
    Ajoute une textbox stylée (dimensions en pouces).
    align : "left" | "center" | "right"
    color : hex string 'RRGGBB'
    """
    txBox = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height),
    )
    tf          = txBox.text_frame
    tf.word_wrap = True

    align_map = {
        "left":   PP_ALIGN.LEFT,
        "center": PP_ALIGN.CENTER,
        "right":  PP_ALIGN.RIGHT,
    }

    # Vider le paragraphe par défaut et le configurer
    p           = tf.paragraphs[0]
    p.alignment = align_map.get(align, PP_ALIGN.LEFT)

    run            = p.add_run()
    run.text       = str(text)
    run.font.name  = str(font)
    run.font.size  = Pt(size_pt)
    run.font.bold  = bold
    run.font.italic = italic
    run.font.color.rgb = _h2_parse_hex(color)

    return txBox


def _h2_kpi(slide,
            left: float, top: float, width: float,
            value: str, label: str, sublabel: str,
            palette: dict):
    """
    Bloc KPI : grande valeur + label + sous-label.
    Conçu pour fond sombre (couleurs blanches/accent).
    width ≈ 2.5–3.0 pouces.
    """
    font     = palette.get("font", "Calibri")
    v_color  = "FFFFFF"
    l_color  = palette.get("accent", "F0A500")
    sl_color = "CCCCCC"

    # Valeur principale (grande, blanche, bold)
    _h2_text(slide, str(value),
             left, top, width, 0.72,
             font, 34, v_color, bold=True, align="center")
    # Ligne de séparation accent
    _h2_rect(slide, left + width * 0.2, top + 0.72, width * 0.6, 0.03, l_color)
    # Label
    _h2_text(slide, str(label),
             left, top + 0.78, width, 0.40,
             font, 12, l_color, bold=False, align="center")
    # Sous-label
    _h2_text(slide, str(sublabel),
             left, top + 1.18, width, 0.50,
             font, 10, sl_color, bold=False, align="center")


def _h2_kpi_light(slide,
                  left: float, top: float, width: float,
                  value: str, label: str, sublabel: str,
                  palette: dict):
    """
    Bloc KPI pour fond clair.
    """
    font     = palette.get("font", "Calibri")
    v_color  = palette.get("primary", "1A3A6B")
    l_color  = palette.get("secondary", "2E6DA4")
    sl_color = "888888"

    _h2_text(slide, str(value),
             left, top, width, 0.72,
             font, 34, v_color, bold=True, align="center")
    _h2_rect(slide, left + width * 0.2, top + 0.72, width * 0.6, 0.03,
             palette.get("accent", "F0A500"))
    _h2_text(slide, str(label),
             left, top + 0.78, width, 0.40,
             font, 12, l_color, bold=False, align="center")
    _h2_text(slide, str(sublabel),
             left, top + 1.18, width, 0.50,
             font, 10, sl_color, bold=False, align="center")


def _h2_divider(slide,
                left: float, top: float, width: float,
                color: str, thickness: float = 0.04):
    """Ligne horizontale fine (en pouces)."""
    _h2_rect(slide, left, top, width, thickness, color)


def _h2_number(slide,
               text: str, left: float, top: float,
               size_in: float, color: str, font: str):
    """
    Grand texte décoratif (numéro de section, chiffre oversize).
    size_in = hauteur approximative de la boîte en pouces.
    Font size calculée automatiquement.
    """
    font_pt = max(24, int(size_in * 68))
    _h2_text(slide, str(text),
             left, top, size_in * 1.8, size_in,
             font, font_pt, color, bold=True)


# ─────────────────────────────────────────────────────────────
# EXTRACTION DE PALETTE NIVEAU 2
# ─────────────────────────────────────────────────────────────

def _h2_extract_palette(brand: dict) -> dict:
    """
    Construit une palette de 6 rôles (primary, secondary, accent, light, text, font)
    depuis la charte extraite du template.
    Fallback sur une palette bleue corporate sobre.
    """
    colors = brand.get("colors", [])

    # Palette de fallback : bleu corporate
    palette = {
        "primary":   "1A3A6B",
        "secondary": "2E6DA4",
        "accent":    "F0A500",
        "light":     "EEF3FA",
        "text":      "1A1A2E",
        "font":      "Calibri",
    }

    # Police principale
    fonts = brand.get("fonts", [])
    if fonts:
        palette["font"] = fonts[0]

    # Couleurs extraites du template
    if len(colors) >= 1:
        palette["primary"] = colors[0]
    if len(colors) >= 2:
        palette["secondary"] = colors[1]
    if len(colors) >= 3:
        palette["accent"] = colors[2]

    # Calcul automatique de "light" (mélange primary + blanc 85%)
    try:
        p = palette["primary"].lstrip("#")
        r = int(p[0:2], 16)
        g = int(p[2:4], 16)
        b = int(p[4:6], 16)
        lr = int(r * 0.12 + 255 * 0.88)
        lg = int(g * 0.12 + 255 * 0.88)
        lb = int(b * 0.12 + 255 * 0.88)
        palette["light"] = f"{lr:02X}{lg:02X}{lb:02X}"
    except Exception:
        pass

    # Calcul de "text" (version très foncée du primary)
    try:
        p = palette["primary"].lstrip("#")
        r = int(p[0:2], 16)
        g = int(p[2:4], 16)
        b = int(p[4:6], 16)
        dr = max(0, int(r * 0.35))
        dg = max(0, int(g * 0.35))
        db = max(0, int(b * 0.35))
        palette["text"] = f"{dr:02X}{dg:02X}{db:02X}"
    except Exception:
        pass

    return palette


# ─────────────────────────────────────────────────────────────
# SÉCURITÉ — Validation du code avant exécution
# ─────────────────────────────────────────────────────────────

# Patterns interdits dans le code généré par Claude
_FORBIDDEN_CODE_PATTERNS = [
    "import ", "__import__", "eval(", "exec(",
    "open(", "os.", "sys.", "subprocess", "socket",
    "urllib", "requests", "http", "shutil",
    "__builtins__", "__globals__", "__locals__", "__class__",
    "getattr", "setattr", "delattr",
    "compile(", "globals(", "locals(",
    "vars(", "dir(",
]

def _validate_code_safety(code: str) -> tuple:
    """
    Vérifie que le code ne contient aucun pattern dangereux.
    Retourne (ok: bool, raison: str).
    """
    for pattern in _FORBIDDEN_CODE_PATTERNS:
        if pattern in code:
            return False, f"Pattern interdit: '{pattern}'"
    return True, ""


# ─────────────────────────────────────────────────────────────
# PROMPTS NIVEAU 2 — Génération de code python-pptx
# ─────────────────────────────────────────────────────────────

# Exemples de code par type de slide (injectés dans le system prompt)
_V2_SLIDE_EXAMPLES = """
[cover — fond coloré] :
slide, W, H = h2_blank_slide(prs)
h2_rect(slide, 0, 0, W, H, brand["primary"])
h2_rect(slide, 0, H*0.78, W, H*0.22, brand["secondary"])
h2_text(slide, "Titre Fort ≤ 7 mots", 0.7, H/2-1.1, W-1.4, 1.4, brand["font"], 42, "FFFFFF", bold=True)
h2_text(slide, "Sous-titre contextuel ≤ 12 mots", 0.7, H/2+0.4, W-1.4, 0.75, brand["font"], 20, "FFFFFF")
h2_text(slide, "Footer · Contexte · 2025", 0.7, H-0.48, W-1.4, 0.38, brand["font"], 11, "FFFFFF")

[cover — fond clair avec barre] :
slide, W, H = h2_blank_slide(prs)
h2_rect(slide, 0, 0, W, H, "FFFFFF")
h2_rect(slide, 0, 0, 0.5, H, brand["primary"])
h2_rect(slide, 0.5, H*0.42, W-0.5, 0.06, brand["accent"])
h2_text(slide, "Titre Fort ≤ 7 mots", 0.85, H/2-1.0, W-1.4, 1.3, brand["font"], 40, brand["primary"], bold=True)
h2_text(slide, "Sous-titre contextuel ≤ 12 mots", 0.85, H/2+0.4, W-1.4, 0.7, brand["font"], 18, brand["text"])
h2_text(slide, "Footer · Contexte · 2025", 0.85, H-0.48, W-1.4, 0.38, brand["font"], 11, brand["text"])

[section] :
slide, W, H = h2_blank_slide(prs)
h2_rect(slide, 0, 0, W, H, brand["primary"])
h2_number(slide, "01", 0.7, H/2-1.35, 1.9, brand["secondary"], brand["font"])
h2_text(slide, "Titre de section ≤ 6 mots", 0.7, H/2+0.52, W-1.2, 0.92, brand["font"], 32, "FFFFFF", bold=True)
h2_divider(slide, 0.7, H/2+1.52, 3.8, brand["accent"])

[kpi — fond sombre, grille 3×2] :
slide, W, H = h2_blank_slide(prs)
h2_rect(slide, 0, 0, W, H, brand["primary"])
h2_text(slide, "Chiffres clés 2024", 0.5, 0.22, W-1.0, 0.78, brand["font"], 30, "FFFFFF", bold=True)
h2_divider(slide, 0.5, 1.12, W-1.0, brand["accent"])
h2_kpi(slide, 0.4, 1.38, 2.8, "600 M€", "contribution", "versée en 2022", brand)
h2_kpi(slide, 3.6, 1.38, 2.8, "~5,6%", "du capital", "détenu par l'État", brand)
h2_kpi(slide, 6.8, 1.38, 2.8, "28 Md€", "CA mondial", "exercice 2023", brand)
h2_kpi(slide, 0.4, 3.22, 2.8, "1er rang", "en France", "par capitalisation boursière", brand)
h2_kpi(slide, 3.6, 3.22, 2.8, "~95k", "employés", "dont 25k en France", brand)
h2_kpi(slide, 6.8, 3.22, 2.8, "80+", "pays", "présence opérationnelle", brand)
h2_text(slide, "Footer · Contexte · 2025", 0.5, H-0.44, W-1.0, 0.36, brand["font"], 10, brand["secondary"])

[kpi — fond clair, ligne de 4] :
slide, W, H = h2_blank_slide(prs)
h2_rect(slide, 0, 0, W, H, "FFFFFF")
h2_rect(slide, 0, 0, 0.08, H, brand["primary"])
h2_text(slide, "Performance 2024", 0.5, 0.22, W-1.0, 0.78, brand["font"], 30, brand["primary"], bold=True)
h2_divider(slide, 0.5, 1.12, W-1.0, brand["accent"])
h2_rect(slide, 0.5, 1.4, W-1.0, 2.2, brand["light"])
for i, (v, l, s) in enumerate([("600 M€","contribution","versée en 2022"),("~5,6%","du capital","État actionnaire"),("28 Md€","CA mondial","2023"),("95k","employés","dans 80 pays")]):
    x = 0.7 + i * ((W-1.4)/3)
    h2_kpi_light(slide, x, 1.55, 2.0, v, l, s, brand)
h2_text(slide, "Footer · Contexte · 2025", 0.5, H-0.44, W-1.0, 0.36, brand["font"], 10, brand["primary"])

[timeline — fond blanc, axe horizontal] :
slide, W, H = h2_blank_slide(prs)
h2_rect(slide, 0, 0, W, H, "FFFFFF")
h2_rect(slide, 0, 0, 0.08, H, brand["primary"])
h2_text(slide, "Chronologie", 0.5, 0.2, W-1.0, 0.72, brand["font"], 28, brand["primary"], bold=True)
h2_divider(slide, 0.5, 1.06, W-1.0, brand["accent"])
h2_rect(slide, 0.5, 2.32, W-1.0, 0.07, brand["primary"])
steps = [("1924","Création","Fondation par décret"),("1985","Privatisation","Entrée en bourse"),("2003","Fusion","Nouveau groupe"),("2024","Centenaire","Repositionnement")]
for i, (date, titre, detail) in enumerate(steps):
    x = 0.5 + i * ((W-1.0)/3)
    h2_rect(slide, x-0.09, 2.17, 0.18, 0.38, brand["primary"])
    h2_text(slide, date, x-0.55, 1.52, 1.1, 0.56, brand["font"], 14, brand["primary"], bold=True, align="center")
    h2_text(slide, titre, x-0.65, 2.72, 1.3, 0.46, brand["font"], 12, brand["text"], bold=True, align="center")
    h2_text(slide, detail, x-0.75, 3.22, 1.5, 0.52, brand["font"], 10, "888888", align="center")
h2_text(slide, "Footer · Contexte · 2025", 0.5, H-0.44, W-1.0, 0.36, brand["font"], 10, brand["primary"])

[two_col — fond blanc] :
slide, W, H = h2_blank_slide(prs)
h2_rect(slide, 0, 0, W, H, "FFFFFF")
h2_rect(slide, 0, 0, 0.08, H, brand["primary"])
h2_text(slide, "Titre comparaison", 0.5, 0.2, W-1.0, 0.72, brand["font"], 28, brand["primary"], bold=True)
h2_divider(slide, 0.5, 1.06, W-1.0, brand["accent"])
col_w = W/2 - 0.8
h2_rect(slide, 0.5, 1.22, col_w, 0.52, brand["primary"])
h2_text(slide, "COLONNE A", 0.62, 1.27, col_w-0.24, 0.42, brand["font"], 13, "FFFFFF", bold=True)
for i, item in enumerate(["Premier point concis", "Deuxième point concis", "Troisième point", "Quatrième point"]):
    h2_text(slide, "•  "+item, 0.62, 1.9+i*0.62, col_w-0.24, 0.56, brand["font"], 12, brand["text"])
h2_rect(slide, W/2+0.3, 1.22, col_w, 0.52, brand["secondary"])
h2_text(slide, "COLONNE B", W/2+0.42, 1.27, col_w-0.24, 0.42, brand["font"], 13, "FFFFFF", bold=True)
for i, item in enumerate(["Aspect premier concis", "Aspect deux concis", "Aspect trois", "Aspect quatre"]):
    h2_text(slide, "•  "+item, W/2+0.42, 1.9+i*0.62, col_w-0.24, 0.56, brand["font"], 12, brand["text"])
h2_text(slide, "Footer · Contexte · 2025", 0.5, H-0.44, W-1.0, 0.36, brand["font"], 10, brand["primary"])

[quote — fond sombre, accent latéral] :
slide, W, H = h2_blank_slide(prs)
h2_rect(slide, 0, 0, W, H, brand["primary"])
h2_rect(slide, 0.55, H/2-1.08, 0.1, 2.12, brand["accent"])
h2_text(slide, "\u00ab\u202fCitation forte et mémorable de vingt mots maximum.\u202f\u00bb", 0.92, H/2-0.95, W-1.72, 1.72, brand["font"], 26, "FFFFFF", bold=True)
h2_text(slide, "\u2014 Auteur ou source, date", 0.92, H/2+0.88, W-1.72, 0.55, brand["font"], 14, brand["accent"])
h2_text(slide, "Footer · Contexte · 2025", 0.5, H-0.44, W-1.0, 0.36, brand["font"], 10, brand["secondary"])

[list — fond blanc, items numérotés] :
slide, W, H = h2_blank_slide(prs)
h2_rect(slide, 0, 0, W, H, "FFFFFF")
h2_rect(slide, 0, 0, 0.08, H, brand["primary"])
h2_text(slide, "Titre de la liste", 0.5, 0.2, W-1.0, 0.72, brand["font"], 28, brand["primary"], bold=True)
h2_divider(slide, 0.5, 1.06, W-1.0, brand["accent"])
items = [("Label court 1","Corps concis, une idée, ≤ 20 mots."),("Label court 2","Corps concis, une idée, ≤ 20 mots."),("Label court 3","Corps concis, une idée, ≤ 20 mots."),("Label court 4","Corps concis, une idée, ≤ 20 mots.")]
for i, (titre, corps) in enumerate(items):
    y = 1.28 + i * 0.86
    h2_rect(slide, 0.5, y, 0.52, 0.52, brand["primary"])
    h2_text(slide, str(i+1), 0.5, y+0.04, 0.52, 0.44, brand["font"], 18, "FFFFFF", bold=True, align="center")
    h2_text(slide, titre, 1.22, y, W-2.2, 0.38, brand["font"], 13, brand["primary"], bold=True)
    h2_text(slide, corps, 1.22, y+0.38, W-2.2, 0.44, brand["font"], 11, brand["text"])
h2_text(slide, "Footer · Contexte · 2025", 0.5, H-0.44, W-1.0, 0.36, brand["font"], 10, brand["primary"])

[image_text — fond split] :
slide, W, H = h2_blank_slide(prs)
h2_rect(slide, 0, 0, W, H, "FFFFFF")
h2_rect(slide, 0, 0, W/2, H, brand["primary"])
h2_text(slide, "Titre ≤ 8 mots", W/2+0.4, 0.35, W/2-0.78, 0.9, brand["font"], 26, brand["primary"], bold=True)
h2_divider(slide, W/2+0.4, 1.35, W/2-0.78, brand["accent"])
for i, point in enumerate(["Point visuel 1, concis et fort.", "Point visuel 2, idée distincte.", "Point visuel 3, conclusion."]):
    h2_text(slide, "\u2192  "+point, W/2+0.4, 1.6+i*0.84, W/2-0.78, 0.74, brand["font"], 13, brand["text"])
h2_text(slide, "Footer · Contexte · 2025", W/2+0.4, H-0.44, W/2-0.78, 0.36, brand["font"], 10, brand["primary"])

[full_text — fond blanc, barre latérale] :
slide, W, H = h2_blank_slide(prs)
h2_rect(slide, 0, 0, W, H, "FFFFFF")
h2_rect(slide, 0, 0, 0.08, H, brand["primary"])
h2_text(slide, "Titre développement ≤ 8 mots", 0.5, 0.2, W-1.0, 0.72, brand["font"], 28, brand["primary"], bold=True)
h2_divider(slide, 0.5, 1.06, W-1.0, brand["accent"])
h2_text(slide, "Premier paragraphe : une idée principale, 2-3 phrases courtes. Ton direct, factuel.", 0.5, 1.26, W-1.0, 0.92, brand["font"], 13, brand["text"])
h2_text(slide, "Deuxième paragraphe : idée distincte et complémentaire. Pas redondant.", 0.5, 2.26, W-1.0, 0.72, brand["font"], 13, brand["text"])
h2_text(slide, "Troisième paragraphe : conclusion ou conséquence pratique.", 0.5, 3.06, W-1.0, 0.72, brand["font"], 13, brand["text"])
h2_text(slide, "Footer · Contexte · 2025", 0.5, H-0.44, W-1.0, 0.36, brand["font"], 10, brand["primary"])

[closing — fond coloré] :
slide, W, H = h2_blank_slide(prs)
h2_rect(slide, 0, 0, W, H, brand["primary"])
h2_rect(slide, 0, H*0.78, W, H*0.22, brand["secondary"])
h2_text(slide, "Merci !", W/2-3.5, H/2-1.0, 7.0, 1.4, brand["font"], 56, "FFFFFF", bold=True, align="center")
h2_text(slide, "Sources · Contact · 2025", W/2-3.5, H/2+0.5, 7.0, 0.65, brand["font"], 15, "FFFFFF", align="center")
h2_text(slide, "Footer · Contexte · 2025", 0.7, H-0.44, W-1.4, 0.36, brand["font"], 11, "FFFFFF")
"""

V2_CODE_GEN_SYSTEM = (
    "Tu es Visual Cortex Level 2 — Générateur de slides python-pptx.\n\n"
    "MISSION : Pour chaque slide du plan narratif, génère du code Python qui crée\n"
    "une slide professionnelle, visuellement forte, respectueuse de la charte graphique.\n\n"
    "══════════════════════════════════════════\n"
    "RÈGLES ABSOLUES\n"
    "══════════════════════════════════════════\n"
    "1. Utilise UNIQUEMENT les fonctions h2_* listées ci-dessous. AUCUN IMPORT.\n"
    "2. Commence CHAQUE slide par : slide, W, H = h2_blank_slide(prs)\n"
    "3. Dimensions en POUCES (float). W ≈ 10.0, H ≈ 5.63 pour 16:9.\n"
    "4. Embed les textes DIRECTEMENT dans le code (strings littérales).\n"
    "5. Densité : respecte les density rules — titre ≤ 8 mots, body ≤ 40 mots, etc.\n"
    "6. Couleurs via brand[\"primary\"], brand[\"secondary\"], brand[\"accent\"],\n"
    "   brand[\"light\"], brand[\"text\"]. Blanc = \"FFFFFF\". Gris = \"888888\".\n"
    "7. Police : brand[\"font\"] pour TOUS les textes.\n"
    "8. Footer : texte identique sur toutes les slides de contenu.\n\n"
    "══════════════════════════════════════════\n"
    "FONCTIONS DISPONIBLES (et UNIQUEMENT celles-ci)\n"
    "══════════════════════════════════════════\n"
    "slide, W, H = h2_blank_slide(prs)\n"
    "→ Crée une slide vierge. Appeler EN PREMIER pour chaque slide.\n\n"
    "h2_rect(slide, left, top, width, height, color)\n"
    "→ Rectangle coloré. Dimensions en pouces. color = brand[\"primary\"] ou \"FFFFFF\".\n\n"
    "h2_text(slide, text, left, top, width, height, font, size_pt, color,\n"
    "        bold=False, italic=False, align=\"left\")\n"
    "→ Textbox. align : \"left\"|\"center\"|\"right\".\n\n"
    "h2_kpi(slide, left, top, width, value, label, sublabel, brand)\n"
    "→ Bloc KPI (fond sombre). width ≈ 2.5–3.0 po.\n\n"
    "h2_kpi_light(slide, left, top, width, value, label, sublabel, brand)\n"
    "→ Bloc KPI (fond clair). width ≈ 2.0–2.5 po.\n\n"
    "h2_divider(slide, left, top, width, color, thickness=0.04)\n"
    "→ Ligne horizontale fine.\n\n"
    "h2_number(slide, text, left, top, size_in, color, font)\n"
    "→ Grand texte décoratif (numéro de section). size_in ≈ 1.5–2.0 po.\n\n"
    "══════════════════════════════════════════\n"
    "PROFIL CLIENT ET STYLE\n"
    "══════════════════════════════════════════\n"
    "{profile_block}\n\n"
    "══════════════════════════════════════════\n"
    "EXEMPLES PAR TYPE DE SLIDE\n"
    "(Adapte le contenu, garde la structure)\n"
    "══════════════════════════════════════════\n"
    + _V2_SLIDE_EXAMPLES +
    "\n══════════════════════════════════════════\n"
    "FORMAT DE SORTIE\n"
    "══════════════════════════════════════════\n"
    "JSON valide uniquement. Clés = plan_index en string (\"0\", \"1\", ...).\n"
    "{\n"
    "  \"0\": \"slide, W, H = h2_blank_slide(prs)\\nh2_rect(...)\\n...\",\n"
    "  \"1\": \"slide, W, H = h2_blank_slide(prs)\\n...\",\n"
    "  ...\n"
    "}\n"
    "Chaque valeur = code Python complet pour UNE slide, en une seule string.\n"
    "Séparateurs de lignes = \\n. AUCUN markdown. AUCUN commentaire hors JSON."
)

V2_CODE_GEN_USER = """PRÉSENTATION : {title}
ARC NARRATIF : {arc}
FOOTER UNIFIÉ : "{footer}"
SUJET : {prompt}

CHARTE (palette extraite du template) :
  brand["primary"]   = "#{primary}"
  brand["secondary"] = "#{secondary}"
  brand["accent"]    = "#{accent}"
  brand["light"]     = "#{light}"
  brand["text"]      = "#{text_color}"
  brand["font"]      = "{font}"
  Dimensions slide : W ≈ {W:.2f} po × H ≈ {H:.2f} po

PROFIL CLIENT : {profile_label}

SLIDES À GÉNÉRER ({n} slides) :
{slides_json}

Instructions :
1. Pour chaque slide, choisis le pattern visuel le plus adapté au slide_type ET au profil client.
2. Génère du code qui embed le contenu réel (dérivé de narrative_angle + key_message + visual_hint).
3. Adapte les textes aux density rules : titre ≤ 8 mots, items ≤ 18 mots, etc.
4. Utilise le footer_text sur toutes les slides sauf cover et closing (où tu peux l'intégrer différemment).
5. Varie les patterns (fond sombre / fond clair) pour créer un rythme visuel.

Génère le JSON avec le code python-pptx complet pour chaque slide."""


# ─────────────────────────────────────────────────────────────
# GÉNÉRATION DE CODE — Appel Claude Niveau 2
# ─────────────────────────────────────────────────────────────

def generate_codes_v2(
    prompt: str,
    plan: dict,
    palette: dict,
    brand: dict,
    profile: str,
) -> dict:
    """
    Phase 3 du pipeline Niveau 2.
    Claude génère un dict { "plan_index": "code_python_string" } pour toutes les slides.
    """
    client  = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    profile_data = CLIENT_PROFILES.get(profile, CLIENT_PROFILES["institutional"])

    profile_block = (
        f"Profil : {profile_data['label']}\n"
        f"Description : {profile_data['description']}\n"
        f"Guide de style : {profile_data['style_guide']}\n"
        f"Layouts préférés : {', '.join(profile_data['layout_prefs'])}\n"
        f"Fonds sombres : {'oui (couverts, KPIs, quotes)' if profile_data['bg_dark'] else 'non (préférer fonds clairs)'}"
    )

    slides_payload = []
    for sp in plan.get("slides", []):
        slides_payload.append({
            "plan_index":      sp.get("plan_index", 0),
            "slide_type":      sp.get("slide_type", "unknown"),
            "narrative_angle": sp.get("narrative_angle", ""),
            "key_message":     sp.get("key_message", ""),
            "visual_hint":     sp.get("visual_hint", ""),
        })

    system = V2_CODE_GEN_SYSTEM.replace("{profile_block}", profile_block)

    user = V2_CODE_GEN_USER.format(
        title       = plan.get("presentation_title", prompt[:60]),
        arc         = plan.get("narrative_arc", ""),
        footer      = plan.get("footer_text", ""),
        prompt      = prompt,
        primary     = palette.get("primary", "1A3A6B"),
        secondary   = palette.get("secondary", "2E6DA4"),
        accent      = palette.get("accent", "F0A500"),
        light       = palette.get("light", "EEF3FA"),
        text_color  = palette.get("text", "1A1A2E"),
        font        = palette.get("font", "Calibri"),
        W           = brand.get("slide_width_in", 10.0),
        H           = brand.get("slide_height_in", 5.63),
        n           = len(slides_payload),
        profile_label = profile_data["label"],
        slides_json = json.dumps(slides_payload, ensure_ascii=False, indent=2),
    )

    # max_tokens adaptatif : ~450 tokens/slide (code python-pptx plus verbeux)
    code_tokens = max(3000, len(slides_payload) * 450)

    msg = client.messages.create(
        model=CLAUDE_MODEL, max_tokens=code_tokens,
        system=system,
        messages=[{"role": "user", "content": user}],
    )

    raw      = _clean_json(msg.content[0].text.strip())
    code_map = json.loads(raw)
    log.info(f"[V2] Code généré pour {len(code_map)} slides.")
    return code_map


# ─────────────────────────────────────────────────────────────
# EXÉCUTION SÉCURISÉE — Sandbox avec namespace restreint
# ─────────────────────────────────────────────────────────────

def _build_safe_namespace(prs: Presentation, palette: dict) -> dict:
    """
    Construit le namespace restreint exposé au code généré par Claude.
    Seules les fonctions h2_* et les builtins sécurisés sont accessibles.
    """
    return {
        # Contexte obligatoire
        "prs":   prs,
        "brand": palette,
        # Helpers Level 2 (closures sur prs/palette selon besoin)
        "h2_blank_slide":  lambda: _h2_blank_slide(prs),
        "h2_rect":         _h2_rect,
        "h2_text":         _h2_text,
        "h2_kpi":          _h2_kpi,
        "h2_kpi_light":    _h2_kpi_light,
        "h2_divider":      _h2_divider,
        "h2_number":       _h2_number,
        # Builtins sécurisés seulement
        "__builtins__": {
            "range":     range,
            "len":       len,
            "int":       int,
            "float":     float,
            "str":       str,
            "bool":      bool,
            "list":      list,
            "dict":      dict,
            "tuple":     tuple,
            "enumerate": enumerate,
            "zip":       zip,
            "round":     round,
            "abs":       abs,
            "min":       min,
            "max":       max,
            "sum":       sum,
            "sorted":    sorted,
            "reversed":  reversed,
            "any":       any,
            "all":       all,
            "print":     lambda *args, **kwargs: None,  # silenced
            "True":      True,
            "False":     False,
            "None":      None,
        },
    }


def _execute_slide_code_v2(code: str, prs: Presentation, palette: dict) -> bool:
    """
    Exécute le code python-pptx généré par Claude dans un sandbox sécurisé.
    - Validation du code (patterns interdits)
    - Timeout 30 secondes via threading
    - En cas d'erreur : log + return False (la slide ne sera pas ajoutée)
    Retourne True si succès, False si erreur.
    """
    # Validation sécurité
    ok, reason = _validate_code_safety(code)
    if not ok:
        log.warning(f"[V2] Code rejeté (sécurité) : {reason}")
        return False

    result = {"success": False, "error": None}
    safe_ns = _build_safe_namespace(prs, palette)

    def _run():
        try:
            exec(code, safe_ns)  # noqa: S102
            result["success"] = True
        except Exception as e:
            result["error"] = str(e)
            log.warning(f"[V2] exec() error: {e}")

    t = threading.Thread(target=_run, daemon=True)
    t.start()
    t.join(timeout=30)

    if t.is_alive():
        log.warning("[V2] exec() timeout (30s)")
        return False

    if result["error"]:
        return False

    return True


def _execute_all_codes_v2(
    code_map: dict,
    plan_slides: list,
    prs: Presentation,
    palette: dict,
) -> int:
    """
    Exécute tous les codes dans l'ordre du plan narratif.
    Retourne le nombre de slides générées avec succès.
    """
    success = 0
    for sp in plan_slides:
        plan_idx = str(sp.get("plan_index", 0))
        code     = code_map.get(plan_idx, "")

        if not code:
            log.warning(f"[V2] Pas de code pour slide plan_index={plan_idx}")
            # Slide de fallback minimale
            _inject_fallback_slide(prs, sp, palette)
            continue

        ok = _execute_slide_code_v2(code, prs, palette)
        if ok:
            success += 1
        else:
            log.warning(f"[V2] Fallback pour slide plan_index={plan_idx}")
            _inject_fallback_slide(prs, sp, palette)

    return success


def _inject_fallback_slide(prs: Presentation, slide_plan: dict, palette: dict):
    """
    Slide de fallback minimaliste si le code Level 2 échoue.
    Crée une slide blanche avec le key_message centré.
    """
    try:
        slide, W, H = _h2_blank_slide(prs)
        _h2_rect(slide, 0, 0, W, H, "FFFFFF")
        _h2_rect(slide, 0, 0, 0.08, H, palette.get("primary", "1A3A6B"))
        key_msg = slide_plan.get("key_message", "Contenu")
        _h2_text(
            slide, key_msg,
            0.5, H / 2 - 0.4, W - 1.0, 0.8,
            palette.get("font", "Calibri"), 24,
            palette.get("primary", "1A3A6B"),
            bold=True,
        )
    except Exception as e:
        log.error(f"[V2] Fallback slide failed: {e}")


def _remove_original_slides_v2(prs: Presentation, n_original: int):
    """
    Supprime les n_original premières slides (template) pour ne garder
    que les slides générées par Level 2.
    """
    xml_slides  = prs.slides._sldIdLst
    all_sld_ids = list(xml_slides)

    # Garder seulement les slides ajoutées après les originales
    new_sld_ids = all_sld_ids[n_original:]

    if not new_sld_ids:
        log.warning("[V2] Aucune slide Level 2 générée, conservation des slides originales.")
        return

    # Nettoyer les relations des slides supprimées
    kept_rids = set()
    for sld_el in new_sld_ids:
        rid = sld_el.get(qn("r:id"))
        if rid:
            kept_rids.add(rid)

    for sld in list(xml_slides):
        xml_slides.remove(sld)
    for sld in new_sld_ids:
        xml_slides.append(sld)

    # Note : nettoyage des rels orphelins supprime (instable avec python-pptx).
    # PowerPoint ignore les rels non references dans sldIdLst.


# ─────────────────────────────────────────────────────────────
# PIPELINE NIVEAU 2
# ─────────────────────────────────────────────────────────────

def run_pipeline_v2(
    pptx_bytes: bytes,
    prompt:     str,
    nb_slides:  int,
    profile:    str = "institutional",
) -> tuple:
    """
    Pipeline Level 2 — 4 phases :
    Phase 1 : Analyse brand + extraction palette (même que L1)
    Phase 2 : Planification narrative (même que L1)
    Phase 3 : Génération code python-pptx (Claude → JSON de code strings)
    Phase 4 : Exécution sécurisée + suppression slides originales + export
    """
    if not ANTHROPIC_API_KEY:
        raise ValueError("Clé API Claude manquante.")

    nb_slides  = max(2, min(nb_slides, 30))
    profile    = profile if profile in CLIENT_PROFILES else "institutional"

    prs        = Presentation(io.BytesIO(pptx_bytes))
    n_original = len(prs.slides)

    # ── Phase 1 ──────────────────────────────────────────────
    log.info(f"[V2] Phase 1 : analyse brand (profil={profile})...")
    brand     = extract_brand(prs)
    palette   = _h2_extract_palette(brand)
    library   = build_layout_library(prs)
    selection = select_template_slides(library, nb_slides)
    log.info(f"[V2] Palette : primary=#{palette['primary']} font={palette['font']}")

    # ── Phase 2 ──────────────────────────────────────────────
    log.info("[V2] Phase 2 : planification narrative...")
    plan = plan_presentation(prompt, nb_slides, selection, brand)

    plan_slides = plan.get("slides", [])
    while len(plan_slides) < nb_slides:
        fallback = selection[min(len(plan_slides), len(selection) - 1)]
        plan_slides.append({
            "plan_index":           len(plan_slides),
            "template_slide_index": fallback["slide_index"],
            "slide_type":           fallback.get("slide_type", "full_text"),
            "narrative_angle":      "Développement complémentaire",
            "key_message":          "Argument additionnel",
            "visual_hint":          "",
        })
    plan["slides"] = plan_slides[:nb_slides]

    # ── Phase 3 ──────────────────────────────────────────────
    log.info("[V2] Phase 3 : génération du code python-pptx...")
    code_map = generate_codes_v2(prompt, plan, palette, brand, profile)

    # ── Phase 4 ──────────────────────────────────────────────
    log.info("[V2] Phase 4 : exécution des codes et assemblage...")
    success_count = _execute_all_codes_v2(code_map, plan["slides"], prs, palette)
    log.info(f"[V2] {success_count}/{nb_slides} slides générées avec succès.")

    # Fallback L1 automatique si taux de succès < 50%
    success_rate = success_count / max(nb_slides, 1)
    if success_rate < 0.5:
        log.warning(f"[V2] Taux d'échec élevé ({success_rate:.0%}) → fallback automatique Level 1")
        final_bytes_l1, plan_l1, brand_l1 = run_pipeline(pptx_bytes, prompt, nb_slides)
        return final_bytes_l1, plan_l1, brand_l1, palette

    # Supprimer les slides originales du template
    _remove_original_slides_v2(prs, n_original)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out.read(), plan, brand, palette


# ══════════════════════════════════════════════════════════════
# ROUTES NIVEAU 2
# ══════════════════════════════════════════════════════════════

@app.get("/profiles")
def get_profiles():
    """Retourne les profils clients disponibles pour le Niveau 2."""
    return {
        key: {
            "label":        p["label"],
            "description":  p["description"],
            "layout_prefs": p["layout_prefs"],
        }
        for key, p in CLIENT_PROFILES.items()
    }


@app.post("/generate-v2")
async def generate_presentation_v2(
    request:       Request,
    template:      UploadFile = File(...),
    prompt:        str        = Form(...),
    nb_slides:     int        = Form(default=8),
    profile:       str        = Form(default="institutional"),
    authorization: str        = Form(default=None),
):
    """
    Pipeline Niveau 2 — Génération créative (Brand Book → Slides nouvelles).
    Mêmes paramètres que /generate + `profile` (clé de CLIENT_PROFILES).
    """
    if not _is_pro(authorization):
        _quota(_ip(request))

    pptx_bytes = await template.read()

    try:
        final_bytes, plan, _, palette = run_pipeline_v2(
            pptx_bytes, prompt, nb_slides, profile
        )
        log.info("[V2] Pipeline Level 2 réussi.")
    except Exception as e:
        log.warning(f"[V2] Échec Level 2 ({e}) → fallback Level 1")
        final_bytes, plan, _ = run_pipeline(pptx_bytes, prompt, nb_slides)

    filename = f"visualcortex-v2-{_safe_name(prompt)}.pptx"
    return StreamingResponse(
        io.BytesIO(final_bytes),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={
            "Content-Disposition": f"attachment; filename={filename}",
            "Content-Length": str(len(final_bytes)),
        },
    )


@app.post("/generate-v2-preview")
async def generate_v2_preview(
    request:       Request,
    template:      UploadFile = File(...),
    prompt:        str        = Form(...),
    nb_slides:     int        = Form(default=8),
    profile:       str        = Form(default="institutional"),
    authorization: str        = Form(default=None),
):
    """
    Aperçu du plan narratif Level 2 (sans générer le fichier, sans consommer quota).
    Retourne le plan + la palette extraite.
    """
    pptx_bytes = await template.read()
    prs        = Presentation(io.BytesIO(pptx_bytes))
    brand      = extract_brand(prs)
    palette    = _h2_extract_palette(brand)
    library    = build_layout_library(prs)
    sel        = select_template_slides(library, nb_slides)

    pro = _is_pro(authorization)
    quota_info = (
        {"plan": "pro"} if pro
        else {"used": _quota(_ip(request))[0], "total": FREE_QUOTA_PER_IP, "plan": "free"}
    )

    plan = plan_presentation(prompt, nb_slides, sel, brand)

    return {
        "success":            True,
        "level":              2,
        "profile":            profile,
        "profile_label":      CLIENT_PROFILES.get(profile, {}).get("label", profile),
        "presentation_title": plan.get("presentation_title", prompt[:60]),
        "narrative_arc":      plan.get("narrative_arc", ""),
        "palette":            palette,
        "slides": [
            {
                "index":           s.get("plan_index"),
                "type":            s.get("slide_type"),
                "narrative_angle": s.get("narrative_angle"),
                "key_message":     s.get("key_message"),
            }
            for s in plan.get("slides", [])
        ],
        "quota": quota_info,
    }


# ══════════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=int(os.environ.get("PORT", 8000)),
        reload=False,
    )
