"""
Visual Cortex — PPTX Generator API v13 (Modèle Cortex — Edition Définitive)
═════════════════════════════════════════════════════════════════════════════
Nouveautés v13 :
  - Normalisation robuste des textes (apostrophes, tirets, espaces insécables,
    sauts de ligne, guillemets typographiques) — fix bug "texte ancien persiste"
  - Correspondance partielle sur textes fragmentés entre plusieurs runs XML
  - Zone emptying : Claude peut retourner "" pour vider une zone excédentaire
    → respiration visuelle sans modifier le design
  - _validate_and_trim préserve les "" intentionnels
  - Couverture totale obligatoire : chaque zone doit être mappée

Philosophie Cortex :
  - Une slide = une idée. Pas un catalogue.
  - Chaque type de slide a ses propres contraintes de densité.
  - Le texte est une accroche, pas un rapport.
  - Le design du template prime — zéro corruption.

Modèle : claude-sonnet-4-6 (configurable via CLAUDE_MODEL)
"""

import os, io, json, time, copy, re, logging
from collections import defaultdict
from typing import Optional

import anthropic
from fastapi import FastAPI, File, Form, UploadFile, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, StreamingResponse
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
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

# Patterns de footers placeholder à remplacer automatiquement
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
    Les GROUP shapes contiennent des sous-shapes avec des text frames
    qui n'étaient jamais extraits ni remplacés.
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

# ─────────────────────────────────────────────────────────────
# RÈGLES DE DENSITÉ PAR TYPE DE SLIDE — Cœur du Modèle Cortex
# ─────────────────────────────────────────────────────────────
# Chaque type de slide a des contraintes précises sur :
# - Le nombre maximum d'items (body, list, kpi, timeline…)
# - La longueur maximum par zone (en mots)
# - Le style éditorial attendu
# - Ce qu'il ne faut jamais faire

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
        "max_title_words":    8,
        "max_col_title_words": 4,
        "max_col_body_words": 30,
        "max_items_per_col":  4,
        "style":              "Titre général + 2 colonnes symétriques. Chaque colonne : label court + 2-4 items courts.",
        "never":              "Jamais de paragraphes dans les colonnes. Jamais d'asymétrie (même nombre d'items par colonne).",
        "example":            "Colonne gauche : 'LIENS FINANCIERS' + 4 items de 10 mots max. Colonne droite : 'LIENS RÉGLEMENTAIRES' + 4 items.",
    },
    "kpi": {
        "max_kpi_count":      6,
        "max_value_words":    2,
        "max_label_words":    5,
        "max_sublabel_words": 12,
        "style":              "Chiffre ou métrique très visible (ex: '~5,6%') + label court (ex: 'du capital') + sous-label contextuel.",
        "never":              "Jamais plus de 6 KPIs. Jamais de phrases complètes pour les valeurs. Jamais de KPI sans unité.",
        "example":            "'600 M€' / 'contribution exceptionnelle' / 'versée en 2022, loi superprofits'",
    },
    "quote": {
        "max_quote_words":    20,
        "max_author_words":   6,
        "style":              "Citation courte et percutante, ton affirmatif. Idéalement entre guillemets. Auteur ou source en dessous.",
        "never":              "Jamais de bullet points. Jamais plus d'une citation par slide. Jamais de citation de plus de 20 mots.",
        "example":            "'La transition énergétique ne se fera pas sans les majors.' — Analyse stratégique 2025",
    },
    "timeline": {
        "max_steps":          6,
        "max_step_title_words": 4,
        "max_step_body_words": 12,
        "style":              "4 à 6 jalons chronologiques. Chaque étape : date/numéro + titre très court + optionnel: phrase courte.",
        "never":              "Jamais plus de 6 étapes. Jamais de paragraphes. Jamais sans repère temporel (date ou numéro).",
        "example":            "1924 / 'Création CFP' / 'Fondation par décret d'État'",
    },
    "list": {
        "max_items":          5,
        "max_item_title_words": 4,
        "max_item_body_words": 20,
        "style":              "3 à 5 items. Chaque item : titre en gras court + phrase de développement concise.",
        "never":              "Jamais plus de 5 items. Jamais sans titre par item. Jamais d'items sans structure parallèle.",
        "example":            "'Lobbying institutionnel' / '~2,3 M€ déclarés/an à la HATVP — contacts réguliers Élysée, Bercy.'",
    },
    "image_text": {
        "max_title_words":    8,
        "max_body_words":     40,
        "max_body_items":     3,
        "style":              "Titre + corps structuré en 2-3 points courts. L'image fait le travail visuel, le texte commente.",
        "never":              "Jamais de mur de texte côté texte. Jamais plus de 3 bullet points. Jamais de corps > 40 mots.",
    },
    "full_text": {
        "max_title_words":    8,
        "max_body_words":     60,
        "max_paragraphs":     3,
        "style":              "Titre + 2-3 paragraphes courts et aérés. Chaque paragraphe = une idée.",
        "never":              "Jamais de corps > 60 mots au total. Jamais plus de 3 paragraphes. Jamais sans titre fort.",
    },
    "closing": {
        "max_title_words":    5,
        "max_subtitle_words": 15,
        "style":              "Message mémorable ou 'Merci !' + sous-titre : sources, contact ou call-to-action.",
        "never":              "Jamais de bullet points. Jamais de corps long. Simple, élégant, mémorable.",
        "example_title":      "Merci !",
        "example_subtitle":   "Sources : Rapport Annuel 2023 · HATVP · Loi de vigilance 2017",
    },
    "complex": {
        "max_label_words":    4,
        "max_body_words":     15,
        "style":              "Labels ultra-courts pour les éléments du diagramme. Titres de section concis.",
        "never":              "Jamais de phrases complètes dans un diagramme. Jamais de body > 15 mots.",
    },
    "unknown": {
        "max_title_words":    8,
        "max_body_words":     35,
        "style":              "Titre + corps aéré. Respecter la structure du template.",
        "never":              "Jamais de surcharge.",
    },
}

# Limites de mots par rôle (fallback universel)
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
    shapes   = list(iter_all_shapes(slide.shapes))
    raw_shapes = list(slide.shapes)

    img_shapes   = [s for s in shapes if s.shape_type in (13, 11)]
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

    # Position absolue
    if idx == 0:
        return "cover"
    if idx == total - 1:
        return "closing"

    # Slides avec beaucoup de GROUP shapes = complexes (diagrammes, organigrammes)
    # Elles restent utilisables mais marquées "complex" pour la sélection
    if len(group_shapes) >= 5:
        return "complex"

    # Image + texte
    if img_shapes and n >= 2:
        return "image_text"

    # Section : peu de texte
    if n <= 3 and total_chars < 100:
        return "section"

    # KPI : textes courts distribués horizontalement
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

    # Timeline : shapes étalés verticalement, textes courts
    if n >= 3 and total_chars < 500:
        try:
            tops = sorted([t["shape"].top for t in texts])
            if _emu(tops[-1]) - _emu(tops[0]) > 2.5:
                return "timeline"
        except Exception:
            pass

    # Deux colonnes : shapes gauche / droite
    if n >= 4:
        try:
            half = w / 2
            lc = sum(1 for t in texts if t["shape"].left < half)
            rc = sum(1 for t in texts if t["shape"].left >= half)
            if lc >= 2 and rc >= 2:
                return "two_col"
        except Exception:
            pass

    # Citation : peu de textes, un long
    if n <= 3 and any(len(t["text"]) > 60 for t in texts):
        return "quote"

    # Liste : plusieurs textes de longueur similaire
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
                            if rgb not in (RGBColor(0xFF,0xFF,0xFF), RGBColor(0,0,0)):
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
            "16:9" if abs(w/h - 16/9) < 0.05 else
            "4:3"  if abs(w/h - 4/3)  < 0.05 else "custom"
        ),
    }


def build_layout_library(prs: Presentation) -> list:
    """
    Construit la bibliothèque de layouts du template.
    Inclut maintenant les textes dans les GROUP shapes (traversée récursive).
    Détecte les footers placeholder pour les remplacer automatiquement.
    """
    library = []
    total = len(prs.slides)
    w, h  = prs.slide_width, prs.slide_height

    for idx, slide in enumerate(prs.slides):
        slide_type = _classify_slide(slide, idx, total, w, h)

        # Compter les GROUP shapes de niveau racine (indicateur de complexité)
        root_groups = sum(1 for s in slide.shapes if s.shape_type == MSO_SHAPE_TYPE.GROUP)
        has_images  = any(s.shape_type in (13,11) for s in iter_all_shapes(slide.shapes))

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

                # Détecter les footers placeholder
                is_placeholder_footer = _is_footer_placeholder(text)

                zones.append({
                    "original_text":          text,
                    "role":                   "footer" if is_placeholder_footer else role,
                    "word_count":             len(text.split()),
                    "word_limit":             WORD_LIMITS.get(
                        "footer" if is_placeholder_footer else role, 30
                    ),
                    "is_placeholder_footer":  is_placeholder_footer,
                    "char_count":             len(text),
                })

        if zones:
            library.append({
                "slide_index":   idx,
                "slide_type":    slide_type,
                "description":   SLIDE_TYPE_DESC.get(slide_type, ""),
                "position":      (
                    "cover"   if idx == 0 else
                    "closing" if idx == total - 1 else
                    f"{idx+1}/{total}"
                ),
                "root_groups":   root_groups,
                "has_images":    has_images,
                "zones":         zones,
                "total_words":   sum(z["word_count"] for z in zones),
                "visual_score":  (
                    has_images * 3 +
                    (1 if slide_type in ("kpi","timeline","two_col","image_text","quote") else 0) * 2 +
                    (1 if slide_type in ("list","full_text") else 0) * 0 +
                    (1 if root_groups > 0 else 0) * 1
                ),
            })

    return library


def select_template_slides(library: list, nb_slides: int) -> list:
    """
    Sélectionne intelligemment nb_slides slides du template.

    Règles :
    1. Slide 0 (cover) toujours en premier.
    2. Dernière slide (closing) toujours en dernier.
    3. Pour les slides du milieu : priorité aux layouts visuels,
       éviter de répéter le même type consécutivement.
    4. Exclure les slides "complex" en priorité si possible.
    5. Si nb_slides > template : signaler qu'il faudra dupliquer.
    """
    if not library:
        return []

    cover   = [s for s in library if s["slide_type"] == "cover"]
    closing = [s for s in library if s["slide_type"] == "closing"]
    middle  = [s for s in library if s["slide_type"] not in ("cover", "closing")]

    # Slides du milieu triées par visual_score desc, complex en dernier
    middle_sorted = sorted(
        middle,
        key=lambda s: (
            0 if s["slide_type"] == "complex" else 1,  # complex = dernier recours
            s["visual_score"],
            -s["total_words"],  # préférer moins de texte
        ),
        reverse=True,
    )

    # Slots disponibles pour le milieu
    n_cover   = min(len(cover), 1)
    n_closing = min(len(closing), 1)
    n_middle  = nb_slides - n_cover - n_closing

    if n_middle < 0:
        n_middle = 0

    # Sélection avec anti-répétition de type
    selected_middle = []
    last_type = None
    pool = middle_sorted.copy()

    while len(selected_middle) < n_middle and pool:
        # Préférer un type différent du précédent
        for i, s in enumerate(pool):
            if s["slide_type"] != last_type or i == len(pool) - 1:
                selected_middle.append(s)
                last_type = s["slide_type"]
                pool.pop(i)
                break

    # Compléter si pas assez de slides uniques → dupliquer les meilleures
    if len(selected_middle) < n_middle and middle_sorted:
        cycle = middle_sorted * 10
        for s in cycle:
            if len(selected_middle) >= n_middle:
                break
            # Créer une copie avec flag "duplicate"
            dup = {**s, "duplicate": True}
            selected_middle.append(dup)

    # Assembler : cover + milieu (trié par slide_index) + closing
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

RÈGLE DES SLIDES DE SECTION — Proportion et sobriété :
Une slide de section est une pause dans la narration. Comme une respiration.
Trop de pauses dans un texte court, et on perd le fil. Pas assez dans un long, et on se noie.

Règle stricte selon le nombre de slides :
  ≤ 6 slides  → 0 slide de section. La cover suffit. Chaque slide compte trop.
  7-10 slides → 1 slide de section maximum, uniquement si le sujet a 2 parties très distinctes.
  11-15 slides → 2 slides de section maximum.
  16-20 slides → 3 slides de section maximum.
  > 20 slides → 1 section pour 6-7 slides de contenu, jamais plus.

Règle de qualité : une slide de section ne se justifie que si ce qui suit
est suffisamment différent de ce qui précède pour mériter une introduction.
Si deux parties se ressemblent, supprimer la section et fusionner.



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
  "footer_text": "Entreprise · Contexte · Année  (≤ 8 mots, ex: 'TotalEnergies · Analyse ONG · 2025')",
  "slides": [
    {{
      "plan_index": 0,
      "template_slide_index": 0,
      "slide_type": "cover",
      "narrative_angle": "Ce que cette slide accomplit dans l'histoire (1 phrase)",
      "key_message": "Le message principal ≤ 10 mots",
      "visual_hint": "Contrainte de densité et de style pour cette slide (ex: '4 KPIs: valeur chiffrée + label court + sous-label 10 mots max')"
    }}
  ]
}}\n"""


async def plan_presentation(prompt: str, nb_slides: int, selection: list, brand: dict) -> dict:
    client = anthropic.AsyncAnthropic(api_key=ANTHROPIC_API_KEY)

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

    msg = await client.messages.create(
        model=CLAUDE_MODEL, max_tokens=4000,
        system=system,
        messages=[{"role": "user", "content": user}],
    )
    plan = json.loads(_clean_json(msg.content[0].text.strip()))
    log.info(f"Plan: {len(plan.get('slides',[]))} slides — {plan.get('narrative_arc','')[:80]}")
    return plan


# ══════════════════════════════════════════════════════════════
# PHASE 3 — GÉNÉRATION DU CONTENU (visuel-first)
# ══════════════════════════════════════════════════════════════

CORTEX_SYSTEM = """Tu es Visual Cortex, expert en présentations B2B professionnelles et visuelles.
Philosophie : une slide = une idée. Le texte est une accroche, pas un rapport.

═══════════════════════════════════════════════════
RÈGLES UNIVERSELLES (s'appliquent à toutes les slides)
═══════════════════════════════════════════════════
1. LIMITES PAR RÔLE — ne jamais dépasser :
   title       → ≤ 8 mots   (percutant, mémorable, verbe ou tension)
   subtitle    → ≤ 12 mots  (précis, complémentaire, pas redondant)
   label       → ≤ 5 mots   (factuel, souvent sans verbe)
   kpi_value   → ≤ 3 mots   (chiffre + unité : "600 M€", "~5,6%", "1er")
   kpi_label   → ≤ 5 mots   (contexte court : "du capital", "de contribution")
   body        → ≤ 40 mots  (jamais de mur de texte)
   list_item   → ≤ 18 mots  (titre en gras + corps concis)
   quote       → ≤ 20 mots  (fort, affirmatif, entre guillemets)
   footer      → ≤ 8 mots   (entreprise · contexte · date)
   page_number → NE PAS MODIFIER
   section_num → 1-2 chars  (01, 02, A, B…)

2. COHÉRENCE : footer identique (ou quasi) sur toutes les slides de contenu.
3. B2B : vocabulaire du secteur, ton direct, orienté valeur, zéro formule creuse.
4. PROGRESSION : chaque slide fait avancer l'histoire selon son narrative_angle.
5. ZÉRO invention de données, chiffres ou noms non fournis dans le prompt.
6. ZONES VIDES — levier de respiration visuelle :
   Si une slide a plus de zones que ce que la density rule autorise,
   tu DOIS vider les zones excédentaires en retournant "" comme valeur.
   Exemples :
   - Slide list avec 5 items dans le template mais density rule dit max 3 → vider les 2 derniers items
   - Slide kpi avec 6 blocs mais le sujet n'a que 4 KPIs pertinents → vider les 2 derniers blocs
   - Slide two_col avec 4 items/col mais le contenu n'en justifie que 3 → vider le 4e item de chaque col
   RÈGLE : mieux vaut une zone vide et une slide aérée qu'une zone remplie avec du contenu artificiel.

7. COUVERTURE TOTALE : chaque zone listée dans "zones" DOIT avoir une clé dans ta réponse JSON.
   Si tu ne veux pas remplacer une zone (page_number, logo), retourne la valeur originale.
   Si tu veux vider une zone, retourne "".
   Ne jamais omettre une zone — sinon l'ancien texte du template persistera.

═══════════════════════════════════════════════════
RÈGLES PAR TYPE DE SLIDE (density rules)
═══════════════════════════════════════════════════

[cover]
  Titre ≤ 7 mots — accroche forte, verbe d'action ou tension.
  Sous-titre ≤ 12 mots — angle ou contexte.
  JAMAIS de bullet points. JAMAIS plus de 2 zones texte.

[section]
  Numéro (01, 02…) + titre ≤ 6 mots. RIEN D'AUTRE.
  JAMAIS de body text sur une slide de section.

[kpi]
  4 à 6 KPIs MAXIMUM.
  Chaque KPI : valeur courte (≤3 mots) + label (≤5 mots) + sous-label contextuel (≤12 mots).
  JAMAIS de phrases complètes pour les valeurs.
  JAMAIS de KPI sans unité ou repère (%, M€, rang, année…).

[timeline]
  4 à 6 jalons MAXIMUM.
  Chaque jalon : repère temporel (date, année, "Étape N") + titre ≤ 4 mots + phrase optionnelle ≤ 12 mots.
  JAMAIS de paragraphes. JAMAIS plus de 6 jalons. JAMAIS sans repère temporel.

[two_col]
  2 colonnes SYMÉTRIQUES — même nombre d'items dans chaque colonne (max 4 items/colonne).
  Label de colonne ≤ 4 mots + chaque item ≤ 18 mots.
  JAMAIS d'asymétrie. JAMAIS de body long dans une colonne.

[quote]
  1 SEULE citation ≤ 20 mots — fort, affirmatif, entre guillemets.
  Source/auteur optionnel ≤ 6 mots en dessous.
  JAMAIS de bullet points. JAMAIS plus d'une citation.

[list]
  3 à 5 items MAXIMUM.
  Chaque item : titre/label ≤ 5 mots (gras) + corps ≤ 20 mots.
  Structure parallèle : tous les items ont la même structure.
  JAMAIS plus de 5 items. JAMAIS d'items sans titre.

[image_text]
  Titre ≤ 8 mots + corps structuré en 2-3 points ≤ 40 mots total.
  L'image porte le message visuel — le texte commente et précise.
  JAMAIS de mur de texte côté texte.

[full_text]
  Titre ≤ 8 mots + 2-3 paragraphes courts = une idée par paragraphe.
  Total body ≤ 60 mots.
  JAMAIS plus de 3 paragraphes. JAMAIS sans structure claire.

[closing]
  Titre simple ≤ 5 mots ("Merci !", "Passons à l'action", "Construisons ensemble").
  Sous-titre ≤ 15 mots (sources, contact, CTA, ou mots de conclusion).
  JAMAIS de bullet points. Simple, élégant, mémorable.

[complex]
  Labels ultra-courts pour les nœuds du diagramme ≤ 4 mots.
  Titres de section ≤ 5 mots.
  JAMAIS de phrases complètes dans un diagramme.

Réponds UNIQUEMENT en JSON valide, sans commentaire ni markdown."""

CORTEX_USER = """PRÉSENTATION : {title}
ARC NARRATIF : {arc}
FOOTER : "{footer}"
SUJET COMPLET : {prompt}
CHARTE : Polices {fonts} | Couleurs {colors}

═══════════════════════════════
SLIDES À GÉNÉRER — {n} slides
═══════════════════════════════
{slides_json}

INSTRUCTIONS DE GÉNÉRATION :
Pour chaque slide :
1. Lis le "slide_type" → applique les density rules correspondantes (voir system prompt).
2. Lis le "narrative_angle" et "key_message" → génère du contenu qui sert cet angle.
3. Lis le "visual_hint" → contrainte de densité précise voulue par le planner.
4. Pour chaque zone dans "zones" :
   - role "page_number" → retourne le texte original tel quel
   - role "footer" ou "is_placeholder_footer: true" → retourne le FOOTER ci-dessus
   - zones excédentaires (density rules dépassées) → retourne ""
   - sinon → génère un texte respectant le "word_limit"
5. OBLIGATION : chaque "original_text" de chaque zone DOIT être une clé dans ta réponse.
   Aucune omission tolérée — un texte non mappé = ancien contenu du template qui reste.

RAPPEL ZONES VIDES : retourner "" pour vider une zone est un choix éditorial fort.
Utilise-le quand le contenu ne justifie pas de remplir toutes les zones disponibles.

FORMAT DE SORTIE (clés = template_slide_index en string) :
{{
  "0": {{"Texte original exact": "Nouveau texte ou vide"}},
  "1": {{"Texte original exact": ""}},
  ...
}}"""


async def generate_content(prompt: str, plan: dict, selection: list, brand: dict) -> dict:
    client = anthropic.AsyncAnthropic(api_key=ANTHROPIC_API_KEY)
    sel_by_idx  = {s["slide_index"]: s for s in selection}
    footer_text = plan.get("footer_text", "")

    slides_payload = []
    for sp in plan.get("slides", []):
        tidx       = sp.get("template_slide_index", 0)
        tmpl       = sel_by_idx.get(tidx, {})
        slide_type = sp.get("slide_type", "unknown")

        # Injecter les density rules du type dans chaque slide
        density = DENSITY_RULES.get(slide_type, DENSITY_RULES["unknown"])

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

    msg = await client.messages.create(
        model=CLAUDE_MODEL, max_tokens=8000,
        system=CORTEX_SYSTEM,
        messages=[{"role": "user", "content": user}],
    )

    raw     = _clean_json(msg.content[0].text.strip())
    mapping = json.loads(raw)

    # Post-validation : tronquer les textes qui dépassent les limites
    mapping = _validate_and_trim(mapping, slides_payload)

    log.info(f"Contenu généré et validé : {len(mapping)} slides.")
    return mapping


def _validate_and_trim(mapping: dict, slides_payload: list) -> dict:
    """
    Validation post-génération :
    - Préserve les zones vides intentionnelles ("" → vider la zone)
    - Tronque intelligemment les textes qui dépassent leur word_limit
    - Protège les page_numbers (remet le texte original)
    """
    zone_limits: dict = {}
    for sp in slides_payload:
        tidx = str(sp["template_slide_index"])
        for z in sp.get("zones", []):
            zone_limits[(tidx, z["original_text"])] = {
                "word_limit": z["word_limit"],
                "role":       z["role"],
            }

    validated = {}
    for slide_key, replacements in mapping.items():
        validated[slide_key] = {}
        for orig, new_text in replacements.items():

            zone_info = zone_limits.get((str(slide_key), orig), {})
            role      = zone_info.get("role", "text")

            # page_number : remettre le texte original sans exception
            if role == "page_number":
                validated[slide_key][orig] = orig
                continue

            # Zone vide intentionnelle : préserver tel quel
            if new_text == "" or new_text is None:
                validated[slide_key][orig] = ""
                continue

            # Tronquer si dépassement
            limit = zone_info.get("word_limit", WORD_LIMITS.get(role, 40))
            words = new_text.split()
            if len(words) > limit:
                trimmed = " ".join(words[:limit])
                for punct in [".", ";", ":", "—", "–", ","]:
                    last = trimmed.rfind(punct)
                    if last > len(trimmed) * 0.6:
                        trimmed = trimmed[:last + 1].strip()
                        break
                log.debug(f"Trim {slide_key}/{role}: {len(words)}→{len(trimmed.split())} mots")
                validated[slide_key][orig] = trimmed
            else:
                validated[slide_key][orig] = new_text

    return validated


# ══════════════════════════════════════════════════════════════
# HYDRATATION — Injection avec traversée récursive
# ══════════════════════════════════════════════════════════════

_NORM_TABLE = str.maketrans({
    "\u2019": "'",  "\u2018": "'",  "\u2032": "'",
    "\u201c": '"',  "\u201d": '"',  "\u201e": '"',
    "\u2013": "-",  "\u2014": "-",  "\u2015": "-",
    "\u00a0": " ",  "\u200b": "",   "\u200c": "",
    "\u000b": " ",  "\u000c": " ",  "\r":     "",
})

def _normalize(text: str) -> str:
    """Normalise un texte pour comparaison robuste."""
    t = text.translate(_NORM_TABLE)
    t = re.sub(r"\s+", " ", t)
    return t.strip()


_SENTINEL_EMPTY = "__EMPTY__"   # Valeur sentinelle : Claude veut vider cette zone


def _replace_text_in_para(para, replacements: dict):
    """
    Remplace le texte d'un paragraphe en préservant le style XML.

    Logique de recherche :
    1. Correspondance exacte
    2. Correspondance normalisée (apostrophes, tirets, espaces…)

    Logique de remplacement :
    - new_text == _SENTINEL_EMPTY ou "" → vider le paragraphe (clear_text)
    - new_text == None (non dans le mapping) → NE PAS TOUCHER
    - sinon → remplacer en préservant le style XML
    """
    # Reconstituer le texte brut du paragraphe
    raw_parts = [r.text for r in para.runs]
    para_text = "".join(raw_parts).strip()
    if not para_text:
        return

    # 1. Correspondance exacte
    new_text = replacements.get(para_text)

    # 2. Correspondance normalisée
    if new_text is None:
        para_norm = _normalize(para_text)
        for k, v in replacements.items():
            if _normalize(k) == para_norm:
                new_text = v
                break

    # 3. Correspondance partielle sur texte long fragmenté
    #    (cas : le texte du template a été coupé par des runs de style différent)
    if new_text is None and len(para_text) > 10:
        for k, v in replacements.items():
            if len(k) > 10 and _normalize(k) in _normalize(para_text):
                new_text = v
                break

    # Clé non trouvée → ne pas toucher
    if new_text is None:
        return

    # Sauvegarder le style du premier run avant toute modification
    rpr_xml = None
    if para.runs:
        rpr_el = para.runs[0]._r.find(qn("a:rPr"))
        if rpr_el is not None:
            rpr_xml = copy.deepcopy(rpr_el)

    # Vider la zone (Claude a retourné "" ou _SENTINEL_EMPTY)
    if new_text in ("", _SENTINEL_EMPTY):
        # Supprimer tous les runs du paragraphe — la zone devient invisible
        for run in list(para.runs):
            run._r.getparent().remove(run._r)
        return

    # Remplacer le texte en préservant le style
    para.text = new_text
    if rpr_xml is not None:
        for run in para.runs:
            ex = run._r.find(qn("a:rPr"))
            if ex is not None:
                run._r.remove(ex)
            run._r.insert(0, copy.deepcopy(rpr_xml))


def _hydrate_slide(slide, replacements: dict):
    """
    Hydrate une slide en traversant TOUS les shapes (y compris GROUP shapes).
    """
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
    """
    Reconstruit le fichier PPTX :
    1. Sélectionne les slides du template correspondant au plan
    2. Hydrate chaque slide avec le contenu généré
    3. Retourne le fichier final avec exactement nb_slides slides
    """
    prs = Presentation(io.BytesIO(pptx_bytes))

    # Construire l'ordre des slides selon le plan
    template_indices = [s.get("template_slide_index", 0) for s in plan_slides]

    # Créer la nouvelle liste de slides dans le bon ordre
    # en utilisant les slides du template
    _reorder_and_hydrate(prs, template_indices, mapping, nb_slides)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out.read()


def _reorder_and_hydrate(prs: Presentation, template_indices: list, mapping: dict, nb_slides: int):
    """
    Réorganise et hydrate les slides du template.

    Approche robuste en 3 étapes :
    1. Hydrater toutes les slides référencées dans le mapping (en place)
    2. Reconstruire la liste sldIdLst dans le bon ordre (avec duplication si besoin)
    3. Supprimer les slides en excès par la méthode sûre (depuis la fin)

    On évite drop_rel() qui est instable selon les versions de python-pptx.
    """
    all_sld_ids = list(prs.slides._sldIdLst)
    total_tmpl  = len(all_sld_ids)

    # ── Étape 1 : hydrater chaque slide référencée ───────────────
    for slide_key, replacements in mapping.items():
        try:
            idx = int(str(slide_key).replace("slide_", ""))
            if idx < total_tmpl:
                _hydrate_slide(prs.slides[idx], replacements)
        except Exception as e:
            log.warning(f"Hydratation slide {slide_key}: {e}")

    # ── Étape 2 : construire l'ordre souhaité ────────────────────
    # Pour chaque position dans le plan, pointer vers la sldId correspondante
    # Si un même index est réutilisé, dupliquer physiquement la slide
    xml_slides  = prs.slides._sldIdLst
    seen_ids    = set()
    final_order = []

    for tidx in template_indices:
        if not (0 <= tidx < total_tmpl):
            tidx = max(0, min(tidx, total_tmpl - 1))

        sld_el  = all_sld_ids[tidx]
        sld_id  = sld_el.get("id")

        if sld_id not in seen_ids:
            # Première utilisation : prendre l'élément original
            seen_ids.add(sld_id)
            final_order.append(sld_el)
        else:
            # Réutilisation : dupliquer la slide via add_slide
            new_el = _safe_duplicate_slide(prs, tidx)
            if new_el is not None:
                final_order.append(new_el)
            else:
                # Fallback : réutiliser la même slide (sans duplication)
                final_order.append(sld_el)

    # Limiter au nb_slides souhaité
    final_order = final_order[:nb_slides]

    # ── Étape 3 : reconstruire sldIdLst proprement ───────────────
    # Vider et reconstruire avec les éléments sélectionnés
    for sld in list(xml_slides):
        xml_slides.remove(sld)
    for sld in final_order:
        xml_slides.append(sld)


def _safe_duplicate_slide(prs: Presentation, src_idx: int):
    """
    Duplique une slide de manière sûre.
    Utilise add_slide() + copie XML, sans toucher aux relations.
    Retourne le sldId element du nouveau slide, ou None si échec.
    """
    try:
        src_slide    = prs.slides[src_idx]
        blank_layout = prs.slide_layouts[-1]
        new_slide    = prs.slides.add_slide(blank_layout)

        # Remplacer le spTree du nouveau slide par celui de la source
        import copy as _copy
        src_sp_tree = src_slide.shapes._spTree
        new_sp_tree = new_slide.shapes._spTree

        # Vider le spTree du nouveau slide
        for child in list(new_sp_tree):
            new_sp_tree.remove(child)

        # Copier tous les éléments de la source
        for child in src_sp_tree:
            new_sp_tree.append(_copy.deepcopy(child))

        # Retourner le dernier sldId (celui qu'on vient d'ajouter)
        return list(prs.slides._sldIdLst)[-1]

    except Exception as e:
        log.warning(f"Duplication slide {src_idx} échouée : {e}")
        return None


# ══════════════════════════════════════════════════════════════
# PIPELINE COMPLET
# ══════════════════════════════════════════════════════════════

async def run_pipeline(pptx_bytes: bytes, prompt: str, nb_slides: int) -> tuple:
    """
    3 phases async :
    Phase 1 → Compréhension (CPU, instantané)
    Phase 2 → Planification narrative (appel Claude async)
    Phase 3 → Génération + hydratation (appel Claude async + CPU)
    """
    if not ANTHROPIC_API_KEY:
        raise ValueError("Clé API Claude manquante.")

    prs = Presentation(io.BytesIO(pptx_bytes))
    nb_slides = max(2, min(nb_slides, 30))

    # Phase 1 — CPU uniquement, instantané
    log.info("Phase 1 : analyse du template...")
    brand     = extract_brand(prs)
    library   = build_layout_library(prs)
    selection = select_template_slides(library, nb_slides)
    log.info(
        f"Template : {len(library)} slides analysées → "
        f"{len(selection)} sélectionnées pour {nb_slides} slides demandées"
    )

    # Phase 2 — appel Claude async (non bloquant)
    log.info("Phase 2 : planification narrative...")
    plan = await plan_presentation(prompt, nb_slides, selection, brand)

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

    # Phase 3 — appel Claude async (non bloquant)
    log.info("Phase 3 : génération du contenu...")
    mapping = await generate_content(prompt, plan, selection, brand)

    log.info("Hydratation PPTX...")
    final_bytes = hydrate_presentation(pptx_bytes, mapping, plan["slides"], nb_slides)

    return final_bytes, plan, brand


# ══════════════════════════════════════════════════════════════
# ROUTES API
# ══════════════════════════════════════════════════════════════

@app.get("/")
def root():
    return {"status": "ok", "version": "13.0.0 - Modèle Cortex Edition Définitive", "model": CLAUDE_MODEL}


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

    nb_slides = max(2, min(nb_slides, 30))
    pptx_bytes = await template.read()
    prs   = Presentation(io.BytesIO(pptx_bytes))
    brand = extract_brand(prs)
    lib   = build_layout_library(prs)
    sel   = select_template_slides(lib, nb_slides)
    plan  = await plan_presentation(prompt, nb_slides, sel, brand)

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
    final_bytes, plan, _ = await run_pipeline(pptx_bytes, prompt, nb_slides)

    filename = f"visualcortex-{_safe_name(prompt)}.pptx"
    return StreamingResponse(
        io.BytesIO(final_bytes),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=int(os.environ.get("PORT", 8000)), reload=False)
