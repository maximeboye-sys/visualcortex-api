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

import os, io, json, time, copy, re, logging, threading, zipfile, base64
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
from layouts import LAYOUT_REGISTRY, LAYOUT_DESCRIPTIONS

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
    """
    Extrait un bloc JSON propre depuis la réponse de Claude.
    Gère : markdown fences, préambules textuels, JSON tronqué.
    """
    s = raw.strip()

    # Cas 1 : bloc ```json ... ``` ou ``` ... ```
    if "```" in s:
        parts = s.split("```")
        for part in parts[1::2]:          # parties entre backticks
            candidate = part.strip()
            if candidate.startswith("json"):
                candidate = candidate[4:].strip()
            if candidate.startswith("{") or candidate.startswith("["):
                s = candidate
                break

    # Cas 2 : extraire depuis le premier { ou [ jusqu'au dernier } ou ]
    start_brace  = s.find("{")
    start_bracket = s.find("[")
    if start_brace == -1 and start_bracket == -1:
        return s  # on laisse planter json.loads avec un message clair

    if start_bracket == -1 or (start_brace != -1 and start_brace < start_bracket):
        start = start_brace
        end   = s.rfind("}")
    else:
        start = start_bracket
        end   = s.rfind("]")

    if start != -1 and end != -1 and end > start:
        s = s[start:end + 1]

    return s.strip()


def _parse_json_robust(raw: str, context: str = "") -> dict:
    """
    Parse JSON avec fallback sur réparation basique.
    Si le JSON est tronqué (max_tokens atteint), tente de le compléter.
    """
    cleaned = _clean_json(raw)

    # Tentative 1 : parse direct
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError as e:
        log.warning(f"JSON parse error [{context}]: {e} — tentative de réparation")

    # Tentative 2 : supprimer trailing comma avant } ou ]
    fixed = re.sub(r",\s*([}\]])", r"\1", cleaned)
    try:
        return json.loads(fixed)
    except json.JSONDecodeError:
        pass

    # Tentative 3 : JSON tronqué — essayer de fermer les accolades manquantes
    attempt = fixed
    open_braces   = attempt.count("{") - attempt.count("}")
    open_brackets = attempt.count("[") - attempt.count("]")
    # Supprimer la dernière entrée potentiellement incomplète
    last_comma = attempt.rfind(",")
    last_brace  = attempt.rfind("}")
    if last_comma > last_brace:
        attempt = attempt[:last_comma]
    attempt += "]" * open_brackets + "}" * open_braces
    try:
        result = json.loads(attempt)
        log.warning(f"JSON réparé (tronqué) [{context}] : {open_braces} accolades + {open_brackets} crochets ajoutés")
        return result
    except json.JSONDecodeError as e:
        raise ValueError(f"JSON irrécupérable [{context}] : {e}\nDébut : {cleaned[:200]}") from e

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
    """
    Extraction complète de la charte depuis le template.
    Source prioritaire : ppt/theme/theme1.xml lu depuis le ZIP PPTX.
    Fallback : scan regex de tout le XML pour les couleurs et polices.
    """
    import re as _re
    import io
    import zipfile as _zipfile
    import lxml.etree as etree
    from collections import Counter

    color_counter = Counter()
    fonts_seen: list = []
    fonts_set: set = set()

    def _add_font(f: str):
        if f and not f.startswith('+') and f.strip() and f not in fonts_set:
            fonts_set.add(f)
            fonts_seen.append(f)

    def _is_neutral(h: str) -> bool:
        try:
            r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
            if r > 235 and g > 235 and b > 235: return True
            if r < 25  and g < 25  and b < 25:  return True
            if abs(r-g) < 25 and abs(g-b) < 25: return True
            return False
        except Exception:
            return True

    # ── Étape 1 : lire ppt/theme/theme1.xml depuis le ZIP ────────────────────
    # Un PPTX est un ZIP. Le thème officiel s'y trouve toujours.
    theme_colors: dict = {}
    try:
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        with _zipfile.ZipFile(buf) as zf:
            theme_files = sorted([n for n in zf.namelist()
                                   if _re.search(r'ppt/theme/theme\d*\.xml$', n, _re.I)])
            for tf in theme_files[:1]:
                xml = zf.read(tf).decode('utf-8', errors='ignore')
                for slot in ['dk1', 'lt1', 'dk2', 'lt2',
                             'accent1', 'accent2', 'accent3',
                             'accent4', 'accent5', 'accent6']:
                    m = _re.search(
                        rf'<a:{slot}[^>]*>\s*<a:srgbClr val="([0-9A-Fa-f]{{6}})"',
                        xml)
                    if m:
                        theme_colors[slot] = m.group(1).upper()
                        continue
                    m = _re.search(
                        rf'<a:{slot}[^>]*>\s*<a:sysClr[^>]*lastClr="([0-9A-Fa-f]{{6}})"',
                        xml)
                    if m:
                        theme_colors[slot] = m.group(1).upper()
                # Polices du thème
                for m in _re.findall(r'typeface="([^"]+)"', xml):
                    _add_font(m)
    except Exception as e:
        log.warning(f'[brand] theme ZIP extraction failed: {e}')

    log.info(f'[brand] theme_colors={theme_colors}')

    # ── Étape 2 : scan regex (slides + layouts + masters) ────────────────────
    sources = list(prs.slides) + list(prs.slide_layouts) + list(prs.slide_masters)
    for source in sources:
        try:
            xml = etree.tostring(source._element, pretty_print=False).decode()
        except Exception:
            continue
        for m in _re.findall(r'srgbClr val="([0-9A-Fa-f]{6})"', xml):
            color_counter[m.upper()] += 1
        for m in _re.findall(r'typeface="([^"]+)"', xml):
            _add_font(m)

    chromatic  = [c for c, _ in color_counter.most_common() if not _is_neutral(c)]
    all_colors = [c for c, _ in color_counter.most_common()]

    w, h = prs.slide_width, prs.slide_height
    log.info(f'[brand] chromatic={chromatic[:4]} fonts={fonts_seen[:3]}')
    return {
        'fonts':           fonts_seen[:5],
        'colors':          chromatic[:8],
        'all_colors':      all_colors[:20],
        'theme_colors':    theme_colors,
        'slide_count':     len(prs.slides),
        'layouts':         [l.name for l in prs.slide_layouts],
        'slide_width_in':  round(_emu(w), 2),
        'slide_height_in': round(_emu(h), 2),
        'aspect_ratio':    (
            '16:9' if abs(w / h - 16 / 9) < 0.05 else
            '4:3'  if abs(w / h - 4 / 3)  < 0.05 else 'custom'
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

    result = result[:nb_slides]

    # Garantie : si closing existe dans le template, elle est TOUJOURS en dernière position
    if closing and (not result or result[-1]["slide_type"] != "closing"):
        if len(result) < nb_slides:
            result.append(closing[0])
        else:
            result[-1] = closing[0]

    return result


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
    planner_tokens = max(2000, nb_slides * 180)

    for attempt in range(3):
        msg = client.messages.create(
            model=CLAUDE_MODEL, max_tokens=planner_tokens,
            system=system,
            messages=[{"role": "user", "content": user}],
        )
        try:
            plan = _parse_json_robust(msg.content[0].text.strip(), context="plan")
            log.info(f"Plan: {len(plan.get('slides', []))} slides — {plan.get('narrative_arc', '')[:80]}")
            return plan
        except (ValueError, KeyError) as e:
            log.warning(f"plan_presentation attempt {attempt+1}/3 échoué : {e}")
            if attempt == 2:
                raise
    raise RuntimeError("plan_presentation : 3 tentatives échouées")


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
7. ZONES VIDES INTENTIONNELLES : si une zone n'a pas de contenu pertinent pour le type de slide,
   retourne "" (chaîne vide) pour la vider. Ne jamais laisser un texte de template générique.
8. SLIDE CLOSING : toujours présente en dernière position. Titre court, mémorable. Jamais de bullet points.

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
    content_tokens = max(3000, len(slides_payload) * 320)

    msg = client.messages.create(
        model=CLAUDE_MODEL, max_tokens=content_tokens,
        system=CORTEX_SYSTEM,
        messages=[{"role": "user", "content": user}],
    )

    try:
        mapping = _parse_json_robust(msg.content[0].text.strip(), context="content")
    except ValueError as e:
        log.error(f"generate_content JSON irrécupérable : {e}")
        raise
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

# ══════════════════════════════════════════════════════════════
# RÉSOLUTION NB_SLIDES (Essentiel / Complet / Approfondi)
# ══════════════════════════════════════════════════════════════

NB_SLIDES_MAP = {
    # Valeurs fixes — le planner s'adapte
    "essentiel":  6,
    "complet":    9,
    "approfondi": 14,
}

def _resolve_nb_slides(value) -> int:
    """
    Accepte un int, un string numérique, ou un label
    ("Essentiel" / "Complet" / "Approfondi").
    """
    if isinstance(value, str):
        v = value.strip().lower()
        if v in NB_SLIDES_MAP:
            return NB_SLIDES_MAP[v]
        try:
            return max(2, min(int(v), 30))
        except ValueError:
            return 8
    try:
        return max(2, min(int(value), 30))
    except (ValueError, TypeError):
        return 8


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
    nb_slides:     str        = Form(default="complet"),
    authorization: str        = Form(default=None),
):
    pro = _is_pro(authorization)
    quota_info = (
        {"plan": "pro"} if pro
        else {"used": _quota(_ip(request))[0], "total": FREE_QUOTA_PER_IP, "plan": "free"}
    )

    n          = _resolve_nb_slides(nb_slides)
    pptx_bytes = await template.read()
    prs        = Presentation(io.BytesIO(pptx_bytes))
    brand      = extract_brand(prs)
    palette    = _h2_extract_palette(brand)
    plan       = plan_presentation_v3(prompt, n, palette)

    return {
        "success":            True,
        "presentation_title": plan.get("presentation_title", prompt[:60]),
        "footer_text":        plan.get("footer_text", ""),
        "slides": [
            {
                "index":  i,
                "layout": s.get("layout"),
                "title":  s.get("content", {}).get("title", ""),
            }
            for i, s in enumerate(plan.get("slides", []))
        ],
        "brand": brand,
        "quota": quota_info,
    }


@app.post("/generate")
async def generate_presentation(
    request:       Request,
    template:      UploadFile = File(...),
    prompt:        str        = Form(...),
    nb_slides:     str        = Form(default="complet"),
    authorization: str        = Form(default=None),
):
    if not _is_pro(authorization):
        _quota(_ip(request))

    n = _resolve_nb_slides(nb_slides)
    pptx_bytes = await template.read()

    try:
        final_bytes, plan, _, _pal = run_pipeline_v3(pptx_bytes, prompt, n)
        log.info("[/generate] Pipeline V3 OK")
    except Exception as e:
        log.warning(f"[/generate] V3 échoué ({e}) → fallback L1")
        final_bytes, plan, _ = run_pipeline(pptx_bytes, prompt, n)

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
        import asyncio, base64 as _b64
        loop = asyncio.get_event_loop()

        try:
            yield _sse("start", {"nb_slides": nb_slides})

            # Phase 1+2 : brand + planning V3
            prs_tmp = Presentation(io.BytesIO(pptx_bytes))
            brand   = extract_brand(prs_tmp)
            palette = _h2_extract_palette(brand)
            plan    = await loop.run_in_executor(
                None,
                lambda: plan_presentation_v3(prompt, nb_slides, palette)
            )
            yield _sse("planned", {"title": plan.get("presentation_title", "")})

            # Phase 3 : application layouts (dans un thread)
            yield _sse("generated", {})
            yield _sse("hydrating", {})

            try:
                final_bytes, _plan, _brand, _pal = await loop.run_in_executor(
                    None,
                    lambda: run_pipeline_v3(pptx_bytes, prompt, nb_slides)
                )
                log.info("[/generate-stream] Pipeline V3 OK")
            except Exception as e_v3:
                log.warning(f"[/generate-stream] V3 échoué ({e_v3}) → fallback L1")
                final_bytes, _plan, _ = await loop.run_in_executor(
                    None,
                    lambda: run_pipeline(pptx_bytes, prompt, nb_slides)
                )

            b64      = _b64.b64encode(final_bytes).decode()
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
#  GÉNÉRATION CRÉATIVE — Claude voit le template, crée de A à Z
#
#  Architecture v14 :
#  Phase 1 — Analyse brand + extraction visuels (thumbnail + palette image)
#  Phase 2 — Planification narrative (commun L1)
#  Phase 3 — Claude VOIT le template → génère code python-pptx par slide
#  Phase 4 — Exécution sandbox + suppression slides originales + export
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
            "Fond blanc ou très clair. Grandes marges. Hiérarchie typographique nette. "
            "Accent doré ou bleu marine. KPIs proéminents. Zéro décoration superflue. "
            "Lignes fines. Espacement aéré. Confiance et autorité."
        ),
        "layout_prefs": ["kpi", "two_col", "full_text", "quote", "timeline"],
        "bg_dark":      False,
    },
    "industrial": {
        "label":       "Industriel / Technique",
        "description": "Clarté, données, schémas, sobriété",
        "style_guide": (
            "Fond blanc ou gris très clair. Sans-serif sobre. Données et faits en premier. "
            "Couleurs : bleu, gris, orange accent. Timelines et schémas favorisés. "
            "Structure claire, pas d'effets visuels."
        ),
        "layout_prefs": ["timeline", "kpi", "two_col", "list"],
        "bg_dark":      False,
    },
    "institutional": {
        "label":       "Institutionnel / Public",
        "description": "Formalismes respectés, sobre, structuré",
        "style_guide": (
            "Fond blanc. Typographie classique. Structure formelle et lisible. "
            "Éviter les effets visuels marqués. Hiérarchie claire. Sobriété maximale. "
            "Bleu institutionnel, blanc, accent discret."
        ),
        "layout_prefs": ["full_text", "list", "two_col", "timeline"],
        "bg_dark":      False,
    },
    "startup": {
        "label":       "Startup / Tech",
        "description": "Moderne, aéré, accents colorés",
        "style_guide": (
            "Fond très clair ou sombre. Sans-serif bold. Grands espaces blancs. "
            "1-2 couleurs vives en accent. Peu de texte, impact fort. "
            "Chiffres oversize, layouts asymétriques."
        ),
        "layout_prefs": ["cover", "quote", "kpi", "image_text", "closing"],
        "bg_dark":      True,
    },
    "creative": {
        "label":       "Créatif / Agence",
        "description": "Audacieux, typographie expressive",
        "style_guide": (
            "Fond sombre ou couleur franche. Typographie oversize. Compositions audacieuses. "
            "Couleurs inattendues. Géométrie forte. "
            "Peu de texte — chaque mot compte. Impact visuel prime."
        ),
        "layout_prefs": ["cover", "quote", "image_text", "section", "closing"],
        "bg_dark":      True,
    },
}


# ─────────────────────────────────────────────────────────────
# EXTRACTION VISUELS DU TEMPLATE (pour vision Claude)
# ─────────────────────────────────────────────────────────────

def _extract_template_thumbnail(pptx_bytes: bytes) -> dict | None:
    """
    Extrait la miniature embarquée dans le PPTX (docProps/thumbnail.*).
    Présente dans ~90% des fichiers créés par PowerPoint / Google Slides.
    Retourne {'media_type': ..., 'data': base64_string} ou None.
    """
    try:
        with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
            for name in z.namelist():
                lower = name.lower()
                if 'thumbnail' in lower and lower.endswith(('.jpeg', '.jpg', '.png')):
                    data = z.read(name)
                    mt = 'image/png' if lower.endswith('.png') else 'image/jpeg'
                    return {
                        'media_type': mt,
                        'data':       base64.standard_b64encode(data).decode(),
                    }
    except Exception as e:
        log.warning(f"[V2] Extraction thumbnail : {e}")
    return None


def _make_palette_swatch(palette: dict) -> dict | None:
    """
    Génère une image de palette couleurs (600×100px) avec Pillow.
    Permet à Claude de voir les vraies couleurs de la charte.
    """
    try:
        from PIL import Image, ImageDraw, ImageFont
        W, H   = 700, 100
        img    = Image.new('RGB', (W, H), (255, 255, 255))
        draw   = ImageDraw.Draw(img)
        roles  = ['primary', 'secondary', 'accent', 'light', 'text']
        labels = ['Primary', 'Secondary', 'Accent', 'Light', 'Text']
        bw     = W // len(roles)
        for i, (role, label) in enumerate(zip(roles, labels)):
            hex_val = palette.get(role, 'CCCCCC').lstrip('#')
            try:
                r, g, b = int(hex_val[0:2], 16), int(hex_val[2:4], 16), int(hex_val[4:6], 16)
            except Exception:
                r, g, b = 128, 128, 128
            draw.rectangle([i * bw, 0, (i + 1) * bw - 2, 72], fill=(r, g, b))
            # Label blanc ou noir selon luminosité
            lum = 0.299 * r + 0.587 * g + 0.114 * b
            fg  = (255, 255, 255) if lum < 140 else (30, 30, 30)
            draw.text((i * bw + 6, 52), f"#{hex_val}", fill=fg)
            draw.text((i * bw + 6, 76), label, fill=(60, 60, 60))
        buf = io.BytesIO()
        img.save(buf, format='JPEG', quality=90)
        return {
            'media_type': 'image/jpeg',
            'data':       base64.standard_b64encode(buf.getvalue()).decode(),
        }
    except Exception as e:
        log.warning(f"[V2] Palette swatch : {e}")
        return None


# ─────────────────────────────────────────────────────────────
# HELPERS PALETTE
# ─────────────────────────────────────────────────────────────

def _lighten(hex_str: str, factor: float) -> str:
    """Éclaircit une couleur hex d'un facteur 0–1 vers le blanc."""
    h = hex_str.lstrip('#')
    try:
        r, g, b = int(h[0:2],16), int(h[2:4],16), int(h[4:6],16)
        return (f"{int(r + (255-r)*factor):02X}"
                f"{int(g + (255-g)*factor):02X}"
                f"{int(b + (255-b)*factor):02X}")
    except Exception:
        return 'AAAAAA'


def _complementary(hex_str: str) -> str:
    """Génère une couleur d'accent complémentaire sobre (décalage 150°)."""
    import colorsys
    h = hex_str.lstrip('#')
    try:
        r, g, b = int(h[0:2],16)/255, int(h[2:4],16)/255, int(h[4:6],16)/255
        hue, sat, val = colorsys.rgb_to_hsv(r, g, b)
        hue = (hue + 150/360) % 1.0
        sat = min(sat, 0.7)
        r2, g2, b2 = colorsys.hsv_to_rgb(hue, sat, min(val * 1.1, 1.0))
        return f"{int(r2*255):02X}{int(g2*255):02X}{int(b2*255):02X}"
    except Exception:
        return 'F0A500'


def _find_darkest_neutral(colors: list) -> str | None:
    """Trouve le gris le plus foncé utilisable comme primary dans un template monochrome."""
    candidates = []
    for c in colors:
        h = c.lstrip('#')
        if len(h) != 6:
            continue
        try:
            r, g, b = int(h[0:2],16), int(h[2:4],16), int(h[4:6],16)
            lum = 0.299*r + 0.587*g + 0.114*b
            if abs(r-g) < 20 and abs(g-b) < 20 and 30 < lum < 180:
                candidates.append((lum, h))
        except Exception:
            pass
    if candidates:
        return min(candidates, key=lambda x: x[0])[1]
    return None


# ─────────────────────────────────────────────────────────────
# EXTRACTION DE PALETTE
# ─────────────────────────────────────────────────────────────

def _h2_extract_palette(brand: dict) -> dict:
    """
    Construit la palette depuis les couleurs réelles du template.
    - CAS 1 : template coloré (≥2 chromatiques) → ses couleurs directement
    - CAS 2 : template semi-coloré (1 chromatique) → complète avec dérivés
    - CAS 3 : template monochrome → palette élégante depuis neutres
    Source de vérité : scan fréquence d'abord, theme_colors en complément.
    """
    def _lum(h: str) -> float:
        try:
            r, g, b = int(h[0:2],16), int(h[2:4],16), int(h[4:6],16)
            return 0.299*r + 0.587*g + 0.114*b
        except Exception:
            return 128.0

    fonts      = brand.get('fonts', [])
    chromatic  = [c.lstrip('#').strip() for c in brand.get('colors', [])
                  if len(c.lstrip('#').strip()) == 6]
    all_colors = [c.lstrip('#').strip() for c in brand.get('all_colors', [])
                  if len(c.lstrip('#').strip()) == 6]
    tc = {k: v.lstrip('#').strip() for k, v in brand.get('theme_colors', {}).items()
          if v and len(v.lstrip('#').strip()) == 6}

    # ── Supplément depuis theme si scan freq trop creux ──────────────────────
    # (templates qui utilisent des références de thème plutôt qu'RGB explicite)
    if len(chromatic) < 2 and tc:
        _seen: set = set(chromatic)
        for _slot in ['accent1', 'dk2', 'accent2']:  # slots principaux seulement
            _c = tc.get(_slot, '')
            if not _c or _c in _seen:
                continue
            _r2, _g2, _b2 = int(_c[0:2],16), int(_c[2:4],16), int(_c[4:6],16)
            _lm = 0.299*_r2 + 0.587*_g2 + 0.114*_b2
            _mx = max(_r2, _g2, _b2)
            _sat = (_mx - min(_r2,_g2,_b2)) / _mx if _mx > 0 else 0
            if _sat > 0.2 and 20 < _lm < 230:
                chromatic.append(_c)
                _seen.add(_c)

    if not all_colors and tc:
        all_colors = [v for v in tc.values() if v]

    palette: dict = {'font': fonts[0] if fonts else 'Calibri', 'theme': tc}

    # ── CAS 1 : template coloré (≥2 chromatiques) ────────────────────────────
    if len(chromatic) >= 2:
        palette['primary']   = chromatic[0]
        palette['secondary'] = chromatic[1]
        palette['accent']    = chromatic[2] if len(chromatic) >= 3 else chromatic[0]

    # ── CAS 2 : 1 couleur chromatique → dérivés ───────────────────────────────
    elif len(chromatic) == 1:
        base = chromatic[0]
        palette['primary']   = base
        palette['secondary'] = _lighten(base, 0.25)
        palette['accent']    = _complementary(base)

    # ── CAS 3 : template monochrome ───────────────────────────────────────────
    else:
        dn = _find_darkest_neutral(all_colors)
        if dn:
            palette['primary']   = dn
            palette['secondary'] = _lighten(dn, 0.35)
            palette['accent']    = _lighten(dn, 0.65)
        else:
            palette['primary']   = '1A1A1A'
            palette['secondary'] = '444444'
            palette['accent']    = '888888'

    # ── Couleur texte : dk1 du thème (le plus fiable), sinon scan ────────────
    if 'dk1' in tc and _lum(tc['dk1']) < 150:
        palette['text'] = tc['dk1']
    else:
        palette['text'] = next((c for c in all_colors if 15 < _lum(c) < 120), '1A1A1A')

    # ── Couleur sombre (fonds dark slides) : dk1 si vraiment sombre ──────────
    if 'dk1' in tc and _lum(tc['dk1']) < 100:
        palette['dark'] = tc['dk1']
    else:
        dark_cands = [(c, _lum(c)) for c in all_colors if 10 < _lum(c) < 100]
        palette['dark'] = min(dark_cands, key=lambda x: x[1])[0] if dark_cands else ''

    # ── Dérivés depuis primary si dark/light non définis ─────────────────────
    p = palette['primary']
    try:
        r, g, b = int(p[0:2],16), int(p[2:4],16), int(p[4:6],16)
        if not palette.get('dark'):
            palette['dark']  = f"{max(0,int(r*0.25)):02X}{max(0,int(g*0.25)):02X}{max(0,int(b*0.25)):02X}"
        palette['light'] = f"{int(r*0.08+255*0.92):02X}{int(g*0.08+255*0.92):02X}{int(b*0.08+255*0.92):02X}"
    except Exception:
        if not palette.get('dark'): palette['dark'] = '1A1A1A'
        palette['light'] = 'F0F4FA'

    # ── Couleur de fond des slides claires : lt1 du thème ────────────────────
    raw_lt1 = tc.get('lt1', 'FFFFFF') if tc else 'FFFFFF'
    # Exclure les valeurs invalides ou trop sombres (lt1 doit être clair)
    try:
        h = raw_lt1.lstrip('#')
        if len(h) == 6 and _lum(h) > 180:
            palette['bg'] = h.upper()
        else:
            palette['bg'] = 'FFFFFF'
    except Exception:
        palette['bg'] = 'FFFFFF'

    log.info(f"[palette] primary=#{palette['primary']} secondary=#{palette['secondary']} "
             f"accent=#{palette['accent']} light=#{palette['light']} "
             f"dark=#{palette['dark']} text=#{palette['text']} "
             f"bg=#{palette['bg']} chromatic={chromatic[:4]} font={palette['font']}")
    return palette


# ─────────────────────────────────────────────────────────────
# HELPERS H2_* — Bibliothèque de génération de shapes
# Chaque helper est exposé dans le sandbox exec() du Niveau 2.
# RÈGLE : aucun des helpers ne prend prs en argument.
#         h2_blank_slide() s'appelle SANS argument (prs est dans la closure).
# ─────────────────────────────────────────────────────────────

def _h2_parse_hex(hex_str: str) -> RGBColor:
    """Parse 'RRGGBB' ou '#RRGGBB' → RGBColor. Fallback bleu corporate si invalide."""
    try:
        h = str(hex_str).lstrip('#').strip()
        if len(h) == 3:
            h = h[0]*2 + h[1]*2 + h[2]*2
        if len(h) != 6:
            return RGBColor(0x1A, 0x3A, 0x6B)
        return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    except Exception:
        return RGBColor(0x1A, 0x3A, 0x6B)


def _h2_blank_slide(prs: Presentation):
    """
    Ajoute une slide entièrement vierge (sans placeholders résiduels).
    Cherche un layout 'Blank' par nom, sinon prend le layout index 6 ou le dernier.
    ⚠ Dans le code généré : appeler h2_blank_slide() SANS argument.
    Retourne (slide, W, H) en pouces.
    """
    target = None
    for layout in prs.slide_layouts:
        if 'blank' in layout.name.lower():
            target = layout
            break
    if target is None:
        idx = min(6, len(prs.slide_layouts) - 1)
        target = prs.slide_layouts[idx]

    slide   = prs.slides.add_slide(target)
    sp_tree = slide.shapes._spTree
    for ph in list(slide.placeholders):
        try:
            sp_tree.remove(ph._element)
        except Exception:
            pass

    W = prs.slide_width  / 914400.0
    H = prs.slide_height / 914400.0
    return slide, W, H


def _h2_rect(slide, left: float, top: float, width: float, height: float,
             color: str, alpha: int = 0):
    """
    Rectangle coloré plein (en pouces). color = hex 'RRGGBB'.
    alpha ignoré (python-pptx ne supporte pas la transparence sur les formes).
    """
    from pptx.enum.shapes import MSO_SHAPE_TYPE as _MST
    shape = slide.shapes.add_shape(1, Inches(left), Inches(top),
                                    Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = _h2_parse_hex(color)
    shape.line.fill.background()
    return shape


def _h2_rounded_rect(slide, left: float, top: float, width: float, height: float,
                     color: str, radius: float = 0.08):
    """
    Rectangle arrondi (pill). radius = proportion d'arrondi (0–0.5).
    """
    from pptx.util import Emu as _Emu
    shape = slide.shapes.add_shape(
        5,  # MSO_SHAPE.ROUNDED_RECTANGLE
        Inches(left), Inches(top), Inches(width), Inches(height),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = _h2_parse_hex(color)
    shape.line.fill.background()
    try:
        shape.adjustments[0] = max(0.0, min(0.5, radius))
    except Exception:
        pass
    return shape


def _h2_circle(slide, cx: float, cy: float, r: float, color: str):
    """
    Cercle centré en (cx, cy) de rayon r (en pouces).
    """
    shape = slide.shapes.add_shape(
        9,  # MSO_SHAPE.OVAL
        Inches(cx - r), Inches(cy - r), Inches(r * 2), Inches(r * 2),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = _h2_parse_hex(color)
    shape.line.fill.background()
    return shape


def _h2_text(slide, text: str,
             left: float, top: float, width: float, height: float,
             font: str, size_pt: float, color: str,
             bold: bool = False, italic: bool = False, align: str = 'left',
             line_spacing: float = 1.0):
    """
    Textbox stylée (dimensions en pouces).
    align : 'left' | 'center' | 'right'
    line_spacing : multiplicateur d'interligne (1.0 = normal, 1.2 = aéré)
    """
    txBox = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height),
    )
    tf           = txBox.text_frame
    tf.word_wrap = True

    align_map = {'left': PP_ALIGN.LEFT, 'center': PP_ALIGN.CENTER, 'right': PP_ALIGN.RIGHT}

    p           = tf.paragraphs[0]
    p.alignment = align_map.get(align, PP_ALIGN.LEFT)

    # Interligne
    try:
        from pptx.util import Pt as _Pt
        from pptx.oxml.ns import qn as _qn
        pPr = p._pPr
        if pPr is None:
            pPr = p._p.get_or_add_pPr()
        import lxml.etree as _etree
        lnSpc = _etree.SubElement(pPr, _qn('a:lnSpc'))
        spcPct = _etree.SubElement(lnSpc, _qn('a:spcPct'))
        spcPct.set('val', str(int(line_spacing * 100000)))
    except Exception:
        pass

    run             = p.add_run()
    run.text        = str(text)
    run.font.name   = str(font)
    run.font.size   = Pt(size_pt)
    run.font.bold   = bold
    run.font.italic = italic
    run.font.color.rgb = _h2_parse_hex(color)
    return txBox


def _h2_multiline_text(slide, lines: list,
                       left: float, top: float, width: float, height: float,
                       font: str, size_pt: float, color: str,
                       bold: bool = False, align: str = 'left',
                       line_spacing: float = 1.15):
    """
    Textbox avec plusieurs paragraphes (une entrée de `lines` par paragraphe).
    Chaque entrée peut être un str ou un dict {'text', 'bold', 'size', 'color'}.
    """
    txBox = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height),
    )
    tf           = txBox.text_frame
    tf.word_wrap = True
    align_map    = {'left': PP_ALIGN.LEFT, 'center': PP_ALIGN.CENTER, 'right': PP_ALIGN.RIGHT}

    for i, line in enumerate(lines):
        if isinstance(line, dict):
            txt    = line.get('text', '')
            b      = line.get('bold', bold)
            sz     = line.get('size', size_pt)
            clr    = line.get('color', color)
        else:
            txt, b, sz, clr = str(line), bold, size_pt, color

        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align_map.get(align, PP_ALIGN.LEFT)

        run             = p.add_run()
        run.text        = txt
        run.font.name   = font
        run.font.size   = Pt(sz)
        run.font.bold   = b
        run.font.color.rgb = _h2_parse_hex(clr)
    return txBox


def _h2_kpi(slide,
            left: float, top: float, width: float,
            value: str, label: str, sublabel: str,
            palette: dict, dark: bool = True):
    """
    Bloc KPI : grande valeur + label + sous-label.
    dark=True  → textes blancs (pour fonds sombres)
    dark=False → textes couleur charte (pour fonds clairs)
    width ≈ 2.0–3.0 pouces.
    """
    font    = palette.get('font', 'Calibri')
    v_clr   = 'FFFFFF' if dark else palette.get('primary', '1A3A6B')
    l_clr   = palette.get('accent', 'F0A500')
    sl_clr  = 'CCCCCC' if dark else '888888'

    _h2_text(slide, str(value),  left, top,        width, 0.75, font, 34, v_clr, bold=True, align='center')
    _h2_rect(slide, left + width*0.18, top + 0.76,  width*0.64, 0.035, l_clr)
    _h2_text(slide, str(label),  left, top + 0.82,  width, 0.38, font, 12, l_clr,  bold=False, align='center')
    _h2_text(slide, str(sublabel), left, top + 1.22, width, 0.52, font, 10, sl_clr, bold=False, align='center')


def _h2_divider(slide, left: float, top: float, width: float,
                color: str, thickness: float = 0.035):
    """Ligne horizontale fine (en pouces)."""
    _h2_rect(slide, left, top, width, thickness, color)


def _h2_number(slide, text: str, left: float, top: float,
               size_in: float, color: str, font: str, opacity_hint: str = ''):
    """
    Grand chiffre / lettre décoratif (numéro de section, oversize stat).
    size_in = hauteur approximative de la boîte en pouces.
    """
    font_pt = max(28, int(size_in * 70))
    _h2_text(slide, str(text), left, top, size_in * 2.0, size_in + 0.2,
             font, font_pt, color, bold=True, align='center')


def _h2_icon_circle(slide, cx: float, cy: float, r: float,
                    label: str, font: str, color_bg: str, color_fg: str,
                    font_size: float = 11):
    """
    Cercle coloré avec un label centré dedans (substitut d'icône).
    """
    _h2_circle(slide, cx, cy, r, color_bg)
    _h2_text(slide, label, cx - r, cy - r * 0.55, r * 2, r * 1.1,
             font, font_size, color_fg, bold=True, align='center')


def _h2_card(slide, left: float, top: float, width: float, height: float,
             bg_color: str, title: str, body: str,
             font: str, title_color: str, body_color: str,
             title_size: float = 13, body_size: float = 11,
             rounded: bool = True):
    """
    Carte avec fond coloré + titre en gras + corps.
    Utile pour grilles, colonnes, items de liste visuels.
    """
    if rounded:
        _h2_rounded_rect(slide, left, top, width, height, bg_color, radius=0.06)
    else:
        _h2_rect(slide, left, top, width, height, bg_color)
    pad = 0.18
    _h2_text(slide, title, left + pad, top + pad, width - pad*2, 0.42,
             font, title_size, title_color, bold=True)
    if body:
        _h2_text(slide, body, left + pad, top + pad + 0.44, width - pad*2,
                 height - pad*2 - 0.44, font, body_size, body_color,
                 line_spacing=1.2)


def _h2_progress_bar(slide, left: float, top: float, width: float,
                     value_pct: float, color_fill: str, color_bg: str = 'EEEEEE',
                     height: float = 0.12):
    """
    Barre de progression horizontale. value_pct entre 0 et 100.
    """
    _h2_rect(slide, left, top, width, height, color_bg)
    fill_w = max(0.0, min(width, width * value_pct / 100.0))
    if fill_w > 0:
        _h2_rect(slide, left, top, fill_w, height, color_fill)


def _h2_tag(slide, text: str, left: float, top: float,
            font: str, size_pt: float, bg_color: str, fg_color: str):
    """
    Pill / tag arrondi avec texte centré. Hauteur auto ≈ size_pt * 0.022 po.
    """
    h = size_pt * 0.028 + 0.08
    w = len(text) * size_pt * 0.014 + 0.3
    _h2_rounded_rect(slide, left, top, w, h, bg_color, radius=0.5)
    _h2_text(slide, text, left, top, w, h, font, size_pt, fg_color,
             bold=True, align='center')


# ─────────────────────────────────────────────────────────────
# SÉCURITÉ — Validation du code généré
# ─────────────────────────────────────────────────────────────

_FORBIDDEN_CODE_PATTERNS = [
    'import ', '__import__', 'eval(', 'exec(',
    'open(', 'os.', 'sys.', 'subprocess', 'socket',
    'urllib', 'requests', '__builtins__', '__globals__',
    '__locals__', '__class__', 'compile(', 'globals(',
    'locals(', 'vars(', 'dir(',
]

def _validate_code_safety(code: str) -> tuple:
    for pattern in _FORBIDDEN_CODE_PATTERNS:
        if pattern in code:
            return False, f"Pattern interdit: '{pattern}'"
    return True, ''


# ─────────────────────────────────────────────────────────────
# NAMESPACE SANDBOX
# ─────────────────────────────────────────────────────────────

def _build_safe_namespace(prs: Presentation, palette: dict) -> dict:
    """
    Namespace restreint exposé au code généré.
    ⚠ h2_blank_slide est une lambda SANS argument (prs dans la closure).
    """
    return {
        'brand':  palette,
        # Helpers — h2_blank_slide() s'appelle SANS argument
        'h2_blank_slide':    lambda: _h2_blank_slide(prs),
        'h2_rect':           _h2_rect,
        'h2_rounded_rect':   _h2_rounded_rect,
        'h2_circle':         _h2_circle,
        'h2_text':           _h2_text,
        'h2_multiline_text': _h2_multiline_text,
        'h2_kpi':            _h2_kpi,
        'h2_divider':        _h2_divider,
        'h2_number':         _h2_number,
        'h2_icon_circle':    _h2_icon_circle,
        'h2_card':           _h2_card,
        'h2_progress_bar':   _h2_progress_bar,
        'h2_tag':            _h2_tag,
        # Builtins sécurisés
        '__builtins__': {
            'range': range, 'len': len, 'int': int, 'float': float,
            'str': str, 'bool': bool, 'list': list, 'dict': dict,
            'tuple': tuple, 'enumerate': enumerate, 'zip': zip,
            'round': round, 'abs': abs, 'min': min, 'max': max,
            'sum': sum, 'sorted': sorted, 'reversed': reversed,
            'any': any, 'all': all, 'print': lambda *a, **k: None,
            'True': True, 'False': False, 'None': None,
        },
    }


# ─────────────────────────────────────────────────────────────
# PROMPTS NIVEAU 2
# ─────────────────────────────────────────────────────────────

_V2_SYSTEM = """\
Tu es Visual Cortex Level 2 — Créateur de présentations PowerPoint professionnelles.

Tu reçois :
1. Une image de miniature du template de l'entreprise (pour voir la charte visuelle réelle)
2. Une image de la palette de couleurs extraite
3. Le plan narratif à implémenter

TA MISSION : générer du code python-pptx pour chaque slide. Le résultat doit être :
→ Professionnel, impactant, visuellement varié
→ Fidèle à la charte de l'entreprise (ses vraies couleurs, ses proportions)
→ Digne d'une agence de communication

══════════════════════════════════════════════
FONCTIONS DISPONIBLES (et UNIQUEMENT celles-ci)
══════════════════════════════════════════════

slide, W, H = h2_blank_slide()
  → TOUJOURS la première ligne de chaque slide. AUCUN ARGUMENT.
  → W ≈ 10.0, H ≈ 5.63 pour 16:9. Dimensions en pouces.

h2_rect(slide, left, top, width, height, color)
  → Rectangle plein. color = brand["primary"] ou "FFFFFF".

h2_rounded_rect(slide, left, top, width, height, color, radius=0.08)
  → Rectangle arrondi. radius entre 0 et 0.5. Parfait pour cartes, tags.

h2_circle(slide, cx, cy, r, color)
  → Cercle centré (cx, cy) de rayon r. Tous en pouces.

h2_text(slide, text, left, top, width, height, font, size_pt, color,
        bold=False, italic=False, align="left", line_spacing=1.0)
  → Textbox. align: "left"|"center"|"right".
  → Utiliser brand["font"] pour TOUS les textes.

h2_multiline_text(slide, lines, left, top, width, height, font, size_pt, color,
                  bold=False, align="left", line_spacing=1.15)
  → Multi-paragraphes. lines = liste de str ou dict {"text","bold","size","color"}.

h2_kpi(slide, left, top, width, value, label, sublabel, brand, dark=True)
  → Bloc KPI. dark=True pour fond sombre, dark=False pour fond clair.
  → width ≈ 2.0–3.0 po.

h2_divider(slide, left, top, width, color, thickness=0.035)
  → Ligne horizontale fine. Utiliser avec parcimonie.

h2_number(slide, text, left, top, size_in, color, font)
  → Grand chiffre/lettre décoratif. size_in ≈ 1.5–2.5 po.

h2_icon_circle(slide, cx, cy, r, label, font, color_bg, color_fg, font_size=11)
  → Cercle coloré avec texte centré (substitut d'icône).

h2_card(slide, left, top, width, height, bg_color, title, body,
        font, title_color, body_color, title_size=13, body_size=11, rounded=True)
  → Carte avec fond + titre + corps. Parfait pour grilles.

h2_progress_bar(slide, left, top, width, value_pct, color_fill, color_bg="EEEEEE", height=0.12)
  → Barre de progression horizontale (0–100%).

h2_tag(slide, text, left, top, font, size_pt, bg_color, fg_color)
  → Pill / tag arrondi avec texte. Largeur automatique.

══════════════════════════════════════════════
ACCÈS À LA CHARTE (via brand dict)
══════════════════════════════════════════════
brand["primary"]   → couleur principale de la charte
brand["secondary"] → couleur secondaire
brand["accent"]    → couleur d'accent (titres, highlights)
brand["light"]     → version très claire du primary (fonds doux)
brand["text"]      → couleur de texte foncé
brand["font"]      → police principale à utiliser PARTOUT

Couleurs supplémentaires utiles (hardcodées) :
"FFFFFF" = blanc   "000000" = noir   "888888" = gris moyen
"F5F5F5" = gris très clair   "333333" = presque noir

══════════════════════════════════════════════
RÈGLES DE DENSITÉ (inviolables)
══════════════════════════════════════════════
cover    → titre ≤ 7 mots  + sous-titre ≤ 12 mots
section  → numéro décoratif + titre ≤ 6 mots UNIQUEMENT
kpi      → 4 à 6 KPIs MAX. Valeur courte + label court + sous-label contextuel
timeline → 4 à 6 jalons. Date + titre ≤ 4 mots + phrase optionnelle ≤ 12 mots
two_col  → 2 colonnes symétriques. Max 4 items/colonne ≤ 18 mots chacun
quote    → 1 citation ≤ 20 mots. Source ≤ 8 mots
list     → 3 à 5 items. Titre ≤ 5 mots + corps ≤ 20 mots
full_text→ titre ≤ 8 mots + 2-3 paragraphes ≤ 60 mots total
closing  → titre ≤ 5 mots + sous-titre ≤ 15 mots

══════════════════════════════════════════════
RÈGLES VISUELLES OBLIGATOIRES
══════════════════════════════════════════════
1. VARIÉTÉ : chaque slide DOIT avoir un look distinct des autres.
   Alterner fonds sombres (primary) et clairs (FFFFFF ou light).
   Ne jamais répéter la même composition deux fois.

2. STRUCTURE SANDWICH : cover et closing sur fond sombre/coloré.
   Slides de contenu sur fond clair. Sections sur fond primary.

3. RESPIRATION : marges min 0.4 po sur tous les bords.
   Espacement entre blocs ≥ 0.25 po.

4. HIÉRARCHIE : titre 28-44pt bold. Corps 12-16pt. Footer 10pt.
   Contraste fort entre titre et corps.

5. FOOTER : présent sur toutes les slides de contenu (pas cover ni closing).
   Toujours en bas, texte 10pt, couleur discrète.

6. COULEURS : utiliser UNIQUEMENT les couleurs brand[] + blanc + gris.
   Jamais inventer une couleur non présente dans la charte.

══════════════════════════════════════════════
PATTERNS PAR TYPE (à adapter à la charte vue)
══════════════════════════════════════════════

[cover — fond primary] :
slide, W, H = h2_blank_slide()
h2_rect(slide, 0, 0, W, H, brand["primary"])
h2_rect(slide, 0, H*0.75, W, H*0.25, brand["secondary"])
h2_text(slide, "Titre fort ≤ 7 mots", 0.7, H/2-1.2, W-1.4, 1.5,
        brand["font"], 44, "FFFFFF", bold=True)
h2_text(slide, "Sous-titre contextuel ≤ 12 mots", 0.7, H/2+0.45, W-1.4, 0.8,
        brand["font"], 20, "FFFFFF")
h2_text(slide, "Footer · Contexte · 2025", 0.7, H-0.52, W-1.4, 0.4,
        brand["font"], 10, brand["secondary"])

[cover — barre latérale sur fond clair] :
slide, W, H = h2_blank_slide()
h2_rect(slide, 0, 0, W, H, "FFFFFF")
h2_rect(slide, 0, 0, 0.55, H, brand["primary"])
h2_rect(slide, 0.55, H*0.45, W-0.55, 0.055, brand["accent"])
h2_text(slide, "Titre fort ≤ 7 mots", 0.9, H*0.45-0.9, W-1.4, 1.3,
        brand["font"], 40, brand["primary"], bold=True)
h2_text(slide, "Sous-titre ≤ 12 mots", 0.9, H*0.45+0.55, W-1.4, 0.75,
        brand["font"], 18, brand["text"])
h2_text(slide, "Footer · Contexte · 2025", 0.9, H-0.52, W-1.4, 0.4,
        brand["font"], 10, brand["text"])

[section] :
slide, W, H = h2_blank_slide()
h2_rect(slide, 0, 0, W, H, brand["primary"])
h2_number(slide, "01", 0.7, H/2-1.5, 2.0, brand["secondary"], brand["font"])
h2_text(slide, "Titre de section ≤ 6 mots", 0.7, H/2+0.6, W-1.2, 1.0,
        brand["font"], 34, "FFFFFF", bold=True)
h2_rect(slide, 0.7, H/2+0.52, 3.5, 0.055, brand["accent"])

[kpi — grille 3×2 sur fond sombre] :
slide, W, H = h2_blank_slide()
h2_rect(slide, 0, 0, W, H, brand["primary"])
h2_text(slide, "Titre KPIs ≤ 7 mots", 0.5, 0.22, W-1.0, 0.75,
        brand["font"], 30, "FFFFFF", bold=True)
h2_rect(slide, 0.5, 1.08, W-1.0, 0.04, brand["accent"])
kpi_data = [("600 M€","contribution","versée en 2022"),
            ("~5.6%","du capital","État actionnaire"),
            ("28 Md€","CA mondial","exercice 2023"),
            ("95k","collaborateurs","dont 25k en France"),
            ("80+","pays","présence opérationnelle"),
            ("1er rang","en France","par capitalisation")]
for i, (v, l, s) in enumerate(kpi_data):
    col = i % 3
    row = i // 3
    x = 0.4 + col * ((W-0.8)/3)
    y = 1.35 + row * 1.95
    h2_kpi(slide, x, y, (W-0.8)/3 - 0.15, v, l, s, brand, dark=True)
h2_text(slide, "Footer · Contexte · 2025", 0.5, H-0.46, W-1.0, 0.38,
        brand["font"], 10, brand["secondary"])

[kpi — ligne de 4 sur fond clair] :
slide, W, H = h2_blank_slide()
h2_rect(slide, 0, 0, W, H, "FFFFFF")
h2_rect(slide, 0, 0, 0.08, H, brand["primary"])
h2_text(slide, "Titre KPIs", 0.5, 0.22, W-1.0, 0.72,
        brand["font"], 28, brand["primary"], bold=True)
h2_rect(slide, 0.5, 1.05, W-1.0, 0.04, brand["accent"])
h2_rect(slide, 0.5, 1.28, W-1.0, 2.35, brand["light"])
kpi_data = [("600 M€","contribution","versée en 2022"),
            ("~5.6%","du capital","État actionnaire"),
            ("28 Md€","CA mondial","2023"),
            ("95k","collaborateurs","80 pays")]
kw = (W-1.0) / len(kpi_data)
for i, (v, l, s) in enumerate(kpi_data):
    h2_kpi(slide, 0.5 + i*kw, 1.45, kw-0.1, v, l, s, brand, dark=False)
h2_text(slide, "Footer · Contexte · 2025", 0.5, H-0.46, W-1.0, 0.38,
        brand["font"], 10, brand["primary"])

[timeline — axe horizontal sur fond clair] :
slide, W, H = h2_blank_slide()
h2_rect(slide, 0, 0, W, H, "FFFFFF")
h2_rect(slide, 0, 0, 0.08, H, brand["primary"])
h2_text(slide, "Chronologie", 0.5, 0.22, W-1.0, 0.72,
        brand["font"], 28, brand["primary"], bold=True)
h2_rect(slide, 0.5, 1.05, W-1.0, 0.04, brand["accent"])
h2_rect(slide, 0.5, 2.45, W-1.0, 0.065, brand["primary"])
steps = [("1924","Création","Fondation par l'État"),
         ("1985","Privatisation","Entrée en bourse"),
         ("2003","Fusion","Nouveau groupe mondial"),
         ("2024","Centenaire","Repositionnement stratégique")]
n = len(steps)
for i, (date, titre, detail) in enumerate(steps):
    x = 0.5 + i * ((W-1.0)/(n-1))
    h2_circle(slide, x, 2.48, 0.16, brand["primary"])
    h2_text(slide, date,  x-0.6, 1.68, 1.2, 0.52,
            brand["font"], 13, brand["primary"], bold=True, align="center")
    h2_text(slide, titre, x-0.7, 2.82, 1.4, 0.44,
            brand["font"], 12, brand["text"], bold=True, align="center")
    h2_text(slide, detail, x-0.8, 3.32, 1.6, 0.56,
            brand["font"], 10, "888888", align="center")
h2_text(slide, "Footer · Contexte · 2025", 0.5, H-0.46, W-1.0, 0.38,
        brand["font"], 10, brand["primary"])

[two_col — fond blanc, colonnes avec header coloré] :
slide, W, H = h2_blank_slide()
h2_rect(slide, 0, 0, W, H, "FFFFFF")
h2_rect(slide, 0, 0, 0.08, H, brand["primary"])
h2_text(slide, "Comparaison / Dualité", 0.5, 0.22, W-1.0, 0.72,
        brand["font"], 28, brand["primary"], bold=True)
h2_rect(slide, 0.5, 1.05, W-1.0, 0.04, brand["accent"])
col_w = (W-1.2) / 2
for ci, (label, items) in enumerate([
    ("COLONNE A", ["Aspect 1 concis", "Aspect 2 concis", "Aspect 3", "Aspect 4"]),
    ("COLONNE B", ["Aspect 1 bis", "Aspect 2 bis", "Aspect 3 bis", "Aspect 4 bis"])
]):
    x = 0.5 + ci * (col_w + 0.2)
    clr = brand["primary"] if ci == 0 else brand["secondary"]
    h2_rect(slide, x, 1.25, col_w, 0.5, clr)
    h2_text(slide, label, x+0.12, 1.30, col_w-0.24, 0.4,
            brand["font"], 13, "FFFFFF", bold=True)
    for j, item in enumerate(items):
        h2_text(slide, "→  " + item, x+0.12, 1.92+j*0.65, col_w-0.24, 0.58,
                brand["font"], 12, brand["text"])
h2_text(slide, "Footer · Contexte · 2025", 0.5, H-0.46, W-1.0, 0.38,
        brand["font"], 10, brand["primary"])

[quote — fond sombre, barre latérale accent] :
slide, W, H = h2_blank_slide()
h2_rect(slide, 0, 0, W, H, brand["primary"])
h2_rect(slide, 0.52, H/2-1.15, 0.1, 2.3, brand["accent"])
h2_text(slide, "\u00ab\u202fCitation forte et mémorable, ≤ 20 mots absolument.\u202f\u00bb",
        0.88, H/2-0.98, W-1.6, 1.85,
        brand["font"], 26, "FFFFFF", bold=True, line_spacing=1.25)
h2_text(slide, "\u2014 Auteur ou source, 2025",
        0.88, H/2+1.0, W-1.6, 0.58,
        brand["font"], 14, brand["accent"])
h2_text(slide, "Footer · Contexte · 2025", 0.5, H-0.46, W-1.0, 0.38,
        brand["font"], 10, brand["secondary"])

[list — cartes numérotées sur fond clair] :
slide, W, H = h2_blank_slide()
h2_rect(slide, 0, 0, W, H, "FFFFFF")
h2_rect(slide, 0, 0, 0.08, H, brand["primary"])
h2_text(slide, "Titre de la liste", 0.5, 0.22, W-1.0, 0.72,
        brand["font"], 28, brand["primary"], bold=True)
h2_rect(slide, 0.5, 1.05, W-1.0, 0.04, brand["accent"])
items = [("Titre 1","Corps concis, une idée forte, ≤ 20 mots."),
         ("Titre 2","Corps concis, une idée distincte."),
         ("Titre 3","Corps concis, conclusion ou conséquence."),
         ("Titre 4","Corps concis, argument complémentaire.")]
for i, (titre, corps) in enumerate(items):
    y = 1.28 + i * 0.88
    h2_circle(slide, 0.82, y+0.26, 0.22, brand["primary"])
    h2_text(slide, str(i+1), 0.60, y+0.06, 0.44, 0.4,
            brand["font"], 16, "FFFFFF", bold=True, align="center")
    h2_text(slide, titre, 1.22, y,        W-2.1, 0.38, brand["font"], 13, brand["primary"], bold=True)
    h2_text(slide, corps, 1.22, y+0.40,   W-2.1, 0.44, brand["font"], 11, brand["text"])
h2_text(slide, "Footer · Contexte · 2025", 0.5, H-0.46, W-1.0, 0.38,
        brand["font"], 10, brand["primary"])

[list — grille de cartes 2×2] :
slide, W, H = h2_blank_slide()
h2_rect(slide, 0, 0, W, H, brand["light"])
h2_text(slide, "Titre liste cartes", 0.5, 0.22, W-1.0, 0.72,
        brand["font"], 28, brand["primary"], bold=True)
h2_rect(slide, 0.5, 1.05, W-1.0, 0.04, brand["accent"])
cards = [("Titre 1","Contenu concis ≤ 20 mots.", brand["primary"], "FFFFFF"),
         ("Titre 2","Contenu concis ≤ 20 mots.", "FFFFFF", brand["primary"]),
         ("Titre 3","Contenu concis ≤ 20 mots.", brand["secondary"], "FFFFFF"),
         ("Titre 4","Contenu concis ≤ 20 mots.", "FFFFFF", brand["secondary"])]
cw = (W-1.4) / 2
ch = 1.55
for i, (titre, corps, bg, fg) in enumerate(cards):
    col = i % 2
    row = i // 2
    x = 0.5 + col*(cw+0.2)
    y = 1.28 + row*(ch+0.18)
    h2_card(slide, x, y, cw, ch, bg, titre, corps,
            brand["font"], fg, fg, title_size=13, body_size=11)
h2_text(slide, "Footer · Contexte · 2025", 0.5, H-0.46, W-1.0, 0.38,
        brand["font"], 10, brand["primary"])

[image_text — split vertical, texte à droite] :
slide, W, H = h2_blank_slide()
h2_rect(slide, 0, 0, W/2, H, brand["primary"])
h2_text(slide, "Titre ≤ 8 mots", W/2+0.4, 0.35, W/2-0.75, 0.9,
        brand["font"], 26, brand["primary"], bold=True)
h2_rect(slide, W/2+0.4, 1.35, W/2-0.75, 0.04, brand["accent"])
points = ["Point clé 1 — concis et fort.", "Point clé 2 — idée distincte.", "Point clé 3 — conclusion."]
for i, pt in enumerate(points):
    h2_circle(slide, W/2+0.62, 1.78+i*0.9, 0.14, brand["accent"])
    h2_text(slide, "  " + pt, W/2+0.85, 1.62+i*0.9, W/2-1.2, 0.75,
            brand["font"], 12, brand["text"])
h2_text(slide, "Footer · Contexte · 2025", W/2+0.4, H-0.46, W/2-0.75, 0.38,
        brand["font"], 10, brand["primary"])

[full_text — fond blanc, layout aéré] :
slide, W, H = h2_blank_slide()
h2_rect(slide, 0, 0, W, H, "FFFFFF")
h2_rect(slide, 0, 0, 0.08, H, brand["primary"])
h2_text(slide, "Titre développement ≤ 8 mots", 0.5, 0.22, W-1.0, 0.72,
        brand["font"], 28, brand["primary"], bold=True)
h2_rect(slide, 0.5, 1.05, W-1.0, 0.04, brand["accent"])
h2_rounded_rect(slide, 0.5, 1.25, W-1.0, 0.9, brand["light"])
h2_text(slide, "Premier paragraphe : une idée principale, 2 phrases courtes. Ton direct, factuel et orienté valeur.",
        0.68, 1.38, W-1.36, 0.68, brand["font"], 12, brand["text"], line_spacing=1.3)
h2_text(slide, "Deuxième paragraphe : idée distincte et complémentaire. Pas redondant avec le premier.",
        0.5, 2.32, W-1.0, 0.68, brand["font"], 12, brand["text"], line_spacing=1.3)
h2_text(slide, "Troisième paragraphe : conclusion ou conséquence pratique. Ce que ça implique concrètement.",
        0.5, 3.12, W-1.0, 0.68, brand["font"], 12, brand["text"], line_spacing=1.3)
h2_text(slide, "Footer · Contexte · 2025", 0.5, H-0.46, W-1.0, 0.38,
        brand["font"], 10, brand["primary"])

[closing — fond sombre centré] :
slide, W, H = h2_blank_slide()
h2_rect(slide, 0, 0, W, H, brand["primary"])
h2_rect(slide, 0, H*0.76, W, H*0.24, brand["secondary"])
h2_text(slide, "Merci !", W/2-3.8, H/2-1.05, 7.6, 1.5,
        brand["font"], 58, "FFFFFF", bold=True, align="center")
h2_text(slide, "Sources · Contact · 2025", W/2-3.8, H/2+0.58, 7.6, 0.7,
        brand["font"], 16, "FFFFFF", align="center")
h2_text(slide, "Footer · Contexte · 2025", 0.7, H-0.46, W-1.4, 0.38,
        brand["font"], 11, "FFFFFF")
"""

_V2_USER_TEMPLATE = """\
PRÉSENTATION : {title}
ARC NARRATIF : {arc}
FOOTER UNIFIÉ (identique sur toutes les slides de contenu) : "{footer}"
SUJET COMPLET : {prompt}

CHARTE EXTRAITE DU TEMPLATE :
  brand["primary"]   = "#{primary}"
  brand["secondary"] = "#{secondary}"
  brand["accent"]    = "#{accent}"
  brand["light"]     = "#{light}"
  brand["text"]      = "#{text_color}"
  brand["font"]      = "{font}"
  Dimensions slide   : W = {W:.2f} po  ×  H = {H:.2f} po

PROFIL CLIENT : {profile_label}
Guide de style : {style_guide}

SLIDES À GÉNÉRER — {n} slides :
{slides_json}

══════════════════════════════════════════════
INSTRUCTIONS DE GÉNÉRATION
══════════════════════════════════════════════
1. Pour chaque slide, écris le code python-pptx COMPLET qui crée la slide.
2. COMMENCE chaque slide par : slide, W, H = h2_blank_slide()  ← SANS argument
3. Adapte le contenu réel au narrative_angle et key_message de chaque slide.
4. RESPECTE les density rules : titre ≤ 8 mots, body ≤ 40 mots, etc.
5. VARIE les compositions : alterne fonds sombres et clairs, change les layouts.
6. Structure SANDWICH : cover et closing sur fond sombre, sections sur fond primary,
   slides de contenu sur fond clair (FFFFFF ou light).
7. Utilise le footer_text sur toutes les slides SAUF cover et closing.
8. Les couleurs brand[] sont visibles dans les images jointes — respecte-les.
9. Utilise des boucles Python pour les grilles/listes répétitives.

FORMAT DE SORTIE : JSON valide uniquement, sans markdown, sans commentaires.
Clés = plan_index en string ("0", "1", ...).
{{
  "0": "slide, W, H = h2_blank_slide()\\nh2_rect(...)\\n...",
  "1": "slide, W, H = h2_blank_slide()\\n...",
  ...
}}
Chaque valeur = code Python complet pour UNE slide (string avec \\n comme séparateurs).
"""


# ─────────────────────────────────────────────────────────────
# GÉNÉRATION DE CODE — Appel Claude avec vision
# ─────────────────────────────────────────────────────────────

def generate_codes_v2(
    prompt:    str,
    plan:      dict,
    palette:   dict,
    brand:     dict,
    profile:   str,
    template_thumbnail: dict | None = None,
    palette_swatch:     dict | None = None,
) -> dict:
    """
    Phase 3 du pipeline Niveau 2.
    Claude reçoit les images du template ET de la palette pour voir la vraie charte.
    Retourne { "plan_index": "code_python_string" }.
    """
    client       = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    profile_data = CLIENT_PROFILES.get(profile, CLIENT_PROFILES['institutional'])

    slides_payload = [
        {
            'plan_index':      sp.get('plan_index', i),
            'slide_type':      sp.get('slide_type', 'unknown'),
            'narrative_angle': sp.get('narrative_angle', ''),
            'key_message':     sp.get('key_message', ''),
            'visual_hint':     sp.get('visual_hint', ''),
        }
        for i, sp in enumerate(plan.get('slides', []))
    ]

    user_text = _V2_USER_TEMPLATE.format(
        title       = plan.get('presentation_title', prompt[:60]),
        arc         = plan.get('narrative_arc', ''),
        footer      = plan.get('footer_text', ''),
        prompt      = prompt,
        primary     = palette.get('primary', '1A3A6B'),
        secondary   = palette.get('secondary', '2E6DA4'),
        accent      = palette.get('accent', 'F0A500'),
        light       = palette.get('light', 'EEF3FA'),
        text_color  = palette.get('text', '1A1A2E'),
        font        = palette.get('font', 'Calibri'),
        W           = brand.get('slide_width_in', 10.0),
        H           = brand.get('slide_height_in', 5.63),
        profile_label = profile_data['label'],
        style_guide   = profile_data['style_guide'],
        n           = len(slides_payload),
        slides_json = json.dumps(slides_payload, ensure_ascii=False, indent=2),
    )

    # Construction du message multimodal (images + texte)
    content_parts = []

    if template_thumbnail:
        content_parts.append({
            'type': 'text',
            'text': 'Voici la miniature du template de l\'entreprise (charte graphique réelle) :',
        })
        content_parts.append({
            'type':   'image',
            'source': {
                'type':       'base64',
                'media_type': template_thumbnail['media_type'],
                'data':       template_thumbnail['data'],
            },
        })

    if palette_swatch:
        content_parts.append({
            'type': 'text',
            'text': 'Palette de couleurs extraite du template :',
        })
        content_parts.append({
            'type':   'image',
            'source': {
                'type':       'base64',
                'media_type': palette_swatch['media_type'],
                'data':       palette_swatch['data'],
            },
        })

    content_parts.append({'type': 'text', 'text': user_text})

    code_tokens = max(4000, len(slides_payload) * 600)

    for attempt in range(3):
        if attempt > 0:
            log.info(f'[V2] Retry génération code ({attempt+1}/3)...')

        msg = client.messages.create(
            model      = CLAUDE_MODEL,
            max_tokens = code_tokens,
            system     = _V2_SYSTEM,
            messages   = [{'role': 'user', 'content': content_parts}],
        )

        try:
            code_map = _parse_json_robust(msg.content[0].text.strip(), context='codes_v2')
            log.info(f'[V2] Code généré pour {len(code_map)} slides (attempt {attempt+1}).')
            return code_map
        except (ValueError, KeyError) as e:
            log.warning(f'[V2] generate_codes_v2 attempt {attempt+1}/3 échoué : {e}')
            if attempt == 2:
                log.warning('[V2] 3 tentatives échouées → dict vide (fallback L1 sera déclenché)')
                return {}

    return {}


# ─────────────────────────────────────────────────────────────
# EXÉCUTION SÉCURISÉE
# ─────────────────────────────────────────────────────────────

def _execute_slide_code_v2(code: str, prs: Presentation, palette: dict) -> bool:
    """
    Exécute un bloc de code python-pptx dans le sandbox.
    Timeout 30s via threading. Retourne True si succès.
    """
    ok, reason = _validate_code_safety(code)
    if not ok:
        log.warning(f'[V2] Code rejeté (sécurité) : {reason}')
        return False

    result   = {'success': False, 'error': None}
    safe_ns  = _build_safe_namespace(prs, palette)

    def _run():
        try:
            exec(code, safe_ns)  # noqa: S102
            result['success'] = True
        except Exception as e:
            result['error'] = str(e)

    t = threading.Thread(target=_run, daemon=True)
    t.start()
    t.join(timeout=30)

    if t.is_alive():
        log.warning('[V2] exec() timeout 30s')
        return False
    if result['error']:
        log.warning(f"[V2] exec() error: {result['error']}")
        return False
    return True


def _inject_fallback_slide(prs: Presentation, slide_plan: dict, palette: dict):
    """Slide de fallback minimaliste si le code L2 échoue."""
    try:
        slide, W, H = _h2_blank_slide(prs)
        _h2_rect(slide, 0, 0, W, H, 'FFFFFF')
        _h2_rect(slide, 0, 0, 0.08, H, palette.get('primary', '1A3A6B'))
        key_msg = slide_plan.get('key_message', 'Contenu')
        _h2_text(slide, key_msg, 0.5, H/2-0.4, W-1.0, 0.8,
                 palette.get('font', 'Calibri'), 24,
                 palette.get('primary', '1A3A6B'), bold=True)
    except Exception as e:
        log.error(f'[V2] Fallback slide failed: {e}')


def _execute_all_codes_v2(
    code_map:    dict,
    plan_slides: list,
    prs:         Presentation,
    palette:     dict,
) -> int:
    """
    Exécute les codes dans l'ordre du plan narratif.
    Retourne le nombre de slides générées avec succès.
    """
    success = 0
    for sp in plan_slides:
        plan_idx = str(sp.get('plan_index', 0))
        code     = code_map.get(plan_idx, '')
        if not code:
            log.warning(f'[V2] Pas de code pour plan_index={plan_idx}')
            _inject_fallback_slide(prs, sp, palette)
            continue
        ok = _execute_slide_code_v2(code, prs, palette)
        if ok:
            success += 1
        else:
            _inject_fallback_slide(prs, sp, palette)
    return success


def _remove_original_slides_v2(prs: Presentation, n_original: int):
    """
    Supprime les n_original premières slides du template.
    Ne conserve que les slides générées par Level 2 (ajoutées après les originales).
    """
    xml_slides  = prs.slides._sldIdLst
    all_sld_ids = list(xml_slides)
    new_sld_ids = all_sld_ids[n_original:]

    if not new_sld_ids:
        log.warning('[V2] Aucune slide L2 générée — conservation des slides originales.')
        return

    log.info(f'[V2] Suppression des {n_original} slides originales, '
             f'conservation de {len(new_sld_ids)} slides L2.')

    # Reconstruire sldIdLst avec uniquement les slides L2
    for sld in list(xml_slides):
        xml_slides.remove(sld)
    for sld in new_sld_ids:
        xml_slides.append(sld)


# ─────────────────────────────────────────────────────────────
# PIPELINE NIVEAU 2
# ─────────────────────────────────────────────────────────────

def run_pipeline_v2(
    pptx_bytes: bytes,
    prompt:     str,
    nb_slides:  int,
    profile:    str = 'institutional',
) -> tuple:
    """
    Pipeline Level 2 — 4 phases.
    """
    if not ANTHROPIC_API_KEY:
        raise ValueError('Clé API Claude manquante.')

    nb_slides = max(2, min(nb_slides, 30))
    profile   = profile if profile in CLIENT_PROFILES else 'institutional'

    prs        = Presentation(io.BytesIO(pptx_bytes))
    n_original = len(prs.slides)

    # ── Phase 1 : Analyse brand + extraction visuels ──────────
    log.info(f'[V2] Phase 1 : analyse brand (profil={profile}, {nb_slides} slides)...')
    brand      = extract_brand(prs)
    palette    = _h2_extract_palette(brand)
    library    = build_layout_library(prs)
    selection  = select_template_slides(library, nb_slides)

    thumbnail  = _extract_template_thumbnail(pptx_bytes)
    swatch     = _make_palette_swatch(palette)
    log.info(f'[V2] Palette: primary=#{palette["primary"]} font={palette["font"]} '
             f'thumbnail={"oui" if thumbnail else "non"} swatch={"oui" if swatch else "non"}')

    # ── Phase 2 : Planification narrative ─────────────────────
    log.info('[V2] Phase 2 : planification narrative...')
    plan = plan_presentation(prompt, nb_slides, selection, brand)

    plan_slides = plan.get('slides', [])
    while len(plan_slides) < nb_slides:
        fb = selection[min(len(plan_slides), len(selection) - 1)]
        plan_slides.append({
            'plan_index':           len(plan_slides),
            'template_slide_index': fb['slide_index'],
            'slide_type':           fb.get('slide_type', 'full_text'),
            'narrative_angle':      'Développement complémentaire',
            'key_message':          'Argument additionnel',
            'visual_hint':          '',
        })
    plan['slides'] = plan_slides[:nb_slides]

    # ── Phase 3 : Génération code python-pptx (avec vision) ───
    log.info('[V2] Phase 3 : génération code python-pptx (Claude avec vision)...')
    code_map = generate_codes_v2(
        prompt, plan, palette, brand, profile,
        template_thumbnail = thumbnail,
        palette_swatch     = swatch,
    )

    # ── Phase 4 : Exécution + assemblage ──────────────────────
    log.info('[V2] Phase 4 : exécution des codes...')
    success_count = _execute_all_codes_v2(code_map, plan['slides'], prs, palette)
    log.info(f'[V2] {success_count}/{nb_slides} slides générées avec succès.')

    # Fallback L1 si taux < 50%
    success_rate = success_count / max(nb_slides, 1)
    if success_rate < 0.5:
        log.warning(f'[V2] Taux d\'échec élevé ({success_rate:.0%}) → fallback Level 1')
        final_bytes_l1, plan_l1, brand_l1 = run_pipeline(pptx_bytes, prompt, nb_slides)
        return final_bytes_l1, plan_l1, brand_l1, palette

    _remove_original_slides_v2(prs, n_original)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out.read(), plan, brand, palette


# ══════════════════════════════════════════════════════════════
# PIPELINE V3 — Layouts pré-testés
# ══════════════════════════════════════════════════════════════

_V3_PLANNER_SYSTEM = """\
Tu es un consultant senior spécialisé en communication visuelle et stratégie.
Tu crées des présentations PowerPoint de niveau agence — niveau Gamma.app ou McKinsey.

PRINCIPES ÉDITORIAUX (non négociables) :
1. Une seule idée forte par slide. Jamais deux messages.
2. Les titres sont des assertions, pas des thèmes. Pas "Contexte marché" mais "Le marché a doublé en 3 ans".
3. Les chiffres sont réels, précis et sourcés. Pas "beaucoup" mais "47 %".
4. Le contenu est celui d'un expert du domaine, pas un résumé Wikipedia.
5. Zéro phrase passive, zéro jargon vide. Chaque mot compte.

FORMAT : retourne UNIQUEMENT du JSON valide, sans markdown ni commentaire.\
"""

_V3_PLANNER_USER = """\
Crée une présentation sur : {prompt}
Nombre de slides : {nb_slides}

Charte graphique :
- Couleur primaire : #{primary}
- Couleur accent   : #{accent}
- Police           : {font}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
LAYOUTS DISPONIBLES + STRUCTURE JSON EXACTE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

cover_dark / cover_split  →  {{"title":"...", "subtitle":"...", "footer":"..."}}
section                   →  {{"number":"01", "title":"..."}}
kpi_grid                  →  {{"title":"...", "kpis":[{{"value":"47 %","label":"Part de marché","sublabel":"2023, source XYZ"}}], "footer":"..."}}
kpi_row                   →  {{"title":"...", "kpis":[{{"value":"3,2 Md€","label":"CA 2023","sublabel":"vs 2,1 Md€ en 2022"}}], "footer":"..."}}
timeline_h                →  {{"title":"...", "steps":[{{"date":"2021","title":"Lancement","body":"Déploiement initial dans 3 pays"}}], "footer":"..."}}
two_col                   →  {{"title":"...", "col_a":{{"title":"POUR","items":["Argument 1","Argument 2"]}}, "col_b":{{"title":"CONTRE","items":["Limite 1","Limite 2"]}}, "footer":"..."}}
quote_dark                →  {{"quote":"Citation percutante ≤ 20 mots", "author":"Prénom NOM, Titre — 2024", "footer":"..."}}
list_numbered             →  {{"title":"...", "items":[{{"title":"Levier 1","body":"Explication concise en 15 mots max."}}], "footer":"..."}}
list_cards                →  {{"title":"...", "cards":[{{"title":"Axe 1","body":"Description en 20 mots max."}}], "footer":"..."}}
image_split               →  {{"title":"...", "points":["Point 1 en 15 mots","Point 2"], "footer":"..."}}
full_text                 →  {{"title":"...", "paragraphs":["Paragraphe 1 (≤ 40 mots)","Paragraphe 2"], "footer":"..."}}
stat_hero                 →  {{"value":"€ 2,3 Md", "label":"Montant des pertes évitées", "context":"Estimation sur 5 ans — rapport Ernst & Young 2023", "footer":"..."}}
closing_dark / closing_split → {{"title":"Merci", "subtitle":"Sources : ... · Contact : ..."}}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STRUCTURE DE RÉPONSE ATTENDUE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{{
  "presentation_title": "Titre accrocheur ≤ 8 mots",
  "footer_text": "Nom entreprise · Sujet · Année",
  "slides": [
    {{"layout": "cover_dark", "content": {{"title": "...", "subtitle": "...", "footer": "..."}}}},
    {{"layout": "kpi_grid",   "content": {{"title": "...", "kpis": [...], "footer": "..."}}}},
    ...
    {{"layout": "closing_dark", "content": {{"title": "Merci", "subtitle": "..."}}}}
  ]
}}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
RÈGLES IMPÉRATIVES
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
- Slide 1 : TOUJOURS cover_dark ou cover_split
- Dernière slide : TOUJOURS closing_dark ou closing_split
- Jamais deux layouts identiques consécutifs
- Alterner fonds sombres (cover/section/quote/kpi_grid) et clairs (list/two_col/full_text/stat_hero)
- kpi_grid : 4 à 6 KPIs avec valeurs chiffrées réelles
- kpi_row : 3 à 4 KPIs maximum
- timeline_h : 4 à 5 jalons avec dates réelles
- list_* : 3 à 5 items maximum
- Titres : ≤ 8 mots, assertifs, jamais nominaux
- Bodies : ≤ 25 mots, concis, factuels
- footer_text : identique sur toutes les slides sauf cover et closing
- Contenu expert : chiffres précis, sources citées, angle analytique
\
"""


def plan_presentation_v3(prompt: str, nb_slides: int, palette: dict) -> dict:
    """
    Planifie une présentation via layouts pré-testés.
    Claude retourne uniquement du JSON : [{layout, content}, ...].
    """
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    layouts_block = '\n'.join(
        f'- {name}: {desc}'
        for name, desc in LAYOUT_DESCRIPTIONS.items()
    )

    user = _V3_PLANNER_USER.format(
        prompt        = prompt,
        nb_slides     = nb_slides,
        primary       = palette.get('primary', '1A3A6B'),
        accent        = palette.get('accent',  'F0A500'),
        font          = palette.get('font',    'Calibri'),
        layouts_block = layouts_block,
    )

    max_tokens = max(4000, nb_slides * 380)

    for attempt in range(3):
        msg = client.messages.create(
            model      = CLAUDE_MODEL,
            max_tokens = max_tokens,
            system     = _V3_PLANNER_SYSTEM,
            messages   = [{'role': 'user', 'content': user}],
        )
        try:
            plan = _parse_json_robust(msg.content[0].text.strip(), context='plan_v3')
            slides = plan.get('slides', [])
            log.info(f'[V3] Plan: {len(slides)} slides, title="{plan.get("presentation_title","")[:60]}"')
            return plan
        except (ValueError, KeyError) as e:
            log.warning(f'plan_presentation_v3 attempt {attempt+1}/3 failed: {e}')
            if attempt == 2:
                raise
    raise RuntimeError('plan_presentation_v3 : 3 tentatives échouées')


def run_pipeline_v3(pptx_bytes: bytes, prompt: str, nb_slides: int) -> tuple:
    """
    Pipeline V3 — layouts pré-testés, zéro génération de code.

    Phase 1 : extraction charte + palette
    Phase 2 : planification narrative par Claude (JSON layouts)
    Phase 3 : application des fonctions de layout pré-testées
    Phase 4 : suppression des slides originales du template
    """
    if not ANTHROPIC_API_KEY:
        raise ValueError('Clé API Claude manquante.')

    nb_slides = max(2, min(nb_slides, 30))

    prs        = Presentation(io.BytesIO(pptx_bytes))
    n_original = len(prs.slides)

    # ── Phase 1 ───────────────────────────────────────────────
    log.info(f'[V3] Phase 1 : extraction charte ({nb_slides} slides)...')
    brand   = extract_brand(prs)
    palette = _h2_extract_palette(brand)
    log.info(f'[V3] Palette : primary=#{palette["primary"]} accent=#{palette["accent"]} font={palette["font"]}')

    # ── Phase 2 ───────────────────────────────────────────────
    log.info('[V3] Phase 2 : planification narrative...')
    plan = plan_presentation_v3(prompt, nb_slides, palette)
    log.info(f'[V3] Plan reçu : {json.dumps(plan, ensure_ascii=False)[:1000]}')

    slides_plan = plan.get('slides', [])

    # Compléter si Claude en a généré moins que demandé
    fallback_layouts = ['full_text', 'list_numbered', 'two_col', 'kpi_row']
    while len(slides_plan) < nb_slides:
        fb_name = fallback_layouts[len(slides_plan) % len(fallback_layouts)]
        slides_plan.append({
            'layout':  fb_name,
            'content': {
                'title':      'Développement complémentaire',
                'paragraphs': ['Contenu additionnel à personnaliser.'],
                'footer':     plan.get('footer_text', ''),
            },
        })
    slides_plan = slides_plan[:nb_slides]

    # Garantie : la dernière slide est TOUJOURS une closing
    closing_layouts = {'closing_dark', 'closing_split'}
    if not slides_plan or slides_plan[-1].get('layout') not in closing_layouts:
        closing_slide = {
            'layout': 'closing_dark',
            'content': {
                'title':    'Merci',
                'subtitle': plan.get('footer_text', ''),
            },
        }
        if len(slides_plan) >= nb_slides and slides_plan:
            slides_plan[-1] = closing_slide   # remplace la dernière si plan complet
        else:
            slides_plan.append(closing_slide)  # ajoute sinon

    log.info(f'[V3] {len(slides_plan)} slides à générer : {[s.get("layout") for s in slides_plan]}')

    # ── Phase 3 ───────────────────────────────────────────────
    log.info('[V3] Phase 3 : application des layouts...')
    success = 0
    slides_before = len(prs.slides)
    for sp in slides_plan:
        layout_name = sp.get('layout', 'full_text')
        content     = sp.get('content', {})

        # Injecter footer_text global si absent du content
        if not content.get('footer') and plan.get('footer_text'):
            content['footer'] = plan['footer_text']

        layout_fn = LAYOUT_REGISTRY.get(layout_name) or LAYOUT_REGISTRY['full_text']
        log.info(f'[V3] → layout "{layout_name}" (total slides avant: {len(prs.slides)})')
        try:
            layout_fn(prs, content, palette)
            success += 1
            log.info(f'[V3] ✓ "{layout_name}" OK — total slides: {len(prs.slides)}')
        except Exception as e:
            log.error(f'[V3] ✗ "{layout_name}" ÉCHOUÉ : {repr(e)}', exc_info=True)
            try:
                LAYOUT_REGISTRY['full_text'](prs, {
                    'title':      content.get('title', ''),
                    'paragraphs': [],
                    'footer':     content.get('footer', ''),
                }, palette)
                success += 1
                log.info(f'[V3] ✓ fallback full_text OK — total slides: {len(prs.slides)}')
            except Exception as e2:
                log.error(f'[V3] ✗ fallback full_text AUSSI ÉCHOUÉ : {repr(e2)}', exc_info=True)

    log.info(f'[V3] Phase 3 terminée : {success}/{nb_slides} OK — '
             f'{len(prs.slides) - slides_before} slides ajoutées au total')

    # ── Phase 4 ───────────────────────────────────────────────
    slides_added = len(prs.slides) - n_original
    log.info(f'[V3] Phase 4 : {len(prs.slides)} slides totales '
             f'({n_original} originales + {slides_added} nouvelles)')

    if slides_added == 0:
        raise RuntimeError(
            f'[V3] Aucune slide créée (success={success}) — fallback L1 déclenché'
        )

    xml_slides = prs.slides._sldIdLst
    to_remove  = list(prs.slides._sldIdLst)[:n_original]
    for sld_id in to_remove:
        xml_slides.remove(sld_id)
    log.info(f'[V3] {n_original} slides originales supprimées — '
             f'{len(prs.slides)} slides finales')

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out.read(), plan, brand, palette


# ══════════════════════════════════════════════════════════════
# ROUTES NIVEAU 2
# ══════════════════════════════════════════════════════════════

@app.get('/profiles')
def get_profiles():
    """Retourne les profils clients disponibles pour le Niveau 2."""
    return {
        key: {'label': p['label'], 'description': p['description']}
        for key, p in CLIENT_PROFILES.items()
    }


@app.post('/generate-v2')
async def generate_presentation_v2(
    request:       Request,
    template:      UploadFile = File(...),
    prompt:        str        = Form(...),
    nb_slides:     str        = Form(default='complet'),
    profile:       str        = Form(default='institutional'),
    authorization: str        = Form(default=None),
):
    """
    Route V3 — pipeline layouts pré-testés (fiable, zéro génération de code).
    nb_slides accepte : "Essentiel"→6 | "Complet"→10 | "Approfondi"→16 ou un entier.
    profile conservé pour compatibilité API future.
    """
    if not _is_pro(authorization):
        _quota(_ip(request))

    n = _resolve_nb_slides(nb_slides)
    pptx_bytes = await template.read()

    try:
        final_bytes, plan, brand, palette = run_pipeline_v3(pptx_bytes, prompt, n)
    except Exception as e:
        log.error(f'[V3] Erreur pipeline : {e}', exc_info=True)
        raise HTTPException(500, f'Erreur génération V3 : {e}')

    filename = f'visualcortex-{_safe_name(prompt)}.pptx'
    return StreamingResponse(
        io.BytesIO(final_bytes),
        media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation',
        headers={
            'Content-Disposition': f'attachment; filename={filename}',
            'Content-Length':      str(len(final_bytes)),
        },
    )


@app.post('/generate-v2-preview')
async def generate_v2_preview(
    request:       Request,
    template:      UploadFile = File(...),
    prompt:        str        = Form(...),
    nb_slides:     str        = Form(default='complet'),
    profile:       str        = Form(default='institutional'),
    authorization: str        = Form(default=None),
):
    """Aperçu plan + palette L2 (sans générer le fichier)."""
    pptx_bytes = await template.read()
    prs        = Presentation(io.BytesIO(pptx_bytes))
    brand      = extract_brand(prs)
    palette    = _h2_extract_palette(brand)
    library    = build_layout_library(prs)
    n          = _resolve_nb_slides(nb_slides)
    sel        = select_template_slides(library, n)

    pro = _is_pro(authorization)
    quota_info = (
        {'plan': 'pro'} if pro
        else {'used': _quota(_ip(request))[0], 'total': FREE_QUOTA_PER_IP, 'plan': 'free'}
    )

    plan = plan_presentation(prompt, n, sel, brand)
    return {
        'success':            True,
        'level':              2,
        'profile':            profile,
        'profile_label':      CLIENT_PROFILES.get(profile, {}).get('label', profile),
        'nb_slides_resolved': n,
        'presentation_title': plan.get('presentation_title', prompt[:60]),
        'narrative_arc':      plan.get('narrative_arc', ''),
        'palette':            palette,
        'slides': [
            {
                'index':           s.get('plan_index'),
                'type':            s.get('slide_type'),
                'narrative_angle': s.get('narrative_angle'),
                'key_message':     s.get('key_message'),
            }
            for s in plan.get('slides', [])
        ],
        'quota': quota_info,
    }


# ══════════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════════

if __name__ == '__main__':
    uvicorn.run(
        'main:app',
        host  = '0.0.0.0',
        port  = int(os.environ.get('PORT', 8000)),
        reload = False,
    )
