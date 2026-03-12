"""
Visual Cortex — PPTX Generator API v2.1
Approche robuste : duplique les slides du template au lieu de les recréer.
"""

import os, io, json, zipfile, time, copy, re
from collections import defaultdict
from lxml import etree

import anthropic
from fastapi import FastAPI, File, Form, UploadFile, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pptx import Presentation
from pptx.util import Pt
import uvicorn

app = FastAPI(title="Visual Cortex API", version="2.1.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=False, allow_methods=["*"], allow_headers=["*"])

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
PRO_SECRET_TOKEN  = os.environ.get("PRO_SECRET_TOKEN", "change-me")
FREE_QUOTA_PER_IP = int(os.environ.get("FREE_QUOTA_PER_IP", "3"))
_usage: dict = defaultdict(list)
DAY = 86400

def _get_ip(req: Request) -> str:
    fwd = req.headers.get("x-forwarded-for")
    return fwd.split(",")[0].strip() if fwd else req.client.host

def _is_pro(auth: str | None) -> bool:
    return bool(auth and auth.replace("Bearer ", "").strip() == PRO_SECRET_TOKEN)

def _check_quota(ip: str) -> tuple[int, int]:
    now = time.time()
    _usage[ip] = [t for t in _usage[ip] if now - t < DAY]
    used = len(_usage[ip])
    if used >= FREE_QUOTA_PER_IP:
        raise HTTPException(429, {"error": "quota_exceeded",
            "message": f"Limite gratuite atteinte ({FREE_QUOTA_PER_IP}/jour). Passez en Pro pour un accès illimité.",
            "used": used, "max": FREE_QUOTA_PER_IP})
    _usage[ip].append(now)
    return used + 1, FREE_QUOTA_PER_IP

# ── EXTRACTION CHARTE ──────────────────────────────────────────────────────────

def extract_brand(pptx_bytes: bytes) -> dict:
    prs = Presentation(io.BytesIO(pptx_bytes))
    fonts, colors, layouts, texts = set(), set(), [], []
    for slide in prs.slides:
        layouts.append(slide.slide_layout.name)
        sc = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if run.font.name: fonts.add(run.font.name)
                        try:
                            if run.font.color and run.font.color.type:
                                colors.add(str(run.font.color.rgb))
                        except: pass
                        if run.text.strip(): sc.append(run.text.strip())
        if sc: texts.append(" | ".join(sc[:5]))
    return {
        "fonts": list(fonts)[:5],
        "colors": list(colors)[:10],
        "theme_colors": _theme_colors(pptx_bytes),
        "layouts": list(dict.fromkeys(layouts)),
        "slide_count": len(prs.slides),
        "sample_texts": texts,
    }

def _theme_colors(pptx_bytes: bytes) -> list:
    try:
        with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
            tf = [f for f in z.namelist() if "theme/theme" in f]
            if tf:
                xml = z.read(tf[0]).decode("utf-8")
                return list(dict.fromkeys(re.findall(r'val="([0-9A-Fa-f]{6})"', xml)))[:8]
    except: pass
    return []

# ── GÉNÉRATION CLAUDE ──────────────────────────────────────────────────────────

def generate_with_claude(prompt: str, brand: dict, nb_slides: int) -> dict:
    if not ANTHROPIC_API_KEY:
        raise HTTPException(500, "Clé API Claude non configurée.")
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    system = "Tu es expert en communication B2B. Réponds UNIQUEMENT en JSON valide, sans backticks."
    user = f"""Génère une présentation PowerPoint professionnelle.

DEMANDE : {prompt}

CHARTE :
- Polices : {', '.join(brand.get('fonts', ['Arial']))}
- Couleurs : {', '.join(brand.get('theme_colors', [])[:5])}
- Exemples de textes : {' | '.join(brand.get('sample_texts', [])[:3])}

CONTRAINTES :
- Exactement {nb_slides} slides
- Slide 1 = couverture (type: cover), slide {nb_slides} = conclusion (type: conclusion)
- Langage B2B, concis, orienté action
- Maximum 5 points par slide

JSON attendu :
{{
  "title": "Titre",
  "slides": [
    {{"index": 1, "type": "cover", "title": "...", "subtitle": "...", "notes": "..."}},
    {{"index": 2, "type": "content", "title": "...", "body": ["point1", "point2"], "notes": "..."}}
  ]
}}"""
    msg = client.messages.create(model="claude-sonnet-4-20250514", max_tokens=4000,
        system=system, messages=[{"role": "user", "content": user}])
    raw = msg.content[0].text.strip().lstrip("```json").lstrip("```").rstrip("```").strip()
    return json.loads(raw)

# ── CONSTRUCTION PPTX ROBUSTE ──────────────────────────────────────────────────

NSMAP = "http://schemas.openxmlformats.org/drawingml/2006/main"

def _set_text_in_shape(shape, text_lines: list[str]):
    """Remplace le texte d'un shape en préservant le formatage existant."""
    if not shape.has_text_frame or not text_lines:
        return
    tf = shape.text_frame
    # Préserve le format du premier paragraphe/run existant
    try:
        first_para = tf.paragraphs[0]
        first_run_xml = None
        if first_para.runs:
            first_run_xml = copy.deepcopy(first_para.runs[0]._r)
    except:
        first_run_xml = None

    tf.clear()

    for i, line in enumerate(text_lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        run = p.add_run()
        run.text = line
        # Réapplique le format si disponible
        if first_run_xml is not None:
            try:
                rpr = first_run_xml.find(f"{{{NSMAP}}}rPr")
                if rpr is not None and run._r.find(f"{{{NSMAP}}}rPr") is None:
                    run._r.insert(0, copy.deepcopy(rpr))
            except:
                pass

def build_pptx(pptx_bytes: bytes, content: dict) -> bytes:
    prs = Presentation(io.BytesIO(pptx_bytes))
    slides_data = content.get("slides", [])
    original_slides = list(prs.slides)
    nb_orig = len(original_slides)

    if nb_orig == 0:
        raise HTTPException(500, "Le template ne contient aucune slide.")

    # Catégorise les slides du template
    def pick_template_slide(slide_type: str, index: int) -> object:
        """Choisit la slide template la plus adaptée au type demandé."""
        if slide_type == "cover" and nb_orig >= 1:
            return original_slides[0]
        if slide_type == "conclusion" and nb_orig >= 1:
            return original_slides[-1]
        if slide_type == "section" and nb_orig >= 3:
            return original_slides[2]
        # Pour content : alterne entre les slides du milieu
        mid_slides = original_slides[1:-1] if nb_orig > 2 else original_slides
        return mid_slides[index % len(mid_slides)]

    # Construit les nouvelles slides en dupliquant depuis le template
    from pptx.oxml.ns import qn
    from pptx.opc.part import Part
    import copy

    result_prs = Presentation(io.BytesIO(pptx_bytes))

    # Supprime toutes les slides existantes proprement
    slide_id_list = result_prs.slides._sldIdLst
    rId_list = []
    for sldId in list(slide_id_list):
        rId_list.append(sldId.get('r:id') or sldId.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'))
        slide_id_list.remove(sldId)

    # Recrée les slides depuis les layouts (approche stable)
    layout_map = {}
    for layout in result_prs.slide_layouts:
        n = layout.name.lower()
        if ("couverture" in n or "cover" in n) and "cover" not in layout_map:
            layout_map["cover"] = layout
        elif ("titre" in n or "title" in n) and "section" not in layout_map:
            layout_map["section"] = layout
        elif ("texte" in n or "content" in n or "bullet" in n or "titre" in n) and "content" not in layout_map:
            layout_map["content"] = layout

    available = list(result_prs.slide_layouts)
    layout_map.setdefault("cover", available[0])
    layout_map.setdefault("content", available[min(1, len(available)-1)])
    layout_map.setdefault("section", layout_map["content"])
    layout_map["conclusion"] = layout_map["cover"]

    for sd in slides_data:
        stype = sd.get("type", "content")
        layout = layout_map.get(stype, layout_map["content"])
        slide = result_prs.slides.add_slide(layout)

        # Titre
        title_text = sd.get("title", "")
        if slide.shapes.title and title_text:
            slide.shapes.title.text = title_text

        # Corps
        body = sd.get("body", [])
        subtitle = sd.get("subtitle", "")
        lines = body if body else ([subtitle] if subtitle else [])

        for ph in slide.placeholders:
            if ph.placeholder_format.idx == 1 and lines:
                tf = ph.text_frame
                tf.clear()
                tf.text = lines[0]
                for pt in lines[1:]:
                    p = tf.add_paragraph()
                    p.text = pt
                break

        # Notes
        notes = sd.get("notes", "")
        if notes:
            try:
                slide.notes_slide.notes_text_frame.text = notes
            except: pass

    out = io.BytesIO()
    result_prs.save(out)
    out.seek(0)
    return out.read()

# ── ROUTES ────────────────────────────────────────────────────────────────────

@app.get("/")
def root(): return {"status": "ok", "service": "Visual Cortex API 2.1 🚀"}

@app.get("/health")
def health(): return {"status": "healthy"}

@app.post("/analyze-template")
async def analyze_template(file: UploadFile = File(...)):
    if not file.filename.endswith(".pptx"):
        raise HTTPException(400, "Fichier .pptx requis")
    brand = extract_brand(await file.read())
    return {"success": True, "brand": brand,
            "message": f"{brand['slide_count']} slides • {len(brand['fonts'])} polices détectées"}

@app.post("/generate-preview")
async def generate_preview(request: Request, template: UploadFile = File(...),
    prompt: str = Form(...), nb_slides: int = Form(8), authorization: str = Form(None)):
    if not template.filename.endswith(".pptx"): raise HTTPException(400, "Fichier .pptx requis")
    if not 3 <= nb_slides <= 30: raise HTTPException(400, "Entre 3 et 30 slides")
    pro = _is_pro(authorization)
    quota = {"plan": "pro"} if pro else dict(zip(["used","max"], _check_quota(_get_ip(request))), plan="free")
    brand = extract_brand(await template.read())
    content = generate_with_claude(prompt, brand, nb_slides)
    return {"success": True, "title": content.get("title"),
        "slides": [{"index": s.get("index"), "type": s.get("type"), "title": s.get("title"),
            "summary": s.get("subtitle") or (s.get("body", [""])[0] if s.get("body") else "")}
            for s in content.get("slides", [])],
        "brand": {"fonts": brand["fonts"], "theme_colors": brand["theme_colors"], "layouts": brand["layouts"]},
        "quota": quota}

@app.post("/generate")
async def generate_presentation(request: Request, template: UploadFile = File(...),
    prompt: str = Form(...), nb_slides: int = Form(8), authorization: str = Form(None)):
    if not template.filename.endswith(".pptx"): raise HTTPException(400, "Fichier .pptx requis")
    if not 3 <= nb_slides <= 30: raise HTTPException(400, "Entre 3 et 30 slides")
    pro = _is_pro(authorization)
    if not pro: _check_quota(_get_ip(request))
    tbytes = await template.read()
    brand = extract_brand(tbytes)
    try:
        content = generate_with_claude(prompt, brand, nb_slides)
    except json.JSONDecodeError as e:
        raise HTTPException(500, f"Erreur parsing IA : {e}")
    try:
        pptx_bytes = build_pptx(tbytes, content)
    except Exception as e:
        raise HTTPException(500, f"Erreur PPTX : {e}")
    fname = "".join(c for c in content.get("title","pres")[:40].replace(" ","_") if c.isalnum() or c in "_-.") + ".pptx"
    return StreamingResponse(io.BytesIO(pptx_bytes),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f"attachment; filename={fname}"})

if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))
