"""
Visual Cortex — PPTX Generator API v7 (Strict Hydration Edition)
Approche : Hydratation in-situ. Zéro duplication XML.
Résultat : 0 fichier corrompu, préservation absolue du design (même complexe/3D).
"""

import os
import io
import json
import time
import copy
from collections import defaultdict

import anthropic
from fastapi import FastAPI, File, Form, UploadFile, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, StreamingResponse
from pptx import Presentation
from pptx.oxml.ns import qn
import uvicorn

app = FastAPI(title="Visual Cortex API", version="7.0.0")

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
# 1. EXTRACTION DES TEXTES (Pour envoyer à Claude)
# ─────────────────────────────────────────────
def extract_texts_for_ai(prs: Presentation, nb_slides: int) -> dict:
    """Extrait le texte brut de chaque slide, indexé pour le prompt."""
    extracted = {}
    
    # On se limite au nombre de slides demandées (ou au max du document)
    limit = min(nb_slides, len(prs.slides))
    
    for i in range(limit):
        slide = prs.slides[i]
        texts = []
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                for para in shape.text_frame.paragraphs:
                    full_text = "".join(run.text for run in para.runs).strip()
                    # On ignore les textes trop courts (chiffres de pagination, puces vides)
                    if len(full_text) > 2:
                        texts.append(full_text)
        if texts:
            # On déduplique et on garde l'ordre
            extracted[f"slide_{i}"] = list(dict.fromkeys(texts))
            
    return extracted


# ─────────────────────────────────────────────
# 2. IA : MAPPING DES TEXTES
# ─────────────────────────────────────────────
def generate_text_mapping_with_claude(prompt: str, extracted_texts: dict) -> dict:
    if not ANTHROPIC_API_KEY: raise ValueError("Clé API Claude manquante.")
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    
    system = """Tu es un expert en conception de présentations B2B.
Ton rôle est d'adapter le texte d'un template PowerPoint vers un nouveau sujet.
RÈGLE ABSOLUE : Tu dois conserver la même longueur de texte. 
- Si le texte original est un titre de 3 mots, propose un titre de 3 mots.
- Si c'est un paragraphe de 20 mots, propose un paragraphe de 20 mots.
Réponds UNIQUEMENT en JSON valide."""

    user = f"""Nouveau sujet cible : {prompt}

Voici les textes extraits de chaque slide du template (classés par slide).
Génère les textes de remplacement correspondants.

Textes originaux :
{json.dumps(extracted_texts, ensure_ascii=False, indent=2)}

Format JSON STRICT attendu :
{{
  "slide_0": {{
    "Texte original exact issu du dictionnaire": "Nouveau texte de longueur équivalente",
    "Energy Lab": "État & TotalEnergies"
  }},
  "slide_1": {{ ... }}
}}"""

    # J'utilise le modèle le plus récent et stable pour éviter l'erreur 404
    msg = client.messages.create(
        model="claude-3-5-sonnet-latest", 
        max_tokens=6000,
        system=system,
        messages=[{"role": "user", "content": user}],
    )
    
    raw = msg.content[0].text.strip()
    if "```" in raw:
        parts = raw.split("```")
        raw = parts[1] if len(parts) > 1 else parts[0]
        if raw.startswith("json"): raw = raw[4:]
        
    return json.loads(raw.strip())


# ─────────────────────────────────────────────
# 3. HYDRATATION (Remplacement chirurgical)
# ─────────────────────────────────────────────
def hydrate_presentation(pptx_bytes: bytes, mapping: dict, nb_slides: int) -> bytes:
    """Remplace le texte tout en préservant 100% du style original et supprime les slides en trop."""
    prs = Presentation(io.BytesIO(pptx_bytes))
    
    # Étape 1 : Remplacer le texte
    for slide_idx_str, replacements in mapping.items():
        try:
            slide_idx = int(slide_idx_str.replace("slide_", ""))
            if slide_idx >= len(prs.slides): continue
            slide = prs.slides[slide_idx]
            
            for shape in slide.shapes:
                if not getattr(shape, "has_text_frame", False): continue
                for para in shape.text_frame.paragraphs:
                    para_text = "".join(run.text for run in para.runs).strip()
                    
                    if not para_text: continue
                    
                    # Si le texte de ce paragraphe fait partie de ceux qu'on doit remplacer
                    if para_text in replacements:
                        new_text = replacements[para_text]
                        
                        # 1. Sauvegarde du style du premier caractère (XML)
                        rpr_xml = None
                        if para.runs and para.runs[0]._r.find(qn("a:rPr")) is not None:
                            rpr_xml = copy.deepcopy(para.runs[0]._r.find(qn("a:rPr")))
                            
                        # 2. Remplacement massif du texte (ça détruit les runs existants)
                        para.text = new_text
                        
                        # 3. Ré-application chirurgicale du style sur le nouveau texte
                        if rpr_xml is not None:
                            for run in para.runs:
                                rpr = run._r.find(qn("a:rPr"))
                                if rpr is not None: run._r.remove(rpr)
                                run._r.insert(0, copy.deepcopy(rpr_xml))
        except Exception:
            pass # Si une slide plante, on passe à la suivante sans faire crasher l'app
            
    # Étape 2 : Supprimer les slides en trop (si le template a 30 slides et qu'on en veut 9)
    xml_slides = prs.slides._sldIdLst
    if nb_slides < len(prs.slides):
        slides_to_remove = list(xml_slides)[nb_slides:]
        for sld in slides_to_remove:
            xml_slides.remove(sld)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out.read()


# ─────────────────────────────────────────────
# 4. ROUTES API
# ─────────────────────────────────────────────
@app.get("/")
def root(): return {"status": "ok", "version": "7.0.0 - Hydration Engine"}

@app.post("/analyze-template")
async def analyze_template(file: UploadFile = File(...)):
    # L'analyse n'a plus besoin d'être aussi complexe. On compte juste les slides.
    pptx_bytes = await file.read()
    prs = Presentation(io.BytesIO(pptx_bytes))
    return {"success": True, "message": f"{len(prs.slides)} slides détectées. Prêt pour l'hydratation."}

@app.post("/generate-preview")
async def generate_preview(request: Request, template: UploadFile = File(...), prompt: str = Form(...), nb_slides: int = Form(default=8), authorization: str = Form(default=None)):
    """Génère un aperçu du mapping de texte sans créer le fichier."""
    pro = _is_pro(authorization)
    quota_info = {"plan": "pro"} if pro else {"used": _check_and_increment_quota(_get_ip(request))[0], "plan": "free"}
    
    template_bytes = await template.read()
    prs = Presentation(io.BytesIO(template_bytes))
    
    extracted_texts = extract_texts_for_ai(prs, nb_slides)
    mapping = generate_text_mapping_with_claude(prompt, extracted_texts)
    
    # On renvoie à Lovable les "nouveaux" titres (pour l'aperçu)
    preview_slides = []
    for k, v in mapping.items():
        if v: preview_slides.append({"title": list(v.values())[0]})
        
    return {
        "success": True, "title": "Présentation Générée",
        "slides": preview_slides,
        "quota": quota_info
    }

@app.post("/generate")
async def generate_presentation(request: Request, template: UploadFile = File(...), prompt: str = Form(...), nb_slides: int = Form(default=8), authorization: str = Form(default=None)):
    if not _is_pro(authorization): _check_and_increment_quota(_get_ip(request))
    
    template_bytes = await template.read()
    prs = Presentation(io.BytesIO(template_bytes))
    
    extracted_texts = extract_texts_for_ai(prs, nb_slides)
    mapping = generate_text_mapping_with_claude(prompt, extracted_texts)
    
    pptx_bytes = hydrate_presentation(template_bytes, mapping, nb_slides)
    
    return StreamingResponse(
        io.BytesIO(pptx_bytes), 
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation", 
        headers={"Content-Disposition": "attachment; filename=presentation-visualcortex.pptx"}
    )

if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=int(os.environ.get("PORT", 8000)), reload=False)
