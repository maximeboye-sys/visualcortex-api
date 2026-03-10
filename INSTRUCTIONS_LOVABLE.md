# Instructions Lovable — Visual Cortex
# Connecter le front à l'API PPTX Generator (modèle freemium)

---

## CONTEXTE IMPORTANT

La clé API Claude est stockée côté serveur (Railway).
L'utilisateur ne la voit jamais. Comme WeTransfer, il utilise l'app sans se soucier de la technique.

Modèle freemium :
- Gratuit : 3 présentations/jour
- Pro : illimité (token envoyé en header Authorization)

---

## PROMPT À DONNER À LOVABLE

```
Je veux mettre à jour l'application Visual Cortex pour connecter le générateur de présentation à une API externe.

FONCTIONNEMENT :
L'utilisateur n'a pas besoin de clé API. Il suffit de :
1. Uploader un fichier .pptx (template de sa charte d'entreprise)
2. Saisir un prompt décrivant la présentation souhaitée
3. Choisir le nombre de slides avec un slider (entre 5 et 20, défaut = 8)
4. Cliquer sur "Générer" pour télécharger le fichier .pptx

FLUX EN 2 ÉTAPES :
- Étape 1 : Bouton "Voir le plan" → affiche les titres des slides générés avant de créer le fichier
- Étape 2 : Bouton "Générer la présentation" → télécharge le .pptx final

GESTION DU QUOTA GRATUIT :
- Afficher un compteur discret : "2/3 générations gratuites utilisées aujourd'hui"
- Si quota dépassé (erreur 429), afficher un message : "Limite atteinte — Passez en Pro pour un accès illimité"
- Prévoir un bouton "Passer en Pro" (lien vers page pricing, à créer plus tard)

MESSAGES D'ÉTAT :
- Upload template → "Analyse de la charte en cours..."
- Génération plan → "L'IA prépare votre présentation..."
- Génération fichier → "Construction du fichier PowerPoint..."
- Succès → "Votre présentation est prête !" + téléchargement automatique

---

VOICI LE CODE JAVASCRIPT À UTILISER :

const API_BASE = "REMPLACER_PAR_TON_URL_RAILWAY";

// Analyser le template uploadé (gratuit, sans quota)
async function analyzeTemplate(file) {
  const formData = new FormData();
  formData.append("file", file);
  const res = await fetch(`${API_BASE}/analyze-template`, {
    method: "POST",
    body: formData,
  });
  return await res.json();
  // Retourne : { brand: { fonts, theme_colors, layouts, slide_count } }
}

// Générer le plan (titres des slides)
async function generatePreview(templateFile, prompt, nbSlides, proToken = null) {
  const formData = new FormData();
  formData.append("template", templateFile);
  formData.append("prompt", prompt);
  formData.append("nb_slides", nbSlides.toString());
  if (proToken) formData.append("authorization", `Bearer ${proToken}`);

  const res = await fetch(`${API_BASE}/generate-preview`, {
    method: "POST",
    body: formData,
  });

  if (res.status === 429) {
    const err = await res.json();
    throw new Error(err.detail?.message || "Quota dépassé");
  }
  if (!res.ok) {
    const err = await res.json();
    throw new Error(err.detail || "Erreur lors de la génération");
  }
  return await res.json();
  // Retourne : { title, slides: [{index, type, title, summary}], quota: {used, max, plan} }
}

// Générer le fichier .pptx complet
async function generatePresentation(templateFile, prompt, nbSlides, proToken = null) {
  const formData = new FormData();
  formData.append("template", templateFile);
  formData.append("prompt", prompt);
  formData.append("nb_slides", nbSlides.toString());
  if (proToken) formData.append("authorization", `Bearer ${proToken}`);

  const res = await fetch(`${API_BASE}/generate`, {
    method: "POST",
    body: formData,
  });

  if (res.status === 429) {
    const err = await res.json();
    throw new Error(err.detail?.message || "Quota dépassé");
  }
  if (!res.ok) {
    const err = await res.json();
    throw new Error(err.detail || "Erreur lors de la génération");
  }

  // Téléchargement automatique du fichier
  const blob = await res.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "presentation-visual-cortex.pptx";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
```

---

INTERFACE SOUHAITÉE :
- Zone de drop pour uploader le .pptx (avec aperçu du nom de fichier une fois uploadé)
- Affichage discret de la charte détectée après upload : "Charte détectée : Nunito, Roboto • 4 couleurs"
- Grand champ textarea pour le prompt avec placeholder "Ex : Présentation commerciale pour notre offre de formation Q1 2025..."
- Slider horizontal pour le nombre de slides (5 à 20)
- Deux boutons d'action : "Voir le plan" (secondaire) et "Générer" (principal, couleur accent)
- Zone de résultat : liste des titres de slides si on a cliqué "Voir le plan"
- Compteur de quota discret en bas de page
```

---

## APRÈS DÉPLOIEMENT SUR RAILWAY

Remplace `REMPLACER_PAR_TON_URL_RAILWAY` par ton URL Railway.
Format : `https://visualcortex-api-production-xxxx.up.railway.app`

---

## VARIABLES D'ENVIRONNEMENT À CONFIGURER SUR RAILWAY

| Variable | Valeur |
|----------|--------|
| `ANTHROPIC_API_KEY` | ta clé Claude (sk-ant-...) |
| `PRO_SECRET_TOKEN` | un mot de passe secret pour les comptes pro (ex: `vc-pro-2025-xxxx`) |
| `FREE_QUOTA_PER_IP` | `3` (modifiable selon ta stratégie) |

---

## FLUX UTILISATEUR FINAL

```
[1] Upload .pptx  →  Analyse automatique de la charte (2 sec)
[2] Saisie du prompt + choix nb slides
[3] "Voir le plan"  →  Aperçu des titres de slides
[4] "Générer"  →  Téléchargement automatique du .pptx
```
