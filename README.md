# Présentation IA — V4 Groq + Kokoro

Cette version ne dépend plus de Gemini.

## Architecture

```text
PowerPoint
   ↓
python-pptx
   ↓
Groq API
   ↓
Narration
   ↓
Kokoro-82M (sur le serveur Streamlit)
   ↓
Audio WAV
   ↓
LibreOffice + FFmpeg
   ↓
MP4
```

## Clé Groq

Créer une clé dans Groq Console.

Pour un test local, tu peux la saisir directement dans l'interface.

Pour Streamlit Community Cloud, ajouter dans **Secrets** :

```toml
GROQ_API_KEY = "ta_cle_groq"
```

Ne mets pas la clé dans GitHub.

## Déploiement Streamlit Community Cloud

Mettre dans GitHub :

- `app.py`
- `requirements.txt`
- `packages.txt`
- `.gitignore`
- `README.md`

`packages.txt` installe :

- LibreOffice
- FFmpeg
- espeak-ng
- libsndfile

Kokoro télécharge ses poids lors du premier chargement du modèle.

## Voix française

Kokoro-82M v1.0 ne propose actuellement qu'une voix française officielle : `ff_siwis`.

## Remarque importante

Streamlit Community Cloud a des ressources limitées.
Kokoro est un modèle léger (82M paramètres), mais le premier chargement peut prendre du temps.
Si le déploiement devient trop lourd ou instable, il faudra envisager un hébergement plus généreux ou un moteur TTS externe.

## Modèles Groq proposés

- `qwen/qwen3.6-27b` — choix par défaut
- `llama-3.3-70b-versatile`
- `llama-3.1-8b-instant`

Les modèles disponibles chez Groq peuvent évoluer. Si Groq retire un modèle, remplace son identifiant dans `TEXT_MODELS`.
