
import io
import json
import re
import hashlib
import wave
import zipfile
import tempfile
import subprocess
import shutil
from pathlib import Path
from collections import Counter
from typing import List, Dict

import fitz
import numpy as np
import soundfile as sf
import streamlit as st
from pptx import Presentation
from groq import Groq
from kokoro import KPipeline


DEFAULT_IGNORED_PATTERNS = [
    r"^\s*Département\s+douane\s*$",
    r"^\s*Date\s*$",
    r"^\s*Titre de la présentation.*Émetteur\s*$",
    r"^\s*\d+\s*$",
]

TEXT_MODELS = [
    "qwen/qwen3.6-27b",
    "llama-3.3-70b-versatile",
    "llama-3.1-8b-instant",
]


def normalize_text(text: str) -> str:
    text = text.replace("\u00a0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def normalize_for_compare(text: str) -> str:
    text = normalize_text(text).lower()
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"[“”«»\"'’]", "", text)
    return text.strip(" .,:;-\n\t")


def is_ignored_block(text: str, custom_ignored: List[str]) -> bool:
    cleaned = normalize_text(text)
    if not cleaned:
        return True

    for pattern in DEFAULT_IGNORED_PATTERNS:
        if re.match(pattern, cleaned, flags=re.IGNORECASE):
            return True

    low = cleaned.lower()
    for item in custom_ignored:
        item = item.strip().lower()
        if item and item in low:
            return True

    return False


def extract_text_from_shape(shape) -> List[str]:
    blocks = []

    if hasattr(shape, "shapes"):
        for subshape in shape.shapes:
            blocks.extend(extract_text_from_shape(subshape))

    if getattr(shape, "has_text_frame", False):
        paragraphs = []
        for p in shape.text_frame.paragraphs:
            txt = normalize_text(p.text)
            if txt:
                paragraphs.append(txt)
        if paragraphs:
            blocks.append("\n".join(paragraphs))

    if getattr(shape, "has_table", False):
        rows = []
        for row in shape.table.rows:
            values = [normalize_text(cell.text) for cell in row.cells]
            values = [v for v in values if v]
            if values:
                rows.append(" | ".join(values))
        if rows:
            blocks.append("\n".join(rows))

    return blocks


def extract_presentation(uploaded_file) -> List[Dict]:
    prs = Presentation(io.BytesIO(uploaded_file.getvalue()))
    slides = []

    for idx, slide in enumerate(prs.slides, start=1):
        blocks = []
        for shape in slide.shapes:
            blocks.extend(extract_text_from_shape(shape))

        seen = set()
        unique_blocks = []
        for block in blocks:
            key = normalize_for_compare(block)
            if key and key not in seen:
                seen.add(key)
                unique_blocks.append(normalize_text(block))

        slides.append({
            "number": idx,
            "raw_blocks": unique_blocks,
        })

    return slides


def detect_repeated_blocks(slides: List[Dict], min_occurrences: int = 3) -> set:
    counter = Counter()

    for slide in slides:
        keys = set()
        for block in slide["raw_blocks"]:
            key = normalize_for_compare(block)
            if len(key) >= 70:
                keys.add(key)
        counter.update(keys)

    return {key for key, count in counter.items() if count >= min_occurrences}


def clean_slides(slides, custom_ignored, remove_repeated=True, min_occurrences=3):
    repeated = (
        detect_repeated_blocks(slides, min_occurrences)
        if remove_repeated else set()
    )

    result = []

    for slide in slides:
        kept, removed = [], []

        for block in slide["raw_blocks"]:
            key = normalize_for_compare(block)

            if is_ignored_block(block, custom_ignored):
                removed.append({
                    "text": block,
                    "reason": "gabarit / élément ignoré",
                })
            elif key in repeated:
                removed.append({
                    "text": block,
                    "reason": "bloc répété sur plusieurs slides",
                })
            else:
                kept.append(block)

        result.append({
            **slide,
            "clean_blocks": kept,
            "removed_blocks": removed,
            "clean_text": "\n\n".join(kept).strip(),
        })

    return result


def get_groq_key():
    try:
        return st.secrets["GROQ_API_KEY"]
    except Exception:
        return ""


def build_prompt(slide, previous_narration, next_slide_text, mode, target_seconds):
    words_min = max(35, int(target_seconds * 1.7))
    words_max = max(words_min + 15, int(target_seconds * 2.2))

    mode_instruction = {
        "Expliquer naturellement": (
            "Explique le contenu comme un formateur qui s'adresse oralement à des salariés. "
            "Ne lis pas simplement les puces : relie les idées et rends le propos fluide."
        ),
        "Résumer": (
            "Résume les idées essentielles de la slide de façon claire et orale. "
            "Supprime les détails secondaires sans inventer d'information."
        ),
        "Lecture reformulée": (
            "Reste très proche du contenu de la slide, mais reformule pour que cela sonne naturel à l'oral."
        ),
    }[mode]

    return f"""
Tu rédiges la narration orale d'une présentation professionnelle destinée à des salariés.

RÈGLES IMPÉRATIVES :
- Utilise UNIQUEMENT les informations présentes dans le contenu fourni.
- N'invente aucun chiffre, aucune règle, aucune date, aucune sanction ni aucun fait.
- Ne dis pas "la slide", "la diapositive" ou "comme vous pouvez le voir".
- Ne lis pas les éléments de pied de page, numéros, dates de modèle ou mentions techniques.
- Le français doit être naturel, simple et professionnel.
- Évite les répétitions avec la narration précédente.
- Une transition courte est autorisée si elle ne crée pas de nouvelle information.
- Vise environ {words_min} à {words_max} mots.
- Retourne UNIQUEMENT le texte à prononcer.
- Aucun titre, aucun guillemet, aucun commentaire, aucun raisonnement.

STYLE :
{mode_instruction}

NUMÉRO DE SLIDE :
{slide["number"]}

CONTENU À EXPLIQUER :
{slide["clean_text"] if slide["clean_text"] else "[Aucun contenu textuel exploitable]"}

NARRATION DE LA SLIDE PRÉCÉDENTE :
{previous_narration[-1200:] if previous_narration else "[Première slide]"}

CONTENU DE LA SLIDE SUIVANTE :
{next_slide_text[:1200] if next_slide_text else "[Dernière slide]"}

Rédige maintenant la narration.
""".strip()


def generate_text(api_key: str, model: str, prompt: str) -> str:
    client = Groq(api_key=api_key)

    kwargs = {
        "model": model,
        "messages": [
            {
                "role": "system",
                "content": (
                    "Tu écris des narrations professionnelles en français. "
                    "Tu respectes strictement le contenu fourni et tu n'inventes rien."
                ),
            },
            {
                "role": "user",
                "content": prompt,
            },
        ],
        "temperature": 0.3,
        "max_tokens": 900,
    }

    # Qwen 3.x peut exposer un mode reasoning. On lui demande explicitement
    # de retourner uniquement la réponse utile via le prompt ci-dessus.
    response = client.chat.completions.create(**kwargs)
    content = response.choices[0].message.content or ""
    return normalize_text(content)


@st.cache_resource(show_spinner="Chargement de la voix française Kokoro…")
def get_kokoro_pipeline():
    # 'f' = français fr-FR dans Kokoro.
    return KPipeline(lang_code="f")


def kokoro_tts(text: str, speed: float = 1.0) -> bytes:
    """
    Synthèse locale/open-source.
    Kokoro fournit actuellement une voix française ff_siwis.
    """
    pipeline = get_kokoro_pipeline()
    generator = pipeline(
        text,
        voice="ff_siwis",
        speed=speed,
        split_pattern=r"\n+|(?<=[.!?;:])\s+",
    )

    chunks = []

    for _graphemes, _phonemes, audio in generator:
        arr = np.asarray(audio, dtype=np.float32)
        if arr.size:
            chunks.append(arr)

    if not chunks:
        raise RuntimeError("Kokoro n'a produit aucun audio pour ce texte.")

    # Petite pause entre les segments pour éviter un débit trop compact.
    pause = np.zeros(int(24000 * 0.10), dtype=np.float32)
    merged_parts = []

    for i, chunk in enumerate(chunks):
        merged_parts.append(chunk)
        if i < len(chunks) - 1:
            merged_parts.append(pause)

    merged = np.concatenate(merged_parts)

    buf = io.BytesIO()
    sf.write(buf, merged, 24000, format="WAV", subtype="PCM_16")
    return buf.getvalue()


def pcm_to_wav_bytes(pcm: bytes, channels=1, rate=24000, sample_width=2) -> bytes:
    buffer = io.BytesIO()

    with wave.open(buffer, "wb") as wf:
        wf.setnchannels(channels)
        wf.setsampwidth(sample_width)
        wf.setframerate(rate)
        wf.writeframes(pcm)

    return buffer.getvalue()


def create_silent_wav(duration_seconds: float, rate: int = 24000) -> bytes:
    duration_seconds = max(0.5, float(duration_seconds))
    frame_count = int(duration_seconds * rate)
    pcm = b"\x00\x00" * frame_count
    return pcm_to_wav_bytes(pcm, channels=1, rate=rate, sample_width=2)


def find_command(names: List[str]) -> str | None:
    for name in names:
        path = shutil.which(name)
        if path:
            return path
    return None


def render_pptx_to_pngs(pptx_bytes: bytes, workdir: Path, dpi: int = 140) -> List[Path]:
    soffice = find_command(["libreoffice", "soffice"])

    if not soffice:
        raise RuntimeError(
            "LibreOffice est introuvable. "
            "Sur Streamlit Community Cloud, vérifie packages.txt."
        )

    pptx_path = workdir / "presentation.pptx"
    pptx_path.write_bytes(pptx_bytes)

    cmd = [
        soffice,
        "--headless",
        "--convert-to", "pdf",
        "--outdir", str(workdir),
        str(pptx_path),
    ]

    result = subprocess.run(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        timeout=180,
    )

    pdf_path = workdir / "presentation.pdf"

    if result.returncode != 0 or not pdf_path.exists():
        raise RuntimeError(
            "LibreOffice n'a pas réussi à convertir le PowerPoint.\n"
            f"{result.stderr[-2000:]}"
        )

    doc = fitz.open(pdf_path)
    png_paths = []

    zoom = dpi / 72.0
    matrix = fitz.Matrix(zoom, zoom)

    for index, page in enumerate(doc, start=1):
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        out = workdir / f"slide_{index:03d}.png"
        pix.save(out)
        png_paths.append(out)

    doc.close()
    return png_paths


def build_video(
    pptx_bytes: bytes,
    audio_by_slide: Dict[int, bytes],
    slide_count: int,
    silent_slide_seconds: float = 2.0,
    resolution: str = "1280x720",
):
    ffmpeg = find_command(["ffmpeg"])

    if not ffmpeg:
        raise RuntimeError(
            "FFmpeg est introuvable. "
            "Sur Streamlit Community Cloud, vérifie packages.txt."
        )

    width, height = [int(x) for x in resolution.split("x")]

    with tempfile.TemporaryDirectory() as tmp:
        workdir = Path(tmp)
        pngs = render_pptx_to_pngs(pptx_bytes, workdir)

        if len(pngs) != slide_count:
            raise RuntimeError(
                f"{slide_count} slides attendues, mais {len(pngs)} slides ont été rendues."
            )

        segment_paths = []

        for index, png_path in enumerate(pngs, start=1):
            wav_bytes = audio_by_slide.get(index)

            if not wav_bytes:
                wav_bytes = create_silent_wav(silent_slide_seconds)

            wav_path = workdir / f"audio_{index:03d}.wav"
            wav_path.write_bytes(wav_bytes)

            segment_path = workdir / f"segment_{index:03d}.mp4"

            vf = (
                f"scale={width}:{height}:force_original_aspect_ratio=decrease,"
                f"pad={width}:{height}:(ow-iw)/2:(oh-ih)/2,"
                "setsar=1"
            )

            cmd = [
                ffmpeg,
                "-y",
                "-loop", "1",
                "-framerate", "25",
                "-i", str(png_path),
                "-i", str(wav_path),
                "-vf", vf,
                "-c:v", "libx264",
                "-preset", "veryfast",
                "-tune", "stillimage",
                "-pix_fmt", "yuv420p",
                "-r", "25",
                "-c:a", "aac",
                "-b:a", "160k",
                "-ar", "48000",
                "-ac", "2",
                "-shortest",
                "-movflags", "+faststart",
                str(segment_path),
            ]

            result = subprocess.run(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                timeout=240,
            )

            if result.returncode != 0 or not segment_path.exists():
                raise RuntimeError(
                    f"FFmpeg a échoué sur la slide {index}.\n"
                    f"{result.stderr[-2500:]}"
                )

            segment_paths.append(segment_path)

        concat_file = workdir / "concat.txt"
        concat_file.write_text(
            "\n".join(f"file '{p.as_posix()}'" for p in segment_paths),
            encoding="utf-8",
        )

        final_path = workdir / "presentation_narree.mp4"

        result = subprocess.run(
            [
                ffmpeg,
                "-y",
                "-f", "concat",
                "-safe", "0",
                "-i", str(concat_file),
                "-c", "copy",
                "-movflags", "+faststart",
                str(final_path),
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=240,
        )

        if result.returncode != 0 or not final_path.exists():
            raise RuntimeError(
                "FFmpeg n'a pas réussi à assembler la vidéo finale.\n"
                f"{result.stderr[-3000:]}"
            )

        return (
            final_path.read_bytes(),
            [p.read_bytes() for p in pngs],
        )


def create_audio_zip(audio_by_slide: Dict[int, bytes]) -> bytes:
    buf = io.BytesIO()

    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for slide_num, wav_bytes in sorted(audio_by_slide.items()):
            z.writestr(
                f"slide_{slide_num:02d}.wav",
                wav_bytes,
            )

    return buf.getvalue()


def create_png_zip(png_bytes: List[bytes]) -> bytes:
    buf = io.BytesIO()

    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for i, data in enumerate(png_bytes, start=1):
            z.writestr(
                f"slide_{i:02d}.png",
                data,
            )

    return buf.getvalue()


st.set_page_config(
    page_title="Présentation IA — V4",
    page_icon="🎬",
    layout="wide",
)

st.title("🎬 Présentation IA — V4")
st.caption(
    "PowerPoint → Groq → Kokoro → vidéo MP4 • Gemini supprimé"
)

with st.sidebar:
    st.header("Narration")

    mode = st.selectbox(
        "Mode",
        [
            "Expliquer naturellement",
            "Résumer",
            "Lecture reformulée",
        ],
    )

    target_seconds = st.slider(
        "Durée cible par slide",
        15, 90, 40, 5,
    )

    remove_repeated = st.toggle(
        "Supprimer les longs blocs répétés",
        value=True,
    )

    min_occurrences = st.slider(
        "Répétition à partir de",
        2, 6, 3, 1,
    )

    custom_ignored_text = st.text_area(
        "Expressions à toujours ignorer",
        value=(
            "Département douane\n"
            "Titre de la présentation\n"
            "Émetteur"
        ),
    )

    custom_ignored = [
        x.strip()
        for x in custom_ignored_text.splitlines()
        if x.strip()
    ]

    st.divider()
    st.header("Groq")

    groq_key = get_groq_key()

    if not groq_key:
        groq_key = st.text_input(
            "Clé API Groq",
            type="password",
            help="Sur Streamlit Cloud, utilise GROQ_API_KEY dans Secrets.",
        )
    else:
        st.success("Clé Groq chargée depuis Secrets")

    text_model = st.selectbox(
        "Modèle de narration",
        TEXT_MODELS,
        index=0,
    )

    st.divider()
    st.header("Voix Kokoro")

    st.write("🇫🇷 Voix : **ff_siwis**")

    speech_speed = st.slider(
        "Vitesse de lecture",
        min_value=0.75,
        max_value=1.25,
        value=0.95,
        step=0.05,
    )

    st.caption(
        "Kokoro tourne directement sur le serveur : "
        "aucune clé API n'est utilisée pour la voix."
    )

    st.divider()
    st.header("Vidéo")

    resolution = st.selectbox(
        "Résolution",
        ["1280x720", "1920x1080"],
        index=0,
    )

    silent_slide_seconds = st.slider(
        "Durée d'une slide sans audio",
        1.0, 8.0, 2.0, 0.5,
    )


uploaded = st.file_uploader(
    "Dépose ton PowerPoint ici",
    type=["pptx"],
)

if uploaded is None:
    st.info("Charge un fichier .pptx pour commencer.")
    st.stop()

pptx_bytes = uploaded.getvalue()
file_hash = hashlib.md5(pptx_bytes).hexdigest()

try:
    raw_slides = extract_presentation(uploaded)
except Exception as exc:
    st.error(f"Impossible de lire le PowerPoint : {exc}")
    st.stop()

slides = clean_slides(
    raw_slides,
    custom_ignored,
    remove_repeated,
    min_occurrences,
)

if st.session_state.get("ppt_hash") != file_hash:
    st.session_state["ppt_hash"] = file_hash
    st.session_state["narrations"] = {}
    st.session_state["audio"] = {}
    st.session_state["video"] = None
    st.session_state["slide_pngs"] = []

narrations = st.session_state.setdefault(
    "narrations",
    {},
)

audio_by_slide = st.session_state.setdefault(
    "audio",
    {},
)

c1, c2, c3, c4 = st.columns(4)

c1.metric("Slides", len(slides))

c2.metric(
    "Narrations",
    sum(
        bool(narrations.get(s["number"]))
        for s in slides
    ),
)

c3.metric(
    "Audios",
    len(audio_by_slide),
)

c4.metric(
    "Vidéo",
    "Prête"
    if st.session_state.get("video")
    else "Non générée",
)

tab_content, tab_narr, tab_audio, tab_video, tab_export = st.tabs(
    [
        "1. Contenu",
        "2. Narrations",
        "3. Voix Kokoro",
        "4. Vidéo",
        "5. Export",
    ]
)


with tab_content:
    for slide in slides:
        with st.expander(
            f"Slide {slide['number']}",
            expanded=slide["number"] <= 2,
        ):
            st.text_area(
                "Texte conservé",
                value=slide["clean_text"],
                height=170,
                key=f"src_{file_hash}_{slide['number']}",
                disabled=True,
            )

            if slide["removed_blocks"]:
                st.caption("Retiré automatiquement")

                for item in slide["removed_blocks"]:
                    preview = item["text"].replace(
                        "\n",
                        " ",
                    )

                    st.write(
                        f"- {item['reason']} : "
                        f"{preview[:220]}"
                    )


with tab_narr:
    if not groq_key:
        st.warning(
            "Ajoute une clé Groq dans la barre latérale."
        )

    if st.button(
        "✨ Générer toutes les narrations",
        type="primary",
        disabled=not groq_key,
        use_container_width=True,
    ):
        progress = st.progress(0)
        previous = ""

        for i, slide in enumerate(slides):
            next_text = (
                slides[i + 1]["clean_text"]
                if i + 1 < len(slides)
                else ""
            )

            if slide["clean_text"]:
                try:
                    narration = generate_text(
                        groq_key,
                        text_model,
                        build_prompt(
                            slide,
                            previous,
                            next_text,
                            mode,
                            target_seconds,
                        ),
                    )

                    narrations[
                        slide["number"]
                    ] = narration

                    previous = narration

                except Exception as exc:
                    st.error(
                        f"Slide {slide['number']} : "
                        f"{exc}"
                    )
                    break

            progress.progress(
                (i + 1) / len(slides)
            )

        st.session_state[
            "narrations"
        ] = narrations

        st.session_state["video"] = None

        st.success(
            "Narrations générées avec Groq."
        )

    for i, slide in enumerate(slides):
        n = slide["number"]

        st.markdown(f"### Slide {n}")

        current = narrations.get(n, "")

        edited = st.text_area(
            "Narration",
            value=current,
            height=160,
            key=(
                f"edit_{file_hash}_"
                f"{n}_{hash(current)}"
            ),
        )

        if edited != current:
            narrations[n] = edited
            audio_by_slide.pop(n, None)
            st.session_state["video"] = None

        if st.button(
            f"Régénérer la slide {n}",
            key=f"regen_{n}",
            disabled=not groq_key,
        ):
            prev = (
                narrations.get(
                    slides[i - 1]["number"],
                    "",
                )
                if i > 0
                else ""
            )

            next_text = (
                slides[i + 1]["clean_text"]
                if i + 1 < len(slides)
                else ""
            )

            try:
                narrations[n] = generate_text(
                    groq_key,
                    text_model,
                    build_prompt(
                        slide,
                        prev,
                        next_text,
                        mode,
                        target_seconds,
                    ),
                )

                audio_by_slide.pop(n, None)

                st.session_state[
                    "narrations"
                ] = narrations

                st.session_state["video"] = None
                st.rerun()

            except Exception as exc:
                st.error(str(exc))

        st.divider()


with tab_audio:
    st.subheader(
        "Transformer les narrations en voix"
    )

    st.info(
        "La première génération peut être plus longue : "
        "Streamlit doit charger le modèle Kokoro. "
        "Les générations suivantes sont plus rapides."
    )

    if not any(narrations.values()):
        st.info(
            "Génère d'abord les narrations."
        )

    else:
        if st.button(
            "🔊 Générer tous les audios avec Kokoro",
            type="primary",
            use_container_width=True,
        ):
            progress = st.progress(0)

            for i, slide in enumerate(slides):
                n = slide["number"]
                narration = (
                    narrations.get(n, "").strip()
                )

                if narration:
                    try:
                        audio_by_slide[n] = (
                            kokoro_tts(
                                narration,
                                speed=speech_speed,
                            )
                        )

                    except Exception as exc:
                        st.error(
                            f"Audio slide {n} : "
                            f"{exc}"
                        )
                        break

                progress.progress(
                    (i + 1) / len(slides)
                )

            st.session_state[
                "audio"
            ] = audio_by_slide

            st.session_state["video"] = None

            st.success(
                "Audios Kokoro générés."
            )

        for slide in slides:
            n = slide["number"]
            narration = (
                narrations.get(n, "").strip()
            )

            if not narration:
                continue

            st.markdown(f"### Slide {n}")
            st.caption(narration)

            if st.button(
                f"🔊 Générer / régénérer slide {n}",
                key=f"tts_{n}",
            ):
                try:
                    with st.spinner(
                        "Synthèse Kokoro…"
                    ):
                        audio_by_slide[n] = (
                            kokoro_tts(
                                narration,
                                speed=speech_speed,
                            )
                        )

                        st.session_state[
                            "audio"
                        ] = audio_by_slide

                        st.session_state[
                            "video"
                        ] = None

                    st.rerun()

                except Exception as exc:
                    st.error(
                        f"Erreur Kokoro : {exc}"
                    )

            if n in audio_by_slide:
                st.audio(
                    audio_by_slide[n],
                    format="audio/wav",
                )

            st.divider()


with tab_video:
    st.subheader(
        "Créer la présentation vidéo"
    )

    missing_audio = [
        s["number"]
        for s in slides
        if narrations.get(
            s["number"],
            "",
        ).strip()
        and s["number"]
        not in audio_by_slide
    ]

    if missing_audio:
        st.warning(
            "Audio manquant pour : "
            + ", ".join(
                f"slide {n}"
                for n in missing_audio
            )
        )

    libreoffice_ok = bool(
        find_command(
            ["libreoffice", "soffice"]
        )
    )

    ffmpeg_ok = bool(
        find_command(["ffmpeg"])
    )

    d1, d2 = st.columns(2)

    d1.write(
        "✅ LibreOffice détecté"
        if libreoffice_ok
        else "❌ LibreOffice non détecté"
    )

    d2.write(
        "✅ FFmpeg détecté"
        if ffmpeg_ok
        else "❌ FFmpeg non détecté"
    )

    if (
        not libreoffice_ok
        or not ffmpeg_ok
    ):
        st.info(
            "Sur Streamlit Community Cloud, "
            "packages.txt les installe automatiquement."
        )

    if st.button(
        "🎬 Générer le MP4 final",
        type="primary",
        use_container_width=True,
        disabled=not (
            libreoffice_ok
            and ffmpeg_ok
        ),
    ):
        try:
            with st.spinner(
                "Création de la vidéo…"
            ):
                video, rendered_pngs = (
                    build_video(
                        pptx_bytes,
                        audio_by_slide,
                        len(slides),
                        silent_slide_seconds,
                        resolution,
                    )
                )

                st.session_state[
                    "video"
                ] = video

                st.session_state[
                    "slide_pngs"
                ] = rendered_pngs

        except Exception as exc:
            st.error(str(exc))

    current_video = st.session_state.get(
        "video"
    )

    if current_video:
        st.success(
            "La présentation narrée est prête."
        )

        st.video(current_video)

        safe_name = re.sub(
            r"[^A-Za-z0-9_-]+",
            "_",
            Path(uploaded.name).stem,
        ).strip("_")

        st.download_button(
            "⬇️ Télécharger le MP4",
            current_video,
            file_name=(
                f"{safe_name}_narree.mp4"
            ),
            mime="video/mp4",
            use_container_width=True,
        )


with tab_export:
    export_obj = {
        "source_file": uploaded.name,
        "model": text_model,
        "tts": "Kokoro-82M / ff_siwis",
        "slides": [
            {
                "slide": s["number"],
                "source_text": s[
                    "clean_text"
                ],
                "narration": narrations.get(
                    s["number"],
                    "",
                ),
            }
            for s in slides
        ],
    }

    json_data = json.dumps(
        export_obj,
        ensure_ascii=False,
        indent=2,
    )

    # Construction de l'export texte.\n    txt_parts = []
    for s in slides:
        n = s["number"]
        txt_parts.append(
            f"SLIDE {n}\n"
            f"{narrations.get(n, '[Aucune narration]')}"
        )
    txt_data = "\n\n".join(txt_parts)

    c1, c2 = st.columns(2)

    with c1:
        st.download_button(
            "⬇️ Narrations JSON",
            json_data.encode("utf-8"),
            "narrations.json",
            "application/json",
            use_container_width=True,
        )

    with c2:
        st.download_button(
            "⬇️ Narrations TXT",
            txt_data.encode("utf-8"),
            "narrations.txt",
            "text/plain",
            use_container_width=True,
        )

    if audio_by_slide:
        st.download_button(
            "⬇️ Tous les audios (.zip)",
            create_audio_zip(
                audio_by_slide
            ),
            "audios_slides.zip",
            "application/zip",
            use_container_width=True,
        )

    if st.session_state.get(
        "slide_pngs"
    ):
        st.download_button(
            "⬇️ Slides rendues (.zip)",
            create_png_zip(
                st.session_state[
                    "slide_pngs"
                ]
            ),
            "slides_png.zip",
            "application/zip",
            use_container_width=True,
        )
