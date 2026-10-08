
import asyncio, hashlib, io, json, re, shutil, subprocess, tempfile, wave, zipfile
from collections import Counter
from pathlib import Path

import edge_tts
import fitz
import streamlit as st
from openai import OpenAI
from pptx import Presentation

IGNORED = [
    r"^\s*Département\s+douane\s*$",
    r"^\s*Date\s*$",
    r"^\s*Titre de la présentation.*Émetteur\s*$",
    r"^\s*\d+\s*$",
]

EDGE_VOICES = {
    "Denise — femme, France": "fr-FR-DeniseNeural",
    "Henri — homme, France": "fr-FR-HenriNeural",
    "Eloise — femme, France": "fr-FR-EloiseNeural",
    "Remy — homme, France": "fr-FR-RemyMultilingualNeural",
    "Vivienne — femme, France": "fr-FR-VivienneMultilingualNeural",
}

def norm(t):
    t = t.replace("\u00a0", " ")
    t = re.sub(r"[ \t]+", " ", t)
    return re.sub(r"\n{3,}", "\n\n", t).strip()

def cmp_norm(t):
    return re.sub(r"\s+", " ", norm(t).lower()).strip(" .,:;-")

def shape_text(shape):
    out = []
    if hasattr(shape, "shapes"):
        for s in shape.shapes:
            out += shape_text(s)
    if getattr(shape, "has_text_frame", False):
        txt = "\n".join(norm(p.text) for p in shape.text_frame.paragraphs if norm(p.text))
        if txt:
            out.append(txt)
    if getattr(shape, "has_table", False):
        rows = []
        for row in shape.table.rows:
            vals = [norm(c.text) for c in row.cells if norm(c.text)]
            if vals:
                rows.append(" | ".join(vals))
        if rows:
            out.append("\n".join(rows))
    return out

def extract_slides(pptx_bytes):
    prs = Presentation(io.BytesIO(pptx_bytes))
    slides = []
    for i, slide in enumerate(prs.slides, 1):
        blocks, seen = [], set()
        for sh in slide.shapes:
            for b in shape_text(sh):
                k = cmp_norm(b)
                if k and k not in seen:
                    seen.add(k)
                    blocks.append(norm(b))
        slides.append({"number": i, "raw_blocks": blocks})
    return slides

def clean_slides(slides, custom_ignored):
    counts = Counter()
    for s in slides:
        for b in set(cmp_norm(x) for x in s["raw_blocks"] if len(cmp_norm(x)) >= 70):
            counts[b] += 1
    repeated = {k for k, v in counts.items() if v >= 3}
    cleaned = []
    for s in slides:
        kept = []
        for b in s["raw_blocks"]:
            if any(re.match(p, b, re.I) for p in IGNORED):
                continue
            if any(x.lower() in b.lower() for x in custom_ignored if x):
                continue
            if cmp_norm(b) in repeated:
                continue
            kept.append(b)
        cleaned.append({"number": s["number"], "clean_text": "\n\n".join(kept).strip()})
    return cleaned

def get_openai_key():
    try:
        return str(st.secrets["OPENAI_API_KEY"]).strip()
    except Exception:
        return ""

def narration_prompt(slide, previous, target_seconds):
    return f"""
Transforme le contenu suivant en narration orale professionnelle en français.

OBJECTIF :
Présenter fidèlement TOUTES les informations utiles présentes dans le contenu.
Il ne s'agit PAS de résumer.

RÈGLES :
- Ne supprime aucune idée importante.
- Conserve tous les chiffres, dates, montants, références, exemples,
  conséquences et sanctions présents dans le contenu.
- Si le contenu contient plusieurs puces, traite chacune d'elles.
- Pour un cas pratique, conserve le contexte, l'erreur commise,
  les conséquences et les sanctions mentionnées.
- Tu peux reformuler pour rendre le discours naturel à l'oral.
- N'invente aucune information.
- Ne complète pas avec tes connaissances générales.
- Ne dis jamais "slide", "diapositive" ou "comme vous pouvez le voir".
- Ne lis pas les pieds de page ou les éléments techniques.
- Si le contenu est simplement une page de titre, reste très bref.
- Si le contenu est dense, prends le temps nécessaire pour tout expliquer.
- La fidélité au contenu est prioritaire sur la durée cible.
- Retourne uniquement le texte qui doit être prononcé.

CONTENU :
{slide["clean_text"]}

NARRATION PRÉCÉDENTE :
{previous[-1000:] if previous else "[Aucune]"}
""".strip()

def generate_narration(api_key, model, prompt):
    client = OpenAI(api_key=api_key)

    response = client.responses.create(
        model=model,
        instructions=(
            "Tu rédiges uniquement le texte final destiné à être prononcé "
            "dans une présentation professionnelle en français. "
            "Tu dois être fidèle au contenu fourni, ne rien inventer et "
            "ne retourner aucune analyse, aucun raisonnement et aucun commentaire."
        ),
        input=prompt,
        reasoning={
            "effort": "none"
        },
        max_output_tokens=1200,
    )

    return norm(response.output_text or "")

async def edge_audio_async(text, voice, rate, pitch):
    c = edge_tts.Communicate(text=text, voice=voice, rate=rate, pitch=pitch)
    data = bytearray()
    async for chunk in c.stream():
        if chunk["type"] == "audio":
            data.extend(chunk["data"])
    if not data:
        raise RuntimeError("Edge TTS n'a produit aucun audio.")
    return bytes(data)

def edge_audio(text, voice, rate, pitch):
    return asyncio.run(edge_audio_async(text, voice, rate, pitch))

def find_cmd(names):
    for n in names:
        p = shutil.which(n)
        if p:
            return p
    return None

def mp3_to_wav(mp3):
    ffmpeg = find_cmd(["ffmpeg"])
    if not ffmpeg:
        raise RuntimeError("FFmpeg introuvable.")
    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        a, w = td/"a.mp3", td/"a.wav"
        a.write_bytes(mp3)
        r = subprocess.run([ffmpeg, "-y", "-i", str(a), "-ar", "24000", "-ac", "1", str(w)],
                           capture_output=True, text=True, timeout=120)
        if r.returncode != 0 or not w.exists():
            raise RuntimeError(r.stderr[-1500:])
        return w.read_bytes()

def silent_wav(seconds=2):
    import wave
    buf = io.BytesIO()
    rate = 24000
    with wave.open(buf, "wb") as wf:
        wf.setnchannels(1); wf.setsampwidth(2); wf.setframerate(rate)
        wf.writeframes(b"\x00\x00" * int(rate * seconds))
    return buf.getvalue()

AVATAR_POSITIONS = [
    "À droite de la diapositive", "À gauche de la diapositive",
    "En bas à droite", "En bas à gauche", "En haut à droite", "En haut à gauche",
]

def wav_duration(data):
    with wave.open(io.BytesIO(data), "rb") as audio:
        return audio.getnframes() / audio.getframerate()

def audio_hash(data):
    return hashlib.sha256(data).hexdigest()

def audio_bundle(audios, narrations, slide_numbers, presentation_sha256=None):
    """Export the exact audio used by the MP4, without generating new speech."""
    buf = io.BytesIO()
    manifest = {"slides": [], "slides_sans_audio": [], "presentation_sha256": presentation_sha256}
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for n in slide_numbers:
            data = audios.get(n)
            if not data:
                manifest["slides_sans_audio"].append(n)
                continue
            filename = f"slide_{n:03d}.wav"
            archive.writestr(filename, data)
            archive.writestr(f"slide_{n:03d}.txt", narrations.get(n, ""))
            manifest["slides"].append({
                "diapositive": n, "audio": filename,
                "duree_secondes": round(wav_duration(data), 3),
                "sha256": audio_hash(data), "avatar_attendu": f"slide_{n:03d}.mp4",
            })
        archive.writestr("manifest.json", json.dumps(manifest, ensure_ascii=False, indent=2))
        archive.writestr("MODE_EMPLOI.txt", (
            "Utilise chaque fichier WAV comme voix de l'avatar, sans regénérer la narration.\n"
            "Exporte une vidéo d'avatar par audio, sans introduction, musique ni changement de vitesse.\n"
            "Nomme la vidéo slide_001.mp4 pour slide_001.wav, et ainsi de suite.\n"
            "Importe les vidéos dans l'onglet Avatar de l'application.\n"
            "Si tu modifies un audio, regénère aussi sa vidéo d'avatar.\n"
        ))
    return buf.getvalue()

def load_audio_bundle(data, slide_numbers, presentation_sha256):
    """Restore only an export made for this exact PowerPoint."""
    restored_audios, restored_narrations = {}, {}
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            if sum(entry.file_size for entry in archive.infolist()) > 200 * 1024 * 1024:
                raise RuntimeError("L'archive audio dépasse 200 Mo une fois décompressée.")
            manifest = json.loads(archive.read("manifest.json"))
            if not isinstance(manifest, dict) or not isinstance(manifest.get("slides"), list):
                raise ValueError("Manifeste audio invalide.")
            if manifest.get("presentation_sha256") != presentation_sha256:
                raise RuntimeError("Ce ZIP a été exporté pour un autre PowerPoint. Recharge la présentation correspondante.")
            for item in manifest["slides"]:
                n = item["diapositive"]
                if n not in slide_numbers or n in restored_audios:
                    raise RuntimeError("L'archive contient un numéro de diapositive invalide ou en double.")
                wav = archive.read(f"slide_{n:03d}.wav")
                if audio_hash(wav) != item["sha256"]:
                    raise RuntimeError(f"L'audio de la diapositive {n} a été modifié depuis l'export.")
                if wav_duration(wav) <= 0:
                    raise RuntimeError(f"L'audio de la diapositive {n} est vide.")
                restored_audios[n] = wav
                restored_narrations[n] = archive.read(f"slide_{n:03d}.txt").decode("utf-8")
    except (zipfile.BadZipFile, KeyError, TypeError, ValueError, UnicodeDecodeError, wave.Error, EOFError) as exc:
        raise RuntimeError("ZIP audio invalide. Utilise l'archive téléchargée depuis l'onglet Voix.") from exc
    if not restored_audios:
        raise RuntimeError("Cette archive ne contient aucun audio.")
    return restored_audios, restored_narrations

def guess_avatar_slide(filename, slide_numbers):
    match = re.fullmatch(r"(?:(?:slide|diapo|avatar)[_ -]?)?(\d+)", Path(filename).stem, re.I)
    number = int(match.group(1)) if match else None
    return number if number in slide_numbers else None

def avatar_binding_issues(avatar_files, audios, bindings):
    issues = []
    for n in avatar_files:
        if n not in audios:
            issues.append(f"Diapositive {n} : génère son audio avant d'utiliser l'avatar.")
        elif bindings.get(n, {}).get("audio_sha") != audio_hash(audios[n]):
            issues.append(
                f"Diapositive {n} : l'audio a changé depuis l'import de l'avatar. "
                "Regénère la vidéo d'avatar avec le nouvel audio, puis remplace le fichier importé."
            )
    return issues

def probe_avatar(path, ffprobe):
    result = subprocess.run(
        [ffprobe, "-v", "error", "-show_streams", "-show_format", "-of", "json", str(path)],
        capture_output=True, text=True, timeout=30,
    )
    if result.returncode:
        raise RuntimeError("Vidéo d'avatar illisible. Utilise un fichier MP4 ou WebM valide.")
    metadata = json.loads(result.stdout)
    videos = [s for s in metadata.get("streams", []) if s.get("codec_type") == "video"]
    if not videos:
        raise RuntimeError("Le fichier d'avatar ne contient pas de vidéo.")
    duration = videos[0].get("duration") or metadata.get("format", {}).get("duration")
    try:
        duration = float(duration)
    except (TypeError, ValueError):
        raise RuntimeError("La durée de la vidéo d'avatar est introuvable.") from None
    if duration <= 0:
        raise RuntimeError("La vidéo d'avatar est vide.")
    return duration

def validate_avatar_duration(path, audio_seconds, ffprobe, slide_number):
    seconds = probe_avatar(path, ffprobe)
    # Allow a little encoder padding, not an unrelated or truncated narration.
    if abs(seconds - audio_seconds) > 0.5:
        raise RuntimeError(
            f"Diapositive {slide_number} : l'avatar dure {seconds:.2f} s et l'audio "
            f"{audio_seconds:.2f} s. Utilise le WAV exporté correspondant, sans "
            "introduction ni changement de vitesse (écart maximal : 0,5 s)."
        )

def avatar_filtergraph(width, height, position, width_percent):
    if position not in AVATAR_POSITIONS:
        raise ValueError("Position d'avatar inconnue.")
    if not 15 <= width_percent <= 35:
        raise ValueError("La largeur de l'avatar doit être comprise entre 15 et 35 %.")
    margin = max(8, int(width * 0.0125))
    avatar_width = int(width * width_percent / 100) // 2 * 2
    beside = "de la diapositive" in position
    panel_width = avatar_width + 2 * margin
    slide_width = width - panel_width if beside else width
    slide_x = panel_width if position == "À gauche de la diapositive" else 0
    avatar_height = (height - 2 * margin) if beside else int(height * 0.42) // 2 * 2
    if beside:
        x = str(margin) if slide_x else f"main_w-overlay_w-{margin}"
        y = "(main_h-overlay_h)/2"
    else:
        x = str(margin) if "gauche" in position else f"main_w-overlay_w-{margin}"
        y = str(margin) if "haut" in position else f"main_h-overlay_h-{margin}"
    return (
        f"[0:v]scale={slide_width}:{height}:force_original_aspect_ratio=decrease,"
        f"pad={width}:{height}:{slide_x}+({slide_width}-iw)/2:(oh-ih)/2,"
        "setsar=1,setpts=PTS-STARTPTS[slide];"
        f"[2:v]scale={avatar_width}:{avatar_height}:force_original_aspect_ratio=decrease,"
        "setsar=1,setpts=PTS-STARTPTS,fps=25[avatar];"
        f"[slide][avatar]overlay=x={x}:y={y}:eof_action=repeat:repeatlast=1[v]"
    )

def render_video_segment(ffmpeg, png, wav, segment, width, height,
                         avatar=None, position=AVATAR_POSITIONS[0], width_percent=22):
    seconds = wav_duration(Path(wav).read_bytes())
    command = [ffmpeg, "-y", "-loop", "1", "-framerate", "25", "-i", str(png), "-i", str(wav)]
    if avatar:
        command += [
            "-i", str(avatar), "-filter_complex_threads", "1",
            "-filter_complex", avatar_filtergraph(width, height, position, width_percent),
            "-map", "[v]", "-map", "1:a:0",
        ]
    else:
        vf = f"scale={width}:{height}:force_original_aspect_ratio=decrease,pad={width}:{height}:(ow-iw)/2:(oh-ih)/2,setsar=1"
        command += ["-vf", vf, "-map", "0:v:0", "-map", "1:a:0"]
    command += [
        "-c:v", "libx264", "-preset", "veryfast", "-pix_fmt", "yuv420p", "-r", "25",
        "-c:a", "aac", "-ar", "24000", "-ac", "1", "-t", f"{seconds:.6f}",
        "-shortest", str(segment),
    ]
    result = subprocess.run(command, capture_output=True, text=True, timeout=max(240, int(seconds * 6)))
    if result.returncode or not Path(segment).exists():
        raise RuntimeError(result.stderr[-1500:])

def render_slides(powerpoint_bytes, workdir, extension=".pptx"):
    soffice = find_cmd(["libreoffice", "soffice"])
    if not soffice:
        raise RuntimeError("LibreOffice introuvable.")

    extension = extension.lower()
    if extension not in [".pptx", ".pptm"]:
        raise RuntimeError(f"Format PowerPoint non pris en charge : {extension}")

    powerpoint_file = workdir / f"presentation{extension}"
    powerpoint_file.write_bytes(powerpoint_bytes)

    r = subprocess.run(
        [
            soffice,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", str(workdir),
            str(powerpoint_file),
        ],
        capture_output=True,
        text=True,
        timeout=180,
    )
    pdf = workdir/"presentation.pdf"
    if r.returncode != 0 or not pdf.exists():
        raise RuntimeError("Conversion PowerPoint → PDF impossible.")
    doc = fitz.open(pdf)
    out = []
    for i, page in enumerate(doc, 1):
        pix = page.get_pixmap(matrix=fitz.Matrix(2,2), alpha=False)
        p = workdir/f"slide_{i:03d}.png"
        pix.save(p)
        out.append(p)
    doc.close()
    return out

def build_video(powerpoint_bytes, audios, slide_count, resolution="1280x720", extension=".pptx",
                avatar_files=None, avatar_position=AVATAR_POSITIONS[0], avatar_width_percent=22,
                progress_callback=None):
    ffmpeg = find_cmd(["ffmpeg"])
    if not ffmpeg:
        raise RuntimeError("FFmpeg introuvable.")
    avatar_files = avatar_files or {}
    ffprobe = find_cmd(["ffprobe"]) if avatar_files else None
    if avatar_files and not ffprobe:
        raise RuntimeError("FFprobe introuvable. Il est fourni avec FFmpeg.")
    w, h = map(int, resolution.split("x"))
    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        avatars = {}
        for n, upload in avatar_files.items():
            if n not in audios or not 1 <= n <= slide_count:
                raise RuntimeError(f"Diapositive {n} : audio ou numéro de diapositive invalide.")
            suffix = Path(upload.name).suffix.lower()
            if suffix not in [".mp4", ".webm"]:
                raise RuntimeError("L'avatar doit être une vidéo MP4 ou WebM.")
            path = td / f"avatar_{n:03d}{suffix}"
            with path.open("wb") as destination:
                destination.write(upload.getbuffer())
            validate_avatar_duration(path, wav_duration(audios[n]), ffprobe, n)
            avatars[n] = path
        pngs = render_slides(powerpoint_bytes, td, extension)
        if len(pngs) != slide_count:
            raise RuntimeError("Le nombre de pages rendues ne correspond pas aux diapositives.")
        segs = []
        for i, png in enumerate(pngs, 1):
            wav = td/f"a_{i}.wav"
            wav.write_bytes(audios.get(i, silent_wav()))
            seg = td/f"s_{i}.mp4"
            render_video_segment(ffmpeg, png, wav, seg, w, h, avatars.get(i),
                                 avatar_position, avatar_width_percent)
            segs.append(seg)
            if progress_callback:
                progress_callback(i, len(pngs))
        concat = td/"concat.txt"
        concat.write_text("\n".join(f"file '{p.as_posix()}'" for p in segs), encoding="utf-8")
        out = td/"presentation_narree.mp4"
        r = subprocess.run([ffmpeg,"-y","-f","concat","-safe","0","-i",str(concat),"-c","copy",str(out)],
                           capture_output=True,text=True,timeout=240)
        if r.returncode != 0 or not out.exists():
            raise RuntimeError(r.stderr[-1500:])
        return out.read_bytes()

st.set_page_config(page_title="Présentation IA V6", page_icon="🎬", layout="wide")
st.title("🎬 Présentation IA — V6")
st.caption("OPENAI + Edge TTS + génération MP4 • avatar vidéo facultatif • formats .pptx et .pptm")

with st.sidebar:
    key = get_openai_key()

    if not key:
        key = st.text_input(
            "Clé API OpenAI",
            type="password"
        )

    model = "gpt-5.6-terra"

    st.sidebar.success(
        "Modèle de narration : GPT-5.6 Terra"
    )
    target_seconds = st.slider("Durée cible par slide", 15, 90, 40, 5)
    ignored_text = st.text_area("Expressions à ignorer", "Département douane\nTitre de la présentation\nÉmetteur")
    ignored = [x.strip() for x in ignored_text.splitlines() if x.strip()]
    voice_label = st.selectbox("Voix française", list(EDGE_VOICES.keys()))
    voice = EDGE_VOICES[voice_label]
    rate_i = st.slider("Vitesse (%)", -30, 30, -5, 5)
    pitch_i = st.slider("Hauteur (Hz)", -20, 20, 0, 5)
    rate, pitch = f"{rate_i:+d}%", f"{pitch_i:+d}Hz"
    resolution = st.selectbox("Résolution", ["1280x720", "1920x1080"])

uploaded = st.file_uploader("Dépose ton PowerPoint", type=["pptx", "pptm"])
if not uploaded:
    st.stop()

powerpoint_bytes = uploaded.getvalue()
extension = Path(uploaded.name).suffix.lower()
file_hash = hashlib.md5(powerpoint_bytes).hexdigest()

if extension not in [".pptx", ".pptm"]:
    st.error("Format non pris en charge. Utilise un fichier .pptx ou .pptm.")
    st.stop()

try:
    slides = clean_slides(extract_slides(powerpoint_bytes), ignored)
except Exception as exc:
    st.error(
        "Impossible de lire cette présentation. "
        "Certains fichiers .pptm très particuliers peuvent poser problème.\n\n"
        f"Détail : {exc}"
    )
    st.stop()

if st.session_state.get("file_hash") != file_hash:
    st.session_state["file_hash"] = file_hash
    st.session_state["narrations"] = {}
    st.session_state["audios"] = {}
    st.session_state["video"] = None
    st.session_state["avatar_bindings"] = {}
    st.session_state["audio_export"] = None
    st.session_state["audio_export_fingerprint"] = None
    st.session_state["video_fingerprint"] = None

narr = st.session_state["narrations"]
audios = st.session_state["audios"]
tabs = st.tabs(["1. Contenu", "2. Narrations", "3. Voix", "4. Avatar", "5. Vidéo"])

with tabs[0]:
    for s in slides:
        with st.expander(f"Slide {s['number']}"):
            st.text_area("Texte", s["clean_text"], height=140, disabled=True, key=f"src{s['number']}")

with tabs[1]:
    if st.button("✨ Générer toutes les narrations", type="primary", disabled=not key):
        audios.clear()
        st.session_state["video"] = None
        prev = ""
        bar = st.progress(0)
        for i, s in enumerate(slides):
            if s["clean_text"]:
                narr[s["number"]] = generate_narration(key, model, narration_prompt(s, prev, target_seconds))
                prev = narr[s["number"]]
            bar.progress((i+1)/len(slides))
        st.session_state["narrations"] = narr
        st.success("Narrations générées.")
    for s in slides:
        n = s["number"]
        cur = narr.get(n, "")
        edit = st.text_area(f"Slide {n}", cur, height=130, key=f"narr{n}_{hash(cur)}")
        if edit != cur:
            narr[n] = edit
            audios.pop(n, None)
            st.session_state["video"] = None

with tabs[2]:
    st.info("Edge TTS ne charge aucun modèle lourd en mémoire.")
    with st.expander("Reprendre les audios sauvegardés"):
        saved_audio_zip = st.file_uploader("ZIP audio exporté depuis cette application", type=["zip"],
                                           key=f"audio_restore_{file_hash}")
        if st.button("Restaurer les audios et les narrations", disabled=saved_audio_zip is None):
            try:
                restored_audios, restored_narrations = load_audio_bundle(
                    saved_audio_zip.getvalue(), [s["number"] for s in slides],
                    hashlib.sha256(powerpoint_bytes).hexdigest(),
                )
                audios.clear(); audios.update(restored_audios)
                narr.clear(); narr.update(restored_narrations)
                st.session_state["video"] = None
                st.session_state["audio_export_fingerprint"] = None
                st.rerun()
            except RuntimeError as exc:
                st.error(str(exc))
    if st.button("🔊 Générer tous les audios", type="primary"):
        st.session_state["video"] = None
        bar = st.progress(0)
        for i, s in enumerate(slides):
            n = s["number"]
            text = narr.get(n, "").strip()
            if text:
                audios[n] = mp3_to_wav(edge_audio(text, voice, rate, pitch))
            bar.progress((i+1)/len(slides))
        st.session_state["audios"] = audios
        st.success("Audios générés.")
    if audios:
        export_fingerprint = tuple((n, audio_hash(data), narr.get(n, "")) for n, data in sorted(audios.items()))
        if st.session_state.get("audio_export_fingerprint") != export_fingerprint:
            st.session_state["audio_export"] = audio_bundle(
                audios, narr, [s["number"] for s in slides], hashlib.sha256(powerpoint_bytes).hexdigest(),
            )
            st.session_state["audio_export_fingerprint"] = export_fingerprint
        st.download_button("⬇️ Télécharger les audios pour l'avatar (ZIP)",
                           st.session_state["audio_export"], "audios_presentation.zip", "application/zip")
        st.caption("Le ZIP contient les WAV exacts, les narrations et les durées par diapositive.")
    for s in slides:
        if s["number"] in audios:
            st.write(f"Diapositive {s['number']}")
            st.audio(audios[s["number"]], format="audio/wav")

avatar_files, avatar_issues = {}, []
avatar_position, avatar_width_percent = AVATAR_POSITIONS[0], 22
with tabs[3]:
    use_avatars = st.checkbox("Ajouter un avatar à la vidéo", key=f"use_avatars_{file_hash}")
    if use_avatars:
        st.info("Utilise les WAV de l'onglet Voix pour créer une vidéo d'avatar par diapositive. "
                "Importe ensuite ces vidéos ici, sans modifier leur vitesse ou ajouter une introduction.")
        avatar_position = st.selectbox("Position de l'avatar", AVATAR_POSITIONS,
                                       key=f"avatar_position_{file_hash}")
        avatar_width_percent = st.slider("Largeur maximale de l'avatar (%)", 15, 35, 22,
                                         key=f"avatar_width_{file_hash}")
        st.caption("À droite ou à gauche : un espace est réservé à l'avatar pour garder tout le texte visible. "
                   "Dans un coin : l'avatar se superpose à la diapositive.")
        uploads = st.file_uploader("Vidéos d'avatar (MP4 ou WebM)", type=["mp4", "webm"],
                                   accept_multiple_files=True, key=f"avatar_uploads_{file_hash}")
        st.caption("Nomme les fichiers slide_001.mp4, slide_002.mp4… pour les associer automatiquement.")
        slide_numbers = [s["number"] for s in slides]
        bindings = st.session_state.setdefault("avatar_bindings", {})
        options = [None] + slide_numbers
        for index, upload in enumerate(uploads):
            clip_sha = hashlib.sha256(upload.getbuffer()).hexdigest()
            guessed = guess_avatar_slide(upload.name, slide_numbers)
            n = st.selectbox(f"Diapositive correspondant à {upload.name}", options,
                             index=options.index(guessed),
                             format_func=lambda x: "À associer" if x is None else f"Diapositive {x}",
                             key=f"avatar_slide_{file_hash}_{index}_{clip_sha}")
            if n is None:
                avatar_issues.append(f"Associe {upload.name} à une diapositive, ou retire ce fichier.")
                continue
            if n in avatar_files:
                avatar_issues.append(f"Deux vidéos sont associées à la diapositive {n}. Garde-en une seule.")
                continue
            avatar_files[n] = upload
            binding = bindings.get(n, {})
            if binding.get("clip_sha") != clip_sha and n in audios:
                bindings[n] = {"clip_sha": clip_sha, "audio_sha": audio_hash(audios[n])}
        avatar_issues.extend(avatar_binding_issues(avatar_files, audios, bindings))
        if not avatar_files:
            avatar_issues.append("Importe au moins une vidéo d'avatar pour activer cette option.")
        for issue in avatar_issues:
            st.error(issue)
        if avatar_files and not avatar_issues:
            st.success(f"Avatar associé à {len(avatar_files)} diapositive(s) sur {len(slides)}.")
            if len(avatar_files) < len(slides):
                st.caption("Les autres diapositives seront générées sans avatar.")
            with st.expander("Prévisualiser un avatar"):
                preview_n = st.selectbox("Diapositive à prévisualiser", sorted(avatar_files),
                                          key=f"avatar_preview_{file_hash}")
                st.video(avatar_files[preview_n].getvalue())
        st.caption("Si tu changes une narration ou sa voix, regénère l'audio et sa vidéo d'avatar.")

with tabs[4]:
    libreoffice_ok = bool(find_cmd(["libreoffice", "soffice"]))
    ffmpeg_ok = bool(find_cmd(["ffmpeg"]))
    ffprobe_ok = bool(find_cmd(["ffprobe"]))
    st.write("LibreOffice :", "✅" if libreoffice_ok else "❌")
    st.write("FFmpeg :", "✅" if ffmpeg_ok else "❌")
    if use_avatars:
        st.write("FFprobe :", "✅" if ffprobe_ok else "❌")
    video_fingerprint = (
        file_hash, resolution, tuple((n, audio_hash(data)) for n, data in sorted(audios.items())),
        use_avatars, avatar_position, avatar_width_percent,
        tuple((n, hashlib.sha256(upload.getbuffer()).hexdigest()) for n, upload in sorted(avatar_files.items())),
    )
    if st.session_state.get("video_fingerprint") != video_fingerprint:
        st.session_state["video"] = None
    can_build = libreoffice_ok and ffmpeg_ok and (not use_avatars or (ffprobe_ok and not avatar_issues))
    if st.button("🎬 Générer le MP4", type="primary", disabled=not can_build):
        try:
            bar = st.progress(0)
            with st.spinner("Création de la vidéo…"):
                st.session_state["video"] = build_video(
                    powerpoint_bytes, audios, len(slides), resolution, extension,
                    avatar_files=avatar_files if use_avatars else None,
                    avatar_position=avatar_position, avatar_width_percent=avatar_width_percent,
                    progress_callback=lambda current, total: bar.progress(current / total),
                )
            st.session_state["video_fingerprint"] = video_fingerprint
            st.success("Vidéo générée.")
        except (RuntimeError, ValueError, subprocess.TimeoutExpired) as exc:
            st.session_state["video"] = None
            st.error(f"La vidéo n'a pas pu être générée : {exc}")
    if st.session_state.get("video"):
        st.video(st.session_state["video"])
        st.download_button("⬇️ Télécharger la vidéo", st.session_state["video"], "presentation_narree.mp4", "video/mp4")
