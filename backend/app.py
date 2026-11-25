from __future__ import annotations

import os
import re
import tempfile
from typing import List

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field

try:
    from sentence_transformers import SentenceTransformer
    from sklearn.metrics.pairwise import cosine_similarity
    import numpy as np
except Exception as exc:  # surfaced clearly when endpoint is called
    SentenceTransformer = None
    cosine_similarity = None
    np = None
    MODEL_IMPORT_ERROR = str(exc)
else:
    MODEL_IMPORT_ERROR = None

app = FastAPI(title="Presentation Semantic Alignment API", version="1.0.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://127.0.0.1:5173"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

MODEL_NAME = "all-MiniLM-L6-v2"
model = None
WHISPER_MODEL_NAME = "base"
whisper_model = None

class CompareRequest(BaseModel):
    slides: List[str] = Field(..., min_length=1)
    transcript: str = Field(..., min_length=1)


def clean_words(text: str) -> list[str]:
    stop = {
        "the", "and", "for", "with", "that", "this", "from", "into", "are", "was",
        "were", "will", "would", "could", "should", "have", "has", "had", "you", "your",
        "about", "there", "their", "they", "them", "then", "than", "also", "can", "not",
        "but", "all", "any", "our", "out", "use", "using", "used", "its", "it's", "a", "an",
        "to", "of", "in", "on", "at", "by", "is", "be", "as", "or", "it", "we", "i"
    }
    words = re.findall(r"[A-Za-z][A-Za-z0-9'-]*", text.lower())
    return [w for w in words if len(w) > 2 and w not in stop]


def split_points(slide_text: str) -> list[str]:
    # Keep bullets/new lines as separate points where possible, otherwise split sentences.
    pieces = re.split(r"[\n•;]+|(?<=[.!?])\s+", slide_text)
    points = [re.sub(r"^\s*[-*\d.)]+\s*", "", p).strip() for p in pieces]
    return [p for p in points if len(clean_words(p)) >= 2][:12]


def get_model():
    global model
    if MODEL_IMPORT_ERROR:
        raise HTTPException(status_code=500, detail=f"Model dependencies could not be loaded: {MODEL_IMPORT_ERROR}")
    if model is None:
        model = SentenceTransformer(MODEL_NAME)
    return model

@app.get("/")
def root():
    return {"message": "Presentation Semantic Alignment API is running", "model": MODEL_NAME, "status": "ready"}

@app.get("/health")
def health():
    return {"status": "ok", "model": MODEL_NAME}

@app.post("/compare")
def compare(payload: CompareRequest):
    transcript = payload.transcript.strip()
    if not transcript:
        raise HTTPException(status_code=400, detail="Transcript cannot be empty.")
    slide_texts = [str(s or "").strip() for s in payload.slides]
    if not slide_texts:
        raise HTTPException(status_code=400, detail="Upload a presentation with readable slide text.")

    sentence_model = get_model()
    transcript_units = [s.strip() for s in re.split(r"(?<=[.!?])\s+|\n+", transcript) if len(clean_words(s)) >= 2]
    if not transcript_units:
        transcript_units = [transcript]
    transcript_embeddings = sentence_model.encode(transcript_units, normalize_embeddings=True)

    results = []
    scores = []
    for idx, slide_text in enumerate(slide_texts, start=1):
        points = split_points(slide_text)
        if not slide_text or not points:
            results.append({
                "slide_number": idx, "slide_text": slide_text, "score": 0,
                "feedback": "No readable key points were found on this slide.",
                "covered_topics": [], "missing_topics": [], "missing_key_terms": [],
                "recommendations": ["Add readable text or explain the main ideas from this slide."],
                "point_analysis": []
            })
            scores.append(0)
            continue

        point_embeddings = sentence_model.encode(points, normalize_embeddings=True)
        matrix = cosine_similarity(point_embeddings, transcript_embeddings)
        point_analysis = []
        covered, missing = [], []
        point_scores = []
        for point, similarities in zip(points, matrix):
            best = float(max(similarities))
            score = max(0, min(100, round(best * 100)))
            point_scores.append(score)
            entry = {"topic": point, "score": score, "covered": score >= 48}
            point_analysis.append(entry)
            (covered if score >= 48 else missing).append(point)

        score = round(sum(point_scores) / len(point_scores)) if point_scores else 0
        scores.append(score)
        slide_terms = list(dict.fromkeys(clean_words(slide_text)))
        transcript_terms = set(clean_words(transcript))
        missing_terms = [w for w in slide_terms if w not in transcript_terms][:10]
        recommendations = []
        if missing:
            recommendations.append("Explain the missing points explicitly: " + "; ".join(missing[:3]) + ("." if len(missing) <= 3 else "; and other points."))
        if missing_terms:
            recommendations.append("Consider mentioning these slide terms where relevant: " + ", ".join(missing_terms[:8]) + ".")
        if score < 45:
            recommendations.append("Make the connection between your spoken explanation and this slide clearer.")
        elif score < 70:
            recommendations.append("Add a little more detail to strengthen the explanation of this slide.")
        else:
            recommendations.append("Good alignment. Keep the explanation focused and support key points with examples where useful.")
        feedback = f"Estimated semantic alignment is {score}%. {len(covered)} point(s) appear covered and {len(missing)} may need more explanation."
        results.append({
            "slide_number": idx, "slide_text": slide_text, "score": score,
            "feedback": feedback, "covered_topics": covered, "missing_topics": missing,
            "missing_key_terms": missing_terms, "recommendations": recommendations,
            "point_analysis": point_analysis
        })

    overall = round(sum(scores) / len(scores)) if scores else 0
    review = [r["slide_number"] for r in results if r["score"] < 50]
    return {
        "overall_score": overall,
        "total_slides": len(slide_texts),
        "results": results,
        "slides_needing_review": review,
        "note": "Scores are estimates from semantic similarity, not a definitive judgement of presentation quality. Missing terms are heuristic suggestions; review them in context."
    }

def get_whisper_model():
    global whisper_model
    if whisper_model is None:
        try:
            from faster_whisper import WhisperModel
        except Exception as exc:
            raise HTTPException(status_code=500, detail=f"faster-whisper is not available: {exc}")
        # CPU int8 is safest default on typical student machines
        whisper_model = WhisperModel(WHISPER_MODEL_NAME, device="cpu", compute_type="int8")
    return whisper_model


@app.post("/transcribe")
async def transcribe_audio(file: UploadFile = File(...)):
    filename = file.filename or "audio.webm"
    suffix = os.path.splitext(filename)[1] or ".webm"
    try:
        content = await file.read()
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Could not read uploaded audio: {exc}")
    if not content:
        raise HTTPException(status_code=400, detail="Uploaded audio file is empty.")

    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(content)
            tmp_path = tmp.name
        wm = get_whisper_model()
        # Prefer decoding WAV via stdlib to avoid PyAV/ffmpeg issues.
        audio_input = tmp_path
        try:
            if suffix.lower() == ".wav":
                import wave
                import numpy as np
                with wave.open(tmp_path, "rb") as wf:
                    n_channels = wf.getnchannels()
                    sampwidth = wf.getsampwidth()
                    framerate = wf.getframerate()
                    n_frames = wf.getnframes()
                    pcm = wf.readframes(n_frames)
                if sampwidth == 2:
                    audio_np = np.frombuffer(pcm, dtype=np.int16).astype(np.float32) / 32768.0
                elif sampwidth == 1:
                    audio_np = (np.frombuffer(pcm, dtype=np.uint8).astype(np.float32) - 128.0) / 128.0
                else:
                    audio_np = None
                if audio_np is not None and n_channels > 1:
                    audio_np = audio_np.reshape(-1, n_channels).mean(axis=1)
                if audio_np is not None and framerate != 16000:
                    # simple linear resample to 16k for whisper
                    duration = audio_np.shape[0] / float(framerate)
                    target_len = int(duration * 16000)
                    if target_len > 0:
                        x_old = np.linspace(0, 1, num=audio_np.shape[0], endpoint=False)
                        x_new = np.linspace(0, 1, num=target_len, endpoint=False)
                        audio_np = np.interp(x_new, x_old, audio_np).astype(np.float32)
                if audio_np is not None:
                    audio_input = audio_np
        except Exception:
            audio_input = tmp_path

        segments_iter, info = wm.transcribe(audio_input, beam_size=1, vad_filter=True)
        segments = []
        texts = []
        for seg in segments_iter:
            text = (seg.text or "").strip()
            if text:
                texts.append(text)
                segments.append({
                    "start": round(float(seg.start), 2),
                    "end": round(float(seg.end), 2),
                    "text": text,
                })
        transcript = " ".join(texts).strip()
        return {
            "transcript": transcript,
            "segments": segments,
            "language": getattr(info, "language", None),
            "duration": round(float(getattr(info, "duration", 0.0) or 0.0), 2),
            "model": WHISPER_MODEL_NAME,
        }
    except HTTPException:
        raise
    except Exception as exc:
        msg = str(exc)
        if "ffmpeg" in msg.lower() or "av" in msg.lower():
            msg = f"{msg}. ffmpeg may be required for this audio format."
        raise HTTPException(status_code=500, detail=f"Transcription failed: {msg}")
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass
