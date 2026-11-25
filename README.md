# Presentation Proficiency Assessment

A React/Vite frontend and FastAPI backend for analyzing presentation slide text, practicing delivery, and comparing a typed speaking transcript with slide content using Sentence-BERT (`all-MiniLM-L6-v2`).

## Current status

This repository is a **working prototype** aligned with the larger "AI-Powered Multimodal Presentation Coach" proposal. It implements the core demo flow and several feasible multimodal features, while clearly marking research-grade components that remain future work.

### Implemented

- **Slide upload & extraction**: `.pptx` and `.pdf` text extraction (JSZip/XML for PPTX, PDF.js for PDF).
- **Original slide preview**: PPTX rendered in-browser via `pptx-browser`; PDF via PDF.js canvas.
- **Slide analysis**: heuristic content/readability/structure/visual scoring and per-slide audit including PPTX structural signals (image count, font-size, bullets, shapes) when available.
- **Practice mode**: independent camera/microphone controls, live preview, audio recording/playback/download, practice timer.
- **Speech capture**: browser SpeechRecognition when available, editable transcript fallback.
- **Delivery metrics**: word count, WPM, filler-word count, keyword overlap with slide content, pause/silence analysis from recorded audio, delivery score.
- **Behavioral estimates (prototype)**: MediaPipe Gesture Recognizer for hand poses, Face Landmarker for camera-facing/head-orientation and facial expression blendshapes (smile, brow tension, jaw open, dominant expression), Pose Landmarker for posture estimate.
- **Semantic alignment**: FastAPI `/compare` endpoint using Sentence-BERT cosine similarity; per-slide covered/missing points, missing key terms, recommendations.
- **Progress tracking**: last 10 sessions saved in browser `localStorage`, results table, and progress chart.
- **Export**: JSON assessment report including slide, practice, behavior, semantic, and session history.

### Remaining research scope (from proposal)

- Whisper timestamped transcription, pause/silence analysis, Wav2Vec2/HuBERT voice features.
- ViT/SlideAudit-based visual slide-quality model; layout/contrast/font/alignment CV checks.
- Real-time slide-speech alignment during delivery.
- Multimodal fusion model (Transformer/MLP) trained and validated against human ratings (MAE/RMSE/R²/correlation).
- Persistent backend (auth, PostgreSQL, video/file storage), multi-user progress, deployment.

## Recommended setup

- Windows 10/11
- Node.js 18+
- Python 3.11+ (Windows `py -3` launcher recommended)
- Internet access for package install and first-time model/MediaPipe asset downloads

## Start the project

### 1. Backend

1. Open `backend/` in a terminal.
2. Run `start_backend.bat` (or create venv, install `requirements.txt`, run uvicorn).
3. Wait for Uvicorn at `http://127.0.0.1:8011`.
4. Check `http://127.0.0.1:8011/` or `http://127.0.0.1:8011/docs`.

`backend/start_backend.bat` creates a venv if needed, installs requirements, and starts the API. First install/model download can take several minutes.

### 2. Frontend

1. Open a second terminal in the project root (`package.json` folder).
2. Run `start_frontend.bat`, or `npm install` then `npm run dev`.
3. Open the Vite URL, normally `http://localhost:5173`.

## Suggested demo sequence

1. **Dashboard** – introduce the prototype and upload a `.pptx` or `.pdf`.
2. **Slide Analysis** – review extracted slide count, heuristic scores, and slide-by-slide audit.
3. **Semantic Alignment** – paste/type the transcript and run Sentence-BERT comparison; inspect covered/missing points and recommendations.
4. **Practice** – start camera and microphone separately; observe live slide preview, MediaPipe gesture/face/posture estimates, transcript/audio capture, and delivery metrics.
5. **Results** – review combined summary, behavioral estimates, progress chart/history, and export JSON report.

## Browser & permission notes

- Use Chrome/Edge on `http://localhost:5173` (not `file://`).
- Allow Camera and Microphone separately in browser site permissions.
- Speech recognition support varies by browser and may require internet; audio recording and manual transcript remain available.
- MediaPipe models load from CDN on first use; camera should have good lighting and visible face/hands.

## Troubleshooting

- **Port 8011 in use**: stop the old backend process before starting a new one.
- **Frontend cannot reach backend**: confirm `http://127.0.0.1:8011/` returns JSON.
- **Model download fails**: check internet access; first `/compare` request may take longer.
- **Camera/mic issues**: check browser permissions, close other apps using devices, use localhost.
- **PPTX preview differs**: browser rendering may not be pixel-identical to PowerPoint.

## Important limitations

- Slide visual score and many slide checks are **heuristic/prototype estimates**, not trained computer-vision assessment.
- Gaze/head/posture/gesture outputs are **prototype estimates** from MediaPipe; they are not human-validated or clinical measurements.
- Semantic alignment score is **cosine-similarity estimate**, not a definitive judgment of presentation quality.
- Browser speech recognition is not Whisper; timestamps/pauses/voice features are not fully implemented.
- No persistent database/auth/multi-user storage; session history is local to the browser.

## Project structure

```text
presentation_proficiency_project_fixed/
  backend/
    app.py                 # FastAPI + Sentence-BERT /compare
    requirements.txt
    start_backend.bat
  src/
    main.jsx               # React app
    styles.css
  index.html
  package.json
  start_frontend.bat
  README.md
```
