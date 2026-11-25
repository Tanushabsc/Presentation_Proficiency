

import React, { useEffect, useMemo, useRef, useState } from 'react';

import { createRoot } from 'react-dom/client';

import * as pdfjsLib from 'pdfjs-dist';
import { FaceLandmarker, FilesetResolver, GestureRecognizer, PoseLandmarker } from '@mediapipe/tasks-vision';

import JSZip from 'jszip';
import PptxGenJS from 'pptxgenjs';
import { PptxRenderer } from 'pptx-browser';
import { CartesianGrid, Line, LineChart, ResponsiveContainer, Tooltip, XAxis, YAxis } from 'recharts';

import './styles.css';



pdfjsLib.GlobalWorkerOptions.workerSrc =

  'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.worker.min.mjs';



const fillers = [

  'um',

  'uh',

  'like',

  'you know',

  'actually',

  'basically',

  'so'

];



function tokenize(text) {

  return text.toLowerCase().match(/[a-z0-9]+/g) || [];

}



function keywordSet(text) {

  const stopWords = new Set([

    'the', 'a', 'an', 'is', 'are', 'of', 'to', 'in',

    'and', 'or', 'for', 'on', 'with', 'this', 'that',

    'as', 'by', 'from', 'be', 'was', 'were', 'it',

    'at', 'we', 'our', 'can', 'will'

  ]);



  return new Set(

    tokenize(text).filter(

      word => word.length > 2 && !stopWords.has(word)

    )

  );

}



function similarity(a, b) {

  const A = keywordSet(a);

  const B = keywordSet(b);



  if (!A.size || !B.size) return 0;



  let common = 0;



  for (const word of A) {

    if (B.has(word)) common++;

  }



  return Math.round(

    (common / Math.max(A.size, B.size)) * 100

  );

}



function countFillers(text) {

  const lower = text.toLowerCase();



  return fillers.reduce((total, filler) => {

    const pattern = new RegExp(

      `\\b${filler.replace(/ /g, '\\s+')}\\b`,

      'g'

    );



    return total + (lower.match(pattern) || []).length;

  }, 0);

}



function countWords(text) {

  return tokenize(text).length;

}



async function extractPdf(file) {

  const buffer = await file.arrayBuffer();

  const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;

  const pages = [];



  for (let i = 1; i <= pdf.numPages; i++) {

    const page = await pdf.getPage(i);

    const content = await page.getTextContent();



    pages.push(

      content.items.map(item => item.str).join(' ')

    );

  }



  return pages;

}



async function extractPptx(file) {

  const buffer = await file.arrayBuffer();

  const zip = await JSZip.loadAsync(buffer);



  const names = Object.keys(zip.files)

    .filter(name => /^ppt\/slides\/slide\d+\.xml$/.test(name))

    .sort((a, b) => {

      const numberA = Number(a.match(/slide(\d+)/)[1]);

      const numberB = Number(b.match(/slide(\d+)/)[1]);

      return numberA - numberB;

    });



  const slides = [];



  for (const name of names) {

    const xml = await zip.file(name).async('text');



    const texts = [

      ...xml.matchAll(/<a:t>([\s\S]*-)<\/a:t>/g)

    ].map(match =>

      match[1]

        .replace(/&amp;/g, '&')

        .replace(/&lt;/g, '<')

        .replace(/&gt;/g, '>')

        .replace(/&quot;/g, '"')

        .replace(/&apos;/g, "'")

    );

    const fontSizes = [...xml.matchAll(/<a:rPr[^>]*\ssz="(\d+)"/g)].map(m => Number(m[1]) / 100).filter(n => n > 0);
    const images = (xml.match(/<p:pic\b/g) || []).length + (xml.match(/<a:blip\b/g) || []).length;
    const bullets = (xml.match(/<a:buChar\b|<a:buAutoNum\b|<a:buBlip\b/g) || []).length;
    const shapes = (xml.match(/<p:sp\b/g) || []).length;
    const textBoxes = (xml.match(/<p:txBody>/g) || []).length;

    slides.push({
      text: texts.join(' '),
      meta: {
        images,
        bullets,
        shapes,
        textBoxes,
        minFontPt: fontSizes.length ? Math.min(...fontSizes) : null,
        maxFontPt: fontSizes.length ? Math.max(...fontSizes) : null,
        avgFontPt: fontSizes.length ? Math.round(fontSizes.reduce((a,b)=>a+b,0) / fontSizes.length) : null
      }
    });

  }



  return slides;

}



function analyzeSlides(slidesInput) {
  const normalized = (slidesInput || []).map(s => typeof s === 'string' ? { text: s, meta: {} } : { text: s?.text || '', meta: s?.meta || {} });
  const texts = normalized.map(s => s.text);
  const text = texts.join(' ');
  const words = countWords(text);
  const average = normalized.length ? Math.round(words / normalized.length) : 0;

  const slideDetails = normalized.map((slide, index) => {
    const clean = String(slide.text || '').trim();
    const meta = slide.meta || {};
    const slideWords = countWords(clean);
    const lines = clean.split(/[\n*]+/).map(line => line.trim()).filter(Boolean);
    const sentences = (clean.match(/[.!-]+/g) || []).length;
    const title = lines[0]?.slice(0, 90) || `Slide ${index + 1}`;
    const excerpt = clean.length > 180 ? `${clean.slice(0, 177).trim()}...` : clean;
    const keyTerms = [...new Set(tokenize(clean).filter(w => w.length > 4))].slice(0, 6);
    const bodyLines = lines.slice(1).filter(Boolean);
    const longLines = bodyLines.filter(l => countWords(l) > 18);
    const suggestions = [];

    if (slideWords > 55) suggestions.push({
      priority: 'High', area: 'Text density',
      issue: `"${title}" contains about ${slideWords} words, which is dense for one slide.`,
      why: 'Your audience will try to read the slide while listening to you, which reduces attention and retention.',
      action: `Reduce this slide to one main message and 3-5 short bullets. Keep key terms like ${keyTerms.slice(0,3).join(', ') || 'the main topic'} and move supporting detail to speaker notes.`,
      example: bodyLines[0] ? `Turn "${bodyLines[0].slice(0, 80)}${bodyLines[0].length > 80 ? '...' : ''}" into a short bullet with only the key point.` : `Use a headline "${title}" and 3 short supporting bullets.`
    });

    if (slideWords < 5) suggestions.push({
      priority: 'Medium', area: 'Content clarity',
      issue: `"${title}" has very little readable text (${slideWords} words detected).`,
      why: 'If this slide is visual-heavy, the audience may not understand the takeaway without a clear cue.',
      action: `Add a clear headline and one takeaway sentence for "${title}". If it is an image/chart slide, state what the audience should notice.`,
      example: `Add: "${title}: key takeaway" and one line explaining what the visual proves.`
    });

    if (longLines.length) suggestions.push({
      priority: 'Medium', area: 'Long bullet/text',
      issue: `${longLines.length} text block(s) on "${title}" are long and may be hard to scan.`,
      why: 'Long bullets look like paragraphs and slow down both reading and delivery.',
      action: `Shorten long bullets on "${title}" to one idea each. Keep the keyword and remove extra explanation.`,
      example: `Rewrite "${longLines[0].slice(0, 70)}${longLines[0].length > 70 ? '...' : ''}" as a 6-12 word bullet.`
    });

    if (!/[.!-]/.test(clean) && slideWords > 20) suggestions.push({
      priority: 'Medium', area: 'Readability',
      issue: `The content on "${title}" reads as unstructured text.`,
      why: 'Structure helps the audience quickly understand the point of the slide.',
      action: `Organize "${title}" into a short heading and bullets. Use one idea per line.`,
      example: keyTerms.length ? `Use bullets around: ${keyTerms.slice(0,3).join(' * ')}.` : 'Use: "* Point 1 * Point 2 * Point 3".'
    });

    if (lines.length === 1 && slideWords > 18) suggestions.push({
      priority: 'Medium', area: 'Structure',
      issue: `"${title}" appears as a single large text block.`,
      why: 'A single block makes it harder for the audience to follow your spoken explanation.',
      action: `Split "${title}" into a title and 2-4 supporting bullets.`,
      example: `Make "${title}" the heading, then add 2-4 bullets underneath it.`
    });

    if (meta.minFontPt && meta.minFontPt < 18) suggestions.push({
      priority: 'High', area: 'Font size',
      issue: `Small font detected on "${title}" (about ${meta.minFontPt}pt).`,
      why: 'Small text is hard to read in live presentations and on shared screens.',
      action: `Increase body text on "${title}" to at least 24pt and keep headings larger. Reduce wording if needed.`,
      example: `Keep the title, then shorten body text so it fits at 24-28pt.`
    });

    if (meta.images === 0 && slideWords > 30) suggestions.push({
      priority: 'Low', area: 'Visual balance',
      issue: `"${title}" is text-heavy with no detected image/visual.`,
      why: 'A relevant visual can reinforce the message and make the slide easier to remember.',
      action: `Add a simple visual that supports "${title}" - a chart, icon, diagram, or image related to ${keyTerms.slice(0,2).join(' / ') || 'the topic'}.`,
      example: `For "${title}", use a simple diagram or icon next to 3 short bullets.`
    });

    if (meta.bullets > 8) suggestions.push({
      priority: 'Medium', area: 'Bullet overload',
      issue: `"${title}" has many bullet items detected (${meta.bullets}).`,
      why: 'Too many bullets make the slide crowded and difficult to present clearly.',
      action: `Split "${title}" into two slides or group bullets into 3-5 categories.`,
      example: keyTerms.length ? `Group bullets into categories such as ${keyTerms.slice(0,3).join(', ')}.` : 'Split into "Slide A: Overview" and "Slide B: Details".'
    });

    if (!suggestions.length) suggestions.push({
      priority: 'Good', area: 'Overall',
      issue: `"${title}" looks reasonably balanced based on text/structure checks.`,
      why: 'No major text-density or structure issue was detected by the prototype checker.',
      action: `When presenting "${title}", emphasize the key takeaway and connect it to the next slide.`,
      example: keyTerms.length ? `Say: "The key takeaway from ${title} is ${keyTerms[0]}..."` : `Say: "The key takeaway from this slide is..."`
    });

    const density = slideWords > 55 ? 'High' : slideWords < 5 ? 'Low' : 'Moderate';
    return {
      slide: index + 1, title, words: slideWords, textBlocks: lines.length, sentences, density,
      images: meta.images ?? null, bullets: meta.bullets ?? null, shapes: meta.shapes ?? null,
      minFontPt: meta.minFontPt ?? null, avgFontPt: meta.avgFontPt ?? null,
      excerpt, keyTerms,
      suggestions,
      issues: suggestions.map(s => `${s.area}: ${s.action}`)
    };
  });

  const highDensity = slideDetails.filter(d => d.density === 'High').length;
  const tinyFontSlides = slideDetails.filter(d => d.minFontPt && d.minFontPt < 18).length;
  const imageHeavyBalance = slideDetails.filter(d => (d.images || 0) > 0 || d.words < 35).length;
  const scores = {
    content: Math.min(100, 55 + Math.min(35, Math.round(average > 10 ? 25 : average * 2))),
    readability: Math.max(45, 100 - Math.max(0, average - 35) * 2 - tinyFontSlides * 3 - highDensity * 2),
    structure: (normalized.length >= 5 ? 88 : normalized.length >= 3 ? 75 : 55) - Math.min(15, highDensity * 3),
    visual: Math.max(45, 70 + Math.round((imageHeavyBalance / Math.max(1, normalized.length)) * 15) - tinyFontSlides * 4)
  };
  scores.overall = Math.round((scores.content + scores.readability + scores.structure + scores.visual) / 4);

  return { ...scores, words, avg: average, slides: normalized.length, slideDetails };

}



function App() {

  const [page, setPage] = useState('dashboard');



  const [file, setFile] = useState(null);

  const [slides, setSlides] = useState([]);
  const [slideEdits, setSlideEdits] = useState([]);
  useEffect(() => { slideEditsRef.current = slideEdits; }, [slideEdits]);
  const [currentSlide, setCurrentSlide] = useState(0);
  const [pptxReady, setPptxReady] = useState(false);
  const [previewVersion, setPreviewVersion] = useState(0);
  const [slideRenderError, setSlideRenderError] = useState('');
  const [slidePreviewBlank, setSlidePreviewBlank] = useState(false);

  const [analysis, setAnalysis] = useState(null);

  const [error, setError] = useState('');



  // Existing practice feature

  const [recording, setRecording] = useState(false);

  const [transcript, setTranscript] = useState('');

  const [seconds, setSeconds] = useState(0);

  const [support, setSupport] = useState(true);
  const [cameraOn, setCameraOn] = useState(false);
  const [micOn, setMicOn] = useState(false);
  const [cameraMessage, setCameraMessage] = useState('Camera is off.');
  const [micMessage, setMicMessage] = useState('Microphone is off.');
  const [speechError, setSpeechError] = useState('');
  const [audioUrl, setAudioUrl] = useState('');
  const [audioBlob, setAudioBlob] = useState(null);
  const [behaviorChecks, setBehaviorChecks] = useState({eyeContact:false,posture:false,gestures:false,faceVisible:false});
  const [sessionHistory, setSessionHistory] = useState(() => { try { return JSON.parse(localStorage.getItem('presentaiPracticeHistory') || '[]'); } catch { return []; } });
  const [gestureLabel, setGestureLabel] = useState('Waiting for hand detection');
  const [gestureCount, setGestureCount] = useState(0);
  const [gestureEvents, setGestureEvents] = useState([]);
  const [gestureStatus, setGestureStatus] = useState('Gesture model not loaded');
  const [behaviorStatus, setBehaviorStatus] = useState('Face/posture analysis not loaded');
  const [behaviorMetrics, setBehaviorMetrics] = useState({
    gazePercent: 0,
    headStability: 0,
    postureScore: 0,
    faceDetectedPercent: 0,
    frames: 0,
    smilePercent: 0,
    eyeContactStability: 0,
    browTensionPercent: 0,
    jawOpenPercent: 0,
    dominantExpression: 'Neutral'
  });
  const [pauseMetrics, setPauseMetrics] = useState({ pauses: 0, avgPauseMs: 0, longPauses: 0, silencePercent: 0 });
  const [transcribing, setTranscribing] = useState(false);



  // New semantic alignment feature

  const [manualTranscript, setManualTranscript] = useState('');

  const [semanticResult, setSemanticResult] = useState(null);

  const [comparing, setComparing] = useState(false);
  const [backendOk, setBackendOk] = useState(null);



  const videoRef = useRef(null);
  const slideCanvasRef = useRef(null);
  const editCanvasRef = useRef(null);
  const pptxRendererRef = useRef(null);
  const modifiedPptxRef = useRef(null);
  const slideEditsRef = useRef([]);
  const rebuildingRef = useRef(false);

  const streamRef = useRef(null);
  const cameraStreamRef = useRef(null);
  const micStreamRef = useRef(null);
  const mediaRecorderRef = useRef(null);
  const audioChunksRef = useRef([]);
  const gestureRecognizerRef = useRef(null);
  const faceLandmarkerRef = useRef(null);
  const poseLandmarkerRef = useRef(null);
  const gestureLoopRef = useRef(null);
  const lastGestureRef = useRef('');
  const behaviorStatsRef = useRef({
    total: 0, face: 0, gazeFacing: 0, postureAligned: 0,
    yawSum: 0, yawSqSum: 0, pitchSum: 0, pitchSqSum: 0,
    smileSum: 0, blinkSum: 0, browDownSum: 0, jawOpenSum: 0, lookAwaySum: 0
  });

  const recognitionRef = useRef(null);

  const timerRef = useRef(null);

  const recordingRef = useRef(false);



  const speechSupport = useMemo(() => {

    return typeof window !== 'undefined' &&

      !!(

        window.SpeechRecognition ||

        window.webkitSpeechRecognition

      );

  }, []);



  useEffect(() => {

    setSupport(speechSupport);

  }, [speechSupport]);

  useEffect(() => {
    let active = true;
    async function checkBackend() {
      try {
        const res = await fetch('http://127.0.0.1:8011/health');
        if (!active) return;
        setBackendOk(res.ok);
      } catch (_) {
        if (active) setBackendOk(false);
      }
    }
    checkBackend();
    const id = setInterval(checkBackend, 30000);
    return () => { active = false; clearInterval(id); };
  }, []);



  async function handleFile(selectedFile) {

    if (!selectedFile) return;



    setError('');

    setFile(selectedFile);

    setSlides([]);

    setAnalysis(null);

    setSemanticResult(null);



    try {

      let extracted = [];



      if (selectedFile.name.toLowerCase().endsWith('.pdf')) {

        extracted = await extractPdf(selectedFile);

      } else if (

        selectedFile.name.toLowerCase().endsWith('.pptx')

      ) {

        extracted = await extractPptx(selectedFile);

      } else {

        throw new Error('Please upload a .pptx or .pdf file.');

      }



      if (!extracted.length) {

        throw new Error('No readable slide text was found.');

      }

      const normalized = extracted.map(item => typeof item === 'string' ? { text: item, meta: {} } : item);
      const texts = normalized.map(item => item.text || '');

      if (!texts.some(text => text.trim())) {

        throw new Error(

          'No readable text was found. The file may contain image-only slides.'

        );

      }



      setSlides(texts);
      setSlideEdits(texts.map(t => String(t || '')));

      setAnalysis(analyzeSlides(normalized));

      setPage('analysis');

    } catch (err) {

      setError(err.message || 'Unable to read this presentation.');

    }

  }



  // Load the uploaded PowerPoint into a browser-side renderer once per file.
  useEffect(() => {
    let cancelled = false;
    let renderer = null;
    async function loadPowerPoint() {
      if (!file || !file.name.toLowerCase().endsWith('.pptx')) {
        setPptxReady(false);
        return;
      }
      setPptxReady(false);
      setSlideRenderError('Loading PowerPoint slides...');
      try {
        renderer = new PptxRenderer();
        await renderer.load(file);
        if (cancelled) {
          renderer.destroy();
          return;
        }
        if (pptxRendererRef.current) pptxRendererRef.current.destroy();
        pptxRendererRef.current = renderer;
        setPptxReady(true);
        setSlideRenderError('');
      } catch (err) {
        if (!cancelled) setSlideRenderError(`PowerPoint preview could not be loaded: ${err.message || err}`);
      }
    }
    loadPowerPoint();
    return () => {
      cancelled = true;
      if (renderer) {
        if (pptxRendererRef.current === renderer) pptxRendererRef.current = null;
        renderer.destroy();
      }
    };
  }, [file]);

  // Draw the selected original slide into the practice viewer (PPTX or PDF).
  useEffect(() => {
    let cancelled = false;
    async function renderSelectedSlide() {
      const canvas = page === 'edit' ? editCanvasRef.current : slideCanvasRef.current;
      if (!['practice', 'edit'].includes(page) || !file || !canvas) return;
      try {
        if (file.name.toLowerCase().endsWith('.pptx')) {
          if (!pptxReady || !pptxRendererRef.current) return;
          await pptxRendererRef.current.renderSlide(currentSlide, canvas, 1440);
          try {
            const ctx = canvas.getContext('2d', { willReadFrequently: true });
            const { width: w, height: h } = canvas;
            const pts = [[10,10],[w/2,h/2],[w-10,10],[10,h-10],[w/2,h/4],[w/2,3*h/4],[w/4,h/2],[3*w/4,h/2]];
            const blank = w > 0 && pts.every(([x,y]) => { const d = ctx.getImageData(Math.floor(x),Math.floor(y),1,1).data; return d[3] === 0 || (d[0] > 250 && d[1] > 250 && d[2] > 250); });
            if (!cancelled) { setSlidePreviewBlank(blank); setSlideRenderError(blank ? 'Exact PPT preview unavailable for this slide in browser; showing editable content preview.' : ''); }
          } catch (_) { if (!cancelled) { setSlidePreviewBlank(false); setSlideRenderError(''); } }
        } else if (file.name.toLowerCase().endsWith('.pdf')) {
          if (!cancelled) setSlidePreviewBlank(false);
          const pdf = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
          if (cancelled) return;
          const pdfPage = await pdf.getPage(currentSlide + 1);
          if (cancelled) return;
          const context = canvas.getContext('2d');
          const viewport = pdfPage.getViewport({ scale: 1.5 });
          canvas.width = viewport.width;
          canvas.height = viewport.height;
          await pdfPage.render({ canvasContext: context, viewport }).promise;
        }
      } catch (err) {
        if (!cancelled) setSlideRenderError(`Slide ${currentSlide + 1} could not be displayed: ${err.message || err}`);
      }
    }
    renderSelectedSlide();
    return () => { cancelled = true; };
  }, [page, file, currentSlide, pptxReady, previewVersion]);

  useEffect(() => {
    if (page !== 'edit' || !slideEdits.length) return;
    const id = setTimeout(() => { saveSlideEdits({ silent: true }); }, 1200);
    return () => clearTimeout(id);
  }, [slideEdits, page]);

  function matrixToEuler(matrixData) {
    // MediaPipe facialTransformationMatrixes is a 4x4 matrix in column-major-like flat array.
    // We extract approximate yaw/pitch/roll for prototype feedback.
    const m = matrixData || [];
    if (m.length < 16) return null;
    const r00 = m[0], r01 = m[1], r02 = m[2];
    const r10 = m[4], r11 = m[5], r12 = m[6];
    const r20 = m[8], r21 = m[9], r22 = m[10];
    const sy = Math.sqrt(r00 * r00 + r10 * r10);
    const singular = sy < 1e-6;
    let pitch, yaw, roll;
    if (!singular) {
      pitch = Math.atan2(r21, r22);
      yaw = Math.atan2(-r20, sy);
      roll = Math.atan2(r10, r00);
    } else {
      pitch = Math.atan2(-r12, r11);
      yaw = Math.atan2(-r20, sy);
      roll = 0;
    }
    const toDeg = v => (v * 180) / Math.PI;
    return { yaw: toDeg(yaw), pitch: toDeg(pitch), roll: toDeg(roll) };
  }

  function analyzePosture(poseLandmarks) {
    // Simple prototype posture score from shoulders/hips/nose alignment.
    const lm = poseLandmarks?.[0];
    if (!lm) return null;
    const nose = lm[0], lShoulder = lm[11], rShoulder = lm[12], lHip = lm[23], rHip = lm[24];
    if (!lShoulder || !rShoulder || !lHip || !rHip) return null;
    const shoulderMidX = (lShoulder.x + rShoulder.x) / 2;
    const hipMidX = (lHip.x + rHip.x) / 2;
    const shoulderYDiff = Math.abs(lShoulder.y - rShoulder.y);
    const torsoLean = Math.abs(shoulderMidX - hipMidX);
    const noseCenter = nose ? Math.abs(nose.x - shoulderMidX) : 0;
    const aligned = shoulderYDiff < 0.08 && torsoLean < 0.08 && noseCenter < 0.12;
    const score = Math.max(0, Math.min(100, Math.round(100 - shoulderYDiff * 500 - torsoLean * 500 - noseCenter * 250)));
    return { aligned, score };
  }

  function updateBehaviorMetrics() {
    const s = behaviorStatsRef.current;
    if (!s.total) return;
    const gazePercent = Math.round((s.gazeFacing / s.total) * 100);
    const faceDetectedPercent = Math.round((s.face / s.total) * 100);
    const postureScore = Math.round((s.postureAligned / Math.max(1, s.total)) * 100);
    const faceFrames = Math.max(1, s.face);
    const yawMean = s.yawSum / faceFrames;
    const pitchMean = s.pitchSum / faceFrames;
    const yawVar = s.yawSqSum / faceFrames - yawMean * yawMean;
    const pitchVar = s.pitchSqSum / faceFrames - pitchMean * pitchMean;
    const headStability = Math.max(0, Math.min(100, Math.round(100 - Math.sqrt(Math.max(0, yawVar) + Math.max(0, pitchVar)) * 2)));
    const smilePercent = Math.round((s.smileSum / faceFrames) * 100);
    const blinkPercent = Math.round((s.blinkSum / faceFrames) * 100);
    const browTensionPercent = Math.round((s.browDownSum / faceFrames) * 100);
    const jawOpenPercent = Math.round((s.jawOpenSum / faceFrames) * 100);
    const lookAwayPercent = Math.round((s.lookAwaySum / faceFrames) * 100);
    const eyeContactStability = Math.max(0, Math.min(100, 100 - lookAwayPercent));
    const expressions = [
      ['Smiling', smilePercent],
      ['Focused / Tense', browTensionPercent],
      ['Speaking / Mouth Open', jawOpenPercent],
      ['Blinking', blinkPercent],
      ['Neutral', Math.max(0, 100 - Math.max(smilePercent, browTensionPercent, jawOpenPercent, blinkPercent))]
    ];
    const dominantExpression = expressions.sort((a, b) => b[1] - a[1])[0][0];
    setBehaviorMetrics({
      gazePercent, headStability, postureScore, faceDetectedPercent, frames: s.total,
      smilePercent, eyeContactStability, browTensionPercent, jawOpenPercent, dominantExpression
    });
  }

  async function startGestureDetection() {
    try {
      setGestureStatus('Loading MediaPipe hand-gesture model...');
      setBehaviorStatus('Loading face/posture models...');
      behaviorStatsRef.current = { total: 0, face: 0, gazeFacing: 0, postureAligned: 0, yawSum: 0, yawSqSum: 0, pitchSum: 0, pitchSqSum: 0, smileSum: 0, blinkSum: 0, browDownSum: 0, jawOpenSum: 0, lookAwaySum: 0 };
      setBehaviorMetrics({ gazePercent: 0, headStability: 0, postureScore: 0, faceDetectedPercent: 0, frames: 0, smilePercent: 0, eyeContactStability: 0, browTensionPercent: 0, jawOpenPercent: 0, dominantExpression: 'Neutral' });

      const vision = await FilesetResolver.forVisionTasks('https://cdn.jsdelivr.net/npm/@mediapipe/tasks-vision@0.10.14/wasm');

      if (!gestureRecognizerRef.current) {
        const modelOptions = {
          baseOptions: { modelAssetPath: 'https://storage.googleapis.com/mediapipe-models/gesture_recognizer/gesture_recognizer/float16/1/gesture_recognizer.task' },
          runningMode: 'VIDEO', numHands: 2
        };
        try {
          gestureRecognizerRef.current = await GestureRecognizer.createFromOptions(vision, { ...modelOptions, baseOptions: { ...modelOptions.baseOptions, delegate: 'GPU' } });
        } catch (_) {
          gestureRecognizerRef.current = await GestureRecognizer.createFromOptions(vision, modelOptions);
        }
      }
      setGestureStatus('Live hand-gesture detection is active.');

      if (!faceLandmarkerRef.current) {
        try {
          faceLandmarkerRef.current = await FaceLandmarker.createFromOptions(vision, {
            baseOptions: { modelAssetPath: 'https://storage.googleapis.com/mediapipe-models/face_landmarker/face_landmarker/float16/1/face_landmarker.task', delegate: 'GPU' },
            runningMode: 'VIDEO', numFaces: 1, outputFaceBlendshapes: true, outputFacialTransformationMatrixes: true
          });
        } catch (_) {
          try {
            faceLandmarkerRef.current = await FaceLandmarker.createFromOptions(vision, {
              baseOptions: { modelAssetPath: 'https://storage.googleapis.com/mediapipe-models/face_landmarker/face_landmarker/float16/1/face_landmarker.task' },
              runningMode: 'VIDEO', numFaces: 1, outputFaceBlendshapes: true, outputFacialTransformationMatrixes: true
            });
          } catch (e) { setBehaviorStatus(`Face model unavailable: ${e.message}`); }
        }
      }

      if (!poseLandmarkerRef.current) {
        try {
          poseLandmarkerRef.current = await PoseLandmarker.createFromOptions(vision, {
            baseOptions: { modelAssetPath: 'https://storage.googleapis.com/mediapipe-models/pose_landmarker/pose_landmarker_lite/float16/1/pose_landmarker_lite.task', delegate: 'GPU' },
            runningMode: 'VIDEO', numPoses: 1
          });
        } catch (_) {
          try {
            poseLandmarkerRef.current = await PoseLandmarker.createFromOptions(vision, {
              baseOptions: { modelAssetPath: 'https://storage.googleapis.com/mediapipe-models/pose_landmarker/pose_landmarker_lite/float16/1/pose_landmarker_lite.task' },
              runningMode: 'VIDEO', numPoses: 1
            });
          } catch (e) { setBehaviorStatus(`Pose model unavailable: ${e.message}`); }
        }
      }

      if (faceLandmarkerRef.current || poseLandmarkerRef.current) {
        setBehaviorStatus('Live face, expression, head, and posture analysis is active (prototype estimate).');
      }

      const detect = () => {
        const video = videoRef.current;
        const recognizer = gestureRecognizerRef.current;
        if (!video || !recognizer || video.readyState < 2) { gestureLoopRef.current = requestAnimationFrame(detect); return; }
        const now = performance.now();
        try {
          const result = recognizer.recognizeForVideo(video, now);
          const gestures = (result.gestures || []).map((hand, i) => ({name: hand[0]?.categoryName || 'Hand detected', score: hand[0]?.score || 0, handedness: result.handedness?.[i]?.[0]?.categoryName || `Hand ${i+1}`}));
          if (!gestures.length) setGestureLabel('No hand detected - move your hands into the camera view');
          else {
            const label = gestures.map(g => `${g.handedness}: ${g.name}`).join(' · ');
            setGestureLabel(label);
            const primary = gestures[0].name;
            if (primary !== lastGestureRef.current && primary !== 'None') {
              lastGestureRef.current = primary;
              setGestureCount(c => c + 1);
              setGestureEvents(prev => [{name: primary, time: new Date().toLocaleTimeString(), confidence: Math.round(gestures[0].score*100)}, ...prev].slice(0,8));
            }
          }
        } catch (e) { setGestureStatus(`Gesture detection issue: ${e.message}`); }

        try {
          const stats = behaviorStatsRef.current;
          stats.total += 1;
          if (faceLandmarkerRef.current) {
            const faceResult = faceLandmarkerRef.current.detectForVideo(video, now);
            const matrix = faceResult.facialTransformationMatrixes?.[0]?.data;
            const blendshapes = faceResult.faceBlendshapes?.[0]?.categories || [];
            if (matrix && matrix.length) {
              stats.face += 1;
              const euler = matrixToEuler(Array.from(matrix));
              if (euler) {
                stats.yawSum += euler.yaw; stats.yawSqSum += euler.yaw * euler.yaw;
                stats.pitchSum += euler.pitch; stats.pitchSqSum += euler.pitch * euler.pitch;
                if (Math.abs(euler.yaw) < 20 && Math.abs(euler.pitch) < 20) stats.gazeFacing += 1;
              }
              if (blendshapes.length) {
                const get = (name) => blendshapes.find(c => c.categoryName === name)?.score || 0;
                const smile = (get('mouthSmileLeft') + get('mouthSmileRight')) / 2;
                const blink = (get('eyeBlinkLeft') + get('eyeBlinkRight')) / 2;
                const browDown = (get('browDownLeft') + get('browDownRight')) / 2;
                const jawOpen = get('jawOpen');
                const lookAway = Math.max(get('eyeLookOutLeft'), get('eyeLookOutRight'), get('eyeLookUpLeft'), get('eyeLookUpRight'), get('eyeLookDownLeft'), get('eyeLookDownRight'));
                stats.smileSum += smile;
                stats.blinkSum += blink;
                stats.browDownSum += browDown;
                stats.jawOpenSum += jawOpen;
                stats.lookAwaySum += lookAway;
              }
            }
          }
          if (poseLandmarkerRef.current) {
            const poseResult = poseLandmarkerRef.current.detectForVideo(video, now);
            const posture = analyzePosture(poseResult.landmarks);
            if (posture?.aligned) stats.postureAligned += 1;
          }
          if (stats.total % 15 === 0) updateBehaviorMetrics();
        } catch (_) { /* keep loop alive */ }

        gestureLoopRef.current = requestAnimationFrame(detect);
      };
      if (gestureLoopRef.current) cancelAnimationFrame(gestureLoopRef.current);
      gestureLoopRef.current = requestAnimationFrame(detect);
    } catch (err) {
      setGestureStatus(`Could not load gesture model: ${err.message}. Check internet access and restart camera.`);
    }
  }

  function stopGestureDetection() {
    if (gestureLoopRef.current) cancelAnimationFrame(gestureLoopRef.current);
    gestureLoopRef.current = null;
    setGestureStatus('Gesture detection stopped.');
    setBehaviorStatus('Face/expression/posture analysis stopped.');
    updateBehaviorMetrics();
  }

  async function startCamera() {
    setError('');
    if (!navigator.mediaDevices?.getUserMedia) {
      setCameraMessage('Camera access requires HTTPS or localhost and a supported browser.');
      return;
    }
    try {
      const stream = await navigator.mediaDevices.getUserMedia({
        video: { facingMode: 'user', width: { ideal: 1280 }, height: { ideal: 720 } },
        audio: false
      });
      cameraStreamRef.current = stream;
      setCameraOn(true);
      setCameraMessage('Camera connected. Live preview is active.');
      beginTimer();
      if (videoRef.current) {
        videoRef.current.srcObject = stream;
        await videoRef.current.play().catch(() => {});
        startGestureDetection();
      }
    } catch (err) {
      const message = err.name === 'NotAllowedError'
        ? 'Camera permission was blocked. Allow camera access in the browser address-bar permissions.'
        : err.name === 'NotFoundError'
          ? 'No camera was detected. Connect a webcam and try again.'
          : err.name === 'NotReadableError'
            ? 'The camera may be in use by another app. Close Teams/Zoom/Camera and retry.'
            : `Camera could not start: ${err.message || err.name}`;
      setCameraMessage(message);
    }
  }

  function stopCamera() {
    if (cameraStreamRef.current) {
      cameraStreamRef.current.getTracks().forEach(track => track.stop());
      cameraStreamRef.current = null;
    }
    if (videoRef.current) videoRef.current.srcObject = null;
    setCameraOn(false);
    stopGestureDetection();
    setCameraMessage('Camera is off.');
    if (!micStreamRef.current && timerRef.current) {
      clearInterval(timerRef.current);
      timerRef.current = null;
      setRecording(false);
    }
  }

  async function startMicrophone() {
    setError('');
    if (!navigator.mediaDevices?.getUserMedia) {
      setMicMessage('Microphone access requires HTTPS or localhost and a supported browser.');
      return;
    }
    try {
      const stream = await navigator.mediaDevices.getUserMedia({
        audio: { echoCancellation: true, noiseSuppression: true, autoGainControl: true },
        video: false
      });
      micStreamRef.current = stream;
      setMicOn(true);
      setMicMessage('Microphone connected. Audio recording is active.');
      setSpeechError('');
      setPauseMetrics({ pauses: 0, avgPauseMs: 0, longPauses: 0, silencePercent: 0 });
      audioChunksRef.current = [];
      const mimeType = MediaRecorder.isTypeSupported('audio/webm;codecs=opus')
        ? 'audio/webm;codecs=opus'
        : '';
      const recorder = new MediaRecorder(stream, mimeType ? { mimeType } : undefined);
      mediaRecorderRef.current = recorder;
      recorder.ondataavailable = event => {
        if (event.data && event.data.size > 0) audioChunksRef.current.push(event.data);
      };
      recorder.onstop = () => {
        if (audioChunksRef.current.length) {
          const blob = new Blob(audioChunksRef.current, { type: recorder.mimeType || 'audio/webm' });
          const url = URL.createObjectURL(blob);
          setAudioBlob(blob);
          setAudioUrl(previous => {
            if (previous) URL.revokeObjectURL(previous);
            return url;
          });
          setMicMessage('Audio captured. Use Play recording to review your practice.');
          analyzeAudioPauses(blob);
        }
      };
      recorder.start(500);

      const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
      if (SR) {
        const recognition = new SR();
        recognition.continuous = true;
        recognition.interimResults = true;
        recognition.lang = 'en-IN';
        recognition.onresult = event => {
          let finalText = '';
          let interimText = '';
          for (let i = 0; i < event.results.length; i++) {
            const phrase = event.results[i][0].transcript;
            if (event.results[i].isFinal) finalText += `${phrase} `;
            else interimText += `${phrase} `;
          }
          setTranscript([finalText.trim(), interimText.trim()].filter(Boolean).join(' '));
        };
        recognition.onerror = event => {
          const messages = {
            'not-allowed': 'Speech recognition permission was blocked. Check microphone permissions.',
            'no-speech': 'No speech detected yet. Speak clearly near the microphone.',
            'audio-capture': 'The browser cannot access an audio input device.',
            'network': 'Browser speech recognition needs an internet connection.'
          };
          setSpeechError(messages[event.error] || `Speech recognition: ${event.error}`);
        };
        recognition.onend = () => {
          if (recordingRef.current) {
            try { recognition.start(); } catch (_) { /* Browser may still be restarting. */ }
          }
        };
        recognitionRef.current = recognition;
        try { recognition.start(); } catch (err) {
          setSpeechError(`Speech recognition could not start: ${err.message}`);
        }
      } else {
        setSpeechError('Live speech-to-text is not supported in this browser. Audio is still being recorded; type or paste your transcript below.');
      }
      beginTimer();
    } catch (err) {
      const message = err.name === 'NotAllowedError'
        ? 'Microphone permission was blocked. Allow microphone access in browser settings.'
        : err.name === 'NotFoundError'
          ? 'No microphone was detected. Connect a microphone and retry.'
          : err.name === 'NotReadableError'
            ? 'The microphone may be in use by another app. Close other recording apps and retry.'
            : `Microphone could not start: ${err.message || err.name}`;
      setMicMessage(message);
    }
  }

  function beginTimer() {
    if (!timerRef.current) {
      setSeconds(0);
      recordingRef.current = true;
      setRecording(true);
      timerRef.current = setInterval(() => setSeconds(value => value + 1), 1000);
    }
  }

  async function analyzeAudioPauses(blob) {
    try {
      const arrayBuffer = await blob.arrayBuffer();
      const AudioCtx = window.AudioContext || window.webkitAudioContext;
      if (!AudioCtx) return;
      const ctx = new AudioCtx();
      const audioBuffer = await ctx.decodeAudioData(arrayBuffer.slice(0));
      const data = audioBuffer.getChannelData(0);
      const sampleRate = audioBuffer.sampleRate;
      const windowSize = Math.floor(sampleRate * 0.05); // 50ms
      let silentWindows = 0;
      const totalWindows = Math.floor(data.length / windowSize);
      const silenceFlags = [];
      for (let i = 0; i < totalWindows; i++) {
        let sum = 0;
        for (let j = 0; j < windowSize; j++) sum += Math.abs(data[i * windowSize + j] || 0);
        const rms = sum / windowSize;
        const silent = rms < 0.015;
        silenceFlags.push(silent);
        if (silent) silentWindows++;
      }
      // group contiguous silent windows into pauses
      const pauses = [];
      let run = 0;
      for (let i = 0; i < silenceFlags.length; i++) {
        if (silenceFlags[i]) run++;
        else if (run) { pauses.push(run * 50); run = 0; }
      }
      if (run) pauses.push(run * 50);
      const significantPauses = pauses.filter(p => p >= 600);
      const longPauses = significantPauses.filter(p => p >= 1500).length;
      const avgPauseMs = significantPauses.length ? Math.round(significantPauses.reduce((a, b) => a + b, 0) / significantPauses.length) : 0;
      const silencePercent = totalWindows ? Math.round((silentWindows / totalWindows) * 100) : 0;
      setPauseMetrics({ pauses: significantPauses.length, avgPauseMs, longPauses, silencePercent });
      ctx.close?.();
    } catch (_) { /* ignore analysis failures */ }
  }

  function audioBufferToWav(buffer) {
    const numChannels = Math.min(2, buffer.numberOfChannels || 1);
    const sampleRate = buffer.sampleRate;
    const format = 1; // PCM
    const bitDepth = 16;
    const samples = buffer.length * numChannels;
    const blockAlign = numChannels * bitDepth / 8;
    const byteRate = sampleRate * blockAlign;
    const dataSize = samples * blockAlign;
    const view = new DataView(new ArrayBuffer(44 + dataSize));
    let offset = 0;
    const writeString = (s) => { for (let i = 0; i < s.length; i++) view.setUint8(offset++, s.charCodeAt(i)); };
    writeString('RIFF'); view.setUint32(offset, 36 + dataSize, true); offset += 4;
    writeString('WAVE'); writeString('fmt '); view.setUint32(offset, 16, true); offset += 4;
    view.setUint16(offset, format, true); offset += 2;
    view.setUint16(offset, numChannels, true); offset += 2;
    view.setUint32(offset, sampleRate, true); offset += 4;
    view.setUint32(offset, byteRate, true); offset += 4;
    view.setUint16(offset, blockAlign, true); offset += 2;
    view.setUint16(offset, bitDepth, true); offset += 2;
    writeString('data'); view.setUint32(offset, dataSize, true); offset += 4;
    const channels = [];
    for (let c = 0; c < numChannels; c++) channels.push(buffer.getChannelData(c));
    for (let i = 0; i < buffer.length; i++) {
      for (let c = 0; c < numChannels; c++) {
        let sample = channels[c][i] || 0;
        sample = Math.max(-1, Math.min(1, sample));
        view.setInt16(offset, sample < 0 ? sample * 0x8000 : sample * 0x7FFF, true);
        offset += 2;
      }
    }
    return new Blob([view.buffer], { type: 'audio/wav' });
  }

  async function blobToWavBlob(blob) {
    const arrayBuffer = await blob.arrayBuffer();
    const AudioCtx = window.AudioContext || window.webkitAudioContext;
    const ctx = new AudioCtx();
    const audioBuffer = await ctx.decodeAudioData(arrayBuffer.slice(0));
    const wav = audioBufferToWav(audioBuffer);
    ctx.close?.();
    return wav;
  }

  async function transcribeRecording() {
    if (!audioBlob) {
      setError('No recorded audio available. Record practice audio first.');
      return;
    }
    setTranscribing(true);
    setError('');
    try {
      let uploadBlob = audioBlob;
      let filename = 'practice.webm';
      try {
        uploadBlob = await blobToWavBlob(audioBlob);
        filename = 'practice.wav';
      } catch (_) {
        // fallback to original blob; backend may still handle it if ffmpeg exists
      }
      const form = new FormData();
      form.append('file', uploadBlob, filename);
      const res = await fetch('http://127.0.0.1:8011/transcribe', { method: 'POST', body: form });
      const data = await res.json();
      if (!res.ok) throw new Error(data.detail || 'Transcription failed.');
      if (data.transcript) setTranscript(data.transcript);
      if (Array.isArray(data.segments) && data.segments.length) {
        // simple pause estimate from whisper segment gaps
        const gaps = [];
        for (let i = 1; i < data.segments.length; i++) {
          const gap = (data.segments[i].start - data.segments[i - 1].end) * 1000;
          if (gap >= 600) gaps.push(gap);
        }
        if (gaps.length) {
          const avgPauseMs = Math.round(gaps.reduce((a, b) => a + b, 0) / gaps.length);
          const longPauses = gaps.filter(g => g >= 1500).length;
          setPauseMetrics(prev => ({ ...prev, pauses: gaps.length, avgPauseMs, longPauses }));
        }
      }
      setMicMessage(`Transcribed with Whisper (${data.model || 'base'}). Review/edit the transcript below.`);
    } catch (err) {
      setError(err.message === 'Failed to fetch' ? 'Cannot connect to backend for transcription.' : err.message);
    } finally {
      setTranscribing(false);
    }
  }

  function stopMicrophone() {
    recordingRef.current = false;
    if (recognitionRef.current) {
      recognitionRef.current.onend = null;
      try { recognitionRef.current.stop(); } catch (_) {}
      recognitionRef.current = null;
    }
    if (mediaRecorderRef.current && mediaRecorderRef.current.state !== 'inactive') {
      mediaRecorderRef.current.stop();
    }
    if (micStreamRef.current) {
      micStreamRef.current.getTracks().forEach(track => track.stop());
      micStreamRef.current = null;
    }
    setMicOn(false);
    setMicMessage('Microphone stopped. Your captured audio is available for playback.');
    if (!cameraStreamRef.current && timerRef.current) {
      clearInterval(timerRef.current);
      timerRef.current = null;
      setRecording(false);
    }
  }

  function stopPractice() {
    if (seconds > 0 || transcript.trim()) {
      const behaviorScore = Math.round((behaviorMetrics.gazePercent + behaviorMetrics.headStability + behaviorMetrics.postureScore) / 3);
      const entry = {
        id: Date.now(), date: new Date().toLocaleString(), durationSeconds: seconds,
        words: practice.words, wpm: practice.wpm, fillerWords: practice.fillerCount,
        keywordOverlap: practice.alignment, deliveryScore: practice.overall,
        gazePercent: behaviorMetrics.gazePercent, postureScore: behaviorMetrics.postureScore,
        headStability: behaviorMetrics.headStability, behaviorScore,
        smilePercent: behaviorMetrics.smilePercent, browTensionPercent: behaviorMetrics.browTensionPercent,
        jawOpenPercent: behaviorMetrics.jawOpenPercent, dominantExpression: behaviorMetrics.dominantExpression,
        pauses: pauseMetrics.pauses, longPauses: pauseMetrics.longPauses, avgPauseMs: pauseMetrics.avgPauseMs, silencePercent: pauseMetrics.silencePercent
      };
      setSessionHistory(previous => {
        const updated = [entry, ...previous].slice(0, 10);
        try { localStorage.setItem('presentaiPracticeHistory', JSON.stringify(updated)); } catch (_) {}
        return updated;
      });
    }
    stopMicrophone();
    stopCamera();
    if (timerRef.current) {
      clearInterval(timerRef.current);
      timerRef.current = null;
    }
    recordingRef.current = false;
    setRecording(false);
  }

  useEffect(() => {
    return () => {
      recordingRef.current = false;
      if (timerRef.current) clearInterval(timerRef.current);
      if (recognitionRef.current) {
        try { recognitionRef.current.stop(); } catch (_) {}
      }
      [cameraStreamRef.current, micStreamRef.current].forEach(stream => {
        if (stream) stream.getTracks().forEach(track => track.stop());
      });
      if (audioUrl) URL.revokeObjectURL(audioUrl);
    };
  }, []);

  // Send the slide text and manually entered transcript to Sentence-BERT.

  async function compareSemanticAlignment() {

    if (!slides.length) {

      setError('Please upload a PPTX or PDF first.');

      return;

    }



    if (!manualTranscript.trim()) {

      setError('Please enter your presentation transcript.');

      return;

    }



    setComparing(true);

    setError('');

    setSemanticResult(null);



    try {

      const response = await fetch(

        'http://127.0.0.1:8011/compare',

        {

          method: 'POST',

          headers: {

            'Content-Type': 'application/json'

          },

          body: JSON.stringify({

            slides,

            transcript: manualTranscript.trim()

          })

        }

      );



      const data = await response.json();



      if (!response.ok) {

        throw new Error(

          data.detail || 'Semantic comparison failed.'

        );

      }



      setSemanticResult(data);

    } catch (err) {

      setError(
        err.message === 'Failed to fetch'
          ? 'Cannot connect to the Python backend. Start the backend on port 8011 and try again.'
          : err.message
      );

    } finally {

      setComparing(false);

    }

  }



  const practice = useMemo(() => {

    const words = countWords(transcript);

    const wpm = seconds > 0
      ? Math.round(words / (seconds / 60))
      : 0;



    const fillerCount = countFillers(transcript);

    const alignment = similarity(slides.join(' '), transcript);



    const pace = wpm === 0

      ? 0

      : wpm >= 100 && wpm <= 160

        ? 100

        : Math.max(45, 100 - Math.abs(130 - wpm));



    const fillerScore = Math.max(40, 100 - fillerCount * 5);



    const overall = Math.round(

      (pace + fillerScore + alignment) / 3

    );



    return {

      words,

      wpm,

      fillerCount,

      alignment,

      pace,

      fillerScore,

      overall

    };

  }, [transcript, seconds, slides]);



  async function buildEditedPptxBlob() {
    if (!file || !file.name.toLowerCase().endsWith('.pptx')) return null;
    const buffer = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(buffer);
    const names = Object.keys(zip.files)
      .filter(name => /^ppt\/slides\/slide\d+\.xml$/.test(name))
      .sort((a, b) => Number(a.match(/slide(\d+)/)[1]) - Number(b.match(/slide(\d+)/)[1]));

    for (let i = 0; i < names.length; i++) {
      const name = names[i];
      let xml = await zip.file(name).async('text');
      const currentEdits = slideEditsRef.current;
      const edited = String(currentEdits[i] ?? '').replace(/\r\n/g, '\n');
      const lines = edited.split('\n').map(l => l.trim()).filter(Boolean);
      const doc = new DOMParser().parseFromString(xml, 'application/xml');
      const paras = Array.from(doc.getElementsByTagNameNS('*', 'p'));
      // only paragraph elements inside txBody; filter by localName p and having a:r or a:endParaR or parent txBody
      const textParas = paras.filter(p => p.localName === 'p' && (p.getElementsByTagNameNS('*', 't').length > 0 || p.parentNode?.localName === 'txBody'));
      const escapeXml = (str) => String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');

      if (textParas.length) {
        for (let pi = 0; pi < textParas.length; pi++) {
          const para = textParas[pi];
          const tNodes = Array.from(para.getElementsByTagNameNS('*', 't'));
          const line = lines[pi] ?? (pi === textParas.length - 1 && lines.length > textParas.length ? lines.slice(pi).join(' ') : '');
          if (tNodes.length === 0) {
            // create a run if needed
            if (line) {
              const nsA = 'http://schemas.openxmlformats.org/drawingml/2006/main';
              const r = doc.createElementNS(nsA, 'a:r');
              const t = doc.createElementNS(nsA, 'a:t');
              t.textContent = line;
              r.appendChild(t);
              para.appendChild(r);
            }
          } else {
            tNodes[0].textContent = line;
            for (let ti = 1; ti < tNodes.length; ti++) tNodes[ti].textContent = '';
          }
        }
      } else if (lines.length) {
        // no existing text paragraphs; try to put into first txBody
        const txBodies = Array.from(doc.getElementsByTagNameNS('*', 'txBody'));
        if (txBodies.length) {
          const nsA = 'http://schemas.openxmlformats.org/drawingml/2006/main';
          const tx = txBodies[0];
          // remove existing p children
          Array.from(tx.childNodes).forEach(n => { if (n.localName === 'p') tx.removeChild(n); });
          lines.forEach(line => {
            const pEl = doc.createElementNS(nsA, 'a:p');
            const r = doc.createElementNS(nsA, 'a:r');
            const t = doc.createElementNS(nsA, 'a:t');
            t.textContent = line;
            r.appendChild(t); pEl.appendChild(r); tx.appendChild(pEl);
          });
        }
      }

      const serializer = new XMLSerializer();
      xml = serializer.serializeToString(doc);
      zip.file(name, xml);
    }
    const out = await zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
    modifiedPptxRef.current = out;
    return out;
  }

  async function saveSlideEdits({ silent = false } = {}) {
    const cleaned = (slideEditsRef.current.length ? slideEditsRef.current : slideEdits).map(s => String(s || '').trim());
    setSlides(cleaned);
    setAnalysis(analyzeSlides(cleaned.map(t => ({ text: t, meta: {} }))));
    setSemanticResult(null);
    if (!silent) setError('');
    if (file?.name.toLowerCase().endsWith('.pptx')) {
      if (rebuildingRef.current) return;
      rebuildingRef.current = true;
      try {
        const blob = await buildEditedPptxBlob();
        if (blob) {
          setPptxReady(false);
          if (pptxRendererRef.current) { pptxRendererRef.current.destroy(); pptxRendererRef.current = null; }
          const renderer = new PptxRenderer();
          await renderer.load(blob);
          pptxRendererRef.current = renderer;
          setPptxReady(true);
          setPreviewVersion(v => v + 1);
        }
      } catch (err) {
        console.error(err);
        if (!silent) setError(`Saved text, but could not update PPTX preview: ${err.message}`);
        setPptxReady(true);
      } finally {
        rebuildingRef.current = false;
      }
    } else {
      setPreviewVersion(v => v + 1);
    }
  }

  function downloadEditedPptx() {
    try {
      const pptx = new PptxGenJS();
      pptx.layout = 'LAYOUT_WIDE';
      const edited = slideEdits.length ? slideEdits : slides;
      (edited.length ? edited : ['']).forEach((text, i) => {
        const slide = pptx.addSlide();
        const lines = String(text || '').split(/\n+/).map(l => l.trim()).filter(Boolean);
        const title = lines[0] || `Slide ${i + 1}`;
        const body = lines.slice(1).join('\n');
        slide.addText(title, { x: 0.6, y: 0.4, w: 12, h: 0.8, fontSize: 28, bold: true, color: '1F2937' });
        slide.addText(body || ' ', { x: 0.8, y: 1.5, w: 11.5, h: 5, fontSize: 18, color: '374151', valign: 'top' });
      });
      pptx.writeFile({ fileName: `${(file?.name || 'presentation').replace(/\.[^.]+$/, '')}-edited.pptx` });
    } catch (err) {
      setError(`Could not generate PPTX: ${err.message}`);
    }
  }

  function downloadReport() {

    const report = {

      project:

        'A Multimodal Deep Learning Framework for Assessing Presentation Proficiency',

      file: file?.name || null,

      generatedAt: new Date().toISOString(),

      slideAnalysis: analysis,

      practiceAnalysis: practice,

      behaviorAnalysis: behaviorMetrics,

      pauseAnalysis: pauseMetrics,

      transcript,

      manualTranscript,

      semanticAlignment: semanticResult,

      sessionHistory

    };



    const blob = new Blob(

      [JSON.stringify(report, null, 2)],

      { type: 'application/json' }

    );



    const url = URL.createObjectURL(blob);

    const anchor = document.createElement('a');



    anchor.href = url;

    anchor.download = 'presentation-assessment-report.json';

    anchor.click();



    URL.revokeObjectURL(url);

  }



  const navigation = [

    ['dashboard', 'Dashboard'],

    ['analysis', 'Slide Analysis'],

    ['edit', 'Edit Slides'],

    ['practice', 'Practice'],

    ['alignment', 'Semantic Alignment'],

    ['results', 'Results']

  ];



  return (

    <div className="app">

      <aside className="sidebar">

        <div className="brand">

          <div className="brandMark">P</div>

          <div>

            <b>PresentAI</b>

            <span>Assessment Studio</span>

          </div>

        </div>



        {navigation.map(([id, label]) => (

          <button

            key={id}

            className={page === id ? 'nav active' : 'nav'}

            onClick={() => setPage(id)}

          >

            {label}

          </button>

        ))}



        <div className="sideNote">

          <b>Research Prototype</b>

          <p>

            Multimodal presentation assessment using slide,

            speech and behavioral signals.

          </p>

        </div>

      </aside>



      <main>

        <header>

          <div>

            <p className="eyebrow">PHASE 2 PROJECT</p>

            <h1>Presentation Proficiency Assessment</h1>

            <p className="sub">

              Analyze your slides, practice delivery, and receive

              structured feedback.

            </p>

          </div>



          <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
            <span className={backendOk ? 'statusGood' : 'statusMuted'} title="Semantic alignment backend status">
              {backendOk === null ? '* Checking backend...' : backendOk ? '* Backend connected' : '* Backend offline'}
            </span>
            <button className="outline" onClick={downloadReport}>Export Report</button>
          </div>

        </header>



        {error && (

          <div className="alert">

            {error}

            <button onClick={() => setError('')}>Ã—</button>

          </div>

        )}



        {page === 'dashboard' && (

          <section className="grid">

            <div className="hero card">

              <div>

                <span className="pill">MULTIMODAL FRAMEWORK</span>

                <h2>Practice, analyze, and improve your presentation.</h2>

                <p>

                  Upload presentation -> Analyze slides -> Get improvement suggestions -> Edit slides -> Practice with camera/microphone -> Analyze delivery -> Get final score and personalized feedback.

                </p>



                <label className="upload">

                  <input

                    type="file"

                    accept=".pptx,.pdf"

                    onChange={event =>

                      event.target.files?.[0] &&

                      handleFile(event.target.files[0])

                    }

                  />

                  <span>ï¼‹ Upload PPTX / PDF</span>

                </label>

              </div>



              <div className="heroVisual">

                <div className="scoreRing">

                  {analysis?.overall ?? '-'}

                  <small>/100</small>

                </div>

                <span>Presentation score</span>

              </div>

            </div>



            <div className="two">

              <div className="card">

                <h3>Project pipeline</h3>

                <div className="pipeline">

                  {[

                    ['01', 'Upload', 'Presentation file'],
                    ['02', 'Analyze', 'Slides & suggestions'],
                    ['03', 'Edit', 'Improve your slides'],
                    ['04', 'Practice', 'Camera & microphone'],
                    ['05', 'Deliver', 'Analyze delivery'],
                    ['06', 'Report', 'Score & feedback']

                  ].map(item => (

                    <div className="step" key={item[0]}>

                      <b>{item[0]}</b>

                      <strong>{item[1]}</strong>

                      <span>{item[2]}</span>

                    </div>

                  ))}

                </div>

              </div>



              <div className="card info">

                <h3>Research direction</h3>

                <p>

                  This prototype integrates Sentence-BERT semantic alignment

                  and MediaPipe gesture/face/pose estimates. Full research scope

                  still includes Whisper, Wav2Vec2, ViT slide vision, multimodal

                  fusion, and human-validated scoring.

                </p>

                <button

                  className="primary"

                  onClick={() => setPage('alignment')}

                >

                  Open Semantic Alignment

                </button>

              </div>

            </div>

          </section>

        )}



        {page === 'analysis' && (

          <section className="content">

            <div className="card sectionHead">

              <div>

                <span className="pill">01 · SLIDE ANALYSIS</span>

                <h2>Presentation content review</h2>

                <p>

                  {file

                    ? file.name

                    : 'Upload a PPTX or PDF to calculate the analysis.'}

                </p>

              </div>

              <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'center' }}>
                <label className="upload small">
                  <input type="file" accept=".pptx,.pdf" onChange={event => event.target.files?.[0] && handleFile(event.target.files[0])} />
                  <span>Upload</span>
                </label>
                {analysis && <button className="primary" onClick={() => setPage('edit')}>Next: Edit Slides</button>}
              </div>

            </div>

            {!analysis && (
              <div className="card">
                <h3>No presentation loaded</h3>
                <p>Upload a PPTX or PDF to see slide content, structure, readability, and heuristic quality checks.</p>
              </div>
            )}



            {analysis && (

              <div className="metricGrid">

                {[

                  ['Overall', analysis.overall],

                  ['Content', analysis.content],

                  ['Readability', analysis.readability],

                  ['Structure', analysis.structure],

                  ['Visual', analysis.visual]

                ].map(([name, value]) => (

                  <div className="metric card" key={name}>

                    <span>{name}</span>

                    <strong>{value}</strong>

                    <div className="bar">

                      <i style={{ width: `${value}%` }} />

                    </div>

                  </div>

                ))}

              </div>

            )}



            {analysis && (

              <>
              <div className="two">

                <div className="card">

                  <h3>Presentation statistics</h3>

                  <p>

                    {analysis.slides} slides · {analysis.words} extracted

                    words · {analysis.avg} average words per slide

                  </p>

                  <ul>

                    <li>

                      Review slides with too much text for readability.

                    </li>

                    <li>

                      Keep headings and body content visually consistent.

                    </li>

                    <li>

                      Use visuals when they improve explanation.

                    </li>

                  </ul>

                </div>



                <div className="card">

                  <h3>Next step: Edit & practice</h3>
                  <p>Apply the suggested improvements to your slides, then practice your delivery with camera and microphone.</p>
                  <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', marginTop: 12 }}>
                    <button className="primary" onClick={() => setPage('edit')}>Edit Slides</button>
                    <button className="outline" onClick={() => setPage('practice')}>Go to Practice</button>
                    <button className="outline" onClick={() => setPage('alignment')}>Check Speech Alignment</button>
                  </div>
                </div>

              </div>

              <div className="card">
                <h3>Slide-by-slide audit & suggested improvements</h3>
                <p>Suggestions below are generated from your actual slide content (title, text, word count, and detected structure), with specific actions for each slide.</p>
                <div className="slideAuditList">
                  {analysis.slideDetails?.map(item => (
                    <div className="slideAuditItem" key={item.slide}>
                      <div className="sectionTitleRow" style={{ alignItems: 'flex-start' }}>
                        <div style={{ minWidth: 0 }}>
                          <strong style={{ fontSize: 17 }}>Slide {item.slide}: {item.title}</strong>
                          <p style={{ margin: '6px 0 0' }}>{item.words} words · {item.textBlocks} text block(s) · {item.sentences} sentence ending(s){item.images != null ? ` · ${item.images} image(s)` : ''}{item.minFontPt ? ` · min font ~${item.minFontPt}pt` : ''}{item.bullets != null ? ` · ${item.bullets} bullet marker(s)` : ''}</p>
                        </div>
                        <span className={item.density === 'High' ? 'tagWarning' : 'tagNeutral'}>{item.density} density</span>
                      </div>

                      {(item.excerpt || item.keyTerms?.length) && (
                        <div style={{ marginTop: 12, padding: 14, borderRadius: 12, background: '#f7fbfa', border: '1px solid #dcecea' }}>
                          {item.excerpt && <p style={{ margin: 0, color: '#334155' }}><strong>Slide content (detected):</strong> {item.excerpt}</p>}
                          {item.keyTerms?.length ? <p style={{ margin: '8px 0 0', color: '#0f766e' }}><strong>Key terms:</strong> {item.keyTerms.join(', ')}</p> : null}
                        </div>
                      )}

                      <div style={{ display: 'grid', gap: 12, marginTop: 14 }}>
                        {(item.suggestions || item.issues?.map(i => ({ priority: 'Note', area: 'Suggestion', issue: i, why: '', action: i, example: '' })) || []).map((s, index) => (
                          <div key={index} style={{ padding: 16, borderRadius: 12, border: '1px solid #e4ecea', background: s.priority === 'High' ? '#fff8f6' : s.priority === 'Good' ? '#f4fbf8' : '#fbfdfc' }}>
                            <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center', marginBottom: 8 }}>
                              <span className={s.priority === 'High' ? 'tagWarning' : 'tagNeutral'}>{s.priority}</span>
                              <strong style={{ fontSize: 14, color: '#1f3f46' }}>{s.area}</strong>
                            </div>
                            <p style={{ margin: '0 0 8px', color: '#334155' }}><strong>Issue:</strong> {s.issue}</p>
                            {s.why && <p style={{ margin: '0 0 8px' }}><strong>Why it matters:</strong> {s.why}</p>}
                            <p style={{ margin: '0 0 8px' }}><strong>What to change:</strong> {s.action}</p>
                            {s.example && <p style={{ margin: 0, color: '#0f766e' }}><strong>Example:</strong> {s.example}</p>}
                          </div>
                        ))}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
              </>
            )}

          </section>

        )}



        {page === 'practice' && (
          <section className="content">
            <div className="card sectionHead">
              <div>
                <span className="pill">02 · LIVE PRACTICE</span>
                <h2>Practice your delivery</h2>
                <p>Camera and microphone can be started independently. Allow browser permissions when prompted.</p>
              </div>
              <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'center' }}>
                <button className="outline" onClick={() => setPage('analysis')}>Back to Analysis</button>
                <button className="primary" onClick={() => { setManualTranscript(transcript); setPage('alignment'); }}>Next: Semantic Alignment</button>
              </div>
            </div>

            <div className="practiceWorkspace">
              <div className="card presentationViewerCard">
                <div className="practicePanelHeader">
                  <div>
                    <h3>Presentation slides</h3>
                    <p className="deviceMessage">{file ? file.name : 'Upload a PPTX or PDF in Slide Analysis to display it here.'}</p>
                  </div>
                  {file && slides.length > 0 && <span className="slidePosition">Slide {currentSlide + 1} / {slides.length}</span>}
                </div>
                {file && slides.length > 0 ? (
                  (file.name.toLowerCase().endsWith('.pptx') && !slidePreviewBlank) || file.name.toLowerCase().endsWith('.pdf') ? (
                    <div className="originalSlideStage">
                      <canvas ref={slideCanvasRef} className="originalSlideCanvas" aria-label={`Presentation slide ${currentSlide + 1}`} />
                      {file.name.toLowerCase().endsWith('.pptx') && !pptxReady && <div className="slideRenderNotice">Loading PPT preview...</div>}
                      {slideRenderError && <div className="slideRenderNotice">{slideRenderError}</div>}
                    </div>
                  ) : (
                    <div className="originalSlideStage" style={{ background: '#fff' }}>
                      {(() => {
                        const text = String(slides[currentSlide] || '');
                        const lines = text.split(/\n+/).map(l => l.trim()).filter(Boolean);
                        const title = lines[0] || `Slide ${currentSlide + 1}`;
                        const body = lines.slice(1);
                        return (
                          <div style={{ width: '100%', height: '100%', padding: '8% 9%', display: 'flex', flexDirection: 'column', background: 'linear-gradient(180deg,#ffffff 0%,#f7fbfa 100%)', color: '#1f2937', overflow: 'auto' }}>
                            <div style={{ fontSize: 'clamp(24px,3vw,42px)', fontWeight: 800, lineHeight: 1.15, color: '#134e4a', marginBottom: 24, fontFamily: 'Plus Jakarta Sans, Inter, sans-serif' }}>{title}</div>
                            <div style={{ display: 'grid', gap: 12, fontSize: 'clamp(14px,1.6vw,22px)', lineHeight: 1.5 }}>
                              {body.length ? body.map((line, i) => (
                                <div key={i} style={{ display: 'flex', gap: 10, alignItems: 'flex-start' }}>
                                  <span style={{ color: '#0f766e', fontWeight: 700, lineHeight: 1.4 }}>*</span>
                                  <span>{line.replace(/^[-**]\s*/, '')}</span>
                                </div>
                              )) : <div style={{ color: '#64748b' }}>No additional content on this slide.</div>}
                            </div>
                          </div>
                        );
                      })()}
                        {slideRenderError && <div className="slideRenderNotice" style={{marginTop:8}}>{slideRenderError}</div>}
                    </div>
                  )
                ) : (
                  <div className="slideEmptyState">
                    <strong>No presentation loaded</strong>
                    <span>Go to Slide Analysis and upload your .pptx or .pdf file first.</span>
                    <button className="outline" onClick={() => setPage('analysis')}>Open Slide Analysis</button>
                  </div>
                )}
                <div className="slideNavControls">
                  <button className="outline" disabled={!slides.length || currentSlide <= 0} onClick={() => setCurrentSlide(value => Math.max(0, value - 1))}>&larr; Previous</button>
                  <button className="outline" disabled={!slides.length || currentSlide >= slides.length - 1} onClick={() => setCurrentSlide(value => Math.min(slides.length - 1, value + 1))}>Next &rarr;</button>
                </div>
                <p className="support">Showing the exact PPT preview (including saved edits applied to the working copy).</p>
              </div>

              <div className="practiceSideColumn">
                <div className="card videoCard practiceCameraCard">
                  <div className="practicePanelHeader"><h3>Live camera</h3><span className={cameraOn ? 'statusGood' : 'statusMuted'}>{cameraOn ? '* Camera live' : 'Camera off'}</span></div>
                  <video ref={videoRef} autoPlay muted playsInline />
                  <div className="videoControls">
                    {!cameraOn ? <button className="primary" onClick={startCamera}>Start Camera</button> : <button className="danger" onClick={stopCamera}>Stop Camera</button>}
                  </div>
                  <p className="deviceMessage">{cameraMessage}</p>
                </div>

                <div className="card practiceMicCard">
                  <h3>Microphone & speech capture</h3>
                  <p className="deviceMessage">{micMessage}</p>
                  <div className="videoControls">
                    {!micOn ? <button className="primary" onClick={startMicrophone}>Start Microphone</button> : <button className="danger" onClick={stopMicrophone}>Stop Microphone</button>}
                    <span className={micOn ? 'statusGood' : 'statusMuted'}>{micOn ? '* Mic active' : 'Mic off'}</span>
                  </div>
                  <p className="support">{support ? 'Browser speech-to-text is available.' : 'Browser speech-to-text is unavailable; audio recording and manual transcript still work.'}</p>
                  {speechError && <div className="notice">{speechError}</div>}
                  <div className="sessionClock">{Math.floor(seconds / 60)}:{String(seconds % 60).padStart(2, '0')} <small>practice time</small></div>
                  <button className="outline" onClick={stopPractice}>Stop All Devices</button>
                  {audioUrl && (
                    <div className="audioReview">
                      <h4>Recorded practice audio</h4>
                      <audio controls src={audioUrl} />
                      <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'center' }}>
                        <a className="downloadAudio" href={audioUrl} download="presentation-practice.webm">Download audio recording</a>
                        <button className="outline" onClick={transcribeRecording} disabled={transcribing}>{transcribing ? 'Transcribing...' : 'Transcribe with Whisper'}</button>
                      </div>
                      <p className="support">Recorded audio is converted to WAV in-browser before Whisper transcription for better compatibility.</p>
                    </div>
                  )}
                </div>

                <div className="card practiceTranscriptCard">
                  <div className="sectionTitleRow">
                    <div><h3>Live transcript</h3><p>Speech recognition fills this area when supported. You can also correct or type the transcript manually.</p></div>
                    <button className="outline" onClick={() => { setManualTranscript(transcript); setPage('alignment'); }}>Use for Slide Alignment</button>
                  </div>
                  <textarea value={transcript} onChange={event => setTranscript(event.target.value)} placeholder="Your recognised speech appears here. If speech-to-text is unavailable, type what you presented..." rows={6} className="transcriptInput" />
                </div>
              </div>
            </div>

            <div className="metricGrid cols-3">
              {[
                ['Words spoken', practice.words],
                ['Speaking rate', `${practice.wpm} WPM`],
                ['Filler words', practice.fillerCount],
                ['Slide keyword overlap', `${practice.alignment}%`],
                ['Pauses (â‰¥0.6s)', pauseMetrics.pauses],
                ['Long pauses (â‰¥1.5s)', pauseMetrics.longPauses]
              ].map(([name, value]) => (
                <div className="metric card" key={name}>
                  <span>{name}</span>
                  <strong>{value}</strong>
                </div>
              ))}
            </div>

            <div className="two">
              <div className="card">
                <h3>Delivery feedback</h3>
                <ul>
                  <li>{practice.wpm === 0 ? 'Start a practice session and speak to calculate your speaking rate.' : practice.wpm < 100 ? 'Your pace is slower than the suggested 100-160 words per minute range.' : practice.wpm > 160 ? 'Your pace is fast. Add pauses between important points.' : 'Your speaking pace is within the suggested 100-160 words per minute range.'}</li>
                  <li>{practice.fillerCount > 3 ? `Reduce repeated filler words. ${practice.fillerCount} were detected in the transcript.` : 'Few filler words were detected in the current transcript.'}</li>
                  <li>{practice.alignment < 40 ? 'Explain the key ideas shown on your slides more directly.' : 'Your transcript shares some key terms with the uploaded slide content.'}</li>
                  <li>{pauseMetrics.pauses === 0 ? 'Pause analysis will appear after you record audio.' : pauseMetrics.longPauses > 2 ? `${pauseMetrics.longPauses} long pauses (â‰¥1.5s) detected; use deliberate pauses, but reduce unintended gaps.` : `Detected ${pauseMetrics.pauses} pause(s), avg ${pauseMetrics.avgPauseMs}ms, ${pauseMetrics.silencePercent}% silence.`}</li>
                  <li>{behaviorMetrics.frames < 30 ? 'Start camera to collect face/posture estimates.' : `Estimated camera-facing gaze ${behaviorMetrics.gazePercent}%, head stability ${behaviorMetrics.headStability}%, posture ${behaviorMetrics.postureScore}%.`}</li>
                </ul>
              </div>
              <div className="card">
                <h3>Non-verbal self-review</h3>
                <p>Mark the behaviours you demonstrated. This is a manual self-review, not automated computer-vision detection.</p>
                <div className="checkList">
                  {[
                    ['eyeContact', 'Maintained eye contact with the audience/camera'],
                    ['posture', 'Maintained an upright, confident posture'],
                    ['gestures', 'Used natural and relevant hand gestures'],
                    ['faceVisible', 'Kept my face clearly visible']
                  ].map(([key, label]) => (
                    <label className="checkRow" key={key}>
                      <input type="checkbox" checked={behaviorChecks[key]} onChange={event => setBehaviorChecks(previous => ({ ...previous, [key]: event.target.checked }))} />
                      <span>{label}</span>
                    </label>
                  ))}
                </div>
                <strong className="selfReviewScore">Self-review: {Math.round(Object.values(behaviorChecks).filter(Boolean).length / 4 * 100)}%</strong>
              </div>
            </div>

            <div className="grid" style={{ marginTop: 16 }}>
              <div className="card gesturePanel" style={{ marginTop: 0 }}>
                <div className="gestureHeading"><h3>AI hand gesture detection</h3><span className={cameraOn ? 'statusGood' : 'statusMuted'}>{cameraOn ? '* Camera input' : 'Camera required'}</span></div>
                <p className="deviceMessage">{gestureStatus}</p>
                <div className="gestureLive"><strong>{gestureLabel}</strong><span>{gestureCount} gesture changes detected</span></div>
                <div className="gestureChips"><span>Open Palm</span><span>Closed Fist</span><span>Pointing Up</span><span>Thumbs Up</span><span>Victory</span><span>ILoveYou</span></div>
                <p className="support">Keep your hands visible and well-lit. Gesture recognition identifies hand shapes; it does not judge whether a gesture is appropriate to your speech.</p>
                {gestureEvents.length > 0 && <div className="gestureEvents"><b>Recently detected</b>{gestureEvents.map((g,i)=><div key={`${g.time}-${i}`}><span>{g.name}</span><small>{g.time} - {g.confidence}% confidence</small></div>)}</div>}
              </div>
              <div className="card gesturePanel" style={{ marginTop: 0 }}>
                <div className="gestureHeading"><h3>Face, gaze, posture & expressions</h3><span className={cameraOn ? 'statusGood' : 'statusMuted'}>{cameraOn ? '* Live' : 'Camera required'}</span></div>
                <p className="deviceMessage">{behaviorStatus}</p>
                <div className="metricGrid cols-2" style={{ marginTop: 10 }}>
                  {[['Camera-facing gaze', `${behaviorMetrics.gazePercent}%`], ['Head stability', `${behaviorMetrics.headStability}%`], ['Posture score', `${behaviorMetrics.postureScore}%`], ['Smile / positivity', `${behaviorMetrics.smilePercent}%`], ['Brow tension', `${behaviorMetrics.browTensionPercent}%`], ['Jaw open (speaking)', `${behaviorMetrics.jawOpenPercent}%`]].map(([name, value]) => (
                    <div className="metric card" key={name} style={{ padding: 12 }}><span>{name}</span><strong style={{ fontSize: 22, margin: '6px 0' }}>{value}</strong></div>
                  ))}
                </div>
                <p className="support"><strong>Dominant expression:</strong> {behaviorMetrics.dominantExpression}. Prototype estimate from MediaPipe Face Landmarker blendshapes/pose; not a clinical or human-validated assessment.</p>
              </div>
            </div>
          </section>
        )}

        {page === 'edit' && (
          <section className="content">
            <div className="card sectionHead">
              <div>
                <span className="pill">02 · EDIT SLIDES</span>
                <h2>Edit your presentation</h2>
                <p>Work on a duplicate slide view: preview the original slide, edit its content, save changes, and move to practice.</p>
              </div>
              <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'center' }}>
                <button className="outline" onClick={() => setPage('analysis')}>Back to Analysis</button>
                <button className="outline" onClick={downloadEditedPptx} disabled={!slideEdits.length}>Download PPTX</button>
                <button className="outline" onClick={saveSlideEdits} disabled={!slideEdits.length}>Save Edits</button>
                <button className="primary" onClick={() => { saveSlideEdits(); setPage('practice'); }} disabled={!slideEdits.length}>Save & Move to Practice</button>
              </div>
            </div>

            {!slideEdits.length ? (
              <div className="card">
                <h3>No slides to edit</h3>
                <p>Upload a PPTX/PDF first, then you can edit slide content here.</p>
                <button className="primary" onClick={() => setPage('dashboard')}>Go to Dashboard</button>
              </div>
            ) : (
              <div className="card">
                <div className="sectionTitleRow" style={{ marginBottom: 16 }}>
                  <div>
                    <h3>Edit slide {currentSlide + 1} of {slideEdits.length}</h3>
                    <p className="support">Edit the extracted slide content below. Save to apply changes for analysis and practice.</p>
                  </div>
                  <span className="statusMuted">{countWords(slideEdits[currentSlide] || '')} words</span>
                </div>

                <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', marginBottom: 16 }}>
                  <button className="outline" disabled={currentSlide <= 0} onClick={() => setCurrentSlide(v => Math.max(0, v - 1))}>- Previous</button>
                  <button className="outline" disabled={currentSlide >= slideEdits.length - 1} onClick={() => setCurrentSlide(v => Math.min(slideEdits.length - 1, v + 1))}>Next -</button>
                </div>

                <textarea
                  className="transcriptInput"
                  value={slideEdits[currentSlide] || ''}
                  rows={14}
                  onChange={e => {
                    const value = e.target.value;
                    setSlideEdits(prev => prev.map((s, i) => i === currentSlide ? value : s));
                  }}
                  placeholder="Enter slide title and content..."
                  style={{ minHeight: 280 }}
                />

                <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', marginTop: 16 }}>
                  <button className="primary" onClick={saveSlideEdits}>Save Edits</button>
                  <button className="outline" onClick={() => { saveSlideEdits(); setPage('practice'); }}>Save & Move to Practice</button>
                </div>

                <div style={{ marginTop: 24 }}>
                  <h4 style={{ margin: '0 0 10px' }}>All slides</h4>
                  <div style={{ display: 'grid', gap: 8, maxHeight: 260, overflow: 'auto' }}>
                    {slideEdits.map((text, idx) => (
                      <button
                        key={idx}
                        className={idx === currentSlide ? 'primary' : 'outline'}
                        style={{ justifyContent: 'flex-start', textAlign: 'left', whiteSpace: 'normal' }}
                        onClick={() => setCurrentSlide(idx)}
                      >
                        {idx + 1}. {String(text || `Slide ${idx + 1}`).split(/\n+/)[0].slice(0, 60) || `Slide ${idx + 1}`}
                      </button>
                    ))}
                  </div>
                </div>
              </div>
            )}
          </section>
        )}

        {page === 'alignment' && (

          <section className="content">

            <div className="card sectionHead">

              <div>
                <span className="pill">03 · SEMANTIC ALIGNMENT</span>
                <h2>Slide and Speech Alignment</h2>
                <p>Compare the meaning of your presentation transcript with each slide using Sentence-BERT.</p>
                {backendOk === false && <div className="notice" style={{ marginTop: 12 }}>Backend is offline. Start the FastAPI server on port 8011 to run semantic alignment.</div>}
              </div>
              <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'center' }}>
                <button className="outline" onClick={() => setPage('practice')}>Back to Practice</button>
                {semanticResult && <button className="primary" onClick={() => setPage('results')}>Next: Results</button>}
              </div>

            </div>



            <div className="card">

              <h3>Presentation</h3>

              <p>

                {file

                  ? `Uploaded file: ${file.name}`

                  : 'No presentation uploaded yet.'}

              </p>

              <p>Slides extracted: {slides.length}</p>



              <label className="upload small">

                <input

                  type="file"

                  accept=".pptx,.pdf"

                  onChange={event =>

                    event.target.files?.[0] &&

                    handleFile(event.target.files[0])

                  }

                />

                <span>Upload PPTX / PDF</span>

              </label>



              <p>

                Uploading a new presentation opens the Slide Analysis

                page. Return to Semantic Alignment afterward.

              </p>

            </div>



            <div className="card">

              <h3>Enter your presentation transcript</h3>

              <p>

                Type or paste what you would say while presenting.

                Microphone recording is not required.

              </p>



              <textarea
                className="transcriptInput"
                value={manualTranscript}
                onChange={event => {
                  setManualTranscript(event.target.value);
                  setSemanticResult(null);
                }}
                placeholder="Example: Climate change refers to long-term changes in temperature and weather patterns..."
                rows={9}
                style={{ marginTop: 10 }}
              />



              <button

                className="primary"

                onClick={compareSemanticAlignment}

                disabled={comparing}

                style={{ marginTop: '14px' }}

              >

                {comparing

                  ? 'Comparing with Sentence-BERT...'

                  : 'Analyze Semantic Alignment'}

              </button>

            </div>



            {semanticResult && (

              <>

                <div className="metricGrid">

                  <div className="metric card">

                    <span>Overall Alignment</span>

                    <strong>

                      {semanticResult.overall_score}%

                    </strong>

                    <div className="bar">

                      <i

                        style={{

                          width: `${semanticResult.overall_score}%`

                        }}

                      />

                    </div>

                  </div>



                  <div className="metric card">

                    <span>Slides Compared</span>

                    <strong>

                      {semanticResult.total_slides}

                    </strong>

                  </div>

                </div>



                <div className="card">
                  <h3>Detailed Slide Feedback</h3>

                  {semanticResult.results.map(item => (
                    <div
                      key={item.slide_number}
                      style={{
                        padding: '20px 0',
                        borderBottom: '1px solid #e5e7eb'
                      }}
                    >
                      <div
                        style={{
                          display: 'flex',
                          justifyContent: 'space-between',
                          alignItems: 'center',
                          gap: '12px'
                        }}
                      >
                        <h3>Slide {item.slide_number}</h3>
                        <strong>{item.score}% alignment</strong>
                      </div>

                      <div
                        style={{
                          background: '#e6eef2',
                          height: '8px',
                          borderRadius: '10px',
                          margin: '12px 0',
                          overflow: 'hidden'
                        }}
                      >
                        <div
                          style={{
                            width: `${item.score}%`,
                            height: '100%',
                            background: '#0f766e',
                            borderRadius: '10px'
                          }}
                        />
                      </div>

                      <p>
                        <strong>Slide content:</strong>{' '}
                        {item.slide_text || 'No slide text available.'}
                      </p>

                      <div style={{ background: '#f0fdf4', padding: '14px', borderRadius: '8px', marginTop: '12px' }}>
                        <h4>Covered Topics</h4>
                        {item.covered_topics?.length ? (
                          <ul>
                            {item.covered_topics.map((topic, index) => (
                              <li key={index}>{topic}</li>
                            ))}
                          </ul>
                        ) : (
                          <p>No topics were confidently identified as covered.</p>
                        )}
                      </div>

                      <div style={{ background: '#fff7ed', padding: '14px', borderRadius: '8px', marginTop: '12px' }}>
                        <h4>Possibly Missing Topics</h4>
                        {item.missing_topics?.length ? (
                          <ul>
                            {item.missing_topics.map((topic, index) => (
                              <li key={index}>{topic}</li>
                            ))}
                          </ul>
                        ) : (
                          <p>No missing slide points were detected.</p>
                        )}
                      </div>

                      <div style={{ background: '#eff6ff', padding: '14px', borderRadius: '8px', marginTop: '12px' }}>
                        <h4>Key Terms to Review</h4>
                        {item.missing_key_terms?.length ? (
                          <ul>
                            {item.missing_key_terms.map((term, index) => (
                              <li key={index}>{term}</li>
                            ))}
                          </ul>
                        ) : (
                          <p>No obviously missing key terms were detected.</p>
                        )}
                      </div>

                      <div style={{ marginTop: '14px' }}>
                        <h4>Suggestions for Improvement</h4>
                        {item.recommendations?.length ? (
                          <ul>
                            {item.recommendations.map((suggestion, index) => (
                              <li key={index}>{suggestion}</li>
                            ))}
                          </ul>
                        ) : (
                          <p>
                            Review the slide and ensure that the important ideas are explained with sufficient detail.
                          </p>
                        )}
                      </div>

                      {item.feedback && (
                        <p style={{ color: '#667085', fontSize: '13px', marginTop: '12px' }}>
                          {item.feedback}
                        </p>
                      )}
                    </div>
                  ))}

                  <p style={{ fontSize: '13px', color: '#667085' }}>
                    {semanticResult.note}
                  </p>
                  <div style={{ marginTop: 12, display: 'flex', gap: 10, flexWrap: 'wrap' }}>
                    <button className="primary" onClick={() => setPage('results')}>View Results Summary</button>
                    <button className="outline" onClick={() => setPage('practice')}>Back to Practice</button>
                  </div>
                </div>

              </>

            )}

          </section>

        )}



        {page === 'results' && (

          <section className="content">

            <div className="card sectionHead">

              <div>
                <span className="pill">04 · RESULTS</span>
                <h2>Assessment summary</h2>
                <p>Summary of the metrics available in this prototype.</p>
              </div>
              <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'center' }}>
                <button className="outline" onClick={() => setPage('practice')}>Back to Practice</button>
                <button className="outline" onClick={() => window.print()}>Print / Save PDF</button>
                <button className="primary" onClick={downloadReport}>Export JSON Report</button>
              </div>

            </div>



            {!analysis && !sessionHistory.length && !semanticResult && (
              <div className="card">
                <h3>No results yet</h3>
                <p>Complete slide analysis, practice, or semantic alignment to generate your assessment summary.</p>
                <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', marginTop: 12 }}>
                  <button className="primary" onClick={() => setPage('dashboard')}>Go to Dashboard</button>
                  <button className="outline" onClick={() => setPage('analysis')}>Upload Slides</button>
                </div>
              </div>
            )}

            <div className="resultHero card">

              <div className="scoreRing big">

                {(() => {
                  const behaviorScore = Math.round((behaviorMetrics.gazePercent + behaviorMetrics.headStability + behaviorMetrics.postureScore) / 3);
                  const parts = [analysis?.overall, practice.overall, behaviorScore, semanticResult?.overall_score].filter(v => typeof v === 'number' && v > 0);
                  return parts.length ? Math.round(parts.reduce((a, b) => a + b, 0) / parts.length) : 0;
                })()}

                <small>/100</small>

              </div>



              <div>

                <h2>Final score & personalized feedback</h2>

                <p>

                  This summary combines slide analysis, delivery, behavior/expression, and semantic alignment (when available). Scores are prototype estimates, not human-validated.

                </p>

                <button

                  className="primary"

                  onClick={downloadReport}

                >

                  Download JSON Report

                </button>

              </div>

            </div>



            <div className="card">
              <h3>Practice progress history</h3>
              <p>Your last 10 saved practice sessions are stored in this browser so you can compare delivery metrics over repeated attempts.</p>
              {sessionHistory.length > 1 && (
                <div style={{ width: '100%', height: 260, margin: '12px 0 20px' }}>
                  <ResponsiveContainer>
                    <LineChart data={[...sessionHistory].reverse().map((s, i) => ({ name: `P${i + 1}`, delivery: s.deliveryScore || 0, gaze: s.gazePercent || 0, posture: s.postureScore || 0 }))}>
                      <CartesianGrid strokeDasharray="3 3" />
                      <XAxis dataKey="name" />
                      <YAxis domain={[0, 100]} />
                      <Tooltip />
                      <Line type="monotone" dataKey="delivery" stroke="#4169e1" name="Delivery" />
                      <Line type="monotone" dataKey="gaze" stroke="#17b6c5" name="Gaze" />
                      <Line type="monotone" dataKey="posture" stroke="#7c5cff" name="Posture" />
                    </LineChart>
                  </ResponsiveContainer>
                </div>
              )}
              {sessionHistory.length ? (
                <div className="historyTableWrap">
                  <table className="historyTable">
                    <thead><tr><th>Session</th><th>Duration</th><th>WPM</th><th>Fillers</th><th>Keyword</th><th>Delivery</th><th>Gaze</th><th>Posture</th></tr></thead>
                    <tbody>{sessionHistory.map(item => (
                      <tr key={item.id}>
                        <td>{item.date}</td><td>{Math.floor(item.durationSeconds / 60)}:{String(item.durationSeconds % 60).padStart(2, '0')}</td><td>{item.wpm}</td><td>{item.fillerWords}</td><td>{item.keywordOverlap}%</td><td>{item.deliveryScore}%</td><td>{item.gazePercent ?? '-'}{item.gazePercent != null ? '%' : ''}</td><td>{item.postureScore ?? '-'}{item.postureScore != null ? '%' : ''}</td>
                      </tr>
                    ))}</tbody>
                  </table>
                </div>
              ) : <p>No saved sessions yet. Complete a practice and select Stop All Devices to save your first session.</p>}
            </div>

            {semanticResult && (

              <div className="card">

                <h3>Semantic alignment result</h3>

                <p>

                  Overall slide-transcript similarity:{' '}

                  <strong>

                    {semanticResult.overall_score}%

                  </strong>

                </p>

                <button

                  className="outline"

                  onClick={() => setPage('alignment')}

                >

                  View Slide-by-Slide Results

                </button>

              </div>

            )}

            <div className="card">
              <h3>Behavioral, facial expression & pause analysis (prototype)</h3>
              <p>Behavior/expression estimated from MediaPipe Face Landmarker + Pose Landmarker; pauses analyzed from recorded practice audio.</p>
              <div className="metricGrid cols-4">
                {[['Camera-facing gaze', `${behaviorMetrics.gazePercent}%`], ['Head stability', `${behaviorMetrics.headStability}%`], ['Posture score', `${behaviorMetrics.postureScore}%`], ['Dominant expression', behaviorMetrics.dominantExpression || 'Neutral'], ['Smile / positivity', `${behaviorMetrics.smilePercent || 0}%`], ['Brow tension', `${behaviorMetrics.browTensionPercent || 0}%`], ['Jaw open (speaking)', `${behaviorMetrics.jawOpenPercent || 0}%`], ['Face detected', `${behaviorMetrics.faceDetectedPercent}%`], ['Pauses (â‰¥0.6s)', pauseMetrics.pauses], ['Long pauses (â‰¥1.5s)', pauseMetrics.longPauses], ['Avg pause', `${pauseMetrics.avgPauseMs}ms`], ['Silence', `${pauseMetrics.silencePercent}%`]].map(([name, value]) => (
                  <div className="metric card" key={name}><span>{name}</span><strong>{value}</strong></div>
                ))}
              </div>
            </div>

            <div className="card">
              <h3>Personalized feedback</h3>
              <ul>
                {!analysis && !transcript && !semanticResult && <li>Complete the flow to receive personalized feedback: upload slides, practice, and run semantic alignment.</li>}
                {analysis && <li><strong>Slides:</strong> {analysis.overall >= 75 ? 'Your slide structure/content looks strong overall.' : 'Improve slide clarity by reducing dense text, balancing visuals, and fixing flagged slide issues.'}</li>}
                {practice.words > 0 && <li><strong>Delivery:</strong> {practice.wpm < 100 ? 'Try increasing pace slightly toward 100-160 WPM.' : practice.wpm > 160 ? 'Slow down slightly and use deliberate pauses.' : 'Your speaking pace is in a good range.'} {practice.fillerCount > 3 ? `Reduce filler words (${practice.fillerCount} detected).` : 'Filler word usage is under control.'}</li>}
                {behaviorMetrics.frames > 30 && <li><strong>Non-verbal:</strong> Gaze {behaviorMetrics.gazePercent}%, posture {behaviorMetrics.postureScore}%, head stability {behaviorMetrics.headStability}%. Dominant expression: {behaviorMetrics.dominantExpression}. {behaviorMetrics.gazePercent < 60 ? 'Try facing the camera/audience more consistently.' : 'Maintain your current camera-facing presence.'}</li>}
                {pauseMetrics.pauses > 0 && <li><strong>Pauses:</strong> {pauseMetrics.longPauses > 2 ? 'Reduce long unintended pauses; keep pauses purposeful.' : 'Pause pattern looks reasonable; use pauses to emphasize key points.'}</li>}
                {semanticResult && <li><strong>Slide-speech alignment:</strong> Overall {semanticResult.overall_score}%. {semanticResult.slides_needing_review?.length ? `Review slides ${semanticResult.slides_needing_review.join(', ')} for clearer explanation.` : 'Your explanation aligns well with most slides.'}</li>}
              </ul>
            </div>



            <div className="two">

              <div className="card">

                <h3>Implemented features</h3>

                <ul>

                  <li>PPTX/PDF text extraction and original-slide preview.</li>

                  <li>Heuristic slide content/readability/structure audit.</li>

                  <li>Live camera/mic practice with audio recording and browser speech-to-text.</li>

                  <li>MediaPipe hand gesture, face/head-orientation, and posture estimation.</li>

                  <li>Sentence-BERT slide-transcript semantic alignment (FastAPI backend).</li>

                  <li>Session history, progress chart, and JSON report export.</li>

                </ul>

              </div>



              <div className="card">

                <h3>Remaining research scope</h3>

                <ul>

                  <li>Whisper timestamped STT, pause detection, Wav2Vec2/HuBERT voice features.</li>

                  <li>ViT/SlideAudit visual slide-quality model and layout/CV checks.</li>

                  <li>Multimodal fusion model trained on human-rated presentation data.</li>

                  <li>Persistent DB/auth, video storage, and validated evaluation metrics (MAE/RMSE/correlation).</li>

                </ul>

              </div>

            </div>

          </section>

        )}

      </main>

    </div>

  );

}



createRoot(document.getElementById('root')).render(<App />);
