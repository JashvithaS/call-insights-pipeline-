// ============================================================
//  IACG CallIQ — Standalone Backend (OpenAI edition)
//  Run with: node server.js
// ============================================================

const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const { createClient } = require('@supabase/supabase-js');
const axios = require('axios');
const cors = require('cors');
const FormData = require('form-data');
require('dotenv').config();

const app = express();
app.use(express.json({ limit: '10mb' }));
app.use(cors({ origin: process.env.FRONTEND_URL || '*' }));

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });
const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_SERVICE_KEY);

const OPENAI_KEY = process.env.OPENAI_API_KEY;
const OPENAI_BASE = 'https://api.openai.com/v1';

app.get('/health', (req, res) => res.json({ status: 'ok', time: new Date().toISOString() }));

app.post('/upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
    const ext = req.file.originalname.split('.').pop().toLowerCase();
    if (!['xlsx','csv','xls'].includes(ext)) {
      return res.status(400).json({ error: 'Only .xlsx, .xls, .csv accepted' });
    }
    const wb = xlsx.read(req.file.buffer, { type: 'buffer' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = xlsx.utils.sheet_to_json(ws, { defval: '' });
    if (rows.length === 0) return res.status(400).json({ error: 'Excel sheet is empty' });

    const normalize = (row) => ({
      sf_number:      row['SF Number']     || row['sf_number']     || '',
      lead_name:      row['Lead Name']     || row['lead_name']     || row['Name']     || '',
      phone_number:   row['Phone Number']  || row['Phone']         || row['phone_number'] || '',
      recording_link: row['Recording Link']|| row['recording_link']|| row['URL']      || '',
      team_name:      row['Team Name']     || row['team_name']     || row['Team']     || '',
      call_type:      row['Call Type']     || row['call_type']     || 'Unknown',
    });
    const normalized = rows.map(normalize).filter(r => r.recording_link);

    const { data: batch, error: batchErr } = await supabase.from('batches')
      .insert({ file_name: req.file.originalname, total_rows: normalized.length, status: 'processing' })
      .select().single();
    if (batchErr) throw batchErr;

    const { data: calls, error: callsErr } = await supabase.from('calls')
      .insert(normalized.map(r => ({ ...r, batch_id: batch.id, status: 'pending' })))
      .select();
    if (callsErr) throw callsErr;

    res.status(202).json({ success: true, batch_id: batch.id, queued: calls.length,
      message: `Processing ${calls.length} calls in background — watch server console` });

    console.log(`\n📦 Batch ${batch.id} — processing ${calls.length} calls`);
    for (const call of calls) {
      try { await processOneCall(call, batch.id); }
      catch (e) {
        console.error(`  ✗ ${call.sf_number} failed:`, e.response?.data || e.message);
        await supabase.from('calls').update({ status: 'failed', error_message: e.message }).eq('id', call.id);
      }
    }
    await supabase.from('batches').update({ status: 'done' }).eq('id', batch.id);
    console.log(`✅ Batch ${batch.id} complete\n`);

  } catch (err) {
    console.error('Upload error:', err);
    if (!res.headersSent) res.status(500).json({ error: err.message });
  }
});

async function processOneCall(call, batchId) {
  console.log(`  → ${call.sf_number} (${call.lead_name})`);
  await supabase.from('calls').update({ status:'downloading', processing_started:new Date() }).eq('id', call.id);

  console.log(`    ↓ downloading audio...`);
  const audio = await axios.get(call.recording_link, { responseType: 'arraybuffer', timeout: 60000 });
  const audioBuffer = Buffer.from(audio.data);

  console.log(`    ↑ uploading to storage...`);
  const fileName = `${call.sf_number}_${call.id}.mp3`;
  const { error: upErr } = await supabase.storage.from('call-recordings').upload(fileName, audioBuffer, {
    contentType: 'audio/mpeg', upsert: true
  });
  if (upErr) throw upErr;
  const { data: pub } = supabase.storage.from('call-recordings').getPublicUrl(fileName);
  await supabase.from('calls').update({ mp3_url: pub.publicUrl, status: 'transcribing' }).eq('id', call.id);

  console.log(`    🎤 transcribing with Whisper...`);
  const form = new FormData();
  form.append('file', audioBuffer, { filename: fileName, contentType: 'audio/mpeg' });
  form.append('model', 'whisper-1');
  form.append('response_format', 'verbose_json');
  const whisperRes = await axios.post(`${OPENAI_BASE}/audio/transcriptions`, form, {
    headers: { ...form.getHeaders(), 'Authorization': `Bearer ${OPENAI_KEY}` },
    maxBodyLength: Infinity, maxContentLength: Infinity, timeout: 300000
  });
  const transcript = whisperRes.data.text || '';
  const language = whisperRes.data.language || 'unknown';
  await supabase.from('calls').update({ transcript, status: 'analyzing' }).eq('id', call.id);

  const { data: rules } = await supabase.from('rules').select('name,description,condition,action').eq('enabled', true);
  const rulesText = (rules || []).map(r => `- ${r.name}: ${r.condition} → ${r.action}`).join('\n');

  console.log(`    🧠 analyzing with GPT-4o...`);
  const systemPrompt = `You are an expert admissions call analyst for IACG College, Hyderabad. Analyze the transcript and return ONLY valid JSON matching the schema. No markdown, no extra text.

ACTIVE RULES:
${rulesText || '(none)'}

SCORING (0-100): 90-100 exceptional, 70-89 good, 50-69 average, 30-49 below avg, 0-29 poor.`;

  const userPrompt = `Call: ${call.sf_number} | Lead: ${call.lead_name} | Team: ${call.team_name} | Type: ${call.call_type}
Language: ${language}

Transcript:
${transcript}

Return JSON exactly matching:
{
  "program": "string or null",
  "course": "string or null",
  "college_interested": boolean,
  "walkin_interested": boolean,
  "walkin_date": "YYYY-MM-DD or null",
  "follow_up_required": boolean,
  "call_status": "Answered|Missed|Busy|Not Connected|Unknown",
  "sentiment": "Positive|Neutral|Negative",
  "quality_category": "High Quality|Average|Poor|Very Poor",
  "score": integer 0-100,
  "confidence_score": number 0-1,
  "interest_level": "High|Medium|Low|None",
  "lead_qualified": boolean,
  "summary": "2-3 sentence summary",
  "action_items": ["array of strings"],
  "discussion_points": ["array of strings"],
  "objections": ["array of strings"]
}`;

  const llmRes = await axios.post(`${OPENAI_BASE}/chat/completions`, {
    model: 'gpt-4o', temperature: 0.2,
    response_format: { type: 'json_object' },
    messages: [{ role: 'system', content: systemPrompt }, { role: 'user', content: userPrompt }]
  }, {
    headers: { 'Authorization': `Bearer ${OPENAI_KEY}`, 'Content-Type': 'application/json' },
    timeout: 120000
  });
  const ai = JSON.parse(llmRes.data.choices[0].message.content);

  await supabase.from('calls').update({
    program: ai.program, course: ai.course,
    college_interested: ai.college_interested,
    walkin_interested: ai.walkin_interested,
    walkin_date: ai.walkin_date,
    follow_up_required: ai.follow_up_required,
    call_status: ai.call_status, sentiment: ai.sentiment,
    quality_category: ai.quality_category, score: ai.score,
    confidence_score: ai.confidence_score,
    interest_level: ai.interest_level, lead_qualified: ai.lead_qualified,
    summary: ai.summary,
    action_items: ai.action_items, discussion_points: ai.discussion_points,
    objections: ai.objections,
    raw_ai_json: ai, status: 'done', processing_done: new Date()
  }).eq('id', call.id);

  const { data: b } = await supabase.from('batches').select('done_rows').eq('id', batchId).single();
  await supabase.from('batches').update({ done_rows: (b?.done_rows || 0) + 1 }).eq('id', batchId);
  console.log(`    ✓ ${call.sf_number} done (score: ${ai.score}, sentiment: ${ai.sentiment})`);
}

app.get('/batch/:id', async (req, res) => {
  const { data, error } = await supabase.from('batches').select('*').eq('id', req.params.id).single();
  if (error) return res.status(404).json({ error: 'Batch not found' });
  res.json(data);
});

app.get('/calls', async (req, res) => {
  const { batch_id, page = 1, limit = 50 } = req.query;
  const offset = (page - 1) * limit;
  let q = supabase.from('calls').select('*', { count: 'exact' }).order('created_at', { ascending: false }).range(offset, offset + parseInt(limit) - 1);
  if (batch_id) q = q.eq('batch_id', batch_id);
  const { data, error, count } = await q;
  if (error) return res.status(500).json({ error: error.message });
  res.json({ calls: data, total: count });
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`✅ CallIQ standalone server running on http://localhost:${PORT}`);
  console.log(`   Using OpenAI (Whisper-1 + GPT-4o)`);
  console.log(`   Frontend should POST Excel files to /upload\n`);
});
