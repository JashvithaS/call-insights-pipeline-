// ============================================================
//  IACG CallIQ — Standalone Backend v8.1
//  Changes over v8:
//   - New /repush-zoho-all endpoint for bulk re-pushing done calls
//     (use this once after removing the ZOHO_TEST_PHONES whitelist
//      to send historical calls to Zoho)
// ============================================================

const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const { createClient } = require('@supabase/supabase-js');
const axios = require('axios');
const cors = require('cors');
const FormData = require('form-data');
const { google } = require('googleapis');

require('dotenv').config();


const app = express();
app.use(express.json({ limit: '10mb' }));
app.use(cors({ origin: process.env.FRONTEND_URL || '*' }));

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_KEY
);


const fs = require('fs');

function loadGoogleCredentials() {
  // Cloud: full JSON pasted as an env var
  if (process.env.GOOGLE_SERVICE_ACCOUNT && process.env.GOOGLE_SERVICE_ACCOUNT.trim()) {
    try {
      return JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT);
    } catch (e) {
      console.error('❌ GOOGLE_SERVICE_ACCOUNT env var is not valid JSON:', e.message);
      return null;
    }
  }
  // Local: file in project root
  if (fs.existsSync('./google-service-account.json')) {
    try {
      return JSON.parse(fs.readFileSync('./google-service-account.json', 'utf8'));
    } catch (e) {
      console.error('❌ google-service-account.json is not valid JSON:', e.message);
      return null;
    }
  }
  console.warn('⚠ No Google credentials found — sheet polling disabled');
  return null;
}

const googleCredentials = loadGoogleCredentials();

const auth = googleCredentials ? new google.auth.GoogleAuth({
  credentials: googleCredentials,
  scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
}) : null;

const sheets = auth ? google.sheets({ version: 'v4', auth }) : null;

const SHEET_ID = process.env.GOOGLE_SHEET_ID;
const SHEET_NAME = process.env.GOOGLE_SHEET_NAME || 'Sheet1';


// Normalize Superfone enum values to what our pipeline expects
const SUPERFONE_CALL_TYPE = { OUTBOUND: 'Outgoing', INBOUND: 'Incoming' };
const SUPERFONE_CALL_STATUS = {
  ANSWER: 'Answered', ANSWERED: 'Answered',
  NO_ANSWER: 'Missed',  MISSED: 'Missed',
  BUSY: 'Busy',
  FAILED: 'Not Connected', NOT_CONNECTED: 'Not Connected',
  ACTIVE: 'Unknown'
};

// Extract the unique call UUID from a Superfone recording URL
function extractCallUuid(url) {
  if (!url) return null;
  const m = url.match(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i);
  return m ? m[0] : null;
}

const OPENAI_KEY = process.env.OPENAI_API_KEY;
const OPENAI_BASE = 'https://api.openai.com/v1';

const ZOHO_CLIENT_ID     = process.env.ZOHO_CLIENT_ID;
const ZOHO_CLIENT_SECRET = process.env.ZOHO_CLIENT_SECRET;
const ZOHO_REFRESH_TOKEN = process.env.ZOHO_REFRESH_TOKEN;
const ZOHO_REGION        = process.env.ZOHO_REGION || 'in';
const ZOHO_ENABLED       = !!(ZOHO_CLIENT_ID && ZOHO_CLIENT_SECRET && ZOHO_REFRESH_TOKEN);

let zohoAccessToken = null;
let zohoTokenExpires = 0;

// ============================================================
//  ZOHO SAFE PHASE 1 CONFIG  (module scope — declared ONCE)
//   • Whitelist: only push for phones in ZOHO_TEST_PHONES
//   • Dry run: ZOHO_DRY_RUN=true logs payload, does NOT send
//   • Field map: ONE place to fix Zoho API names if rejected
// ============================================================
const ZOHO_DRY_RUN = process.env.ZOHO_DRY_RUN === 'true';
const ZOHO_TEST_PHONES = (process.env.ZOHO_TEST_PHONES || '')
  .split(',').map(s => s.trim()).filter(Boolean);

const ZOHO_FIELD_MAP = {
  Customer:          (c, ai) => c.lead_name || null,
  call_date:         (c, ai) => c.call_date || null,
  Interested_or_not: (c, ai) => (ai.interest_level && ai.interest_level !== 'None') ? 'Yes' : 'No',
  Call_Status:       (c, ai) => ai.call_status || null,
  Start_Time:        (c, ai) => c.start_time || null,
  SF_Number:         (c, ai) => c.sf_number || null,
  Recording:         (c, ai, mp3) => mp3 || null,
  Full_conversation: (c, ai, mp3, tr) => tr || null,
  summary:           (c, ai) => ai.summary || null,
  AI_Call:           (c, ai) => ai.quality_category || null,
  AI_Tags:           (c, ai) => {
    const t = [];
    if (ai.walkin_interested) t.push('walkin_interested');
    if (ai.follow_up_required) t.push('follow_up_required');
    if (ai.college_interested) t.push('college_interested');
    if (typeof ai.score === 'number' && ai.score >= 75) t.push('high_score');
    if (typeof ai.score === 'number' && ai.score === 0) t.push('no_engagement');
    return t.length ? t.join(', ') : null;
  },
  AI_Call_Type:      (c, ai) => c.call_type || null,
  Ringing_Duration:  (c, ai) => c.ringing_duration || null,
  Team_Name:         (c, ai) => c.team_name || null,
  Call_Handled_by:   (c, ai) => c.call_handled_by || null,
  sentiment:         (c, ai) => ai.sentiment || null,
};

app.get('/health', (req, res) => res.json({
  status: 'ok', time: new Date().toISOString(),
  zoho_enabled: ZOHO_ENABLED, zoho_dry_run: ZOHO_DRY_RUN,
  zoho_test_phones: ZOHO_TEST_PHONES
}));

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

    const normalize = (row) => {
      const recording = row['Recording'] || row['Recording Link'] || '';
      const rawCallType = (row['Call Type'] || '').toUpperCase();
      const rawCallStatus = (row['Call Status'] || '').toUpperCase();
      return {
        sf_number:        row['SF Number'] || '',
        lead_name:        row['Customer'] || row['Lead Name'] || row['Student Name'] || row['Parent Name'] || '',
        phone_number:     row['Phone'] || row['Phone Number'] || '',
        recording_link:   recording,
        team_name:        row['Team Name'] || '',
        call_type:        SUPERFONE_CALL_TYPE[rawCallType] || row['Call Type'] || 'Unknown',
        call_status:      SUPERFONE_CALL_STATUS[rawCallStatus] || row['Call Status'] || '',
        duration:         row['Duration'] || '',
        ringing_duration: row['Ringing Duration'] || '',
        start_time:       row['Start Time'] || '',
        call_handled_by:  row['Call Handled by'] || '',
        tags:             row['Tags'] || '',
        ivr:              row['IVR'] || '',
        call_date:        row['call_date'] || row['Call Date'] || '',
        superfone_call_id: extractCallUuid(recording),
      };
    };
    const normalized = rows.map(normalize).filter(r => r.recording_link);

    const { data: batch, error: batchErr } = await supabase.from('batches')
      .insert({ file_name: req.file.originalname, total_rows: normalized.length, status: 'processing' })
      .select().single();
    if (batchErr) throw batchErr;

    const { data: calls, error: callsErr } = await supabase.from('calls')
      .insert(normalized.map(r => ({ ...r, batch_id: batch.id, status: 'pending' })))
      .select();
    if (callsErr) throw callsErr;

    res.status(202).json({ success: true, batch_id: batch.id, queued: calls.length });

    console.log(`\n📦 Batch ${batch.id} — processing ${calls.length} calls`);
    for (const call of calls) {
      try { await processOneCall(call, batch.id); }
      catch (e) {
        console.error(`  ✗ ${call.sf_number} failed:`, e.response?.data || e.message);
        await supabase.from('calls').update({ status: 'failed' }).eq('id', call.id);
      }
    }
    await supabase.from('batches').update({ status: 'done' }).eq('id', batch.id);
    console.log(`✅ Batch ${batch.id} complete\n`);

  } catch (err) {
    console.error('Upload error:', err);
    if (!res.headersSent) res.status(500).json({ error: err.message });
  }
});

// ============================================================
//  POST /reprocess — retries any pending/failed calls
// ============================================================
app.post('/reprocess', async (req, res) => {
  try {
    const { data: pendingCalls, error } = await supabase.from('calls')
      .select('*')
      .in('status', ['pending', 'failed', 'analyzing', 'transcribing'])
      .order('created_at', { ascending: false })
      .limit(100);

    if (error) throw error;
    if (!pendingCalls || pendingCalls.length === 0) {
      return res.json({ message: 'No pending/failed calls to reprocess', count: 0 });
    }

    res.status(202).json({
      success: true, count: pendingCalls.length,
      message: `Re-processing ${pendingCalls.length} calls in background`
    });

    console.log(`\n🔁 Re-processing ${pendingCalls.length} calls`);
    for (const call of pendingCalls) {
      try { await processOneCall(call, call.batch_id); }
      catch (e) {
        console.error(`  ✗ ${call.sf_number} failed:`, e.response?.data || e.message);
        await supabase.from('calls').update({ status: 'failed' }).eq('id', call.id);
      }
    }
    console.log(`✅ Reprocess complete\n`);

  } catch (err) {
    console.error('Reprocess error:', err);
    if (!res.headersSent) res.status(500).json({ error: err.message });
  }
});

async function processOneCall(call, batchId) {
  console.log(`  → ${call.sf_number} (${call.lead_name})`);
  await supabase.from('calls').update({ status: 'downloading' }).eq('id', call.id);

  console.log(`    ↓ downloading audio...`);
  let audioBuffer;
  try {
    const audio = await axios.get(call.recording_link, {
      responseType: 'arraybuffer',
      timeout: 120000
    });
    audioBuffer = Buffer.from(audio.data);
  } catch (e) {
    console.log(`    ⏭ ${call.sf_number} skipped — expired or invalid recording link`);
    await supabase.from('calls').update({
      status: 'skipped',
      call_status: 'Not Connected'
    }).eq('id', call.id);
    return;
  }

  // Detect S3 XML error page returned as 200 (expired pre-signed URL)
  const head = audioBuffer.slice(0, 64).toString('utf8');
  if (head.includes('<?xml') || head.includes('<Error>')) {
    await supabase.from('calls').update({ status: 'failed', call_status: 'Not Connected' }).eq('id', call.id);
    console.log(`    ⚠ ${call.sf_number} skipped — recording URL expired (S3 returned an error page)`);
    if (batchId) await incrementDoneCount(batchId);
    return;
  }

  if (audioBuffer.length < 5000) {
    await supabase.from('calls').update({ status: 'failed', call_status: 'Not Connected' }).eq('id', call.id);
    console.log(`    ⚠ ${call.sf_number} skipped — empty audio (${audioBuffer.length} bytes)`);
    if (batchId) await incrementDoneCount(batchId);
    return;
  }

  console.log(`    ↑ uploading to storage...`);
  const fileName = `${call.sf_number}_${call.id}.mp3`;
  const { error: upErr } = await supabase.storage.from('call-recordings').upload(fileName, audioBuffer, {
    contentType: 'audio/mpeg', upsert: true
  });
  if (upErr) throw upErr;
  const { data: pub } = supabase.storage.from('call-recordings').getPublicUrl(fileName);
  await supabase.from('calls').update({ mp3_url: pub.publicUrl, status: 'transcribing' }).eq('id', call.id);

  console.log(`    🎤 transcribing with whisper-1...`);
  const makeWhisperCall = () => {
    const form = new FormData();
    form.append('file', audioBuffer, { filename: fileName, contentType: 'audio/mpeg' });
    form.append('model', 'whisper-1');
    form.append('language', 'en');
    form.append('temperature', '0');
    form.append('prompt', `This is an admissions call from IACG / IAG Multimedia / IACG College in Hyderabad, India. The customer name is ${call.lead_name}. The agent is an IACG counselor (often a female name like Roja). Conversation discusses Intermediate programs (MPC, MEC, CGA, PCA, Humanities, Arts, Manga/Anime), campus visits at Jubilee Hills or Dilsukhnagar, fees, scholarships, walk-ins.`);
    form.append('response_format', 'verbose_json');
    return axios.post(`${OPENAI_BASE}/audio/transcriptions`, form, {
      headers: { ...form.getHeaders(), 'Authorization': `Bearer ${OPENAI_KEY}` },
      maxBodyLength: Infinity, maxContentLength: Infinity, timeout: 300000
    });
  };
  const transcriptRes = await withRetry(makeWhisperCall, 'whisper transcription');
  const rawTranscript = transcriptRes.data.text || '';

  if (rawTranscript.trim().length < 30) {
    await supabase.from('calls').update({
      status: 'failed', transcript: rawTranscript, call_status: 'Not Connected'
    }).eq('id', call.id);
    console.log(`    ⚠ ${call.sf_number} skipped — no meaningful speech`);
    if (batchId) await incrementDoneCount(batchId);
    return;
  }

  // ============================================================
  //  Speaker labeling — uses LEAD NAME to identify Customer
  // ============================================================
  console.log(`    👥 formatting as speaker dialogue...`);

  const callDirectionContext = call.call_type === 'Outgoing'
    ? `This was an OUTGOING call: the IACG agent called the customer first. The very first turn (asking "Am I speaking with ${call.lead_name}?" or similar) is the AGENT, and the customer responds confirming their identity.`
    : call.call_type === 'Incoming'
    ? `This was an INCOMING call: the customer called IACG first. The very first turn (introducing themselves or asking for info) is the CUSTOMER.`
    : `Direction unclear. Use context to determine roles.`;

  let formattedTranscript = rawTranscript;
  try {
    const dialogueRes = await withRetry(() => axios.post(`${OPENAI_BASE}/chat/completions`, {
      model: 'gpt-4o', temperature: 0,
      messages: [
        {
          role: 'system',
          content: `You are labeling speakers in a phone call transcript between an IACG admissions counselor (Agent) and a prospective student/parent (Customer).

RULES:
1. The CRM lead name is "${call.lead_name}", but the actual audio may contain a different person or name.
2. The AGENT is an IACG / IAG Multimedia / IACG College counselor (typically a female name like Roja, Priya, etc.). They are NEVER the customer.
3. ${callDirectionContext}
4. When someone says "Am I speaking with ${call.lead_name}?" — that speaker is the AGENT (asking for the customer).
5. When someone says "Yes, speaking" or confirms their name as ${call.lead_name} — that speaker is the CUSTOMER.
6. Do NOT mechanically alternate speakers. Use the actual content of each line. Either party can say "Sorry, could you say that again?", "Who is this?", or "Hello?" — assign based on context, not rigid turn-taking.
7. On garbled/confused calls where roles are unclear for a line, prefer the reading that keeps the conversation coherent.
8. Do NOT hallucinate introductions or confirmations.
9. Preserve the exact wording from the transcript.
10. If speaker identity is genuinely unclear, keep labels generic: "Agent:" and "Customer:" — only attach a customer name if the audio clearly mentions it.

Output format: each line as "Speaker: utterance"
Label the customer as "Customer (${call.lead_name})" when their identity is confirmed, otherwise just "Customer:".
Label the agent as "Agent" or "Agent (their name)" if a name is mentioned in the call.

Example (customer name clearly spoken):
Agent: Hello, am I speaking with Najeeb sir?
Customer (Najeeb): Yes, speaking.
Agent: Hi Najeeb, this is Roja calling from IACG.

Do NOT translate. Preserve all English / Telugu / Hindi code-switching exactly.`
        },
        { role: 'user', content: rawTranscript }
      ]
    }, {
      headers: { 'Authorization': `Bearer ${OPENAI_KEY}`, 'Content-Type': 'application/json' },
      timeout: 90000
    }), 'dialogue formatting', 2);

    const candidate = dialogueRes.data.choices[0].message.content || '';
    const customerMentions = (candidate.match(/Customer\s*[:(]/gi) || []).length;
    const agentMentions = (candidate.match(/Agent\s*[:(]/gi) || []).length;

    if (
      customerMentions === 0 ||
      agentMentions === 0 ||
      candidate.length < rawTranscript.length * 0.5
    ) {
      console.log(`    ⚠ bad speaker formatting detected — using raw transcript`);
      formattedTranscript = rawTranscript;
    } else {
      formattedTranscript = candidate;
    }
  } catch (e) {
    console.log(`    ⚠ dialogue formatting failed — using raw transcript instead (${e.message})`);
  }

  await supabase.from('calls').update({
    raw_transcript: rawTranscript,
    transcript: formattedTranscript,
    status: 'analyzing'
  }).eq('id', call.id);

  const { data: rules } = await supabase.from('rules').select('name,description,condition,action').eq('enabled', true);
  const rulesText = (rules || []).map(r => `- ${r.name}: ${r.condition} → ${r.action}`).join('\n');

  console.log(`    🧠 analyzing with GPT-4o...`);

  const today = new Date().toISOString().slice(0, 10);

  const systemPrompt = `You are an expert admissions call analyst for IACG (Indian Academy of Creative Gaming / IACG College) in Hyderabad. You score calls between counselors and prospective student leads.

## IACG BUSINESS CONTEXT

**Campuses**: Jubilee Hills (Madhapur) and Dilsukhnagar, Hyderabad.

**Program structure** — TWO separate fields:

The **program** is the qualification LEVEL. Use EXACTLY one of:
- Intermediate, B.Tech, B.Com, MBA

The **course** is the specific STREAM within the program. Use EXACTLY one of:
- MPC (Math, Physics, Chemistry)
- MEC (Math, Economics, Commerce)
- CGA (Computer Graphics & Animation)
- PCA
- Humanities
- Arts
- Manga (Kyoto Seika University, Japan partnership)
- Anime (Kyoto Seika University, Japan partnership)

**Key features**: 360-degree psychometric career assessment, scholarships up to 50%, Telangana State Board.

**Counselor goal hierarchy** (best → worst outcome):
1. Visit confirmed with date/time → BEST
2. Customer agrees to follow-up / info → GOOD
3. Interest shown, no commitment → MODERATE
4. Not interested / wrong number → POOR
5. Call not connected → FAILED

## SCORING RUBRIC (apply strictly):

- **90–100 (Exceptional)**: Visit confirmed with specific date + lead qualified + counselor handled barriers excellently
- **75–89 (High Quality)**: Visit booked OR strong qualification + good rapport + clear program discussion
- **60–74 (Good)**: Strong interest, follow-up agreed, programs explained
- **45–59 (Average)**: Info exchanged but no commitment
- **30–44 (Below Average)**: Confusion, poor delivery
- **0–29 (Poor)**: Failed, wrong person, no useful interaction

## QUALITY_CATEGORY mapping (MUST follow):
- score 75+ → "High Quality"
- score 50–74 → "Average"
- score 30–49 → "Poor"
- score 0–29 → "Very Poor"

## EXTRACTION RULES:

**program** (qualification LEVEL only — NEVER a stream name):
- Must be EXACTLY one of: Intermediate, B.Tech, B.Com, MBA
- For IACG calls this is almost always "Intermediate"
- If only a stream (MPC/CGA/etc.) is mentioned, INFER "Intermediate"
- null only if no studies were discussed at all

**course** (specific STREAM only — NEVER the qualification level):
- Must be EXACTLY one of: MPC, MEC, CGA, PCA, Humanities, Arts, Manga, Anime
- null if no specific stream was named
- "Commerce" is NOT a valid course — use MEC (if Intermediate) or B.Com (if degree-level)

⚠️ CRITICAL: MPC / MEC / CGA / PCA / Humanities / Arts / Manga / Anime are COURSES, never programs.

**walkin_interested**: true if customer expressed willingness to visit OR asked about campus, even without date.

**walkin_date**: ONLY set if a SPECIFIC date was agreed (e.g. "Friday", "December 15th", "tomorrow"). NEVER invent. Today is ${today} — interpret relative dates from this.

**interest_level**:
- "High" → visit booked OR clear strong intent
- "Medium" → engaged, considering options
- "Low" → just listening, no commitment
- "None" → not interested, wrong number

**call_status**: Answered | Missed | Busy | Not Connected | Unknown

## ACTIVE RULES:
${rulesText || '(none configured)'}

Return ONLY valid JSON. No markdown.`;

  const userPrompt = `Lead: ${call.lead_name} | Phone: ${call.phone_number || 'N/A'} | Team: ${call.team_name || 'N/A'} | Type: ${call.call_type}

Speaker-labeled transcript:
${formattedTranscript}

Analyze this call. Return JSON:
{
  "program": "one of: Intermediate | B.Tech | B.Com | MBA | null",
  "course": "one of: MPC | MEC | CGA | PCA | Humanities | Arts | Manga | Anime | null",
  "college_interested": boolean,
  "walkin_interested": boolean,
  "walkin_date": "YYYY-MM-DD or null",
  "follow_up_required": boolean,
  "call_status": "Answered|Missed|Busy|Not Connected|Unknown",
  "sentiment": "Positive|Neutral|Negative",
  "quality_category": "High Quality|Average|Poor|Very Poor",
  "score": integer 0-100,
  "interest_level": "High|Medium|Low|None",
  "summary": "2-3 sentence summary"
}`;

  const llmRes = await withRetry(() => axios.post(`${OPENAI_BASE}/chat/completions`, {
    model: 'gpt-4o', temperature: 0.1,
    response_format: { type: 'json_object' },
    messages: [{ role: 'system', content: systemPrompt }, { role: 'user', content: userPrompt }]
  }, {
    headers: { 'Authorization': `Bearer ${OPENAI_KEY}`, 'Content-Type': 'application/json' },
    timeout: 180000
  }), 'gpt-4o analysis');
  const ai = JSON.parse(llmRes.data.choices[0].message.content);

  await supabase.from('calls').update({
    program: ai.program, course: ai.course,
    college_interested: ai.college_interested,
    walkin_interested: ai.walkin_interested,
    walkin_date: ai.walkin_date,
    follow_up_required: ai.follow_up_required,
    call_status: ai.call_status, sentiment: ai.sentiment,
    quality_category: ai.quality_category, score: ai.score,
    interest_level: ai.interest_level,
    summary: ai.summary,
    status: 'done'
  }).eq('id', call.id);

  if (batchId) await incrementDoneCount(batchId);

  if (ZOHO_ENABLED) {
    try {
      await pushToZoho(call, ai, pub.publicUrl, formattedTranscript);
    } catch (e) {
      console.error(`    ⚠ Zoho push failed:`, e.response?.data || e.message);
    }
  }

  console.log(`    ✓ ${call.sf_number} done (score: ${ai.score}, sentiment: ${ai.sentiment}, quality: ${ai.quality_category})`);
}

// Retry wrapper for flaky network calls
async function withRetry(fn, label, attempts = 3) {
  let lastErr;
  for (let i = 1; i <= attempts; i++) {
    try { return await fn(); }
    catch (e) {
      lastErr = e;
      const msg = e.response?.data?.error?.message || e.message;
      console.log(`    ↻ ${label} attempt ${i}/${attempts} failed: ${msg}`);
      if (i < attempts) await new Promise(r => setTimeout(r, 2000 * i));
    }
  }
  throw lastErr;
}

async function incrementDoneCount(batchId) {
  const { data: b } = await supabase.from('batches').select('done_rows').eq('id', batchId).single();
  if (b) await supabase.from('batches').update({ done_rows: (b.done_rows || 0) + 1 }).eq('id', batchId);
}

async function getZohoAccessToken() {
  if (zohoAccessToken && Date.now() < zohoTokenExpires) return zohoAccessToken;
  const res = await axios.post(`https://accounts.zoho.${ZOHO_REGION}/oauth/v2/token`, null, {
    params: { refresh_token: ZOHO_REFRESH_TOKEN, client_id: ZOHO_CLIENT_ID, client_secret: ZOHO_CLIENT_SECRET, grant_type: 'refresh_token' }
  });
  zohoAccessToken = res.data.access_token;
  zohoTokenExpires = Date.now() + (res.data.expires_in - 60) * 1000;
  return zohoAccessToken;
}

// Search a Zoho Lead by phone. Tries Zoho's native phone search (scans ALL
// phone-type fields incl. Mobile), then a Phone-or-Mobile criteria fallback.
async function findZohoLeadByPhone(apiBase, token, phone) {
  const digits = (phone || '').replace(/\D/g, '');
  if (!digits) return null;
  const last10 = digits.slice(-10);
  const get = (url) => axios.get(url, {
    headers: { 'Authorization': `Zoho-oauthtoken ${token}` },
    validateStatus: s => s < 500
  });

  // Strategy 1 — Zoho native phone search (searches Phone, Mobile, etc.)
  const phoneVariants = [digits, last10, '+' + digits, '+91' + last10];
  for (const p of [...new Set(phoneVariants)]) {
    const r = await get(`${apiBase}/crm/v8/Leads/search?phone=${encodeURIComponent(p)}`);
    if (r.status === 200 && r.data?.data?.length) {
      console.log(`    🔎 matched Zoho lead via phone search "${p}"`);
      return r.data.data[0];
    }
  }

  // Strategy 2 — explicit criteria on both Phone and Mobile fields
  const critVariants = [phone, '+' + digits, digits, '+91' + last10, '91' + last10, last10];
  for (const p of [...new Set(critVariants)]) {
    if (!p) continue;
    const crit = encodeURIComponent(`((Phone:equals:${p})or(Mobile:equals:${p}))`);
    const r = await get(`${apiBase}/crm/v8/Leads/search?criteria=${crit}`);
    if (r.status === 200 && r.data?.data?.length) {
      console.log(`    🔎 matched Zoho lead via criteria "${p}"`);
      return r.data.data[0];
    }
  }

  console.log(`    🔎 no Zoho match for any format of ${phone} (tried Phone + Mobile fields)`);
  return null;
}

async function pushToZoho(call, ai, mp3Url, transcript) {
  // GUARDRAIL 1 — Whitelist: skip phones not explicitly approved for testing
  if (ZOHO_TEST_PHONES.length > 0 && !ZOHO_TEST_PHONES.includes(call.phone_number)) {
    console.log(`    ⏭ Zoho skipped — ${call.phone_number} not in ZOHO_TEST_PHONES`);
    return;
  }

  // Build payload from config map — drops null/empty values
  const payload = {};
  for (const [field, valueFn] of Object.entries(ZOHO_FIELD_MAP)) {
    const v = valueFn(call, ai, mp3Url, transcript);
    if (v !== null && v !== undefined && v !== '') payload[field] = v;
  }

  // GUARDRAIL 2 — Dry run: log what WOULD be sent, do not call Zoho
  if (ZOHO_DRY_RUN) {
    console.log(`    🧪 ZOHO_DRY_RUN — payload for ${call.phone_number} (NOT sent):`);
    console.log('    ' + JSON.stringify(payload, null, 2).split('\n').join('\n    '));
    return;
  }

  const token = await getZohoAccessToken();
  const apiBase = ZOHO_REGION === 'in' ? 'https://www.zohoapis.in' :
                  ZOHO_REGION === 'eu' ? 'https://www.zohoapis.eu' : 'https://www.zohoapis.com';

  // Find existing lead by phone (across Phone and Mobile fields)
  const lead = await findZohoLeadByPhone(apiBase, token, call.phone_number);

  if (!lead) {
    // Create a new lead if none exists
    console.log(`    ➕ Creating new Zoho lead`);
    const createPayload = {
      Last_Name: call.lead_name || 'Unknown',
      Phone: call.phone_number || '',
      Mobile: call.phone_number || '',
      ...payload
    };
    const createRes = await axios.post(
      `${apiBase}/crm/v8/Leads`,
      { data: [createPayload] },
      { headers: { 'Authorization': `Zoho-oauthtoken ${token}` }, validateStatus: s => s < 500 }
    );
    const result = createRes.data?.data?.[0];
    if (result?.status === 'success') {
      console.log(`    ✓ New Zoho lead created (id: ${result.details?.id})`);
    } else {
      console.log(`    ⚠ Zoho create issue:`, JSON.stringify(createRes.data));
    }
    return;
  }

  // Update ONLY AI Calling fields — never touches Last_Name / Lead_Source / Owner / Stage / Notes / Tags
  const updateRes = await axios.put(
    `${apiBase}/crm/v8/Leads/${lead.id}`,
    { data: [payload] },
    { headers: { 'Authorization': `Zoho-oauthtoken ${token}` }, validateStatus: s => s < 500 }
  );
  const result = updateRes.data?.data?.[0];
  if (result?.status === 'success') {
    console.log(`    ✓ Zoho lead ${lead.id} updated (${Object.keys(payload).length} fields)`);
  } else {
    console.log(`    ⚠ Zoho update issue:`, JSON.stringify(updateRes.data));
  }
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

app.post('/repush-zoho/:id', async (req, res) => {
  try {
    const { data: call } = await supabase
      .from('calls')
      .select('*')
      .eq('id', req.params.id)
      .single();

    if (!call) return res.status(404).json({ error: 'Call not found' });

    const ai = {
      program: call.program,
      course: call.course,
      college_interested: call.college_interested,
      walkin_interested: call.walkin_interested,
      walkin_date: call.walkin_date,
      follow_up_required: call.follow_up_required,
      call_status: call.call_status,
      sentiment: call.sentiment,
      quality_category: call.quality_category,
      score: call.score,
      interest_level: call.interest_level,
      summary: call.summary
    };

    await pushToZoho(call, ai, call.mp3_url, call.transcript);
    res.json({ success: true });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

// ============================================================
//  POST /repush-zoho-all — bulk re-push every "done" row to Zoho
//  Use this once after removing the ZOHO_TEST_PHONES whitelist
//  to send historical calls that were skipped during testing.
// ============================================================
app.post('/repush-zoho-all', async (req, res) => {
  try {
    const { data: calls, error } = await supabase
      .from('calls').select('*')
      .eq('status', 'done')
      .not('summary', 'is', null)
      .order('updated_at', { ascending: false })
      .limit(500);

    if (error) return res.status(500).json({ error: error.message });
    if (!calls || calls.length === 0) {
      return res.json({ queued: 0, message: 'No done calls with summaries to repush' });
    }

    res.status(202).json({
      queued: calls.length,
      message: `Bulk repush started for ${calls.length} calls — watch server logs for progress`
    });

    console.log(`\n🔁 Bulk repush — ${calls.length} done calls`);
    let pushed = 0, failed = 0, skipped = 0;

    for (const call of calls) {
      if (!call.phone_number) { skipped++; continue; }
      const ai = {
        program: call.program,
        course: call.course,
        college_interested: call.college_interested,
        walkin_interested: call.walkin_interested,
        walkin_date: call.walkin_date,
        follow_up_required: call.follow_up_required,
        call_status: call.call_status,
        sentiment: call.sentiment,
        quality_category: call.quality_category,
        score: call.score,
        interest_level: call.interest_level,
        summary: call.summary
      };
      try {
        await pushToZoho(call, ai, call.mp3_url, call.transcript);
        pushed++;
      } catch (e) {
        failed++;
        console.error(`  ✗ ${call.sf_number || call.id} repush failed:`, e.response?.data || e.message);
      }
    }
    console.log(`✅ Bulk repush done — ${pushed} pushed, ${failed} failed, ${skipped} skipped (no phone)\n`);

  } catch (err) {
    console.error('Bulk repush error:', err);
    if (!res.headersSent) res.status(500).json({ error: err.message });
  }
});

// ============================================================
//  pollGoogleSheet — fetch sheet, dedup by UUID, queue new rows
// ============================================================
async function pollGoogleSheet() {
  if (!sheets || !SHEET_ID) return;
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID, range: SHEET_NAME,
    });
    const rows = res.data.values;
    if (!rows || rows.length < 2) {
      console.log('📄 Sheet check — no rows');
      return;
    }

    const headers = rows[0];
    let newCount = 0, existingCount = 0, noRecording = 0, tooShort = 0;

    for (let i = 1; i < rows.length; i++) {
      const row = {};
      headers.forEach((h, idx) => { row[h] = rows[i][idx] || ''; });

      const recording = (row['Recording'] || '').trim();
      if (!recording || recording === '-') { noRecording++; continue; }

      const uuid = extractCallUuid(recording);
      if (!uuid) {
        console.log(`⏭ ${row['Customer'] || '(unnamed)'} — recording URL has no parseable call ID`);
        continue;
      }

      // Skip calls that clearly had no conversation (saves Whisper credits)
      const durationStr = (row['Duration'] || '').toString().trim();
      const durSec = parseInt(durationStr) || 0;
      if (!durationStr || durSec < 10) {
        tooShort++;
        continue;
      }

      // Dedup by the REAL unique call ID, not by the company-phone SF Number
      const { data: existing } = await supabase
        .from('calls').select('id')
        .eq('superfone_call_id', uuid)
        .maybeSingle();
      if (existing) { existingCount++; continue; }

      const rawCallType = (row['Call Type'] || '').toUpperCase();
      const rawCallStatus = (row['Call Status'] || '').toUpperCase();

      const normalized = {
        sf_number:        row['SF Number'] || '',
        lead_name:        row['Customer'] || row['Student Name'] || row['Parent Name'] || '',
        phone_number:     row['Phone'] || '',
        recording_link:   recording,
        team_name:        row['Team Name'] || '',
        call_type:        SUPERFONE_CALL_TYPE[rawCallType] || row['Call Type'] || 'Unknown',
        call_status:      SUPERFONE_CALL_STATUS[rawCallStatus] || '',
        duration:         durationStr,
        ringing_duration: row['Ringing Duration'] || '',
        start_time:       row['Start Time'] || '',
        call_handled_by:  row['Call Handled by'] || '',
        tags:             row['Tags'] || '',
        call_date:        row['call_date'] || '',
        superfone_call_id: uuid,
        status: 'pending'
      };

      const { data: inserted, error: insertErr } = await supabase
        .from('calls').insert(normalized).select().single();

      if (insertErr) {
        // Unique index will reject any sneaky duplicate
        if (insertErr.code === '23505') { existingCount++; continue; }
        console.error(`❌ Insert failed for ${uuid.slice(0,8)}:`, insertErr.message);
        continue;
      }

      newCount++;
      console.log(`🆕 ${uuid.slice(0,8)} — ${normalized.lead_name} (${normalized.phone_number}, ${durSec}s)`);
      try {
        await processOneCall(inserted);
      } catch (e) {
        console.error(`  ✗ ${uuid.slice(0,8)} processing failed:`, e.message);
      }
    }

    console.log(`📄 Sheet check — ${newCount} new, ${existingCount} already done, ${noRecording} no-recording, ${tooShort} too-short`);
  } catch (err) {
    console.error('Google Sheet Error:', err.message);
  }
}

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`✅ CallIQ standalone server v8.1 running on http://localhost:${PORT}`);
  console.log(`   Whisper-1 + GPT-4o with IACG-specific analysis`);
  console.log(`   Zoho CRM push: ${ZOHO_ENABLED ? 'ON' : 'OFF'}  |  DRY_RUN: ${ZOHO_DRY_RUN ? 'ON (nothing sent)' : 'OFF (live)'}`);
  console.log(`   Zoho test phones: ${ZOHO_TEST_PHONES.length ? ZOHO_TEST_PHONES.join(', ') : '(none — all phones allowed)'}`);
  if (sheets && SHEET_ID) {
    console.log(`   Google Sheet polling every 30s\n`);
    setInterval(async () => { await pollGoogleSheet(); }, 30000);
  } else {
    console.log(`   ⚠ Google Sheet polling disabled (missing credentials or GOOGLE_SHEET_ID)\n`);
  }
});