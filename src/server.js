require('dotenv').config();
const express  = require('express');
const cors     = require('cors');
const path     = require('path');
const { joinMeeting } = require('./bot');
const logger   = require('./logger');

const app  = express();
const PORT = parseInt(process.env.PORT || '3000', 10);
const AUTO_REFRESH_MS = parseInt(process.env.AUTO_REFRESH_MS || '0', 10);

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, '../public')));

// ── In-memory session store ────────────────────────────────────────────────
// Key: meetingUrl string → Value: SessionManager instance
const sessions = new Map();

// ── Helper ─────────────────────────────────────────────────────────────────
function sessionSummary(session) {
  return {
    status:       session.status,
    meetingUrl:   session.meetingUrl,
    participants: session.participants,
    count:        session.participants.length,
  };
}

// ── Routes ─────────────────────────────────────────────────────────────────

/**
 * GET /api/health
 * Health check — confirm server is running.
 */
app.get('/api/health', (_req, res) => {
  res.json({ ok: true, activeSessions: sessions.size });
});

/**
 * POST /api/join
 * Body: { meetingUrl: string }
 *
 * Launches the Playwright bot and joins the Teams meeting in a browser.
 * Scrapes and returns the initial participant list.
 */
app.post('/api/join', async (req, res) => {
  const { meetingUrl } = req.body;

  if (!meetingUrl || typeof meetingUrl !== 'string') {
    return res.status(400).json({ error: 'meetingUrl is required.' });
  }
  if (!meetingUrl.includes('teams.microsoft.com') && !meetingUrl.includes('teams.live.com')) {
    return res.status(400).json({ error: 'URL does not look like a Teams meeting link.' });
  }
  if (sessions.has(meetingUrl)) {
    return res.status(409).json({
      error: 'Already joined this meeting.',
      session: sessionSummary(sessions.get(meetingUrl)),
    });
  }

  try {
    logger.info(`JOIN request received for: ${meetingUrl}`);
    const session = await joinMeeting(meetingUrl);
    sessions.set(meetingUrl, session);

    // Optional: auto-refresh participant list on an interval
    if (AUTO_REFRESH_MS > 0) {
      const timer = setInterval(async () => {
        if (!sessions.has(meetingUrl)) { clearInterval(timer); return; }
        await session.scrapeParticipants().catch(() => {});
      }, AUTO_REFRESH_MS);
      session._refreshTimer = timer;
    }

    res.json({ success: true, session: sessionSummary(session) });
  } catch (err) {
    logger.error(`JOIN failed: ${err.message}`);
    res.status(500).json({ error: err.message });
  }
});

/**
 * GET /api/participants?meetingUrl=...
 *
 * Re-scrapes the participant list from an active session.
 */
app.get('/api/participants', async (req, res) => {
  const { meetingUrl } = req.query;

  if (!meetingUrl || !sessions.has(meetingUrl)) {
    return res.status(404).json({ error: 'No active session. Join a meeting first.' });
  }

  try {
    const session = sessions.get(meetingUrl);
    const participants = await session.scrapeParticipants();
    res.json({ participants, count: participants.length });
  } catch (err) {
    logger.error(`REFRESH failed: ${err.message}`);
    res.status(500).json({ error: err.message });
  }
});

/**
 * GET /api/sessions
 *
 * Lists all active bot sessions.
 */
app.get('/api/sessions', (_req, res) => {
  const list = [...sessions.values()].map(sessionSummary);
  res.json({ sessions: list });
});

/**
 * POST /api/pick
 * Body: { meetingUrl: string, exclude?: string[] }
 *
 * Picks a random participant from the current list.
 * Optionally exclude certain names (e.g. the host, already-picked people).
 */
app.post('/api/pick', async (req, res) => {
  const { meetingUrl, exclude = [] } = req.body;

  if (!meetingUrl || !sessions.has(meetingUrl)) {
    return res.status(404).json({ error: 'No active session. Join a meeting first.' });
  }

  try {
    const session = sessions.get(meetingUrl);

    await session.scrapeParticipants();

    const pool = session.participants.filter((n) => !exclude.includes(n));
    if (pool.length === 0) {
      return res.status(422).json({
        error: 'No participants available after exclusions.',
      });
    }

    const picked = pool[Math.floor(Math.random() * pool.length)];
    logger.info(`Picked: "${picked}" from ${pool.length} eligible participant(s)`);
    res.json({ picked, pool: pool.length, total: session.participants.length });
  } catch (err) {
    logger.error(`PICK failed: ${err.message}`);
    res.status(500).json({ error: err.message });
  }
});

/**
 * POST /api/pick-and-post
 * Body: { meetingUrl: string }
 *
 * Automatically picks a random participant (without repeating) and posts
 * their name to the Teams meeting chat as "Next speaker: [Name]"
 */
app.post('/api/pick-and-post', async (req, res) => {
  const { meetingUrl } = req.body;

  if (!meetingUrl || !sessions.has(meetingUrl)) {
    return res.status(404).json({ error: 'No active session. Join a meeting first.' });
  }

  try {
    const session = sessions.get(meetingUrl);

    const result = await session.pickAndPostToChat();

    if (!result.success) {
      return res.status(500).json({ error: result.error || 'Failed to pick and post' });
    }

    res.json({
      success: true,
      picked: result.picked,
      totalPicked: result.totalPicked,
      totalParticipants: result.totalParticipants,
      message: `Posted "${result.picked}" to chat`,
    });
  } catch (err) {
    logger.error(`PICK-AND-POST failed: ${err.message}`);
    res.status(500).json({ error: err.message });
  }
});

/**
 * POST /api/reset-picks
 * Body: { meetingUrl: string }
 *
 * Resets the list of picked participants, allowing names to be picked again.
 */
app.post('/api/reset-picks', (req, res) => {
  const { meetingUrl } = req.body;

  if (!meetingUrl || !sessions.has(meetingUrl)) {
    return res.status(404).json({ error: 'No active session. Join a meeting first.' });
  }

  try {
    const session = sessions.get(meetingUrl);
    session.resetPickedNames();
    res.json({
      success: true,
      message: 'Picked names reset — all participants available again',
      totalParticipants: session.participants.length,
    });
  } catch (err) {
    logger.error(`RESET-PICKS failed: ${err.message}`);
    res.status(500).json({ error: err.message });
  }
});

/**
 * POST /api/leave
 * Body: { meetingUrl: string }
 *
 * Closes the browser and cleans up the session.
 */
app.post('/api/leave', async (req, res) => {
  const { meetingUrl } = req.body;

  if (!meetingUrl || !sessions.has(meetingUrl)) {
    return res.status(404).json({ error: 'No active session for this URL.' });
  }

  try {
    const session = sessions.get(meetingUrl);
    if (session._refreshTimer) clearInterval(session._refreshTimer);
    await session.leave();
    sessions.delete(meetingUrl);
    logger.info('Session ended and cleaned up.');
    res.json({ success: true });
  } catch (err) {
    logger.error(`LEAVE failed: ${err.message}`);
    res.status(500).json({ error: err.message });
  }
});

// ── Graceful shutdown ──────────────────────────────────────────────────────
async function shutdown() {
  logger.info('Shutting down — closing all sessions...');
  for (const [url, session] of sessions) {
    if (session._refreshTimer) clearInterval(session._refreshTimer);
    await session.leave().catch(() => {});
    sessions.delete(url);
  }
  process.exit(0);
}
process.on('SIGINT', shutdown);
process.on('SIGTERM', shutdown);

// ── Start ──────────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  logger.info(`Teams Picker server running → http://localhost:${PORT}`);
  logger.info(`Headless mode: ${process.env.HEADLESS === 'true' ? 'ON' : 'OFF (browser will be visible)'}`);
});
