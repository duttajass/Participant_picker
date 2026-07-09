/* Teams Picker — frontend app */

const API = '';  // same origin
const PARTICIPANT_POLL_MS = 10000;

// ── State ──────────────────────────────────────────────────────────────────
const state = {
  meetingUrl:   null,
  participants: [],
  excluded:     new Set(),
  pickedEver:   new Set(),
  history:      [],          // [{ name, time }]
  autoExclude:  false,
};
let participantPollTimer = null;
let isParticipantPollInFlight = false;

// ── DOM refs ───────────────────────────────────────────────────────────────
const $ = (id) => document.getElementById(id);
const panelJoin        = $('panelJoin');
const panelPicker      = $('panelPicker');
const btnJoin          = $('btnJoin');
const btnPick          = $('btnPick');
const btnAutoPickPost  = $('btnAutoPickPost');
const btnRefresh       = $('btnRefresh');
const btnLeave         = $('btnLeave');
const btnClearHistory  = $('btnClearHistory');
const meetingUrlInput  = $('meetingUrl');
const participantCount = $('participantCount');
const participantsGrid = $('participantsGrid');
const resultPlaceholder= $('resultPlaceholder');
const resultWinner     = $('resultWinner');
const winnerAvatar     = $('winnerAvatar');
const winnerName       = $('winnerName');
const winnerSub        = $('winnerSub');
const statusDot        = $('statusDot');
const statusText       = $('statusText');
const historyList      = $('historyList');
const chkExcludePicked = $('chkExcludePicked');

// ── Toast ──────────────────────────────────────────────────────────────────
let toastTimer = null;
function toast(msg, type = '') {
  const el = $('toast');
  el.textContent = msg;
  el.className = 'toast show' + (type ? ' toast-' + type : '');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { el.className = 'toast'; }, 3200);
}

// ── Status bar ─────────────────────────────────────────────────────────────
function setStatus(status, label) {
  statusDot.dataset.status = status;
  statusText.textContent   = label;
}

// ── Navigation ─────────────────────────────────────────────────────────────
document.querySelectorAll('.nav-item').forEach((btn) => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.nav-item').forEach((b) => b.classList.remove('active'));
    document.querySelectorAll('.view').forEach((v) => v.classList.remove('active'));
    btn.classList.add('active');
    const view = $('view-' + btn.dataset.view);
    if (view) view.classList.add('active');
  });
});

// ── Join ───────────────────────────────────────────────────────────────────
btnJoin.addEventListener('click', async () => {
  const url = meetingUrlInput.value.trim();
  if (!url) { toast('Paste a Teams meeting link first.', 'error'); return; }

  setBtnLoading(btnJoin, true, 'Joining meeting…');
  setStatus('joining', 'Joining…');

  try {
    const res = await apiFetch('/api/join', 'POST', { meetingUrl: url });
    state.meetingUrl   = url;
    state.participants = res.session.participants;
    startParticipantPolling();
    setStatus('in-meeting', 'In meeting');
    showPicker();
    renderParticipants();
    toast(`Joined! ${state.participants.length} participant(s) found.`);
  } catch (err) {
    setStatus('error', 'Error');
    toast(err.message || 'Failed to join meeting.', 'error');
  } finally {
    setBtnLoading(btnJoin, false, 'Join &amp; fetch participants');
  }
});

// ── Refresh participants ───────────────────────────────────────────────────
btnRefresh.addEventListener('click', async () => {
  if (!state.meetingUrl) return;
  setBtnLoading(btnRefresh, true, 'Refreshing…');
  participantsGrid.classList.add('loading');

  try {
    const res = await fetchParticipants();
    state.participants = res.participants;
    renderParticipants();
    toast(`Updated — ${res.count} participant(s).`);
  } catch (err) {
    toast(err.message || 'Refresh failed.', 'error');
  } finally {
    setBtnLoading(btnRefresh, false, 'Refresh');
    participantsGrid.classList.remove('loading');
  }
});

// ── Pick ───────────────────────────────────────────────────────────────────
btnPick.addEventListener('click', async () => {
  if (!state.meetingUrl) return;

  const exclude = [...state.excluded];
  setBtnLoading(btnPick, true, 'Picking…');

  try {
    const res = await apiFetch('/api/pick', 'POST', {
      meetingUrl: state.meetingUrl,
      exclude,
    });

    showWinner(res.picked, res.pool, res.total);

    // Add to history
    state.history.unshift({ name: res.picked, time: new Date() });
    state.pickedEver.add(res.picked);

    // Auto-exclude if toggled
    if (state.autoExclude) {
      state.excluded.add(res.picked);
    }

    renderParticipants();
    renderHistory();
  } catch (err) {
    toast(err.message || 'Could not pick.', 'error');
  } finally {
    setBtnLoading(btnPick, false, 'Pick a name');
  }
});

// ── Auto-pick & post to chat ────────────────────────────────────────────────
btnAutoPickPost.addEventListener('click', async () => {
  if (!state.meetingUrl) return;

  setBtnLoading(btnAutoPickPost, true, 'Posting…');

  try {
    const res = await apiFetch('/api/pick-and-post', 'POST', {
      meetingUrl: state.meetingUrl,
    });

    showWinner(res.picked, res.totalPicked, res.totalParticipants);

    // Add to history
    state.history.unshift({ name: res.picked, time: new Date() });
    state.pickedEver.add(res.picked);
    renderParticipants();
    renderHistory();

    toast(`🎤 Posted "${res.picked}" to chat (${res.totalPicked}/${res.totalParticipants} picked)`);
  } catch (err) {
    toast(err.message || 'Could not auto-pick.', 'error');
  } finally {
    setBtnLoading(btnAutoPickPost, false, '🎤 Auto-pick & post to chat');
  }
});

// ── Leave ──────────────────────────────────────────────────────────────────
btnLeave.addEventListener('click', async () => {
  if (!state.meetingUrl) return;
  if (!confirm('Leave the meeting and close the bot?')) return;

  try {
    await apiFetch('/api/leave', 'POST', { meetingUrl: state.meetingUrl });
    resetState();
    toast('Bot left the meeting.');
  } catch (err) {
    toast(err.message || 'Could not leave.', 'error');
  }
});

// ── Auto-exclude toggle ────────────────────────────────────────────────────
chkExcludePicked.addEventListener('change', () => {
  state.autoExclude = chkExcludePicked.checked;
});

// ── Clear history ──────────────────────────────────────────────────────────
btnClearHistory.addEventListener('click', () => {
  state.history = [];
  state.pickedEver.clear();
  renderHistory();
});

// ── Render participants ────────────────────────────────────────────────────
function renderParticipants() {
  participantCount.textContent =
    `${state.participants.length} participant${state.participants.length !== 1 ? 's' : ''} loaded`;

  participantsGrid.innerHTML = '';

  if (state.participants.length === 0) {
    participantsGrid.innerHTML = '<p style="color:var(--text-3);font-size:13px;">No participants found. Try refreshing.</p>';
    return;
  }

  state.participants.forEach((name) => {
    const chip = document.createElement('span');
    chip.className = 'participant-chip';
    if (state.excluded.has(name))   chip.classList.add('excluded');
    if (state.pickedEver.has(name)) chip.classList.add('picked-chip');

    chip.innerHTML = `<span class="chip-dot"></span>${escHtml(name)}`;
    chip.title = state.excluded.has(name) ? 'Click to include' : 'Click to exclude';

    chip.addEventListener('click', () => {
      if (state.excluded.has(name)) state.excluded.delete(name);
      else state.excluded.add(name);
      renderParticipants();
    });

    participantsGrid.appendChild(chip);
  });
}

// ── Render history ─────────────────────────────────────────────────────────
function renderHistory() {
  if (state.history.length === 0) {
    historyList.innerHTML = '<p class="empty-state">No picks yet this session.</p>';
    return;
  }
  historyList.innerHTML = state.history.map(({ name, time }) => `
    <div class="history-item">
      <span class="history-name">${escHtml(name)}</span>
      <span class="history-time">${formatTime(time)}</span>
    </div>
  `).join('');
}

// ── Show winner ────────────────────────────────────────────────────────────
function showWinner(name, pool, total) {
  resultPlaceholder.classList.add('hidden');
  resultWinner.classList.remove('hidden');
  resultWinner.classList.remove('pop');
  void resultWinner.offsetWidth; // force reflow
  resultWinner.classList.add('pop');

  winnerAvatar.textContent = initials(name);
  winnerName.textContent   = name;
  winnerSub.textContent    = `Picked from ${pool} eligible of ${total} total`;
}

// ── UI state helpers ───────────────────────────────────────────────────────
function showPicker() {
  panelJoin.classList.add('hidden');
  panelPicker.classList.remove('hidden');
}

function resetState() {
  stopParticipantPolling();
  state.meetingUrl   = null;
  state.participants = [];
  state.excluded.clear();
  state.pickedEver.clear();
  panelJoin.classList.remove('hidden');
  panelPicker.classList.add('hidden');
  setStatus('idle', 'Not connected');
  resultWinner.classList.add('hidden');
  resultPlaceholder.classList.remove('hidden');
  participantsGrid.innerHTML = '';
}

function startParticipantPolling() {
  stopParticipantPolling();
  participantPollTimer = setInterval(async () => {
    if (!state.meetingUrl || isParticipantPollInFlight) {
      return;
    }

    isParticipantPollInFlight = true;
    try {
      const res = await fetchParticipants();
      if (res.count === 0 && state.participants.length > 0) {
        return;
      }
      const prevSerialized = JSON.stringify(state.participants);
      const nextSerialized = JSON.stringify(res.participants);
      if (prevSerialized !== nextSerialized) {
        state.participants = res.participants;
        renderParticipants();
      }
    } catch {
      // Avoid noisy UI errors during background refresh loops.
    } finally {
      isParticipantPollInFlight = false;
    }
  }, PARTICIPANT_POLL_MS);
}

function stopParticipantPolling() {
  if (participantPollTimer) {
    clearInterval(participantPollTimer);
    participantPollTimer = null;
  }
}

async function fetchParticipants() {
  return await apiFetch(`/api/participants?meetingUrl=${encodeURIComponent(state.meetingUrl)}`, 'GET');
}

function setBtnLoading(btn, loading, label) {
  btn.disabled = loading;
  const labelEl   = btn.querySelector('.btn-label');
  const spinnerEl = btn.querySelector('.btn-spinner');
  if (labelEl)   labelEl.innerHTML = label;
  if (spinnerEl) spinnerEl.hidden  = !loading;
  if (!spinnerEl && !loading) btn.innerHTML = label;
}

// ── Utilities ──────────────────────────────────────────────────────────────
function initials(name) {
  return name
    .split(/\s+/)
    .slice(0, 2)
    .map((w) => w[0]?.toUpperCase() || '')
    .join('');
}

function escHtml(str) {
  return str.replace(/[&<>"']/g, (c) =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])
  );
}

function formatTime(date) {
  return date.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
}

async function apiFetch(url, method = 'GET', body = null) {
  const opts = {
    method,
    headers: { 'Content-Type': 'application/json' },
  };
  if (body) opts.body = JSON.stringify(body);

  const res = await fetch(API + url, opts);
  const data = await res.json();

  if (!res.ok) throw new Error(data.error || `HTTP ${res.status}`);
  return data;
}
