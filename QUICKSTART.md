# Teams Picker - Quick Start Guide

## ✅ What's Been Implemented

### Core Features
- ✅ **Playwright Browser Automation** - Join Teams meetings automatically
- ✅ **Participant Scraping** - Extract attendee names from meeting roster
- ✅ **Manual Picking** - Click to pick a random participant
- ✅ **Auto-Pick & Post to Chat** - Pick participants and post to Teams chat automatically
- ✅ **No-Repeat Logic** - Names never repeat in one session
- ✅ **Web UI** - Beautiful, responsive interface at http://localhost:3000
- ✅ **REST API** - Full-featured API for automation

### Helper Tools
- ✅ **test-api.js** - Test API endpoints without a real meeting
- ✅ **debug-selectors.js** - Find correct Teams chat selectors
- ✅ **Detailed Logging** - Clear error messages for debugging

---

## 🚀 Quick Start

### 1. Install & Setup
```bash
npm install
npm run install:browsers
cp .env.example .env
```

### 2. Start the Server
```bash
npm start
```

### 3. Access the UI
Open your browser: **http://localhost:3000**

---

## 🧪 Testing

### Option A: Test API Without a Meeting
```bash
node test-api.js http://localhost:3000
```
✅ This tests all endpoints - use this first to verify setup!

### Option B: Test with a Real Teams Meeting

1. **Start a Teams meeting** and get the join link
2. Go to http://localhost:3000
3. Paste the meeting URL
4. Click **"Join & fetch participants"**
5. Wait 10-20 seconds for bot to join
6. Test features:
   - Click **"Pick a name"** - manual random pick
   - Click **"🎤 Auto-pick & post to chat"** - auto-picks and posts to Teams chat

---

## 🔍 Troubleshooting Chat Posting

### If "Auto-pick & post to chat" doesn't post to Teams chat:

**Step 1: Check Server Logs**
- Look at the terminal where `npm start` is running
- Look for error messages like `✗ Failed to send chat message`

**Step 2: Find Correct Selectors**
```bash
node debug-selectors.js "https://teams.microsoft.com/l/meetup-join/YOUR_MEETING_URL"
```

This interactive script will:
- Join the meeting
- Guide you to inspect chat elements
- Test your selectors
- Output the correct CSS selectors

**Step 3: Update selectors in src/bot.js**
```javascript
const SEL = {
  // ... other selectors ...
  chatInput:  'YOUR_NEW_SELECTOR',
  sendButton: 'YOUR_NEW_SELECTOR',
};
```

**Step 4: Restart & Test**
```bash
npm start
# Then test again
```

---

## 📚 API Endpoints

| Endpoint | Method | Purpose |
|----------|--------|---------|
| `/api/health` | GET | Check server status |
| `/api/join` | POST | Join a Teams meeting |
| `/api/participants` | GET | Get participant list |
| `/api/pick` | POST | Pick random (manual) |
| `/api/pick-and-post` | POST | **Auto-pick & post to chat** |
| `/api/reset-picks` | POST | Reset picked names |
| `/api/leave` | POST | Leave meeting |
| `/api/sessions` | GET | List active sessions |

---

## 💡 Environment Variables (.env)

```env
PORT=3000                    # Server port
BOT_NAME=PickerBot          # Name shown in Teams
HEADLESS=false              # false = see browser, true = background
AUTO_REFRESH_MS=15000       # Refresh participants every 15s (0 = off)
JOIN_TIMEOUT_MS=30000       # Timeout for joining meeting
EXCLUDED_PARTICIPANT_TERMS=recording,transcription  # Optional extra roster entries to ignore
```

---

## 📂 Project Files

- `src/server.js` - Express API server
- `src/bot.js` - Playwright automation logic
- `public/index.html` - Web UI
- `debug-selectors.js` - ⭐ Use this to find chat selectors
- `test-api.js` - ⭐ Use this to test API endpoints
- `.env` - Configuration file

---

## 🎯 Feature Status

### ✅ Implemented & Tested
- Bot joins Teams meetings via web browser
- Scrapes participant list
- Manual picking (no repeats)
- Web UI with real-time updates
- API endpoints
- Chat input/send button interaction (selectors may need updating)

### ⚠️ Requires Selector Updates
- **Chat posting to Teams** - The DOM selectors might need updating for your Teams version
  - Use `debug-selectors.js` to find the correct selectors
  - Update `src/bot.js` with new selectors
  - Chat posting should then work

### 📋 Known Limitations
- Requires browser-joinable Teams meetings
- Chat selectors may change when Teams updates their UI
- One bot session per meeting URL at a time
- Participants panel must be visible

---

## 🎓 Next Steps

1. **Test API first**: `node test-api.js http://localhost:3000`
2. **Try a real meeting**: Paste a Teams URL into the web UI
3. **If chat posting fails**: Run `node debug-selectors.js "<MEETING_URL>"`
4. **Update selectors**: Copy output from debug script to `src/bot.js`
5. **Test again**: Restart server and verify chat posting works

---

## 💬 Support

If you encounter issues:

1. Check server logs (`npm start` terminal)
2. Run `node test-api.js` to verify basic functionality
3. Run `node debug-selectors.js` to find correct chat selectors
4. Verify Teams meeting is browser-joinable
5. Check .env configuration

Good luck! 🚀
