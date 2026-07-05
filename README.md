# Teams Random Picker

Pick a random participant from a live Microsoft Teams meeting — automatically, via browser automation. No admin access or Azure app registration required.

**Built with [Playwright](https://playwright.dev/)** — headless browser automation for reliable Teams meeting interaction.

> 🚀 **New to this project?** Check out [QUICKSTART.md](QUICKSTART.md) for a quick setup guide and testing instructions!

## How it works

1. You paste a Teams meeting link into the web UI
2. **Playwright** launches a headless Chromium browser instance
3. The bot uses Playwright to automate browser actions:
   - Navigates to the Teams meeting URL
   - Enters a display name and joins silently (camera & microphone **OFF**)
   - Opens the participants panel
   - Scrapes attendee names from the DOM
4. You click **Pick** — a random name is selected and displayed

The host uses the Teams desktop app as normal. Only the Playwright-driven bot joins via the browser. **The bot joins without video or audio**, so it doesn't consume bandwidth or create any noise.

---

## Prerequisites

- **Node.js 18+** — https://nodejs.org
- **Playwright** — installed via npm (see Setup)
- The Teams meeting must be joinable via browser (standard scheduled meetings support this)

---

## Setup

```bash
# 1. Clone / copy the project folder
cd teams-picker

# 2. Install dependencies (including Playwright)
npm install

# 3. Install Playwright's Chromium browser (required for automation)
npm run install:browsers

# 4. Copy and customize config
cp .env.example .env
```

Open `.env` and adjust settings:

```env
PORT=3000
BOT_NAME=PickerBot        # Name shown in the Teams meeting
HEADLESS=false            # false = see the browser (good for debugging)
                          # true  = runs in background (production)
AUTO_REFRESH_MS=15000     # Auto-refresh participant list every 15s (0 = off)
```

---

## Run

```bash
npm start
```

Then open **http://localhost:3000** in your browser.

---

## Usage

1. Start a Teams meeting (or get the join link for one already in progress)
2. Open http://localhost:3000
3. Paste the meeting URL and click **Join & fetch participants**
4. Wait ~10–20 seconds for the bot to join and load names
5. Click **Pick a name** to randomly select a participant (manual mode)
6. **OR** Click **🎤 Auto-pick & post to chat** to automatically pick a name and post it to the Teams meeting chat
7. Click any name chip to exclude/include them from the pool
8. Toggle **Auto-exclude picked** to prevent the same person being picked twice
9. Click **Refresh** at any time to re-scrape the live attendee list
10. Click **Leave** when done — the bot exits the meeting

**Note:** Names are tracked and never repeated in the same session. Each time you pick or auto-pick, a unique participant is selected. Once all are picked, the list resets automatically.

---

## Troubleshooting

### No participants found
Teams updates their web client DOM regularly. If the Playwright bot joins but returns 0 names:

1. Open `teams.microsoft.com` in Chrome
2. Join any meeting and open the participants panel
3. Right-click a participant name → **Inspect**
4. Find the stable selector (look for `data-tid` attributes first)
5. Update `SEL.participantName` in `src/bot.js`

### Bot gets stuck on the lobby
- Try setting `HEADLESS=false` in `.env` so you can watch the Playwright browser automation in action
- Check that the meeting URL is a standard `teams.microsoft.com` join link
- Some tenant policies block anonymous guest joins — ask the organiser to allow it

### Bot is removed from the meeting
Some tenants auto-remove unrecognised guests. The Playwright context includes anti-detection flags, but if this persists, try signing the bot into a real Microsoft account by adding `storageState` to the Playwright context (see [Playwright docs](https://playwright.dev/docs/api/class-browsercontext#browser-context-storage-state)).

### Auto-pick & post to chat not working
If the **🎤 Auto-pick & post to chat** button works but the message doesn't appear in Teams:

1. Ensure the Teams meeting chat panel is visible on the right side
2. Try setting `HEADLESS=false` in `.env` to watch the browser automation
3. Open `teams.microsoft.com` in Chrome and inspect the message input element:
   - Right-click on the message input field → **Inspect**
   - Find the element and its `data-tid` or `role` attributes
   - Update `SEL.chatInput` and `SEL.sendButton` in `src/bot.js`
4. Check server logs for error messages — they will indicate which selector failed
5. Verify the bot has permission to send messages in the meeting chat

---

## API reference

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET    | `/api/health` | Health check — server status |
| POST   | `/api/join` | Join a Teams meeting and fetch initial participants `{ meetingUrl: string }` |
| GET    | `/api/participants` | Re-scrape participant list from active session `?meetingUrl=...` |
| POST   | `/api/pick` | Pick a random participant `{ meetingUrl: string, exclude?: string[] }` |
| POST   | `/api/pick-and-post` | **Auto-pick a participant and post to Teams chat** `{ meetingUrl: string }` — *no name repeats* |
| POST   | `/api/reset-picks` | Reset picked names list to allow re-picking `{ meetingUrl: string }` |
| POST   | `/api/leave` | Leave meeting and close browser `{ meetingUrl: string }` |
| GET    | `/api/sessions` | List all active bot sessions |

---

## Project structure

```
teams-picker/
├── src/
│   ├── server.js          — Express API server
│   ├── bot.js             — **Playwright automation** (browser control, meeting join, DOM scraping)
│   └── logger.js          — Winston logger
├── public/
│   ├── index.html         — Web UI
│   ├── css/
│   │   └── style.css
│   └── js/
│       └── app.js
├── debug-selectors.js     — Helper script to find correct chat selectors
├── test-api.js            — Test API endpoints without a real Teams meeting
├── .env.example
└── package.json
```

---

## Debugging & Selector Updates

### Quick Debug Helper: Find Chat Selectors Automatically

Instead of manually inspecting, use the built-in debug script:

```bash
# Start with a Teams meeting URL
node debug-selectors.js "https://teams.microsoft.com/l/meetup-join/19:meeting@thread.skype"
```

This script will:
1. ✓ Join a Teams meeting automatically
2. ✓ Guide you to inspect the chat input element
3. ✓ Test your selectors
4. ✓ Output the correct selectors to add to bot.js

### Manual: Finding Chat Input Selector (for posting messages)

If the auto-pick & post feature isn't working:

1. **Start a Teams meeting** and open it in your browser
2. **Right-click on the message input box** → **Inspect element**
3. Look for attributes like:
   - `data-tid="message-input"` or `data-tid="compose-input"`
   - `role="textbox"` with `contenteditable="true"`
   - Class names like `messageBox` or similar
4. Update these in `src/bot.js`:
   ```javascript
   const SEL = {
     // ... other selectors ...
     chatInput:   'YOUR_NEW_SELECTOR_HERE',
     sendButton:  'YOUR_NEW_SEND_BUTTON_SELECTOR',
   };
   ```
5. Restart the server and test again

### Enable Browser Visibility for Debugging

To watch what the bot is doing in real-time:

1. Set `HEADLESS=false` in `.env`
2. Restart the server
3. When you join a meeting, a Chrome window will open showing the bot's actions

### Check Server Logs

Run with `HEADLESS=false` and watch the terminal output for error messages:

```bash
npm start
```

Look for lines like:
- `✓ Chat message sent successfully` — success
- `✗ Failed to send chat message` — see the error message for which selector failed

---

## Testing with a Real Teams Meeting

### Test API Endpoints (without a real meeting)

First, verify the API endpoints are working:

```bash
node test-api.js http://localhost:3000
```

This will test:
- ✅ Health check
- ✅ Sessions endpoint
- ✅ Error handling

### Test Auto-Pick & Chat Posting Feature

To verify the auto-pick and chat posting features work:

1. **Start the server:**
   ```bash
   npm start
   ```

2. **Open the web UI:**
   - Navigate to http://localhost:3000

3. **Start or join a Teams meeting**

4. **Test in the web UI:**
   - Paste your Teams meeting URL
   - Click **"Join & fetch participants"**
   - Wait for the bot to join (10-20 seconds)
   - Once you see the participant list:
     - Click **"Pick a name"** — this picks manually without posting to chat
     - Click **"🎤 Auto-pick & post to chat"** — this picks and attempts to post to Teams chat
   - Check the Teams meeting chat to see if the message was posted

5. **If chat posting fails:**
   - Check the server terminal for error messages
   - Run `node debug-selectors.js "YOUR_MEETING_URL"` to find correct selectors
   - Update `src/bot.js` with the new selectors
   - Restart and test again

---

## Limitations

- Requires the Teams meeting to be joinable via browser (anonymous guest join must be allowed)
- Teams DOM selectors can break when Microsoft updates their web client — see Troubleshooting
- One bot session per meeting URL at a time
- Participants panel must be visible for scraping; some tenants hide it for guests
