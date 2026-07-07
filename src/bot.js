require('dotenv').config();
const { chromium } = require('playwright');
const logger = require('./logger');

const JOIN_TIMEOUT = parseInt(process.env.JOIN_TIMEOUT_MS || '30000', 10);
const HEADLESS = process.env.HEADLESS === 'true';
const BOT_NAME = process.env.BOT_NAME || 'PickerBot';

/**
 * DOM selectors for the Teams web client.
 *
 * Teams updates their frontend regularly. If the bot stops working,
 * open teams.microsoft.com in Chrome DevTools and update these selectors
 * by inspecting the relevant elements in the live DOM.
 *
 * Priority order: data-tid attributes are most stable (Teams' own test IDs),
 * followed by aria-label, then class-based selectors.
 */
const SEL = {
  // Pre-join / lobby screen
  continueInBrowser: 'button[data-tid="joinOnWeb"], a[data-tid="joinOnWeb"], button:has-text("Continue on this browser"), a:has-text("Continue on this browser")',
  nameInput:         'input[data-tid="prejoin-display-name-input"], input[placeholder*="name" i], input[aria-label*="name" i]',
  cameraToggle:      'button[data-tid="toggle-av"], button[aria-label*="camera" i], button[title*="camera" i]',
  micToggle:         'button[data-tid="toggle-mute-button"], button[aria-label*="microphone" i], button[title*="mic" i]',
  joinBtn:           'button[data-tid="prejoin-join-button"], button:has-text("Join now"), button[aria-label*="Join now"], button:has-text("Join meeting"), button[type="button"][class*="primary"]',

  // In-meeting controls
  participantsBtn:   '[data-tid="calling-roster-button"], button[aria-label*="People" i], button[title*="People" i]',

  // Participant roster panel
  rosterPanel:       '[role="tree"][aria-label*="Attendees"], [data-fui-tree-item-value*="roster"], [class*="fui-FlatTree"]',
  participantName:   'span[id^="roster-avatar-img-"], [class*="fui-TreeItemPersonaLayout__main"] span[title], span[title][dir="auto"]',

  // Meeting loaded indicator - updated selectors for current Teams UI
  meetingToolbar:    '[data-tid="calling-component-container"], [class*="calling-unified-bar"], [role="region"][class*="call"], [data-testid*="call"], [class*="meetingToolbar"], button[data-tid="calling-roster-button"]',

  // Chat controls (for posting participant names)
  chatBtn:           'button[data-tid="chat-button"], button[aria-label*="Chat" i], button[title*="Chat" i], [data-tid="callingButtons-showChatButton"]',
  chatInput:         '[data-tid="ckeditor"], [data-tid="message-input"], [data-tid="compose-input"], div[role="textbox"][contenteditable="true"][placeholder*="message" i]',
  sendButton:        '[data-tid="send-message-button"], button[aria-label*="Send" i], button[title*="Send" i], [data-tid="send-button"], button[aria-label="Send"]',
  chatPanel:         '[data-tid="message-pane"], [class*="messagePane"], [class*="chatPanel"]',
};

/**
 * SessionManager wraps a single Playwright browser/page for one meeting.
 */
class SessionManager {
  constructor(browser, page, meetingUrl) {
    this.browser    = browser;
    this.page       = page;
    this.meetingUrl = meetingUrl;
    this.participants = [];
    this.pickedNames = new Set(); // Track picked participants to avoid repeats
    this.status     = 'idle'; // idle | joining | in-meeting | error | left
    this.error      = null;
  }

  /**
   * Scrapes the currently visible participant names from the roster panel.
   * Returns a deduplicated, sorted array of display name strings.
   */
  async scrapeParticipants() {
    try {
      // Ensure panel is open
      const panelVisible = await this.page.locator(SEL.rosterPanel).isVisible().catch(() => false);
      if (!panelVisible) {
        logger.info('Roster panel not visible — reopening...');
        await this.page.click(SEL.participantsBtn, { timeout: 5000 });
        await this.page.waitForTimeout(1500);
      }

      // Try to find and wait for participant elements
      try {
        await this.page.waitForSelector(SEL.participantName, { timeout: 8000 });
      } catch {
        logger.warn('Primary participant selector timed out — trying fallback methods...');
        await this.page.waitForTimeout(2000);
      }

      // Extract names using $$eval
      let names = await this.page.$$eval(SEL.participantName, (els) =>
        els
          .map((el) => el.textContent?.trim())
          .filter((n) => n && n.length > 0 && n !== 'You')
      ).catch(() => []);

      // If no names found, try alternative extraction
      if (names.length === 0) {
        logger.info('No names found with selector — trying alternative extraction...');
        // Try getting all text content from the roster panel
        names = await this.page.$$eval('[role="listitem"]', (els) =>
          els
            .map((el) => el.textContent?.trim())
            .filter((n) => n && n.length > 2 && n !== 'You')
        ).catch(() => []);
      }

      // Deduplicate and sort
      this.participants = [...new Set(names)].sort((a, b) => a.localeCompare(b));
      logger.info(`Scraped ${this.participants.length} participant(s)`);
      
      if (this.participants.length === 0) {
        logger.warn('⚠️  No participants found. Possible reasons:');
        logger.warn('   - No other participants in the meeting yet');
        logger.warn('   - Participants panel is empty or hidden');
        logger.warn('   - DOM selectors need updating for current Teams version');
      }
      
      return this.participants;
    } catch (err) {
      logger.warn(`Scrape failed: ${err.message}`);
      logger.warn('Tip: update SEL.participantName in bot.js using DevTools on teams.microsoft.com');
      return this.participants; // return last known list
    }
  }

  /**
   * Picks a random participant from the available pool.
   * Excludes already-picked names to prevent repeats.
   * 
   * @returns {string|null} The picked name, or null if all have been picked or no participants available
   */
  pickRandomParticipant() {
    const available = this.participants.filter(name => !this.pickedNames.has(name));
    
    if (available.length === 0) {
      logger.warn('All participants already picked. Resetting the picked list.');
      this.pickedNames.clear(); // Reset if all picked
      const resetAvailable = this.participants.filter(name => !this.pickedNames.has(name));
      if (resetAvailable.length === 0) return null;
      return resetAvailable[Math.floor(Math.random() * resetAvailable.length)];
    }
    
    const picked = available[Math.floor(Math.random() * available.length)];
    this.pickedNames.add(picked);
    logger.info(`Picked: "${picked}" (${this.pickedNames.size}/${this.participants.length} picked)`);
    return picked;
  }

  /**
   * Sends a message to the Teams meeting chat.
   * 
   * @param {string} message The message to send
   * @returns {Promise<boolean>} True if message sent successfully
   */
  async sendChatMessage(message) {
    try {
      logger.info(`Attempting to send chat message: "${message}"`);

      // Open the chat panel first (it is hidden until the chat button is clicked)
      const chatBtnLocator = this.page.locator(SEL.chatBtn).first();
      const chatBtnVisible = await chatBtnLocator.isVisible({ timeout: 3000 }).catch(() => false);
      if (chatBtnVisible) {
        logger.info('Opening chat panel...');
        await chatBtnLocator.click();
        await this.page.waitForTimeout(1000);
      } else {
        logger.warn('Chat button not found — panel may already be open');
      }

      // Wait for chat input to become visible
      const chatInputLocator = this.page.locator(SEL.chatInput).first();
      try {
        await chatInputLocator.waitFor({ state: 'visible', timeout: 8000 });
      } catch {
        logger.warn('Chat input still not visible after opening panel — proceeding anyway...');
      }

      // Click and focus the input
      await chatInputLocator.click({ timeout: 5000 });
      await this.page.waitForTimeout(300);

      // Type the message
      await chatInputLocator.fill(message);
      logger.info(`Message typed: "${message}"`);
      await this.page.waitForTimeout(200);

      // Send the message
      const sendBtnLocator = this.page.locator(SEL.sendButton).first();
      const sendBtnVisible = await sendBtnLocator.isVisible({ timeout: 2000 }).catch(() => false);
      
      if (!sendBtnVisible) {
        logger.warn('Send button not visible — trying Enter key...');
        await chatInputLocator.press('Enter');
      } else {
        await sendBtnLocator.click({ timeout: 5000 });
      }
      
      await this.page.waitForTimeout(500);
      logger.info('✓ Chat message sent successfully');
      return true;
    } catch (err) {
      logger.error(`✗ Failed to send chat message: ${err.message}`);
      logger.error('Tip: Update chat selectors (SEL.chatInput, SEL.sendButton) in bot.js using DevTools');
      return false;
    }
  }

  /**
   * Automatically picks a random participant and posts their name to the meeting chat.
   * Ensures no name repeats.
   * 
   * @returns {Promise<object>} Result with picked name and status
   */
  async pickAndPostToChat() {
    try {
      // Refresh participant list
      await this.scrapeParticipants();

      // Pick a random participant (no repeats)
      const pickedName = this.pickRandomParticipant();
      
      if (!pickedName) {
        logger.warn('No participants available to pick');
        return { success: false, error: 'No participants available' };
      }

      // Post to chat
      const chatMessage = `🎤 Next speaker: ${pickedName}`;
      const messageSent = await this.sendChatMessage(chatMessage);

      return {
        success: messageSent,
        picked: pickedName,
        totalPicked: this.pickedNames.size,
        totalParticipants: this.participants.length,
      };
    } catch (err) {
      logger.error(`pickAndPostToChat failed: ${err.message}`);
      return { success: false, error: err.message };
    }
  }

  /**
   * Reset the picked names list to allow re-picking.
   */
  resetPickedNames() {
    this.pickedNames.clear();
    logger.info('Picked names reset — all participants available again');
  }

  async leave() {
    this.status = 'left';
    await this.browser.close().catch(() => {});
    logger.info('Browser closed — session ended.');
  }
}

/**
 * Launches Chromium, navigates to the Teams web meeting,
 * completes the pre-join lobby, and returns a SessionManager.
 *
 * @param {string} meetingUrl
 * @returns {Promise<SessionManager>}
 */
async function joinMeeting(meetingUrl) {
  logger.info(`Launching browser (headless=${HEADLESS})...`);

  const browser = await chromium.launch({
    headless: HEADLESS,
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-dev-shm-usage',
      '--use-fake-ui-for-media-stream',    // auto-grant camera/mic in browser UI
      '--use-fake-device-for-media-stream', // use virtual devices — no real camera/mic needed
      '--mute-audio',                       // prevent any audio output through machine speakers
      '--disable-blink-features=AutomationControlled', // reduces bot detection
    ],
  });

  const context = await browser.newContext({
    permissions: ['camera', 'microphone'],
    userAgent:
      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 ' +
      '(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    viewport: { width: 1280, height: 800 },
  });

  const page = await context.newPage();
  const session = new SessionManager(browser, page, meetingUrl);
  session.status = 'joining';

  // ── Step 1: Navigate ──────────────────────────────────────────────────────
  logger.info('Navigating to meeting URL...');
  await page.goto(meetingUrl, { waitUntil: 'domcontentloaded', timeout: JOIN_TIMEOUT });
  await page.waitForTimeout(2000);

  // ── Step 2: "Continue in browser" prompt ──────────────────────────────────
  logger.info('Looking for "Continue on this browser" button...');
  try {
    const continueBtn = await page.waitForSelector(SEL.continueInBrowser, { timeout: 5000 });
    if (continueBtn) {
      logger.info('✓ Found "Continue on this browser" button, clicking...');
      await continueBtn.click();
      logger.info('✓ Clicked "Continue on this browser"');
      // Wait for page navigation and new content to load
      await page.waitForTimeout(5000);
      logger.info('✓ Waiting for page to transition...');
    }
  } catch (err) {
    logger.info('"Continue in browser" button not found — may already be on web client');
  }

  // Check if we're now on the pre-join screen or in an iframe
  await page.waitForTimeout(2000);

  // ── Step 3: Enter display name ────────────────────────────────────────────
  try {
    await page.waitForSelector(SEL.nameInput, { timeout: 8000 });
    await page.fill(SEL.nameInput, BOT_NAME);
    logger.info(`Display name set to "${BOT_NAME}"`);
  } catch {
    logger.info('Name input not found — user may already be signed in');
  }

  // ── Step 4: Mute camera & mic before joining ──────────────────────────────
  logger.info('Disabling camera and microphone...');
  let cameraDisabled = false;
  let micDisabled = false;

  // Try to disable camera
  try {
    const cameraBtnLocator = page.locator(SEL.cameraToggle).first();
    if (await cameraBtnLocator.isVisible({ timeout: 2000 }).catch(() => false)) {
      const ariaPressed = await cameraBtnLocator.getAttribute('aria-pressed').catch(() => null);
      // If aria-pressed is "true", the camera is ON — click to turn it OFF
      if (ariaPressed === 'true' || ariaPressed === null) {
        await cameraBtnLocator.click();
        await page.waitForTimeout(300);
        logger.info('✓ Camera disabled');
        cameraDisabled = true;
      } else {
        logger.info('✓ Camera already disabled');
        cameraDisabled = true;
      }
    }
  } catch (err) {
    logger.warn(`Could not disable camera: ${err.message}`);
  }

  // Try to disable microphone
  try {
    const micBtnLocator = page.locator(SEL.micToggle).first();
    if (await micBtnLocator.isVisible({ timeout: 2000 }).catch(() => false)) {
      const ariaPressed = await micBtnLocator.getAttribute('aria-pressed').catch(() => null);
      // If aria-pressed is "true", the mic is ON — click to turn it OFF
      if (ariaPressed === 'true' || ariaPressed === null) {
        await micBtnLocator.click();
        await page.waitForTimeout(300);
        logger.info('✓ Microphone disabled');
        micDisabled = true;
      } else {
        logger.info('✓ Microphone already disabled');
        micDisabled = true;
      }
    }
  } catch (err) {
    logger.warn(`Could not disable microphone: ${err.message}`);
  }

  if (!cameraDisabled || !micDisabled) {
    logger.warn('⚠️  Could not confirm camera/mic are disabled — bot may join with media enabled');
  }

  // ── Step 5: Join the meeting ───────────────────────────────────────────────
  logger.info('Looking for join button...');
  let joinButtonFound = false;
  
  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      logger.info(`Join button search (attempt ${attempt}/3)...`);
      const joinBtn = await page.waitForSelector(SEL.joinBtn, { timeout: 8000 });
      if (joinBtn) {
        logger.info('✓ Found join button, clicking...');
        await joinBtn.click();
        logger.info('✓ Clicked join button');
        joinButtonFound = true;
        break;
      }
    } catch (err) {
      logger.warn(`Join button not found on attempt ${attempt}: ${err.message}`);
      if (attempt < 3) {
        logger.info('Waiting before retry...');
        await page.waitForTimeout(3000);
      }
    }
  }

  if (!joinButtonFound) {
    logger.error('Could not find join button after 3 attempts');
    throw new Error('Failed to locate join button. The meeting page may have changed.');
  }

  // ── Step 6: Wait for meeting toolbar (confirms we are in the meeting) ──────
  logger.info('Waiting for meeting to load...');
  try {
    await page.waitForSelector(SEL.meetingToolbar, { timeout: JOIN_TIMEOUT });
    logger.info('In meeting ✓');
  } catch (err) {
    logger.warn('Meeting toolbar selector timeout — trying alternative detection...');
    // Try to detect if we're in a meeting by looking for common elements
    try {
      await page.waitForSelector(SEL.participantsBtn, { timeout: 10000 });
      logger.info('✓ Found participants button — likely in meeting');
    } catch {
      logger.error('Could not detect meeting entry. Teams UI may have changed.');
      throw new Error('Failed to load meeting. Check if the meeting URL is valid and you have permission to join.');
    }
  }
  session.status = 'in-meeting';

  // ── Step 7: Open participants panel and scrape ────────────────────────────
  await page.waitForSelector(SEL.participantsBtn, { timeout: 10000 });
  await page.click(SEL.participantsBtn);
  await page.waitForTimeout(1500);

  await session.scrapeParticipants();

  return session;
}

module.exports = { joinMeeting, SessionManager };
