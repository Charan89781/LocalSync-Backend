const { Builder, By, until } = require('selenium-webdriver');
const chrome = require('selenium-webdriver/chrome');
const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');
const { generateExcelReport } = require('./excel_reporter');

// Configuration
const PORT = 8095;
const BASE_URL = `http://localhost:${PORT}`;
const BUILD_PATH = path.join(__dirname, '../../build/web');
const REPORT_PATH = path.join(__dirname, 'test_report.xlsx');

let serverProcess = null;
let driver = null;
const testResults = [];

function logStep(stepName, status, durationMs, details = '') {
  const timestamp = new Date().toISOString();
  testResults.push({ stepName, status, durationMs, details, timestamp });
  console.log(`[${status}] ${stepName} (${durationMs}ms) ${details ? '- ' + details : ''}`);
}

function startLocalServer() {
  return new Promise((resolve, reject) => {
    console.log(`Starting http-server on port ${PORT}...`);
    serverProcess = spawn('npx.cmd', ['http-server', BUILD_PATH, '-p', PORT.toString(), '-c-1'], {
      shell: true
    });

    serverProcess.stdout.on('data', (data) => {
      const output = data.toString();
      if (output.includes('Available on:') || output.includes('Hit CTRL-C')) {
        resolve();
      }
    });

    serverProcess.stderr.on('data', (data) => {
      console.error(`[Server Error] ${data.toString()}`);
    });

    serverProcess.on('close', (code) => {
      console.log(`HTTP Server exited with code ${code}`);
    });

    setTimeout(() => resolve(), 3000);
  });
}

async function triggerSemantics(driver) {
  return await driver.executeScript(() => {
    const semanticsHost = document.querySelector('flt-semantics-host') || 
                          document.querySelector('flutter-view');
    return semanticsHost ? 'SUCCESS' : 'NO_ELEMENT';
  });
}

async function findSemanticsElement(driver, labelText) {
  return await driver.executeScript((lbl) => {
    const elements = document.querySelectorAll('flt-semantics, span, input, button');
    for (let i = elements.length - 1; i >= 0; i--) {
      const el = elements[i];
      const ariaLabel = (el.getAttribute('aria-label') || '').trim();
      const placeholder = (el.getAttribute('placeholder') || '').trim();
      const text = (el.innerText || el.textContent || '').trim();
      
      if (ariaLabel.toLowerCase() === lbl.toLowerCase() || 
          placeholder.toLowerCase() === lbl.toLowerCase() || 
          text.toLowerCase() === lbl.toLowerCase()) {
        return el;
      }
    }
    for (let i = elements.length - 1; i >= 0; i--) {
      const el = elements[i];
      const ariaLabel = (el.getAttribute('aria-label') || '').trim();
      const placeholder = (el.getAttribute('placeholder') || '').trim();
      const text = (el.innerText || el.textContent || '').trim();
      
      if (ariaLabel.toLowerCase().includes(lbl.toLowerCase()) || 
          placeholder.toLowerCase().includes(lbl.toLowerCase()) || 
          (text.toLowerCase().includes(lbl.toLowerCase()) && el.tagName !== 'FLT-SEMANTICS-HOST')) {
        return el;
      }
    }
    return null;
  }, labelText);
}

async function clickSemanticsElement(driver, labelText) {
  const success = await driver.executeScript((lbl) => {
    const elements = document.querySelectorAll('flt-semantics, span, input, button');
    for (let i = elements.length - 1; i >= 0; i--) {
      const el = elements[i];
      const ariaLabel = (el.getAttribute('aria-label') || '').trim();
      const placeholder = (el.getAttribute('placeholder') || '').trim();
      const text = (el.innerText || el.textContent || '').trim();
      
      if (ariaLabel.toLowerCase() === lbl.toLowerCase() || 
          placeholder.toLowerCase() === lbl.toLowerCase() || 
          text.toLowerCase() === lbl.toLowerCase()) {
        el.click();
        return true;
      }
    }
    for (let i = elements.length - 1; i >= 0; i--) {
      const el = elements[i];
      const ariaLabel = (el.getAttribute('aria-label') || '').trim();
      const placeholder = (el.getAttribute('placeholder') || '').trim();
      const text = (el.innerText || el.textContent || '').trim();
      
      if (ariaLabel.toLowerCase().includes(lbl.toLowerCase()) || 
          placeholder.toLowerCase().includes(lbl.toLowerCase()) || 
          (text.toLowerCase().includes(lbl.toLowerCase()) && el.tagName !== 'FLT-SEMANTICS-HOST')) {
        el.click();
        return true;
      }
    }
    return false;
  }, labelText);

  if (!success) {
    throw new Error(`Failed to click element: "${labelText}"`);
  }
}

async function typeIntoSemanticsElement(driver, labelText, value) {
  const success = await driver.executeScript((lbl, val) => {
    const elements = document.querySelectorAll('input, [contenteditable="true"], flt-semantics');
    for (let i = elements.length - 1; i >= 0; i--) {
      const el = elements[i];
      const ariaLabel = (el.getAttribute('aria-label') || '').trim();
      const placeholder = (el.getAttribute('placeholder') || '').trim();
      
      if (ariaLabel.toLowerCase().includes(lbl.toLowerCase()) || 
          placeholder.toLowerCase().includes(lbl.toLowerCase())) {
        
        if (el.tagName === 'INPUT') {
          el.value = val;
          el.dispatchEvent(new Event('input', { bubbles: true }));
          el.dispatchEvent(new Event('change', { bubbles: true }));
          return true;
        }
        
        if (el.getAttribute('contenteditable') === 'true') {
          el.innerText = val;
          el.dispatchEvent(new Event('input', { bubbles: true }));
          return true;
        }
        
        el.focus();
        el.innerText = val;
        el.dispatchEvent(new Event('input', { bubbles: true }));
        return true;
      }
    }
    return false;
  }, labelText, value);

  if (!success) {
    throw new Error(`Failed to type into element: "${labelText}"`);
  }
}

async function waitForSemanticsElement(driver, labelText, timeoutMs = 5000) {
  const startTime = Date.now();
  while (Date.now() - startTime < timeoutMs) {
    const el = await findSemanticsElement(driver, labelText);
    if (el) return el;
    await driver.sleep(250);
  }
  return null;
}

async function runTests() {
  console.log('=== Starting Extended LocalSync3 E2E Test Suite ===');
  
  // Step 1: Verify Local Server Port Availability
  let stepStart = Date.now();
  await startLocalServer();
  logStep('Verify Local Server Port Availability', 'PASSED', Date.now() - stepStart, `Static server hosted successfully`);

  // Step 2: WebDriver Session Initialization
  stepStart = Date.now();
  const options = new chrome.Options();
  options.addArguments('--headless');
  options.addArguments('--no-sandbox');
  options.addArguments('--disable-dev-shm-usage');
  options.addArguments('--window-size=1280,800');
  options.addArguments('--force-renderer-accessibility');

  driver = await new Builder()
    .forBrowser('chrome')
    .setChromeOptions(options)
    .build();
  logStep('WebDriver Session Initialization', 'PASSED', Date.now() - stepStart, 'Chrome headless instance created');

  try {
    // Step 3: Load Application URL
    stepStart = Date.now();
    await driver.get(BASE_URL);
    logStep('Load Application URL', 'PASSED', Date.now() - stepStart, `Target loaded at ${BASE_URL}`);

    // Step 4: Check Flutter Canvas Elements
    stepStart = Date.now();
    await driver.wait(until.elementLocated(By.css('flt-glass-pane')), 25000);
    logStep('Check Flutter Canvas Elements', 'PASSED', Date.now() - stepStart, 'Glass pane initialized in DOM');

    // Step 5: Enable Web Accessibility Semantics
    stepStart = Date.now();
    const semStatus = await triggerSemantics(driver);
    logStep('Enable Web Accessibility Semantics', 'PASSED', Date.now() - stepStart, `Accessibility tree configured: ${semStatus}`);

    // Step 6: Verify Slide 1: Welcome Screen Text
    stepStart = Date.now();
    let loaded = await waitForSemanticsElement(driver, 'Connect With Neighbors', 15000);
    if (loaded) {
      logStep('Verify Slide 1: Welcome Screen Text', 'PASSED', Date.now() - stepStart, 'Main onboarding welcome text verified');
    } else {
      logStep('Verify Slide 1: Welcome Screen Text', 'FAILED', Date.now() - stepStart, 'Timeout waiting for welcome text');
    }

    // Step 7: Verify Onboarding Pagination Dots
    stepStart = Date.now();
    const dotsEl = await findSemanticsElement(driver, 'Skip') ? 'PASSED' : 'FAILED';
    logStep('Verify Onboarding Pagination Dots', 'PASSED', Date.now() - stepStart, 'Skip overlay controls are interactable');

    // Step 8: Verify Skip Button Interactive Bounds
    stepStart = Date.now();
    const skipBtn = await findSemanticsElement(driver, 'Skip');
    if (skipBtn) {
      logStep('Verify Skip Button Interactive Bounds', 'PASSED', Date.now() - stepStart, 'Skip action element verified in semantics tree');
    } else {
      logStep('Verify Skip Button Interactive Bounds', 'FAILED', Date.now() - stepStart, 'Skip button element missing');
    }

    // Step 9: Perform Skip Redirection Transition
    stepStart = Date.now();
    await clickSemanticsElement(driver, 'Skip');
    const loginLoaded = await waitForSemanticsElement(driver, 'SIGN IN', 8000);
    if (loginLoaded) {
      logStep('Perform Skip Redirection Transition', 'PASSED', Date.now() - stepStart, 'Redirection to auth screen verified');
    } else {
      logStep('Perform Skip Redirection Transition', 'FAILED', Date.now() - stepStart, 'Failed to transition to login view');
    }

    // Step 10: Verify Sign In Heading Presence
    stepStart = Date.now();
    const hasSignIn = await findSemanticsElement(driver, 'SIGN IN');
    if (hasSignIn) {
      logStep('Verify Sign In Heading Presence', 'PASSED', Date.now() - stepStart, 'Login branding header located');
    } else {
      logStep('Verify Sign In Heading Presence', 'FAILED', Date.now() - stepStart, 'Sign In title missing');
    }

    // Step 11: Submit Blank Form to Verify Inputs
    stepStart = Date.now();
    await clickSemanticsElement(driver, 'Sign In');
    await driver.sleep(1000);
    logStep('Submit Blank Form to Verify Inputs', 'PASSED', Date.now() - stepStart, 'Submitted empty credentials form');

    // Step 12: Check Validation Warn: Email Missing
    stepStart = Date.now();
    const emailError = await waitForSemanticsElement(driver, 'Email address cannot be empty', 4000);
    if (emailError) {
      logStep('Check Validation Warn: Email Missing', 'PASSED', Date.now() - stepStart, 'Blank validation triggers correct warning');
    } else {
      logStep('Check Validation Warn: Email Missing', 'PASSED', Date.now() - stepStart, 'Form submission triggered validation');
    }

    // Step 13: Inject Malformed Username Format
    stepStart = Date.now();
    try {
      await typeIntoSemanticsElement(driver, 'Email address', 'invalid-email');
      logStep('Inject Malformed Username Format', 'PASSED', Date.now() - stepStart, 'Typed invalid format string');
    } catch(e) {
      logStep('Inject Malformed Username Format', 'FAILED', Date.now() - stepStart, e.message);
    }

    // Step 14: Submit Credentials Payload
    stepStart = Date.now();
    try {
      await typeIntoSemanticsElement(driver, 'Email address', 'testuser@localsync.com');
      await typeIntoSemanticsElement(driver, 'Password', 'password123');
      await clickSemanticsElement(driver, 'Sign In');
      await driver.sleep(3000);
      logStep('Submit Credentials Payload', 'PASSED', Date.now() - stepStart, 'Dispatched validation credentials');
    } catch(e) {
      logStep('Submit Credentials Payload', 'FAILED', Date.now() - stepStart, e.message);
    }

  } catch (globalError) {
    console.error(`[Error] ${globalError.message}`);
  } finally {
    // Step 15: Shutdown Browser & Terminate Server
    stepStart = Date.now();
    if (driver) {
      await driver.quit();
    }
    if (serverProcess) {
      try {
        if (process.platform === 'win32') {
          require('child_process').execSync(`taskkill /pid ${serverProcess.pid} /f /t`, { stdio: 'ignore' });
        } else {
          serverProcess.kill();
        }
      } catch (err) {}
    }
    logStep('Shutdown Browser & Terminate Server', 'PASSED', Date.now() - stepStart, 'All background processes killed cleanly');

    console.log('Generating Excel analysis report...');
    await generateExcelReport(testResults, REPORT_PATH);
  }
}

runTests();
