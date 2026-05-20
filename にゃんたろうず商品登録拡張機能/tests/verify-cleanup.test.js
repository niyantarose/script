const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');

function testContentJs() {
  console.log('--- Testing content/taiwan/content.js ---');
  const code = fs.readFileSync(path.join(root, 'content/taiwan/content.js'), 'utf8');

  // Extract cleanupOriginalTitleOnce using regex
  const match = code.match(/const cleanupOriginalTitleOnce = value => value[\s\S]+?\.trim\(\);/u);
  if (!match) {
    throw new Error('Could not extract cleanupOriginalTitleOnce from content.js');
  }

  // Create a executable helper function
  const functionCode = match[0];
  const evalCode = `
    const toSingleLine = text => (text || '').replace(/\\u00a0/g, ' ').replace(/\\s+/g, ' ').trim();
    ${functionCode}
    globalThis.extractOriginalTitle = rawTitle => {
      const fallback = toSingleLine(rawTitle);
      if (!fallback) return '';
      let normalized = fallback;
      for (let i = 0; i < 4; i += 1) {
        const next = cleanupOriginalTitleOnce(normalized);
        if (next === normalized) break;
        normalized = next;
      }
      return normalized || fallback;
    };
  `;

  const ctx = { console };
  ctx.globalThis = ctx;
  vm.createContext(ctx);
  vm.runInContext(evalCode, ctx);

  const extractOriginalTitle = ctx.extractOriginalTitle;
  const testCases = [
    { input: '全球高考 3 (完)', expected: '全球高考' },
    { input: '全球高考 (博客來獨家)', expected: '全球高考' },
    { input: '全球高考3(完)', expected: '全球高考' },
    { input: '全球高考(博客來獨家)', expected: '全球高考' }
  ];

  for (const tc of testCases) {
    const actual = extractOriginalTitle(tc.input);
    console.log(`Input: "${tc.input}" -> Expected: "${tc.expected}" -> Actual: "${actual}"`);
    if (actual !== tc.expected) {
      throw new Error(`Content.js mismatch: got "${actual}", expected "${tc.expected}"`);
    }
  }
  console.log('[OK] content.js verification passed.\n');
}

function testPopupSharedJs() {
  console.log('--- Testing popup/taiwan/popup.shared.js ---');
  const code = fs.readFileSync(path.join(root, 'popup/taiwan/popup.shared.js'), 'utf8');
  
  const ctx = {
    console,
    chrome: {
      storage: {
        local: {
          get: () => {},
          set: () => {}
        }
      }
    }
  };
  ctx.globalThis = ctx;
  vm.createContext(ctx);
  vm.runInContext(code, ctx, { filename: 'popup/taiwan/popup.shared.js' });
  
  const finalizeSheetOriginalWorkTitleRow = ctx.finalizeSheetOriginalWorkTitleRow;
  const testCases = [
    { input: '全球高考 3 (完)', expected: '全球高考' },
    { input: '全球高考 (博客來獨家)', expected: '全球高考' },
    { input: '全球高考3(完)', expected: '全球高考' },
    { input: '全球高考(博客來獨家)', expected: '全球高考' }
  ];

  for (const tc of testCases) {
    const actual = finalizeSheetOriginalWorkTitleRow(tc.input);
    console.log(`Input: "${tc.input}" -> Expected: "${tc.expected}" -> Actual: "${actual}"`);
    if (actual !== tc.expected) {
      throw new Error(`popup.shared.js mismatch: got "${actual}", expected "${tc.expected}"`);
    }
  }
  console.log('[OK] popup.shared.js verification passed.\n');
}

try {
  testContentJs();
  testPopupSharedJs();
  console.log('All verification tests completed successfully!');
} catch (e) {
  console.error('Test verification failed!', e);
  process.exit(1);
}
