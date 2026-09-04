const { test } = require('node:test');
const assert = require('node:assert/strict');
const { STRINGS } = require('../src/renderer/i18n');

test('every key in the English table exists in the Chinese table and vice versa', () => {
  const enKeys = Object.keys(STRINGS.en).sort();
  const zhKeys = Object.keys(STRINGS.zh).sort();
  assert.deepEqual(zhKeys, enKeys, 'the two language tables have drifted apart');
});

test('every entry is a non-empty string or a function', () => {
  for (const lang of ['en', 'zh']) {
    for (const [key, value] of Object.entries(STRINGS[lang])) {
      const ok = (typeof value === 'string' && value.length > 0) || typeof value === 'function';
      assert.ok(ok, `${lang}.${key} should be a non-empty string or a function`);
    }
  }
});

test('interpolating entries produce the same argument count in both languages', () => {
  const sample = { en: STRINGS.en, zh: STRINGS.zh };
  for (const key of Object.keys(sample.en)) {
    const enVal = sample.en[key];
    const zhVal = sample.zh[key];
    if (typeof enVal === 'function' || typeof zhVal === 'function') {
      assert.equal(typeof enVal, typeof zhVal, `${key} is a function in one language but not the other`);
      if (typeof enVal === 'function') {
        assert.equal(enVal.length, zhVal.length, `${key} takes a different number of arguments per language`);
      }
    }
  }
});
