const { test } = require('node:test');
const assert = require('node:assert/strict');
const { validateRecipeFields } = require('../src/engine/recipe');

test('accepts a well-formed recipe', () => {
  const result = validateRecipeFields([
    { name: '營收', keywords: '營收, 營業收入' },
    { name: '費用', keywords: '支出, 虧損' },
  ]);
  assert.equal(result.valid, true);
  assert.equal(result.fields.length, 2);
});

test('rejects a file whose top level is not an array', () => {
  const result = validateRecipeFields({ name: '營收', keywords: '營收' });
  assert.equal(result.valid, false);
  assert.match(result.error, /陣列/);
});

test('rejects null, a plain string, and a number', () => {
  for (const bad of [null, 'not json', 42]) {
    assert.equal(validateRecipeFields(bad).valid, false);
  }
});

test('drops individual malformed entries but keeps the valid ones', () => {
  const result = validateRecipeFields([
    { name: '營收', keywords: '營收' },
    { name: 123, keywords: '壞資料' }, // name is not a string
    { keywords: '缺 name' },
    null,
    'not an object',
    { name: '費用', keywords: '支出' },
  ]);
  assert.equal(result.valid, true);
  assert.deepEqual(
    result.fields.map((f) => f.name),
    ['營收', '費用']
  );
});

test('rejects an empty array and an array of only garbage', () => {
  assert.equal(validateRecipeFields([]).valid, false);
  assert.equal(validateRecipeFields([null, 1, 'x', {}]).valid, false);
});
