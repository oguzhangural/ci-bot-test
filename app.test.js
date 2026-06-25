const { add } = require('./app');

test('bu test kasıtlı olarak fail verecek', () => {
  expect(add(1, 2)).toBe(99);
});
