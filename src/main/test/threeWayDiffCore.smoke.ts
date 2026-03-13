import assert from 'node:assert/strict';
import { classifyThreeWayCell, diffArraysToOps, getThreeWayRuntimeConfig } from '../threeWayDiffCore';

const diffConfig = getThreeWayRuntimeConfig('diff');
assert.deepStrictEqual(diffConfig, {
  alignBaseSide: 'ours',
  useTwoWayContentAlignment: true,
  shortCircuitWhenOursEqualsTheirs: true,
});

const mergeConfig = getThreeWayRuntimeConfig('merge');
assert.deepStrictEqual(mergeConfig, {
  alignBaseSide: 'base',
  useTwoWayContentAlignment: false,
  shortCircuitWhenOursEqualsTheirs: false,
});

assert.deepStrictEqual(classifyThreeWayCell('A', 'A', 'A'), {
  status: 'unchanged',
  mergedValue: 'A',
});

assert.deepStrictEqual(classifyThreeWayCell('A', 'B', 'A'), {
  status: 'ours-changed',
  mergedValue: 'B',
});

assert.deepStrictEqual(classifyThreeWayCell('A', 'A', 'C'), {
  status: 'theirs-changed',
  mergedValue: 'C',
});

assert.deepStrictEqual(classifyThreeWayCell('A', 'B', 'B'), {
  status: 'both-changed-same',
  mergedValue: 'B',
});

assert.deepStrictEqual(classifyThreeWayCell('A', 'B', 'C'), {
  status: 'conflict',
  mergedValue: 'B',
});

assert.deepStrictEqual(
  diffArraysToOps(['row-1', 'row-2', 'row-3'], ['row-1', 'row-2b', 'row-3', 'row-4']),
  [
    { type: 'equal', aIndex: 0, bIndex: 0 },
    { type: 'delete', aIndex: 1 },
    { type: 'insert', bIndex: 1 },
    { type: 'equal', aIndex: 2, bIndex: 2 },
    { type: 'insert', bIndex: 3 },
  ],
);

console.log('threeWayDiffCore smoke test passed');
