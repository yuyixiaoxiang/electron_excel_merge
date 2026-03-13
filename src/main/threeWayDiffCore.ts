import { diffArrays } from 'diff';

export type ThreeWayCompareMode = 'diff' | 'merge';
export type ComparableCellValue = string | number | null;
export type MergeCellStatus = 'unchanged' | 'ours-changed' | 'theirs-changed' | 'both-changed-same' | 'conflict';
export type DiffOp =
  | { type: 'equal'; aIndex: number; bIndex: number }
  | { type: 'delete'; aIndex: number }
  | { type: 'insert'; bIndex: number };

export interface ThreeWayRuntimeConfig {
  alignBaseSide: 'base' | 'ours';
  useTwoWayContentAlignment: boolean;
  shortCircuitWhenOursEqualsTheirs: boolean;
}

export const normalizeComparableCellValue = (value: ComparableCellValue): string => {
  if (value === null || value === undefined) return '';
  if (typeof value === 'string') return value.trim();
  if (typeof value === 'number') return String(value);
  return String(value);
};

export const sameComparableCellValue = (a: ComparableCellValue, b: ComparableCellValue): boolean =>
  normalizeComparableCellValue(a) === normalizeComparableCellValue(b);

export const getThreeWayRuntimeConfig = (compareMode: ThreeWayCompareMode): ThreeWayRuntimeConfig => {
  if (compareMode === 'diff') {
    return {
      alignBaseSide: 'ours',
      useTwoWayContentAlignment: true,
      shortCircuitWhenOursEqualsTheirs: true,
    };
  }
  return {
    alignBaseSide: 'base',
    useTwoWayContentAlignment: false,
    shortCircuitWhenOursEqualsTheirs: false,
  };
};

export const classifyThreeWayCell = (
  baseValue: ComparableCellValue,
  oursValue: ComparableCellValue,
  theirsValue: ComparableCellValue,
): { status: MergeCellStatus; mergedValue: ComparableCellValue } => {
  const equalBO = sameComparableCellValue(baseValue, oursValue);
  const equalBT = sameComparableCellValue(baseValue, theirsValue);
  const equalOT = sameComparableCellValue(oursValue, theirsValue);

  if (equalBO && equalBT) {
    return { status: 'unchanged', mergedValue: oursValue };
  }
  if (!equalBO && equalBT) {
    return { status: 'ours-changed', mergedValue: oursValue };
  }
  if (equalBO && !equalBT) {
    return { status: 'theirs-changed', mergedValue: theirsValue };
  }
  if (!equalBO && !equalBT && equalOT) {
    return { status: 'both-changed-same', mergedValue: oursValue };
  }
  return { status: 'conflict', mergedValue: oursValue };
};

export const diffArraysToOps = (a: string[], b: string[]): DiffOp[] => {
  const changes = diffArrays(a, b, { oneChangePerToken: true });
  const ops: DiffOp[] = [];
  let aIndex = 0;
  let bIndex = 0;

  for (const change of changes) {
    const count = change.count ?? change.value.length;
    if (change.removed) {
      for (let i = 0; i < count; i += 1) {
        ops.push({ type: 'delete', aIndex });
        aIndex += 1;
      }
      continue;
    }
    if (change.added) {
      for (let i = 0; i < count; i += 1) {
        ops.push({ type: 'insert', bIndex });
        bIndex += 1;
      }
      continue;
    }
    for (let i = 0; i < count; i += 1) {
      ops.push({ type: 'equal', aIndex, bIndex });
      aIndex += 1;
      bIndex += 1;
    }
  }

  return ops;
};
