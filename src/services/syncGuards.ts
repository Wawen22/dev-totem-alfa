export const MAX_CONSECUTIVE_SYNC_WRITE_FAILURES = 3;

export function getMissingRequiredColumns(
  existingColumnNames: Iterable<string>,
  requiredColumnNames: readonly string[]
): string[] {
  const existing = new Set(existingColumnNames);
  return Array.from(new Set(requiredColumnNames)).filter((name) => !existing.has(name));
}

export function recordWriteOutcome(
  previousConsecutiveFailures: number,
  succeeded: boolean
): number {
  return succeeded ? 0 : previousConsecutiveFailures + 1;
}
