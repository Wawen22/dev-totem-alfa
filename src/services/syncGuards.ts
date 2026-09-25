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

export function shouldApplyExcelAuthoritativeUpdate(
  changedBusinessFields: number,
  identityNeedsAlignment: boolean
): boolean {
  return changedBusinessFields > 0 || identityNeedsAlignment;
}

export function buildExcelAuthoritativePatch(
  fields: Record<string, unknown>,
  comparedFieldKeys: Iterable<string>,
  hasBusinessChanges: boolean
): Record<string, unknown> {
  if (!hasBusinessChanges) {
    return { IdentLotto: fields.IdentLotto };
  }

  const patchKeys = new Set(["Title", "IdentLotto", ...comparedFieldKeys]);
  return Object.fromEntries(
    Array.from(patchKeys)
      .filter((fieldKey) => Object.prototype.hasOwnProperty.call(fields, fieldKey))
      .map((fieldKey) => [fieldKey, fields[fieldKey]])
  );
}
