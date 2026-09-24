export type SyncDetailKey =
  | "updated"
  | "created"
  | "unchanged"
  | "skipped"
  | "duplicates"
  | "sharepoint-only";

export type SyncFieldChange = {
  field: string;
  previous: string;
  next: string;
};

export type SyncDetailItem = {
  label: string;
  code?: string;
  reference?: string;
  detail?: string;
  changes?: SyncFieldChange[];
};

export type SyncDetailSection = {
  key: SyncDetailKey;
  label: string;
  items: SyncDetailItem[];
};

export type SyncCleanupCandidate = {
  itemId: string;
  label: string;
  reason: "identical-duplicate" | "incomplete-duplicate" | "empty-item";
};

export type SyncCleanupPlan = {
  listKind: "FORGIATI" | "TUBI";
  safeDelete: SyncCleanupCandidate[];
  requiresReview: number;
};

export type SyncResult = {
  success: boolean;
  message: string;
  details?: SyncDetailSection[];
  cleanupPlan?: SyncCleanupPlan;
};
