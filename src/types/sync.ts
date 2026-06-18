export type SyncDetailKey = "updated" | "created" | "unchanged" | "skipped" | "duplicates";

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

export type SyncResult = {
  success: boolean;
  message: string;
  details?: SyncDetailSection[];
};
