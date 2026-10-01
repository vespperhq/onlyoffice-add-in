declare const Asc: {
  scope: Record<string, unknown>;
  plugin: {
    init: () => void;
    onDestroy?: () => void;
    executeMethod(
      name: string,
      params: unknown[] | null,
      callback?: (result: any) => void
    ): boolean;
    callCommand(
      command: () => unknown,
      close?: boolean,
      recalculate?: boolean,
      callback?: (result: any) => void
    ): void;
    attachEditorEvent(name: string, callback: () => void): void;
    detachEditorEvent(name: string): void;
  };
};

interface DocumentUpdateState {
  status: "pending" | "done" | "error";
  error: number;
  cancel?: () => void;
}

// Inside callCommand, `Api` is the Office JavaScript API and `editor` the
// editor's internal API (the same object before 9.4), whose undocumented
// methods were verified from 8.3.3 through 9.4.0. Minified internals are
// looked up by name, hence the index signatures.
declare const Api: {
  GetDocument(): {
    IsTrackRevisions(): boolean;
  };
};

declare const editor: {
  asc_coAuthoringGetUsers(): void;
  asc_CompareDocumentUrl(url: string, options: ComparisonOptions): void;
  asc_GetTrackRevisionsReportByAuthors(): Record<string, unknown[]>;
  asc_AcceptChanges(change: unknown): void;
  asc_registerCallback(name: string, callback: (...args: any[]) => void): void;
  asc_unregisterCallback(
    name: string,
    callback: (...args: any[]) => void
  ): void;
  vespperDocumentUpdate?: DocumentUpdateState;
  [minified: string]: any;
};

interface ComparisonOptions {}

declare const AscCommonWord: {
  ComparisonOptions: new () => ComparisonOptions;
  CompareBinary: (...args: unknown[]) => void;
  [minified: string]: any;
};

declare const AscCommon: Record<string, any>;
declare const AscFormat: Record<string, any>;
declare const AscWord: Record<string, any>;
