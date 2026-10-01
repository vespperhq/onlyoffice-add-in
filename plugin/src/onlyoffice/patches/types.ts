// Replaces editor functions so that ONLYOFFICE's Compare keeps a document's
// existing tracked changes. An implementing class is written into command
// source and instantiated inside the editor, so it may reference editor
// globals but nothing else outside its own class body; not even a base class.
export interface CompareFunctionPatcher {
  // Whether this editor build has every internal the patcher relies on.
  isSupported(): boolean;
  // Changes how the next Compare runs. A non-null author replaces the
  // revision's author on the tracked changes Compare creates.
  patch(author: string | null): void;
  // Restores everything patch() replaced.
  unpatch(): void;
}

// The class itself: commands receive it and construct their own instance.
export type CompareFunctionPatcherClass = new () => CompareFunctionPatcher;
