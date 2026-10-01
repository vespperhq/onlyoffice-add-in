import type { CompareFunctionPatcher } from "./types";

// The editor internals the patcher replaces, by their sdkjs source names.
type Names = {
  compareBinary: string; // AscCommonWord.CompareBinary
  comparison: string; // AscCommonWord.CDocumentComparison
  keepReviewMarks: string; // CDocumentComparison.bSaveCustomReviewType
  getUserName: string; // CDocumentComparison.getUserName
  compareRoots: string; // CDocumentComparison.compareRoots
  compareSectPr: string; // CDocumentComparison.compareSectPr
  createNodeFromRun: string; // CDocumentComparison._createNodeFromRun
  runContent: string; // ParaRun.Content
  elementType: string; // CRunElementBase.Type
  setReviewInfo: string; // CDocumentComparison.setReviewInfoForArray
  getReviewInfo: string; // CDocumentComparison.getReviewInfo
  getReviewType: string; // ParaRun.GetReviewType
  setReviewTypeWithInfo: string; // ParaRun.SetReviewTypeWithInfo
  reviewInfo: string; // ParaRun.ReviewInfo, and the AscWord.ReviewInfo class
  copy: string; // ReviewInfo.Copy
  prevType: string; // ReviewInfo.PrevType
  prevInfo: string; // ReviewInfo.PrevInfo
  node: string; // AscCommonWord.CNode
  setCommonReviewType: string; // CNode.setCommonReviewTypeWithInfo
  wordControl: string; // asc_docs_api.WordControl
  logicDocument: string; // WordControl.m_oLogicDocument
  trackRevisions: string; // CDocument.TrackRevisionsManager
  changedDocuments: string; // TrackRevisionsManager.Get_AllChangesLogicDocuments
  drawingProps: string; // AscFormat.CNvPr
  hasSameNameAndId: string; // CNvPr.hasSameNameAndId
};

// Adapts ONLYOFFICE's Compare feature (Collaboration > Compare) so it can apply
// an edited copy of the open document in place. Stock Compare (the official compare feature)
// diffs the editor against the copy and writes the differences in as tracked changes, but it
// was built for comparing two clean documents, so while patched:
//
// - Existing tracked changes are kept. Compare no longer accepts them before
//   diffing, nor asks the user to, nor re-credits them to the current user.
// - New changes are credited to `author` when one is given, so they can be
//   told apart and accepted afterwards (Editing mode).
// - Deleted text is left out when diffing, as it is from the copy, so the
//   words around it still match and are not inserted a second time, and
//   formatting stays on the text it belongs to.
// - Text that is still a pending insertion can be deleted, so rewording it
//   replaces it instead of adding the new words next to it, and the rest of a
//   partly edited change stays tracked.
// - Images are matched by name only, since exporting the document renumbers
//   their ids and every image would otherwise be inserted again.
//
// Patch right before a compare and unpatch once it ends; every replaced
// function is restored as it was.
//
// Most builds minify sdkjs, renaming these internals differently in every
// release, so their names are read off the code that uses them. Verified on
// the minified 9.2 build.
export class RuntimeCompareFunctionPatcher implements CompareFunctionPatcher {
  private readonly restores: (() => void)[] = [];
  private readonly hidden: [run: any, content: unknown[]][] = [];

  isSupported(): boolean {
    return this.getNames() !== null;
  }

  patch(author: string | null): void {
    enum ReviewType {
      Common = 0,
      Remove = 1,
      Add = 2,
    }
    enum RunElementType {
      ParagraphMark = 4,
    }

    const n = this.getNames()!;
    const comparison = AscCommonWord[n.comparison].prototype;
    const patcher = this;

    // Forcing CompareBinary skips the prompt to accept existing tracked
    // changes before comparing. The flag is its last parameter, which later
    // builds moved by dropping a leading api argument.
    for (const key of new Set(["CompareBinary", n.compareBinary])) {
      this.replace(
        AscCommonWord,
        key,
        (compareBinary) =>
          function (this: unknown, ...args: unknown[]) {
            const params = args.slice(0, compareBinary.length - 1);
            params.length = compareBinary.length - 1;
            return compareBinary.apply(this, [...params, true]);
          },
      );
    }
    // Compare accepts every tracked change before it diffs, and re-credits
    // the ones it touches. Both are skipped so existing changes stay as is.
    this.replace(
      editor[n.wordControl][n.logicDocument][n.trackRevisions],
      n.changedDocuments,
      () => () => ({}),
    );
    this.replace(
      comparison,
      "compare",
      (compare) =>
        function (this: any, callback: unknown) {
          this[n.keepReviewMarks] = true;
          if (author) this[n.getUserName] = () => author;
          return compare.call(this, callback);
        },
    );
    // Compare opens the revision without its tracked changes, so it reads
    // the revision's final text, while the editor still holds deleted text.
    // Words touching it ("business[, but]technical") would differ and be
    // inserted again, and formatting, which is matched up character by
    // character, would land next to where it was. Deleted runs are emptied
    // while the documents are diffed, and refilled right after, before
    // anything lays the document out. A paragraph mark always counts.
    this.replace(
      comparison,
      n.createNodeFromRun,
      (createNodeFromRun) =>
        function (this: unknown, run: any, ...rest: unknown[]) {
          const content: any[] = run[n.runContent];
          if (
            content.length &&
            run[n.getReviewType]?.() === ReviewType.Remove &&
            !content.some(
              (element) =>
                element[n.elementType] === RunElementType.ParagraphMark,
            )
          ) {
            patcher.hidden.push([run, content]);
            run[n.runContent] = [];
          }
          return createNodeFromRun.call(this, run, ...rest);
        },
    );
    // Diffing ends when the outermost compareSectPr, which follows
    // compareRoots and calls into it for headers and footers, returns.
    let depth = 0;
    for (const key of [n.compareRoots, n.compareSectPr]) {
      this.replace(
        comparison,
        key,
        (original) =>
          function (this: unknown, ...args: unknown[]) {
            depth += 1;
            try {
              return original.apply(this, args);
            } catch (error) {
              patcher.unhide();
              throw error;
            } finally {
              depth -= 1;
              if (!depth && key === n.compareSectPr) patcher.unhide();
            }
          },
      );
    }
    // Keeping existing changes also keeps Compare from deleting text that is
    // itself a pending insertion, so rewording one would add the new words
    // next to the old. Such text is deleted by the current author, and
    // rejecting that deletion restores the insertion.
    this.replace(
      comparison,
      n.setReviewInfo,
      (setReviewInfo) =>
        function (
          this: any,
          elements: any[],
          type: ReviewType,
          ...rest: unknown[]
        ) {
          const inserted =
            type === ReviewType.Remove
              ? elements.filter(
                  (element) =>
                    element[n.setReviewTypeWithInfo] &&
                    element[n.getReviewType]?.() === ReviewType.Add,
                )
              : [];
          setReviewInfo.call(
            this,
            elements.filter((element) => !inserted.includes(element)),
            type,
            ...rest,
          );
          for (const element of inserted) {
            const info = this[n.getReviewInfo](...rest);
            info[n.prevType] = ReviewType.Add;
            info[n.prevInfo] = element[n.reviewInfo][n.copy]();
            element[n.setReviewTypeWithInfo](ReviewType.Remove, info, false);
          }
        },
    );
    // Text left over from splitting a run is reset to unchanged, which would
    // accept the rest of a pending change ("October 15[, 2025]") on the spot.
    // It keeps its change instead, as in Combine.
    this.replace(
      AscCommonWord[n.node].prototype,
      n.setCommonReviewType,
      () =>
        function (element: any, info: unknown) {
          element[n.setReviewTypeWithInfo](
            element[n.getReviewType]?.() || ReviewType.Common,
            info,
          );
        },
    );
    // Compare pairs images by their drawing id, but exporting the document
    // renumbers those ids, so no image in the revision would match the
    // editor's and each one would be inserted again.
    this.replace(
      AscFormat[n.drawingProps].prototype,
      n.hasSameNameAndId,
      () =>
        function (this: any, other: any) {
          return Boolean(other) && this.name === other.name;
        },
    );
  }

  unpatch(): void {
    this.unhide();
    for (const restore of this.restores.splice(0).reverse()) restore();
  }

  private unhide(): void {
    const { runContent } = this.getNames()!;
    for (const [run, content] of this.hidden.splice(0)) {
      run[runContent] = content.concat(run[runContent]);
    }
  }

  private replace(
    target: any,
    key: string,
    wrap: (original: any) => unknown,
  ): void {
    const original = target[key];
    const own = Object.prototype.hasOwnProperty.call(target, key);
    target[key] = wrap(original);
    this.restores.push(() => {
      if (own) target[key] = original;
      else delete target[key];
    });
  }

  // Detection reads thousands of functions, so its result is kept on the
  // editor for the rest of the session.
  private getNames(): Names | null {
    if (editor.vespperCompareNames === undefined) {
      let names: Names | null = null;
      try {
        names = this.findNames();
      } catch {}
      editor.vespperCompareNames = names && this.hasAll(names) ? names : null;
    }
    return editor.vespperCompareNames;
  }

  private hasAll(n: Names): boolean {
    const comparison = AscCommonWord[n.comparison]?.prototype;
    const revisions =
      editor[n.wordControl]?.[n.logicDocument]?.[n.trackRevisions];
    return (
      typeof AscCommonWord[n.compareBinary] === "function" &&
      typeof comparison?.compare === "function" &&
      typeof comparison[n.getUserName] === "function" &&
      typeof comparison[n.compareRoots] === "function" &&
      typeof comparison[n.compareSectPr] === "function" &&
      typeof comparison[n.createNodeFromRun] === "function" &&
      typeof comparison[n.setReviewInfo] === "function" &&
      typeof comparison[n.getReviewInfo] === "function" &&
      typeof AscCommonWord[n.node]?.prototype?.[n.setCommonReviewType] ===
        "function" &&
      typeof revisions?.[n.changedDocuments] === "function" &&
      typeof AscFormat[n.drawingProps]?.prototype?.[n.hasSameNameAndId] ===
        "function"
    );
  }

  private findNames(): Names {
    if (typeof AscCommonWord.CDocumentComparison === "function") {
      return {
        compareBinary: "CompareBinary",
        comparison: "CDocumentComparison",
        keepReviewMarks: "bSaveCustomReviewType",
        getUserName: "getUserName",
        compareRoots: "compareRoots",
        compareSectPr: "compareSectPr",
        createNodeFromRun: "_createNodeFromRun",
        runContent: "Content",
        elementType: "Type",
        setReviewInfo: "setReviewInfoForArray",
        getReviewInfo: "getReviewInfo",
        getReviewType: "GetReviewType",
        setReviewTypeWithInfo: "SetReviewTypeWithInfo",
        reviewInfo: "ReviewInfo",
        copy: "Copy",
        prevType: "PrevType",
        prevInfo: "PrevInfo",
        node: "CNode",
        setCommonReviewType: "setCommonReviewTypeWithInfo",
        wordControl: "WordControl",
        logicDocument: "m_oLogicDocument",
        trackRevisions: "TrackRevisionsManager",
        changedDocuments: "Get_AllChangesLogicDocuments",
        drawingProps: "CNvPr",
        hasSameNameAndId: "hasSameNameAndId",
      };
    }

    const id = "[\\w$]+";
    // The compiler wraps long lines next to punctuation.
    const source = (fn: unknown) =>
      String(fn).replace(
        /\n(?=[=,;{}()[\]:?&|.])|(?<=[=,;{}()[\]:?&|])\n/g,
        "",
      );
    const escape = (text: string) => text.replace(/[$]/g, "\\$");
    const only = (pattern: string, text: string, group = 1): string => {
      const found = new Set(
        Array.from(text.matchAll(new RegExp(pattern, "g")), (m) => m[group]),
      );
      if (found.size !== 1) throw new Error(`Ambiguous: ${pattern}`);
      return [...found][0];
    };
    const paramsOf = (fn: string) =>
      /^function\s*[\w$]*\(([^)]*)\)/.exec(fn)![1].split(",");

    const compareBinary = AscCommonWord.CompareBinary;
    const binarySource = source(compareBinary);
    const alias = Object.keys(AscCommonWord).find(
      (key) => key !== "CompareBinary" && AscCommonWord[key] === compareBinary,
    )!;
    const comparison = only(
      `new AscCommonWord\\.(${id})\\([^;]*?\\.compare\\(\\)`,
      binarySource,
    );
    const control = new RegExp(
      `\\{var ${id}=(?:${id}|Asc\\.editor)\\.(${id})\\.(${id})[;,]`,
    ).exec(binarySource)!;

    const Comparison = AscCommonWord[comparison];
    const constructorSource = source(Comparison);
    const params = paramsOf(constructorSource);
    const originalDocument = only(
      `this\\.(${id})=${escape(params[0])};`,
      constructorSource,
    );
    const keepReviewMarks = only(
      `this\\.(${id})=!1;this\\.${id}=\\{${id}:!1,${id}:this`,
      constructorSource,
    );
    const wordsByOneSymbol = only(
      `this\\.(${id})=${escape(params[3])}[;,]`,
      constructorSource,
    );
    const methods = new Map<string, string>();
    for (const key of Object.keys(Comparison.prototype)) {
      const value = Comparison.prototype[key];
      if (typeof value === "function") methods.set(key, source(value));
    }
    const methodWith = (...needles: string[]): string => {
      const found = [...methods].filter(([, body]) =>
        needles.every((needle) => body.includes(needle)),
      );
      if (found.length !== 1) throw new Error(`Ambiguous: ${needles}`);
      return found[0][0];
    };

    const compare = methods.get("compare")!;
    const original = only(
      `var (${id})=this\\.${escape(originalDocument)}[,;]`,
      compare,
    );
    const revisions = new RegExp(
      `${escape(original)}\\.(${id})\\.(${id})\\(\\);for\\(`,
    ).exec(compare)!;
    const diff = new RegExp(
      `(${id})\\.(${id})\\((${id}),(${id})\\);\\1\\.(${id})\\(\\3,\\4,!\\1\\.options\\.${id}\\)`,
    ).exec(compare)!;
    const createNodeFromRun = methodWith(
      `||this.${wordsByOneSymbol})`,
      "createNode(",
      ".elements.length",
    );
    const createNodeFromRunSource = methods.get(createNodeFromRun)!;
    const [, run, runContent] = new RegExp(
      `^function\\((${id})[^)]*\\)\\{if\\(0<\\1\\.(${id})\\.length\\)`,
    ).exec(createNodeFromRunSource)!;
    const element = new RegExp(
      `var (${id})=${escape(run)}\\.${escape(runContent)}\\[${id}\\];`,
    ).exec(createNodeFromRunSource)![1];
    const setReviewInfo = methodWith(`this.${keepReviewMarks}&&`);
    const setReviewInfoSource = methods.get(setReviewInfo)!;
    const reviewInfoSetter = new RegExp(
      `if\\((${id})\\.(${id})\\)\\{var ${id}=this\\.(${id})\\(`,
    ).exec(setReviewInfoSource)!;

    // Every class, among those the editor exports, with a method whose
    // source matches.
    const findMethods = (pattern: RegExp) => {
      const found: { owner: string; method: string; body: string }[] = [];
      for (const [space, exports] of Object.entries({
        AscCommon,
        AscCommonWord,
        AscFormat,
        AscWord,
      })) {
        for (const key of Object.keys(exports)) {
          const cls = exports[key];
          if (typeof cls !== "function" || !cls.prototype) continue;
          for (const method of Object.keys(cls.prototype)) {
            let body = "";
            try {
              const value = cls.prototype[method];
              if (typeof value === "function") body = source(value);
            } catch {
              continue;
            }
            if (pattern.test(body)) {
              found.push({ owner: `${space}.${key}`, method, body });
            }
          }
        }
      }
      return found;
    };
    // A class can be exported under several names.
    const unique = (found: ReturnType<typeof findMethods>, what: string) => {
      const methods = new Set(
        found.map((item) => `${item.method}${item.body}`),
      );
      if (methods.size !== 1) throw new Error(`Ambiguous: ${what}`);
      return found[0];
    };

    // The Combine feature's comparison deletes pending insertions itself,
    // copying their review info first.
    const combine = unique(
      findMethods(
        new RegExp(
          `\\{var ${id}=this\\.${escape(reviewInfoSetter[3])}\\(\\);if\\(this\\.${escape(keepReviewMarks)}\\)`,
        ),
      ),
      "Combine's setReviewInfoForArray",
    );
    const [, reviewInfo, copy, savePrev] = new RegExp(
      `=${id}\\.(${id})\\.(${id})\\(\\),${id}\\.(${id})\\(`,
    ).exec(combine.body)!;
    const prevFields = unique(
      findMethods(
        new RegExp(
          `^function\\((${id})\\)\\{this\\.(${id})=\\1;this\\.(${id})=this\\.${escape(copy)}\\(\\)\\}$`,
        ),
      ).filter((found) => found.method === savePrev),
      "ReviewInfo.SavePrev",
    );
    const [, , prevType, prevInfo] = new RegExp(
      `^function\\((${id})\\)\\{this\\.(${id})=\\1;this\\.(${id})=`,
    ).exec(prevFields.body)!;

    const setCommonReviewType = unique(
      findMethods(
        new RegExp(
          `^function\\((${id}),(${id})\\)\\{\\1\\.${escape(reviewInfoSetter[2])}\\(${id},\\2\\)\\}$`,
        ),
      ).filter((found) => found.owner.startsWith("AscCommonWord.")),
      "CNode.setCommonReviewTypeWithInfo",
    );

    const drawing = unique(
      findMethods(
        /^function\(([\w$]+)\)\{return \1\?this\.id===\1\.id&&this\.name===\1\.name:!1\}$/,
      ).filter((found) => found.owner.startsWith("AscFormat.")),
      "hasSameNameAndId",
    );

    return {
      compareBinary: alias,
      comparison,
      keepReviewMarks,
      getUserName: methodWith('split(";")[0]'),
      compareRoots: diff[2],
      compareSectPr: diff[5],
      createNodeFromRun,
      runContent,
      elementType: only(
        `[^\\w$]${escape(element)}\\.(${id})[!=]==`,
        createNodeFromRunSource,
      ),
      setReviewInfo,
      getReviewInfo: reviewInfoSetter[3],
      getReviewType: new RegExp(`(${id})\\.(${id})&&\\1\\.\\2\\(\\)`).exec(
        setReviewInfoSource,
      )![2],
      setReviewTypeWithInfo: reviewInfoSetter[2],
      reviewInfo,
      copy,
      prevType,
      prevInfo,
      node: setCommonReviewType.owner.slice("AscCommonWord.".length),
      setCommonReviewType: setCommonReviewType.method,
      wordControl: control[1],
      logicDocument: control[2],
      trackRevisions: revisions[1],
      changedDocuments: revisions[2],
      drawingProps: drawing.owner.slice("AscFormat.".length),
      hasSameNameAndId: drawing.method,
    };
  }
}
