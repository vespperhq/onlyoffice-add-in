import { sleep } from "../utils/time";
import { RuntimeCompareFunctionPatcher } from "./patches/runtime";
import type { CompareFunctionPatcherClass } from "./patches/types";

const METHOD_TIMEOUT_MS = 60_000;
const COMMAND_TIMEOUT_MS = 60_000;
const COMMAND_RETRY_MS = 100;

export function executeMethod<T>(name: string, params: unknown[]): Promise<T> {
  return new Promise((resolve, reject) => {
    const timer = window.setTimeout(
      () => reject(new Error(`${name} timed out after ${METHOD_TIMEOUT_MS}ms`)),
      METHOD_TIMEOUT_MS
    );

    const started = Asc.plugin.executeMethod(name, params, (result: T) => {
      window.clearTimeout(timer);
      resolve(result);
    });
    if (!started) {
      window.clearTimeout(timer);
      reject(new Error(`ONLYOFFICE rejected ${name}`));
    }
  });
}

let commandQueue: Promise<unknown> = Promise.resolve();

// The plugin SDK keeps a single callCommand callback, so an overlapping call
// would steal the pending result; commands run one at a time instead.
function runCommand<T>(command: () => T): Promise<T | undefined> {
  const result = commandQueue.then(
    () =>
      new Promise<T | undefined>((resolve) => {
        Asc.plugin.callCommand(command, false, false, resolve);
      })
  );
  commandQueue = result.catch(() => undefined);
  return result;
}

// Runs a command inside the editor, which only receives its source text, so
// it may use editor globals and Asc.scope but nothing from this module.
// While a long action such as saving or loading images blocks the editor, it
// skips commands and answers undefined, so commands must always return a value
// and are retried until the editor is idle.
export async function callCommand<T>(command: () => T): Promise<T> {
  const deadline = Date.now() + COMMAND_TIMEOUT_MS;
  for (;;) {
    const result = await runCommand(command);
    if (result !== undefined) return result;
    if (Date.now() > deadline) throw new Error("The editor stayed busy");
    await sleep(COMMAND_RETRY_MS);
  }
}

// Editors from 9.4 on let a command eval only once, for itself, so the
// patcher class is written into the command's source instead.
export function withCompareFunctionPatcher<T>(
  command: (Patcher: CompareFunctionPatcherClass) => T
): () => T {
  const source = `function () { return (${command})(${RuntimeCompareFunctionPatcher}); }`;
  return { toString: () => source } as unknown as () => T;
}
