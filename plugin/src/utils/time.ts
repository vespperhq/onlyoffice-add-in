export function formatMilliseconds(milliseconds: number): string {
  return milliseconds < 1_000
    ? `${Math.round(milliseconds)}ms`
    : `${(milliseconds / 1_000).toFixed(1)}s`;
}

export function sleep(milliseconds: number): Promise<void> {
  return new Promise((resolve) => window.setTimeout(resolve, milliseconds));
}
