/* global clearInterval, console, CustomFunctions, setInterval */

/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
export function add(first: number, second: number): number {
  return first + second;
}

/**
 * Displays the current time once a second.
 * @customfunction
 * @param invocation Custom function handler
 */
export function clock(invocation: CustomFunctions.StreamingInvocation<string>): void {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time.
 * @returns String with the current time formatted for the current locale.
 */
export function currentTime(): string {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param incrementBy Amount to increment
 * @param invocation Custom function handler
 */
export function increment(
  incrementBy: number,
  invocation: CustomFunctions.StreamingInvocation<number>
): void {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param message String to write.
 * @returns String to write.
 */
export function logMessage(message: string): string {
  console.log(message);

  return message;
}

/* -----------------------------------------------------------------------
   NEW  –  TESTVELIXO.FACTORIALROW
------------------------------------------------------------------------ */

/**
 * FACTORIALROW(N) → spill of 1! … N!
 * Row/column orientation is read from localStorage ("row" | "column").
 * Results are cached and returned as strings to preserve precision.
 *
 * @customfunction FACTORIALROW
 * @param n Largest integer (1 – 500)
 */
export function factorialRow(n: number): string[] | string[][] {
  if (!Number.isFinite(n) || n < 1) throw new Error("N must be a positive integer");
  if (n > 500) throw new Error("N too large – max is 500");

  const orientation =
    (globalThis.localStorage?.getItem("orientation") ?? "row").toLowerCase();
  const vertical = orientation === "column";

  // ---- session-wide factorial cache (shared runtime) -------------------
  const key = "__factorial_cache__";
  // @ts-ignore
  const cache: bigint[] = (globalThis[key] ??= [0n, 1n]); // seed 0!,1!

  for (let i = cache.length; i <= n; i++) cache[i] = cache[i - 1] * BigInt(i);

  const slice = cache.slice(1, n + 1).map((v) => v.toString());
  return vertical ? slice.map((v) => [v]) : slice;
}
