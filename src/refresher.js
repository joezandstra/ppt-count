// Decides when to re-read the selection. It debounces bursts of triggers, runs
// one read at a time, retries a failed read once, and double-checks a sudden
// empty selection (PowerPoint occasionally reports one for a moment). This is pure
// scheduling logic: the caller supplies read() and the callbacks, and timers can
// be injected for tests.

/**
 * @template T
 * @param {{
 *   read: () => Promise<T>,
 *   onResult: (result: T) => void,
 *   onError: (error: unknown) => void,
 *   isEmpty?: (result: T) => boolean,
 *   timers?: Pick<typeof globalThis, "setTimeout" | "clearTimeout" | "setInterval" | "clearInterval">,
 *   debounceMs?: number,
 *   retryMs?: number,
 *   confirmEmptyMs?: number,
 * }} options
 */
export function createRefresher({ read, onResult, onError, isEmpty = () => false, timers = globalThis, debounceMs = 120, retryMs = 300, confirmEmptyMs = 150 }) {
  let debounceTimer = null;
  let pollTimer = null;
  let running = false;
  let dirty = false; // a trigger arrived while a read was running
  let stopped = false;
  let last;
  let hasLast = false;

  const wait = (ms) => new Promise((resolve) => timers.setTimeout(resolve, ms));

  async function readWithRetry() {
    try {
      return await read();
    } catch {
      await wait(retryMs);
      return read();
    }
  }

  async function run() {
    if (stopped) return;
    if (running) {
      dirty = true;
      return;
    }
    running = true;
    dirty = false;
    try {
      let result = await readWithRetry();
      if (!dirty && hasLast && isEmpty(result) && !isEmpty(last)) {
        await wait(confirmEmptyMs);
        if (!dirty) result = await readWithRetry();
      }
      // If the selection changed during the read, this result is already stale.
      if (!dirty && !stopped) {
        last = result;
        hasLast = true;
        onResult(result);
      }
    } catch (error) {
      if (!dirty && !stopped) onError(error);
    } finally {
      running = false;
    }
    if (dirty && !stopped) run();
  }

  return {
    schedule() {
      if (stopped) return;
      timers.clearTimeout(debounceTimer);
      debounceTimer = timers.setTimeout(() => {
        debounceTimer = null;
        run();
      }, debounceMs);
    },

    refreshNow() {
      if (stopped) return undefined;
      timers.clearTimeout(debounceTimer);
      debounceTimer = null;
      return run();
    },

    poll(intervalMs, shouldPoll = () => true) {
      timers.clearInterval(pollTimer);
      pollTimer = timers.setInterval(() => {
        if (!running && debounceTimer === null && shouldPoll()) run();
      }, intervalMs);
    },

    stop() {
      stopped = true;
      timers.clearTimeout(debounceTimer);
      timers.clearInterval(pollTimer);
    },
  };
}
