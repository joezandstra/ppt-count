import assert from "node:assert/strict";
import { afterEach, beforeEach, mock, test } from "node:test";

import { createRefresher } from "../src/refresher.js";

// Lets pending promise callbacks run (setImmediate is not mocked).
const flush = async () => {
  for (let i = 0; i < 10; i++) await new Promise((resolve) => setImmediate(resolve));
};

function deferred() {
  let resolve;
  let reject;
  const promise = new Promise((res, rej) => {
    resolve = res;
    reject = rej;
  });
  return { promise, resolve, reject };
}

// Each read() returns the next entry: a value, an Error (rejects), or a deferred.
function setup(reads, options = {}) {
  const log = { reads: 0, results: [], errors: [] };
  const read = () => {
    const next = reads[Math.min(log.reads, reads.length - 1)];
    log.reads++;
    if (next instanceof Error) return Promise.reject(next);
    if (next && typeof next.promise?.then === "function") return next.promise;
    return Promise.resolve(next);
  };
  const refresher = createRefresher({
    read,
    onResult: (result) => log.results.push(result),
    onError: (error) => log.errors.push(error),
    ...options,
  });
  return { refresher, log };
}

beforeEach(() => mock.timers.enable({ apis: ["setTimeout", "setInterval"] }));
afterEach(() => mock.timers.reset());

test("schedule() debounces a burst of triggers into one read", async () => {
  const { refresher, log } = setup(["A"]);
  refresher.schedule();
  refresher.schedule();
  refresher.schedule();
  mock.timers.tick(119);
  await flush();
  assert.equal(log.reads, 0);
  mock.timers.tick(1);
  await flush();
  assert.equal(log.reads, 1);
  assert.deepEqual(log.results, ["A"]);
});

test("refreshNow() reads straight away", async () => {
  const { refresher, log } = setup(["A"]);
  refresher.refreshNow();
  await flush();
  assert.deepEqual(log.results, ["A"]);
});

test("only one read runs at a time; a trigger during a read causes exactly one re-read and the stale result is dropped", async () => {
  const first = deferred();
  const { refresher, log } = setup([first, "second"]);
  refresher.refreshNow();
  await flush();
  refresher.schedule();
  mock.timers.tick(120);
  refresher.refreshNow();
  await flush();
  assert.equal(log.reads, 1);
  first.resolve("first");
  await flush();
  assert.equal(log.reads, 2);
  assert.deepEqual(log.results, ["second"]);
});

test("a failed read is retried once after 300 ms", async () => {
  const { refresher, log } = setup([new Error("transient"), "ok"]);
  refresher.refreshNow();
  await flush();
  assert.equal(log.reads, 1);
  mock.timers.tick(300);
  await flush();
  assert.equal(log.reads, 2);
  assert.deepEqual(log.results, ["ok"]);
  assert.deepEqual(log.errors, []);
});

test("two failures report an error and keep the last good result", async () => {
  const { refresher, log } = setup(["good", new Error("a"), new Error("b")]);
  refresher.refreshNow();
  await flush();
  refresher.refreshNow();
  await flush();
  mock.timers.tick(300);
  await flush();
  assert.deepEqual(log.results, ["good"]);
  assert.equal(log.errors.length, 1);
});

test("a sudden empty result is confirmed by a second read before it is shown", async () => {
  const { refresher, log } = setup(["full", "empty", "empty"], { isEmpty: (r) => r === "empty" });
  refresher.refreshNow();
  await flush();
  refresher.refreshNow();
  await flush();
  assert.deepEqual(log.results, ["full"]);
  mock.timers.tick(150);
  await flush();
  assert.equal(log.reads, 3);
  assert.deepEqual(log.results, ["full", "empty"]);
});

test("if the confirming read has text again, that is shown instead", async () => {
  const { refresher, log } = setup(["full", "empty", "full again"], { isEmpty: (r) => r === "empty" });
  refresher.refreshNow();
  await flush();
  refresher.refreshNow();
  await flush();
  mock.timers.tick(150);
  await flush();
  assert.deepEqual(log.results, ["full", "full again"]);
});

test("empty after empty needs no confirmation", async () => {
  const { refresher, log } = setup(["empty", "empty"], { isEmpty: (r) => r === "empty" });
  refresher.refreshNow();
  await flush();
  refresher.refreshNow();
  await flush();
  assert.equal(log.reads, 2);
  assert.deepEqual(log.results, ["empty", "empty"]);
});

test("poll() reads on an interval only while shouldPoll() is true", async () => {
  let visible = true;
  const { refresher, log } = setup(["A"]);
  refresher.poll(1000, () => visible);
  mock.timers.tick(1000);
  await flush();
  assert.equal(log.reads, 1);
  visible = false;
  mock.timers.tick(1000);
  await flush();
  assert.equal(log.reads, 1);
});

test("poll() skips a tick while a read is still running", async () => {
  const slow = deferred();
  const { refresher, log } = setup([slow]);
  refresher.poll(1000, () => true);
  mock.timers.tick(1000);
  await flush();
  mock.timers.tick(1000);
  await flush();
  assert.equal(log.reads, 1);
});

test("stop() cancels pending and future work", async () => {
  const { refresher, log } = setup(["A"]);
  refresher.poll(1000, () => true);
  refresher.schedule();
  refresher.stop();
  mock.timers.tick(5000);
  await flush();
  assert.equal(refresher.refreshNow(), undefined);
  await flush();
  assert.equal(log.reads, 0);
});
