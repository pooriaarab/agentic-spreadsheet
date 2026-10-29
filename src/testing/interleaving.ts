import { onTestFinished } from "vitest";

// Deterministic interleaving for tests. A sequential test (call A, then call B, assert) passes
// with or without a fix when the defect only exists while A's commit is in flight. This helper
// parks an operation at a named point, lets the test drive a second operation to completion,
// then releases the first. No timers and no wall-clock waits.
//
//   const lane = interleaving();
//   const wire = lane.checkpoint("update travels");
//   const send = lane.start("update reaches B", () => wire.wrap(deliver)(update));
//   await lane.whileParkedAt(wire, () => b.writeRange(...)); // B acts before the update lands
//   await send;
//
// A lane fails its test when a checkpoint never gets a window, an operation is still in flight at
// the end, or a window drives nothing. Each failure message carries the timeline.

export class InterleavingError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "InterleavingError";
  }
}

export interface Checkpoint {
  readonly name: string;
  arrive(): Promise<void>;
  wrap<A extends unknown[], R>(fn: (...args: A) => R | Promise<R>): (...args: A) => Promise<R>;
  arrivals(): number;
}

export interface OpenWindow {
  quiesce(): Promise<void>;
  settledSoFar(): readonly string[];
  stillInFlight(): readonly string[];
}

export interface ParkedWindow {
  readonly checkpoint: string;
  readonly startedInside: readonly string[];
  readonly settledInside: readonly string[];
  readonly marksInside: readonly string[];
}

export interface Interleaving {
  checkpoint(name: string): Checkpoint;
  start<T>(name: string, run: () => Promise<T>): Promise<T>;
  mark(label: string): void;
  whileParkedAt(checkpoint: Checkpoint, drive: (open: OpenWindow) => void | Promise<void>): Promise<ParkedWindow>;
  timeline(): readonly string[];
  checkFixture(): void;
}

type OpState = "pending" | "fulfilled" | "rejected";

const isThenable = (value: unknown): value is Promise<unknown> =>
  (typeof value === "object" || typeof value === "function") && value !== null && typeof (value as { then?: unknown }).then === "function";

class HeldCheckpoint implements Checkpoint {
  private armed = true;
  private readonly parked: (() => void)[] = [];
  private readonly waiting: (() => void)[] = [];
  private reached = 0;
  claimed = false;

  constructor(readonly name: string) {}

  arrive(): Promise<void> {
    this.reached += 1;
    if (!this.armed) return Promise.resolve();
    const held = new Promise<void>((resolve) => this.parked.push(resolve));
    for (const notify of this.waiting.splice(0)) notify();
    return held;
  }

  wrap<A extends unknown[], R>(fn: (...args: A) => R | Promise<R>): (...args: A) => Promise<R> {
    return async (...args: A): Promise<R> => {
      await this.arrive();
      return await fn(...args);
    };
  }

  arrivals(): number {
    return this.reached;
  }

  isParked(): boolean {
    return this.parked.length > 0;
  }

  parkedHere(): Promise<void> {
    return this.isParked() ? Promise.resolve() : new Promise<void>((resolve) => this.waiting.push(resolve));
  }

  releaseAll(): number {
    this.armed = false;
    const released = this.parked.splice(0);
    for (const resume of released) resume();
    return released.length;
  }
}

class Lane implements Interleaving {
  private readonly checkpoints: HeldCheckpoint[] = [];
  private readonly ops = new Map<string, OpState>();
  private readonly marks: string[] = [];
  private readonly events: string[] = [];
  private inFlight = 0;
  private readonly idleWaiters: (() => void)[] = [];

  checkpoint(name: string): Checkpoint {
    if (this.checkpoints.some((existing) => existing.name === name)) {
      throw new InterleavingError(`two checkpoints are both named "${name}"; a failure message has to tell them apart`);
    }
    const held = new HeldCheckpoint(name);
    this.checkpoints.push(held);
    return held;
  }

  start<T>(name: string, run: () => Promise<T>): Promise<T> {
    if (this.ops.has(name)) throw new InterleavingError(`two operations are both named "${name}"; a failure message has to tell them apart`);
    this.ops.set(name, "pending");
    this.inFlight += 1;
    this.events.push(`start ${name}`);
    let running: unknown;
    try {
      running = run();
    } catch (cause) {
      this.settle(name, "rejected");
      throw cause;
    }
    if (!isThenable(running)) {
      this.settle(name, "rejected");
      throw new InterleavingError(`operation "${name}" did not return a promise; use lane.mark for a synchronous effect`);
    }
    running.then(
      () => this.settle(name, "fulfilled"),
      () => this.settle(name, "rejected"),
    );
    return running as Promise<T>;
  }

  mark(label: string): void {
    this.marks.push(label);
    this.events.push(`mark ${label}`);
  }

  async whileParkedAt(checkpoint: Checkpoint, drive: (open: OpenWindow) => void | Promise<void>): Promise<ParkedWindow> {
    const held = this.checkpoints.find((candidate) => candidate === checkpoint);
    if (held === undefined) throw new InterleavingError(`checkpoint "${checkpoint.name}" belongs to a different lane`);
    if (held.claimed) throw new InterleavingError(`checkpoint "${held.name}" already had its window; declare a second checkpoint for a second one`);
    held.claimed = true;
    await this.awaitArrival(held);
    const startedBefore = new Set(this.ops.keys());
    const settledBefore = new Set([...this.ops].filter(([, state]) => state !== "pending").map(([name]) => name));
    const marksBefore = this.marks.length;
    this.events.push(`parked ${held.name}`);
    const open: OpenWindow = {
      quiesce: () => this.quiesce(),
      settledSoFar: () => this.settledSince(settledBefore),
      stillInFlight: () => [...this.ops].filter(([name, state]) => state === "pending" && !startedBefore.has(name)).map(([name]) => name),
    };
    try {
      const driven = drive(open);
      if (driven !== undefined) {
        if (!isThenable(driven)) throw new InterleavingError(`the callback for "${held.name}" must return nothing or a promise`);
        await driven;
      }
    } finally {
      this.events.push(`released ${held.name} (${held.releaseAll()} parked)`);
    }
    const window: ParkedWindow = {
      checkpoint: held.name,
      startedInside: [...this.ops.keys()].filter((name) => !startedBefore.has(name)),
      settledInside: this.settledSince(settledBefore),
      marksInside: this.marks.slice(marksBefore),
    };
    if (window.startedInside.length === 0 && window.settledInside.length === 0 && window.marksInside.length === 0) {
      throw new InterleavingError(`the window at "${held.name}" drove nothing, so it proves only that the operation works alone. ${this.trail()}`);
    }
    return window;
  }

  timeline(): readonly string[] {
    return [...this.events];
  }

  checkFixture(): void {
    const problems: string[] = [];
    for (const held of this.checkpoints) {
      if (!held.claimed) problems.push(`checkpoint "${held.name}" was created but no window was ever opened over it`);
    }
    const stranded = [...this.ops].filter(([, state]) => state === "pending").map(([name]) => `"${name}"`);
    if (stranded.length > 0) problems.push(`operation(s) ${stranded.join(", ")} were still in flight when the test ended; await them after the window closes`);
    if (problems.length > 0) throw new InterleavingError(`${problems.join("; ")}. ${this.trail()}`);
  }

  private async awaitArrival(held: HeldCheckpoint): Promise<void> {
    if (held.isParked()) return;
    if (this.inFlight === 0) throw new InterleavingError(`nothing is in flight, so nothing can park at "${held.name}". ${this.trail()}`);
    const arrived = await Promise.race([held.parkedHere().then(() => true), this.nothingInFlight().then(() => false)]);
    if (!arrived) throw new InterleavingError(`no operation reached "${held.name}"; every one settled first. ${this.trail()}`);
  }

  private settledSince(settledBefore: ReadonlySet<string>): string[] {
    return [...this.ops].filter(([name, state]) => state !== "pending" && !settledBefore.has(name)).map(([name]) => name);
  }

  private async quiesce(): Promise<void> {
    let seen = -1;
    for (let turn = 0; turn < 1000; turn += 1) {
      if (seen === this.events.length) return;
      seen = this.events.length;
      await new Promise<void>((resolve) => setImmediate(resolve));
    }
    throw new InterleavingError(`quiesce turned the event loop 1000 times and the lane kept changing. ${this.trail()}`);
  }

  private nothingInFlight(): Promise<void> {
    return this.inFlight === 0 ? Promise.resolve() : new Promise<void>((resolve) => this.idleWaiters.push(resolve));
  }

  private settle(name: string, state: "fulfilled" | "rejected"): void {
    if (this.ops.get(name) !== "pending") return;
    this.ops.set(name, state);
    this.events.push(`${state} ${name}`);
    this.inFlight -= 1;
    if (this.inFlight === 0) for (const notify of this.idleWaiters.splice(0)) notify();
  }

  private trail(): string {
    return `Timeline: ${this.events.length === 0 ? "(nothing happened)" : this.events.join(" -> ")}.`;
  }
}

// Make a lane for one test. It checks the fixture when the test ends, unless the test already failed.
export function interleaving(): Interleaving {
  const lane = new Lane();
  onTestFinished((context) => {
    if ((context as { task?: { result?: { state?: unknown } } }).task?.result?.state === "fail") return;
    lane.checkFixture();
  });
  return lane;
}
