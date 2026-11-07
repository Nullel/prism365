import { MAX_REPORT_DOWNLOADS, type Pause, REPORT_WINDOW_MILLISECONDS } from "@prism365/audits";

// Keeps a full review inside Microsoft's limit on usage-report downloads: 14 in ten minutes, of
// which a review starts no more than 12. Each audit that downloads reports is charged what it can
// start at most, when it ends. An audit may start only when the audits that ended in the last ten
// minutes, with its own charge, come to no more than the limit. Its downloads fall between its
// start and its end, so every ten minutes that holds some of them also holds only audits counted
// here: the bound is conservative, never short.

interface Charge {
  downloads: number;
  endedAt: number;
}

export class ReportBudget {
  readonly #now: () => number;
  readonly #pause: Pause;
  readonly #limit: number;
  readonly #window: number;
  #charges: Charge[] = [];

  constructor(options: { now: () => number; pause: Pause; limit?: number; window?: number }) {
    this.#now = options.now;
    this.#pause = options.pause;
    this.#limit = options.limit ?? MAX_REPORT_DOWNLOADS;
    this.#window = options.window ?? REPORT_WINDOW_MILLISECONDS;
  }

  // Waits, if it must, until an audit that downloads `downloads` reports may start. Rejects when
  // `signal` aborts while waiting. After a wait the audits that were holding the window are
  // forgotten without asking the clock again, so a pause that returns early cannot make this loop.
  async waitFor(downloads: number, signal: AbortSignal): Promise<void> {
    const need = Math.min(downloads, this.#limit);
    for (;;) {
      signal.throwIfAborted();
      const at = this.#now();
      this.#forgetBefore(at);
      const used = this.#charges.reduce((sum, charge) => sum + charge.downloads, 0);
      const oldest = this.#charges[0];
      if (!oldest || used + need <= this.#limit) return;
      // The audits are in the order they ended, so the first to leave the window is the first.
      const free = oldest.endedAt + this.#window;
      await this.#pause(Math.max(0, free - at), signal);
      this.#forgetBefore(free);
    }
  }

  // An audit that downloaded reports has ended.
  record(downloads: number): void {
    if (downloads > 0)
      this.#charges.push({ downloads: Math.min(downloads, this.#limit), endedAt: this.#now() });
  }

  // How many downloads the audits that ended in the last ten minutes can have started.
  used(): number {
    this.#forgetBefore(this.#now());
    return this.#charges.reduce((sum, charge) => sum + charge.downloads, 0);
  }

  #forgetBefore(at: number): void {
    this.#charges = this.#charges.filter((charge) => charge.endedAt + this.#window > at);
  }
}
