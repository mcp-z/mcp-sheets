// Bookkeeping for the spreadsheets one test run creates. A delete is attempted once per id: an id
// whose delete failed is never tried again and is reported at the end.

type Entry = { title: string; attempted: boolean; failure?: string };

export class RunRegistry {
  private entries = new Map<string, Entry>();

  add(id: string, title: string): void {
    this.entries.set(id, { title, attempted: false });
  }

  get size(): number {
    return this.entries.size;
  }

  /** Attempt the delete of one id. The id leaves the registry only if the delete succeeds; a failure is recorded and rethrown. */
  async remove(id: string, del: (id: string) => Promise<void>): Promise<void> {
    const entry = this.entries.get(id);
    if (entry) entry.attempted = true;
    try {
      await del(id);
      this.entries.delete(id);
    } catch (error) {
      if (entry) entry.failure = error instanceof Error ? error.message : String(error);
      throw error;
    }
  }

  /** Delete every id no delete has been attempted on, then throw naming each id still held (id, title, failure). */
  async cleanup(del: (id: string) => Promise<void>): Promise<void> {
    const untried = [...this.entries].filter(([, entry]) => !entry.attempted).map(([id]) => id);
    await Promise.allSettled(untried.map((id) => this.remove(id, del)));
    if (this.entries.size > 0) {
      const left = [...this.entries].map(([id, entry]) => `${id} "${entry.title}" (${entry.failure ?? 'not deleted'})`);
      throw new Error(`run left ${left.length} spreadsheet(s) behind, remove by hand: ${left.join('; ')}`);
    }
  }
}
