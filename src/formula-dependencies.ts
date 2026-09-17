/** Observed scalar calculation dependencies, keyed by stable sheet/cell identity.
 * Overflow disables reuse on the next revision instead of retaining incomplete edges.
 */
export class FormulaDependencies {
  private precedents = new Map<string, Set<string>>();
  private dependents = new Map<string, Set<string>>();
  private edges = 0;
  overflow = false;
  constructor(private limit = 100000) {}
  read(owner: string | undefined, source: string) {
    if (!owner || this.overflow) return;
    const sources = this.precedents.get(owner) || new Set<string>();
    if (sources.has(source)) return;
    if (this.edges >= this.limit) {
      this.overflow = true;
      return;
    }
    sources.add(source);
    this.precedents.set(owner, sources);
    const consumers = this.dependents.get(source) || new Set<string>();
    consumers.add(owner);
    this.dependents.set(source, consumers);
    this.edges++;
  }
  forget(owner: string) {
    for (const source of this.precedents.get(owner) || []) {
      const consumers = this.dependents.get(source)!;
      consumers.delete(owner);
      this.edges--;
      if (!consumers.size) this.dependents.delete(source);
    }
    this.precedents.delete(owner);
  }
  affected(changed: Iterable<string>) {
    const dirty = new Set(changed),
      queue = [...dirty];
    for (let i = 0; i < queue.length; i++)
      for (const target of this.dependents.get(queue[i]) || [])
        if (!dirty.has(target)) {
          dirty.add(target);
          queue.push(target);
        }
    return dirty;
  }
  clear() {
    this.precedents.clear();
    this.dependents.clear();
    this.edges = 0;
    this.overflow = false;
  }
  get size() {
    return this.edges;
  }
}
