/** Give cancellation and other port traffic a task boundary without a chain
 * of nested timers (which browsers clamp). No source data enters this queue. */
export function createWorkbookTaskQueue(): { yield(): Promise<void>; dispose(): void } {
  const channel = new MessageChannel(); let disposed = false, scheduled = false;
  let pending: (() => void)[] = [];
  const drain = (): void => { scheduled = false; const batch = pending; pending = []; for (const resolve of batch) resolve(); };
  channel.port1.addEventListener("message", drain); channel.port1.start();
  return {
    yield() {
      if (disposed) return Promise.resolve();
      return new Promise<void>(resolve => {
        pending.push(resolve);
        if (!scheduled) { scheduled = true; channel.port2.postMessage(0); }
      });
    },
    dispose() {
      if (disposed) return; disposed = true;
      channel.port1.removeEventListener("message", drain); channel.port1.close(); channel.port2.close();
      // Owners resume and observe their existing abort/disposal guards.
      drain();
    },
  };
}
