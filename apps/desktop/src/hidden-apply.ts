/** Debounced native apply. Skip while IME composition is active. Never treat the preview as file authority. */
export function createHiddenApplyScheduler({
  delayMs = 80,
  composing,
  apply,
}: {
  delayMs?: number;
  composing(): boolean;
  apply(): Promise<void> | void;
}) {
  let timer: ReturnType<typeof setTimeout> | undefined;
  let running = false;
  let pending = false;

  const fire = () => {
    timer = undefined;
    if (running || composing()) return;
    pending = false;
    running = true;
    Promise.resolve(apply()).finally(() => {
      running = false;
      if (pending) queue();
    });
  };

  const queue = () => {
    pending = true;
    if (running || timer || composing()) return;
    timer = setTimeout(fire, delayMs);
  };

  return {
    schedule: queue,
    async flush() {
      if (timer) { clearTimeout(timer); timer = undefined; }
      if (composing()) return;
      if (!running && pending) fire();
    },
    cancel() {
      pending = false;
      if (timer) { clearTimeout(timer); timer = undefined; }
    },
  };
}
