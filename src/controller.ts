export interface WordState {
  readonly ready: boolean;
  readonly readOnly: boolean;
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  readonly bold: boolean;
  readonly italic: boolean;
}

export interface WordEngineApi {
  asc_registerCallback(name: string, callback: (value: boolean) => void): void;
  asc_unregisterCallback(name: string, callback: (value: boolean) => void): void;
  put_TextPrBold(value: boolean): void;
  put_TextPrItalic(value: boolean): void;
  Undo(): void;
  Redo(): void;
  asc_Save(): unknown;
}

export interface WordController {
  getSnapshot(): WordState;
  getServerSnapshot(): WordState;
  subscribe(listener: () => void): () => void;
  setSession(state: {ready: boolean; readOnly: boolean}): void;
  toggleBold(): boolean;
  toggleItalic(): boolean;
  undo(): boolean;
  redo(): boolean;
  /**
   * True means the command was dispatched to the SDK.
   * It is not a server persistence acknowledgement.
   */
  save(): boolean;
  /** Unsubscribes this bridge only. Does not destroy the SDK or disconnect sockets. */
  dispose(): void;
}

const initial: WordState = Object.freeze({
  ready: false,
  readOnly: true,
  canUndo: false,
  canRedo: false,
  bold: false,
  italic: false
});

const requiredMethods = [
  'asc_registerCallback',
  'asc_unregisterCallback',
  'put_TextPrBold',
  'put_TextPrItalic',
  'Undo',
  'Redo',
  'asc_Save'
] as const;

type RequiredMethod = (typeof requiredMethods)[number];
type WordStateKey = keyof WordState;
type EngineCallback = (value: boolean) => void;
type CallbackSubscription = readonly [name: string, callback: EngineCallback];

function assertWordEngineApi(api: unknown): asserts api is WordEngineApi {
  const candidate = api as Partial<Record<RequiredMethod, unknown>> | null | undefined;

  for (const method of requiredMethods) {
    if (typeof candidate?.[method] !== 'function') {
      throw new TypeError(`Word engine is missing ${method}`);
    }
  }
}

/** Attach to an already initialized asc_docs_api. This does not load or own the engine. */
export function createWordController(api: WordEngineApi): WordController {
  assertWordEngineApi(api);

  let snapshot = initial;
  let disposed = false;
  const listeners = new Set<() => void>();
  const callbacks: CallbackSubscription[] = [];

  const update = (changes: Partial<WordState>) => {
    if (disposed) return;

    const keys = Object.keys(changes) as WordStateKey[];
    if (keys.every(key => snapshot[key] === changes[key])) return;

    snapshot = Object.freeze({...snapshot, ...changes});
    for (const listener of [...listeners]) listener();
  };

  const subscribeEvent = (name: string, key: WordStateKey) => {
    const callback: EngineCallback = value => update({[key]: Boolean(value)});
    api.asc_registerCallback(name, callback);
    callbacks.push([name, callback]);
  };

  try {
    subscribeEvent('asc_onCanUndo', 'canUndo');
    subscribeEvent('asc_onCanRedo', 'canRedo');
    subscribeEvent('asc_onBold', 'bold');
    subscribeEvent('asc_onItalic', 'italic');
  } catch (error) {
    for (const [name, callback] of callbacks) api.asc_unregisterCallback(name, callback);
    throw error;
  }

  const invoke = (operation: () => void, allowed = true): boolean => {
    if (disposed) throw new Error('Word controller is disposed');
    if (!snapshot.ready || snapshot.readOnly || !allowed) return false;
    operation();
    return true;
  };

  return Object.freeze({
    getSnapshot: () => snapshot,
    getServerSnapshot: () => initial,
    subscribe(listener: () => void) {
      if (disposed) throw new Error('Word controller is disposed');
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    // The runtime calls this after document readiness and permission changes.
    // This is UI state, never a replacement for engine/server authorization.
    setSession({ready, readOnly}: {ready: boolean; readOnly: boolean}) {
      if (disposed) throw new Error('Word controller is disposed');
      if (typeof ready !== 'boolean' || typeof readOnly !== 'boolean') {
        throw new TypeError('ready/readOnly must be booleans');
      }
      update({ready, readOnly});
    },
    toggleBold: () => invoke(() => api.put_TextPrBold(!snapshot.bold)),
    toggleItalic: () => invoke(() => api.put_TextPrItalic(!snapshot.italic)),
    undo: () => invoke(() => api.Undo(), snapshot.canUndo),
    redo: () => invoke(() => api.Redo(), snapshot.canRedo),
    save: () => invoke(() => {
      api.asc_Save();
    }),
    dispose() {
      if (disposed) return;
      // Notify mounted controls before detaching their external store.
      update({ready: false, readOnly: true});
      disposed = true;
      for (const [name, callback] of callbacks) api.asc_unregisterCallback(name, callback);
      listeners.clear();
    }
  });
}
