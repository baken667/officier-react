const initial = Object.freeze({ready: false, readOnly: true, canUndo: false, canRedo: false, bold: false, italic: false});

/** Attach to an already initialized asc_docs_api. This does not load or own the engine. */
export function createWordController(api) {
  for (const method of ['asc_registerCallback', 'asc_unregisterCallback', 'put_TextPrBold', 'put_TextPrItalic', 'Undo', 'Redo', 'asc_Save']) {
    if (typeof api?.[method] !== 'function') throw new TypeError(`Word engine is missing ${method}`);
  }
  let snapshot = initial;
  let disposed = false;
  const listeners = new Set();
  const callbacks = [];
  const update = changes => {
    if (disposed) return;
    if (Object.entries(changes).every(([key, value]) => snapshot[key] === value)) return;
    snapshot = Object.freeze({...snapshot, ...changes});
    for (const listener of [...listeners]) listener();
  };
  const subscribeEvent = (name, key) => {
    const callback = value => update({[key]: Boolean(value)});
    api.asc_registerCallback(name, callback);
    callbacks.push([name, callback]);
  };
  try {
    for (const [name, key] of [['asc_onCanUndo', 'canUndo'], ['asc_onCanRedo', 'canRedo'], ['asc_onBold', 'bold'], ['asc_onItalic', 'italic']]) {
      subscribeEvent(name, key);
    }
  } catch (error) {
    for (const [name, callback] of callbacks) api.asc_unregisterCallback(name, callback);
    throw error;
  }
  const invoke = (method, args = [], allowed = true) => {
    if (disposed) throw new Error('Word controller is disposed');
    if (!snapshot.ready || snapshot.readOnly || !allowed) return false;
    api[method](...args);
    return true;
  };
  return Object.freeze({
    getSnapshot: () => snapshot,
    getServerSnapshot: () => initial,
    subscribe(listener) {
      if (disposed) throw new Error('Word controller is disposed');
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    // The runtime calls this after document readiness and permission changes.
    // This is UI state, never a replacement for engine/server authorization.
    setSession({ready, readOnly}) {
      if (disposed) throw new Error('Word controller is disposed');
      if (typeof ready !== 'boolean' || typeof readOnly !== 'boolean') throw new TypeError('ready/readOnly must be booleans');
      update({ready, readOnly});
    },
    toggleBold: () => invoke('put_TextPrBold', [!snapshot.bold]),
    toggleItalic: () => invoke('put_TextPrItalic', [!snapshot.italic]),
    undo: () => invoke('Undo', [], snapshot.canUndo),
    redo: () => invoke('Redo', [], snapshot.canRedo),
    save: () => invoke('asc_Save'),
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
