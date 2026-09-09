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
  /** True means the command was dispatched, not that server persistence has completed. */
  save(): boolean;
  /** Unsubscribes this bridge only. Does not destroy the SDK or disconnect its sockets. */
  dispose(): void;
}
export function createWordController(api: WordEngineApi): WordController;
