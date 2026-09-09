'use client';

import {createElement, type MouseEvent, type ReactElement, useSyncExternalStore} from 'react';
import type {WordController, WordState} from './controller.js';

export {createWordController} from './controller.js';
export type {WordController, WordEngineApi, WordState} from './controller.js';
export {OfficierEditor} from './editor.js';
export type {
  OfficierEditorErrorEvent,
  OfficierEditorProps,
  OfficierEditorReadyEvent,
  OfficierEditorStatus
} from './editor.js';
export {
  createOfficierWopiSession,
  createMountedWordEditor,
  loadOfficierRuntime,
  OfficierError
} from './runtime.js';
export type {
  CreateOfficierWopiSessionOptions,
  LoadOfficierRuntimeOptions,
  MountedOfficierWordEditor,
  OfficierErrorCode,
  OfficierRuntime,
  OfficierRuntimeAsset,
  OfficierRuntimeAssetKind,
  OfficierRuntimeManifest,
  OfficierWopiSession,
  OfficierWordMountContext,
  OfficierWordRuntimeAdapter
} from './runtime.js';

export interface OfficierToolbarLabels {
  toolbar: string;
  bold: string;
  italic: string;
  undo: string;
  redo: string;
  save: string;
}

export interface OfficierToolbarProps {
  controller: WordController;
  className?: string;
  labels?: Partial<OfficierToolbarLabels>;
}

export function useWordState(controller: WordController): WordState {
  return useSyncExternalStore(
    controller.subscribe,
    controller.getSnapshot,
    controller.getServerSnapshot
  );
}

/** A native React toolbar. The engine surface and its lifecycle belong to the runtime. */
export function OfficierToolbar({
  controller,
  className,
  labels = {}
}: OfficierToolbarProps): ReactElement {
  const state = useWordState(controller);
  const disabled = !state.ready || state.readOnly;
  const keepDocumentSelection = (event: MouseEvent<HTMLButtonElement>) => event.preventDefault();
  const button = (
    key: keyof Omit<OfficierToolbarLabels, 'toolbar'>,
    label: string,
    action: () => boolean,
    extra: Record<string, unknown> = {}
  ) => createElement('button', {
    key,
    type: 'button',
    disabled,
    onClick: action,
    onMouseDown: keepDocumentSelection,
    ...extra
  }, labels[key] ?? label);

  return createElement('div', {
    role: 'group',
    'aria-label': labels.toolbar ?? 'Document formatting',
    className
  },
    button('bold', 'Bold', controller.toggleBold, {'aria-pressed': state.bold}),
    button('italic', 'Italic', controller.toggleItalic, {'aria-pressed': state.italic}),
    button('undo', 'Undo', controller.undo, {disabled: disabled || !state.canUndo}),
    button('redo', 'Redo', controller.redo, {disabled: disabled || !state.canRedo}),
    button('save', 'Save', controller.save)
  );
}
