'use client';

import {createElement, useSyncExternalStore} from 'react';
export {createWordController} from './controller.js';

export function useWordState(controller) {
  return useSyncExternalStore(controller.subscribe, controller.getSnapshot, controller.getServerSnapshot);
}

/** A native React toolbar. The engine surface and its lifecycle belong to the runtime. */
export function OfficierToolbar({controller, className, labels = {}}) {
  const state = useWordState(controller);
  const disabled = !state.ready || state.readOnly;
  const button = (key, label, action, extra = {}) => createElement('button', {
    key, type: 'button', disabled, onClick: action,
    // Keep the document selection for pointer-driven commands; keyboard navigation remains available.
    onMouseDown: event => event.preventDefault(), ...extra
  }, labels[key] ?? label);
  return createElement('div', {role: 'group', 'aria-label': labels.toolbar ?? 'Document formatting', className},
    button('bold', 'Bold', controller.toggleBold, {'aria-pressed': state.bold}),
    button('italic', 'Italic', controller.toggleItalic, {'aria-pressed': state.italic}),
    button('undo', 'Undo', controller.undo, {disabled: disabled || !state.canUndo}),
    button('redo', 'Redo', controller.redo, {disabled: disabled || !state.canRedo}),
    button('save', 'Save', controller.save)
  );
}
