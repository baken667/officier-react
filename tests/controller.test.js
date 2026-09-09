import {test} from 'node:test';
import assert from 'node:assert/strict';
import {createElement} from 'react';
import {renderToStaticMarkup} from 'react-dom/server';
import {createWordController, OfficierToolbar} from '../src/index.js';

function fixture() {
  const callbacks = new Map();
  const calls = [];
  const api = {
    asc_registerCallback(name, fn) { callbacks.set(name, fn); },
    asc_unregisterCallback(name, fn) { assert.equal(callbacks.get(name), fn); callbacks.delete(name); }
  };
  for (const method of ['put_TextPrBold', 'put_TextPrItalic', 'Undo', 'Redo', 'asc_Save']) {
    api[method] = (...args) => calls.push([method, ...args]);
  }
  return {api, callbacks, calls};
}

test('dispatches directly to initialized engine and reflects selection events', () => {
  const {api, callbacks, calls} = fixture();
  const controller = createWordController(api);
  assert.equal(controller.toggleBold(), false);
  controller.setSession({ready: true, readOnly: false});
  assert.equal(controller.toggleBold(), true);
  callbacks.get('asc_onBold')(true);
  controller.toggleBold();
  assert.deepEqual(calls, [['put_TextPrBold', true], ['put_TextPrBold', false]]);
  assert.equal(controller.undo(), false);
  callbacks.get('asc_onCanUndo')(true);
  controller.undo();
  controller.save();
  assert.deepEqual(calls.slice(-2), [['Undo'], ['asc_Save']]);
});

test('readonly state blocks all mutation dispatch', () => {
  const {api, callbacks, calls} = fixture();
  const controller = createWordController(api);
  callbacks.get('asc_onCanUndo')(true);
  callbacks.get('asc_onCanRedo')(true);
  controller.setSession({ready: true, readOnly: true});
  for (const command of ['toggleBold', 'toggleItalic', 'undo', 'redo', 'save']) assert.equal(controller[command](), false);
  assert.deepEqual(calls, []);
});

test('snapshots remain stable and dispose detaches callbacks without owning the engine', () => {
  const {api, callbacks} = fixture();
  const controller = createWordController(api);
  const snapshot = controller.getSnapshot();
  let changes = 0;
  const unsubscribe = controller.subscribe(() => changes++);
  callbacks.get('asc_onBold')(false);
  assert.equal(controller.getSnapshot(), snapshot);
  assert.equal(changes, 0);
  callbacks.get('asc_onBold')(true);
  assert.equal(changes, 1);
  unsubscribe();
  callbacks.get('asc_onItalic')(true);
  assert.equal(changes, 1);
  controller.dispose();
  controller.dispose();
  assert.equal(callbacks.size, 0);
  assert.throws(() => controller.save(), /disposed/);
  assert.equal(controller.getSnapshot().readOnly, true);
});

test('validates engine compatibility and rolls back incomplete subscriptions', () => {
  assert.throws(() => createWordController({}), /missing/);
  const {api, callbacks} = fixture();
  const register = api.asc_registerCallback;
  api.asc_registerCallback = (name, fn) => {
    if (name === 'asc_onBold') throw new Error('registration failed');
    register(name, fn);
  };
  assert.throws(() => createWordController(api), /registration failed/);
  assert.equal(callbacks.size, 0);
});

test('toolbar supports SSR with disabled initial controls and localized labels', () => {
  const controller = createWordController(fixture().api);
  const html = renderToStaticMarkup(createElement(OfficierToolbar, {controller, labels: {save: 'Сохранить'}}));
  assert.match(html, /Сохранить/);
  assert.equal((html.match(/<button/g) ?? []).length, 5);
  assert.equal((html.match(/disabled=""/g) ?? []).length, 5);
  assert.doesNotMatch(html, /<iframe|<script/);
});
