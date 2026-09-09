import {test} from 'node:test';
import assert from 'node:assert/strict';
import {createElement} from 'react';
import {renderToStaticMarkup} from 'react-dom/server';
import {
  createOfficierWopiSession,
  createMountedWordEditor,
  loadOfficierRuntime,
  OfficierEditor,
  OfficierError
} from '../dist/index.js';

function manifest(overrides = {}) {
  return {
    protocolVersion: 1,
    buildId: 'officier-test',
    sdkVersion: 'sdk-test',
    assets: [],
    capabilities: {word: true, wopi: true, noIframe: true},
    ...overrides
  };
}

function session() {
  return {
    id: 'session-1',
    mode: 'wopi',
    documentTitle: 'Document.docx',
    fileType: 'docx',
    canEdit: true,
    bootstrap: {document: {key: 'doc-1'}}
  };
}

function wordApi() {
  return {
    asc_registerCallback() {},
    asc_unregisterCallback() {},
    put_TextPrBold() {},
    put_TextPrItalic() {},
    Undo() {},
    Redo() {},
    asc_Save() {}
  };
}

function container(defaultView = {}) {
  return {
    ownerDocument: {defaultView},
    querySelector() {
      return null;
    }
  };
}

function fakeDocument(loaded) {
  return {
    head: {
      append(element) {
        loaded.push(element);
        queueMicrotask(() => element.onload?.());
      }
    },
    createElement(tagName) {
      return {
        tagName: tagName.toUpperCase(),
        remove() {}
      };
    }
  };
}

test('loads runtime manifest from document server URL', async () => {
  const requests = [];
  const runtime = await loadOfficierRuntime({
    documentServerUrl: 'https://docs.example/officier/',
    fetch: async url => {
      requests.push(url);
      return Response.json(manifest());
    }
  });

  assert.deepEqual(requests, ['https://docs.example/officier/runtime/manifest.json']);
  assert.equal(runtime.manifest.buildId, 'officier-test');
  assert.equal(runtime.manifest.capabilities.noIframe, true);
});

test('creates WOPI session through the Officier JSON endpoint', async () => {
  const requests = [];
  const created = await createOfficierWopiSession({
    documentServerUrl: 'https://docs.example/officier/',
    wopiSrc: 'https://app.example/wopi/files/1',
    accessToken: 'secret-token',
    accessTokenTtl: 1234,
    userSessionId: 'user-session',
    lang: 'ru',
    docsApiConfig: {editorConfig: {lang: 'ru'}},
    fetch: async (url, init) => {
      requests.push({url, init});
      return Response.json(session());
    }
  });

  assert.equal(created.id, 'session-1');
  assert.equal(
    requests[0].url,
    'https://docs.example/officier/sessions/wopi/word/edit?wopisrc=https%3A%2F%2Fapp.example%2Fwopi%2Ffiles%2F1&usid=user-session&lang=ru'
  );
  assert.equal(requests[0].init.method, 'POST');
  assert.equal(requests[0].init.headers['Content-Type'], 'application/x-www-form-urlencoded');
  assert.equal(requests[0].init.body.get('access_token'), 'secret-token');
  assert.equal(requests[0].init.body.get('access_token_ttl'), '1234');
  assert.equal(requests[0].init.body.get('docs_api_config'), '{"editorConfig":{"lang":"ru"}}');
});

test('validates manifest shape before exposing runtime', async () => {
  await assert.rejects(
    () => loadOfficierRuntime({
      documentServerUrl: 'https://docs.example/officier/',
      fetch: async () => Response.json(manifest({protocolVersion: 2}))
    }),
    error => error instanceof OfficierError && error.code === 'OFFICIER_BAD_MANIFEST'
  );
});

test('loads manifest assets in order against the runtime base URL', async () => {
  const loaded = [];
  await loadOfficierRuntime({
    documentServerUrl: 'https://docs.example/officier/',
    document: fakeDocument(loaded),
    fetch: async () => Response.json(manifest({
      assets: [
        {kind: 'stylesheet', url: 'sdk/editor.css'},
        {kind: 'script', url: 'sdk/runtime.js', crossOrigin: 'anonymous'}
      ]
    }))
  });

  assert.equal(loaded[0].tagName, 'LINK');
  assert.equal(loaded[0].href, 'https://docs.example/officier/sdk/editor.css');
  assert.equal(loaded[1].tagName, 'SCRIPT');
  assert.equal(loaded[1].src, 'https://docs.example/officier/sdk/runtime.js');
  assert.equal(loaded[1].crossOrigin, 'anonymous');
});

test('reports missing direct SDK adapter explicitly', async () => {
  const runtime = await loadOfficierRuntime({
    documentServerUrl: 'https://docs.example/officier/',
    fetch: async () => Response.json(manifest())
  });

  await assert.rejects(
    () => runtime.mountWord(container(), session()),
    error => error instanceof OfficierError && error.code === 'OFFICIER_SDK_ADAPTER_MISSING'
  );
});

test('allows a single mounted word editor and releases the guard on dispose', async () => {
  const runtime = await loadOfficierRuntime({
    documentServerUrl: 'https://docs.example/officier/',
    fetch: async () => Response.json(manifest()),
    adapter: {
      async mountWord() {
        return createMountedWordEditor(wordApi());
      }
    }
  });

  const first = await runtime.mountWord(container(), session());
  await assert.rejects(
    () => runtime.mountWord(container(), session()),
    error => error instanceof OfficierError && error.code === 'OFFICIER_EDITOR_ALREADY_MOUNTED'
  );
  await first.dispose();
  const second = await runtime.mountWord(container(), session());
  await second.dispose();
});

test('OfficierEditor supports SSR without creating iframe or script markup', () => {
  const html = renderToStaticMarkup(createElement(OfficierEditor, {
    documentServerUrl: '/officier/',
    session: session()
  }));

  assert.match(html, /data-officier-status="idle"/);
  assert.doesNotMatch(html, /<iframe|<script/);
});
