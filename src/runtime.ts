import {createWordController, type WordController, type WordEngineApi} from './controller.js';

export type OfficierRuntimeAssetKind = 'script' | 'stylesheet';

export interface OfficierRuntimeAsset {
  readonly kind: OfficierRuntimeAssetKind;
  readonly url: string;
  readonly integrity?: string;
  readonly crossOrigin?: 'anonymous' | 'use-credentials';
}

export interface OfficierRuntimeManifest {
  readonly protocolVersion: 1;
  readonly buildId: string;
  readonly sdkVersion: string;
  readonly assets: readonly OfficierRuntimeAsset[];
  readonly capabilities?: {
    readonly word?: boolean;
    readonly wopi?: boolean;
    readonly noIframe?: boolean;
  };
}

export interface OfficierWopiSession {
  readonly id: string;
  readonly mode: 'wopi';
  readonly documentTitle: string;
  readonly fileType: string;
  readonly canEdit: boolean;
  readonly bootstrap: unknown;
}

export interface MountedOfficierWordEditor {
  readonly controller: WordController;
  dispose(): void | Promise<void>;
}

export interface OfficierWordMountContext {
  readonly container: HTMLElement;
  readonly session: OfficierWopiSession;
  readonly manifest: OfficierRuntimeManifest;
  readonly signal?: AbortSignal;
}

export interface OfficierWordRuntimeAdapter {
  mountWord(context: OfficierWordMountContext): Promise<MountedOfficierWordEditor>;
}

export interface OfficierRuntime {
  readonly manifest: OfficierRuntimeManifest;
  mountWord(
    container: HTMLElement,
    session: OfficierWopiSession,
    options?: {signal?: AbortSignal; onError?: (error: OfficierError) => void}
  ): Promise<MountedOfficierWordEditor>;
}

export type OfficierErrorCode =
  | 'OFFICIER_ABORTED'
  | 'OFFICIER_BAD_MANIFEST'
  | 'OFFICIER_EDITOR_ALREADY_MOUNTED'
  | 'OFFICIER_IFRAME_DETECTED'
  | 'OFFICIER_RUNTIME_LOAD_FAILED'
  | 'OFFICIER_SDK_ADAPTER_MISSING';

export class OfficierError extends Error {
  readonly code: OfficierErrorCode;
  readonly cause?: unknown;

  constructor(code: OfficierErrorCode, message: string, options: {cause?: unknown} = {}) {
    super(message);
    this.name = 'OfficierError';
    this.code = code;
    this.cause = options.cause;
  }
}

export interface LoadOfficierRuntimeOptions {
  readonly documentServerUrl: string | URL;
  readonly manifestUrl?: string | URL;
  readonly fetch?: typeof fetch;
  readonly document?: Document;
  readonly adapter?: OfficierWordRuntimeAdapter;
  readonly signal?: AbortSignal;
}

export interface CreateOfficierWopiSessionOptions {
  readonly documentServerUrl: string | URL;
  readonly wopiSrc: string | URL;
  readonly accessToken: string;
  readonly accessTokenTtl?: number;
  readonly documentType?: 'word';
  readonly mode?: 'edit' | 'view';
  readonly userSessionId?: string;
  readonly lang?: string;
  readonly ui?: string;
  readonly sc?: string;
  readonly docsApiConfig?: unknown;
  readonly fetch?: typeof fetch;
  readonly signal?: AbortSignal;
}

declare global {
  interface Window {
    OfficierDirectRuntime?: OfficierWordRuntimeAdapter;
  }
}

let activeWordMount: MountedOfficierWordEditor | undefined;

function abortIfNeeded(signal: AbortSignal | undefined): void {
  if (signal?.aborted) {
    throw new OfficierError('OFFICIER_ABORTED', 'Officier runtime loading was aborted');
  }
}

function resolveRuntimeUrl(url: string | URL, base: URL): string {
  return new URL(url, base).toString();
}

function runtimeBaseUrl(documentServerUrl: string | URL): URL {
  const base = new URL(documentServerUrl.toString(), globalThis.location?.href ?? 'http://localhost/');
  if (!base.pathname.endsWith('/')) base.pathname = `${base.pathname}/`;
  return base;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function parseAsset(value: unknown, index: number): OfficierRuntimeAsset {
  if (!isRecord(value)) {
    throw new OfficierError('OFFICIER_BAD_MANIFEST', `Runtime asset ${index} must be an object`);
  }

  if (value.kind !== 'script' && value.kind !== 'stylesheet') {
    throw new OfficierError('OFFICIER_BAD_MANIFEST', `Runtime asset ${index} has unsupported kind`);
  }

  if (typeof value.url !== 'string' || value.url.length === 0) {
    throw new OfficierError('OFFICIER_BAD_MANIFEST', `Runtime asset ${index} must have a URL`);
  }

  return {
    kind: value.kind,
    url: value.url,
    ...(typeof value.integrity === 'string' ? {integrity: value.integrity} : {}),
    ...(value.crossOrigin === 'anonymous' || value.crossOrigin === 'use-credentials'
      ? {crossOrigin: value.crossOrigin}
      : {})
  };
}

function parseManifest(value: unknown): OfficierRuntimeManifest {
  if (!isRecord(value)) {
    throw new OfficierError('OFFICIER_BAD_MANIFEST', 'Runtime manifest must be an object');
  }

  if (value.protocolVersion !== 1) {
    throw new OfficierError('OFFICIER_BAD_MANIFEST', 'Runtime manifest protocolVersion must be 1');
  }

  if (typeof value.buildId !== 'string' || value.buildId.length === 0) {
    throw new OfficierError('OFFICIER_BAD_MANIFEST', 'Runtime manifest must have a buildId');
  }

  if (typeof value.sdkVersion !== 'string' || value.sdkVersion.length === 0) {
    throw new OfficierError('OFFICIER_BAD_MANIFEST', 'Runtime manifest must have an sdkVersion');
  }

  if (!Array.isArray(value.assets)) {
    throw new OfficierError('OFFICIER_BAD_MANIFEST', 'Runtime manifest assets must be an array');
  }

  return Object.freeze({
    protocolVersion: 1,
    buildId: value.buildId,
    sdkVersion: value.sdkVersion,
    assets: Object.freeze(value.assets.map(parseAsset)),
    ...(isRecord(value.capabilities) ? {
      capabilities: Object.freeze({
        ...(typeof value.capabilities.word === 'boolean' ? {word: value.capabilities.word} : {}),
        ...(typeof value.capabilities.wopi === 'boolean' ? {wopi: value.capabilities.wopi} : {}),
        ...(typeof value.capabilities.noIframe === 'boolean' ? {noIframe: value.capabilities.noIframe} : {})
      })
    } : {})
  });
}

function parseWopiSession(value: unknown): OfficierWopiSession {
  if (!isRecord(value)) {
    throw new OfficierError('OFFICIER_RUNTIME_LOAD_FAILED', 'WOPI session response must be an object');
  }

  if (typeof value.id !== 'string' || value.id.length === 0) {
    throw new OfficierError('OFFICIER_RUNTIME_LOAD_FAILED', 'WOPI session response must have an id');
  }

  if (value.mode !== 'wopi') {
    throw new OfficierError('OFFICIER_RUNTIME_LOAD_FAILED', 'WOPI session response mode must be wopi');
  }

  if (typeof value.documentTitle !== 'string') {
    throw new OfficierError('OFFICIER_RUNTIME_LOAD_FAILED', 'WOPI session response must have a documentTitle');
  }

  if (typeof value.fileType !== 'string' || value.fileType.length === 0) {
    throw new OfficierError('OFFICIER_RUNTIME_LOAD_FAILED', 'WOPI session response must have a fileType');
  }

  if (typeof value.canEdit !== 'boolean') {
    throw new OfficierError('OFFICIER_RUNTIME_LOAD_FAILED', 'WOPI session response must have canEdit');
  }

  return Object.freeze({
    id: value.id,
    mode: 'wopi',
    documentTitle: value.documentTitle,
    fileType: value.fileType,
    canEdit: value.canEdit,
    bootstrap: value.bootstrap
  });
}

function appendQuery(params: URLSearchParams, key: string, value: string | URL | undefined): void {
  if (value !== undefined) params.set(key, value.toString());
}

export async function createOfficierWopiSession(
  options: CreateOfficierWopiSessionOptions
): Promise<OfficierWopiSession> {
  abortIfNeeded(options.signal);

  const base = runtimeBaseUrl(options.documentServerUrl);
  const documentType = options.documentType ?? 'word';
  const mode = options.mode ?? 'edit';
  const url = new URL(`sessions/wopi/${documentType}/${mode}`, base);
  appendQuery(url.searchParams, 'wopisrc', options.wopiSrc);
  appendQuery(url.searchParams, 'usid', options.userSessionId);
  appendQuery(url.searchParams, 'lang', options.lang);
  appendQuery(url.searchParams, 'ui', options.ui);
  appendQuery(url.searchParams, 'sc', options.sc);

  const body = new URLSearchParams();
  body.set('access_token', options.accessToken);
  if (options.accessTokenTtl !== undefined) body.set('access_token_ttl', String(options.accessTokenTtl));
  if (options.docsApiConfig !== undefined) body.set('docs_api_config', JSON.stringify(options.docsApiConfig));

  const fetchSession = options.fetch ?? globalThis.fetch;
  if (typeof fetchSession !== 'function') {
    throw new OfficierError('OFFICIER_RUNTIME_LOAD_FAILED', 'A fetch implementation is required');
  }

  let response: Response;
  try {
    const requestInit: RequestInit = {
      method: 'POST',
      headers: {'Content-Type': 'application/x-www-form-urlencoded'},
      body,
      ...(options.signal ? {signal: options.signal} : {})
    };
    response = await fetchSession(url.toString(), requestInit);
  } catch (error) {
    throw new OfficierError('OFFICIER_RUNTIME_LOAD_FAILED', 'Failed to create Officier WOPI session', {cause: error});
  }

  if (!response.ok) {
    throw new OfficierError('OFFICIER_RUNTIME_LOAD_FAILED', `Officier WOPI session returned HTTP ${response.status}`);
  }

  return parseWopiSession(await response.json());
}

function loadElement(element: HTMLElement, parent: HTMLElement, signal: AbortSignal | undefined): Promise<void> {
  abortIfNeeded(signal);

  return new Promise((resolve, reject) => {
    const cleanup = () => {
      element.onload = null;
      element.onerror = null;
      signal?.removeEventListener('abort', onAbort);
    };
    const onAbort = () => {
      cleanup();
      element.remove();
      reject(new OfficierError('OFFICIER_ABORTED', 'Officier runtime asset loading was aborted'));
    };

    element.onload = () => {
      cleanup();
      resolve();
    };
    element.onerror = () => {
      cleanup();
      reject(new OfficierError('OFFICIER_RUNTIME_LOAD_FAILED', 'Officier runtime asset failed to load'));
    };
    signal?.addEventListener('abort', onAbort, {once: true});
    parent.append(element);
  });
}

async function loadAsset(
  asset: OfficierRuntimeAsset,
  base: URL,
  ownerDocument: Document,
  signal: AbortSignal | undefined
): Promise<void> {
  const url = resolveRuntimeUrl(asset.url, base);

  if (asset.kind === 'stylesheet') {
    const link = ownerDocument.createElement('link');
    link.rel = 'stylesheet';
    link.href = url;
    if (asset.integrity) link.integrity = asset.integrity;
    if (asset.crossOrigin) link.crossOrigin = asset.crossOrigin;
    await loadElement(link, ownerDocument.head, signal);
    return;
  }

  const script = ownerDocument.createElement('script');
  script.async = false;
  script.src = url;
  if (asset.integrity) script.integrity = asset.integrity;
  if (asset.crossOrigin) script.crossOrigin = asset.crossOrigin;
  await loadElement(script, ownerDocument.head, signal);
}

function assertNoIframe(container: HTMLElement): void {
  if (container.querySelector('iframe')) {
    throw new OfficierError('OFFICIER_IFRAME_DETECTED', 'Officier direct runtime created an iframe');
  }
}

function observeIframes(container: HTMLElement, onError: (error: OfficierError) => void): MutationObserver | undefined {
  const ownerWindow = container.ownerDocument.defaultView;
  if (!ownerWindow?.MutationObserver) return undefined;

  const observer = new ownerWindow.MutationObserver(records => {
    for (const record of records) {
      for (const node of [...record.addedNodes]) {
        if (node instanceof ownerWindow.HTMLIFrameElement) {
          onError(new OfficierError('OFFICIER_IFRAME_DETECTED', 'Officier direct runtime created an iframe'));
          return;
        }
        if (node instanceof ownerWindow.Element && node.querySelector('iframe')) {
          onError(new OfficierError('OFFICIER_IFRAME_DETECTED', 'Officier direct runtime created an iframe'));
          return;
        }
      }
    }
  });
  observer.observe(container, {childList: true, subtree: true});
  return observer;
}

function adapterFromWindow(ownerDocument: Document): OfficierWordRuntimeAdapter | undefined {
  return ownerDocument.defaultView?.OfficierDirectRuntime;
}

export function createMountedWordEditor(api: WordEngineApi): MountedOfficierWordEditor {
  const controller = createWordController(api);
  return {
    controller,
    dispose: () => controller.dispose()
  };
}

export async function loadOfficierRuntime(options: LoadOfficierRuntimeOptions): Promise<OfficierRuntime> {
  abortIfNeeded(options.signal);

  const base = runtimeBaseUrl(options.documentServerUrl);
  const manifestUrl = options.manifestUrl
    ? resolveRuntimeUrl(options.manifestUrl, base)
    : resolveRuntimeUrl('runtime/manifest.json', base);
  const fetchManifest = options.fetch ?? globalThis.fetch;

  if (typeof fetchManifest !== 'function') {
    throw new OfficierError('OFFICIER_RUNTIME_LOAD_FAILED', 'A fetch implementation is required');
  }

  let response: Response;
  try {
    const requestInit: RequestInit = options.signal ? {signal: options.signal} : {};
    response = await fetchManifest(manifestUrl, requestInit);
  } catch (error) {
    throw new OfficierError('OFFICIER_RUNTIME_LOAD_FAILED', 'Failed to fetch Officier runtime manifest', {cause: error});
  }

  if (!response.ok) {
    throw new OfficierError('OFFICIER_RUNTIME_LOAD_FAILED', `Officier runtime manifest returned HTTP ${response.status}`);
  }

  const manifest = parseManifest(await response.json());
  const ownerDocument = options.document ?? globalThis.document;

  if (manifest.assets.length > 0 && !ownerDocument) {
    throw new OfficierError('OFFICIER_RUNTIME_LOAD_FAILED', 'A document is required to load runtime assets');
  }

  for (const asset of manifest.assets) {
    abortIfNeeded(options.signal);
    await loadAsset(asset, base, ownerDocument, options.signal);
  }

  const runtime: OfficierRuntime = Object.freeze({
    manifest,
    async mountWord(
      container: HTMLElement,
      session: OfficierWopiSession,
      mountOptions: {signal?: AbortSignal; onError?: (error: OfficierError) => void} = {}
    ) {
      const signal = mountOptions.signal ?? options.signal;
      abortIfNeeded(signal);

      if (activeWordMount) {
        throw new OfficierError('OFFICIER_EDITOR_ALREADY_MOUNTED', 'Only one Officier word editor can be mounted per window');
      }

      assertNoIframe(container);
      const adapter = options.adapter ?? adapterFromWindow(container.ownerDocument);
      if (!adapter) {
        throw new OfficierError(
          'OFFICIER_SDK_ADAPTER_MISSING',
          'Officier direct SDK adapter is not available in loaded runtime assets'
        );
      }

      let mounted: MountedOfficierWordEditor | undefined;
      let mountedWrapper: MountedOfficierWordEditor | undefined;
      const observer = observeIframes(container, error => {
        void mountedWrapper?.dispose();
        mountOptions.onError?.(error);
      });

      try {
        mounted = await adapter.mountWord({
          container,
          session,
          manifest,
          ...(signal ? {signal} : {})
        });
        assertNoIframe(container);
        mountedWrapper = {
          controller: mounted.controller,
          async dispose() {
            observer?.disconnect();
            if (activeWordMount === mountedWrapper) activeWordMount = undefined;
            await mounted?.dispose();
            assertNoIframe(container);
          }
        };
        activeWordMount = mountedWrapper;
        return mountedWrapper;
      } catch (error) {
        observer?.disconnect();
        if (activeWordMount === mountedWrapper) activeWordMount = undefined;
        throw error;
      }
    }
  });
  return runtime;
}
