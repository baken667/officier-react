'use client';

import {
  createElement,
  type CSSProperties,
  type ReactElement,
  useEffect,
  useRef,
  useState
} from 'react';
import {
  loadOfficierRuntime,
  OfficierError,
  type LoadOfficierRuntimeOptions,
  type MountedOfficierWordEditor,
  type OfficierErrorCode,
  type OfficierRuntime,
  type OfficierWopiSession,
  type OfficierWordRuntimeAdapter
} from './runtime.js';

export type OfficierEditorStatus = 'idle' | 'loading' | 'ready' | 'error';

export interface OfficierEditorReadyEvent {
  readonly runtime: OfficierRuntime;
  readonly editor: MountedOfficierWordEditor;
}

export interface OfficierEditorErrorEvent {
  readonly code: OfficierErrorCode;
  readonly error: OfficierError;
}

export interface OfficierEditorProps {
  readonly documentServerUrl: string | URL;
  readonly session: OfficierWopiSession;
  readonly manifestUrl?: string | URL;
  readonly adapter?: OfficierWordRuntimeAdapter;
  readonly className?: string;
  readonly style?: CSSProperties;
  readonly onReady?: (event: OfficierEditorReadyEvent) => void;
  readonly onError?: (event: OfficierEditorErrorEvent) => void;
  readonly onStatusChange?: (status: OfficierEditorStatus) => void;
}

function toOfficierError(error: unknown): OfficierError {
  if (error instanceof OfficierError) return error;
  return new OfficierError('OFFICIER_RUNTIME_LOAD_FAILED', 'Officier editor failed to start', {cause: error});
}

function notifyStatus(
  status: OfficierEditorStatus,
  setStatus: (status: OfficierEditorStatus) => void,
  onStatusChange: ((status: OfficierEditorStatus) => void) | undefined
): void {
  setStatus(status);
  onStatusChange?.(status);
}

export function OfficierEditor({
  documentServerUrl,
  session,
  manifestUrl,
  adapter,
  className,
  style,
  onReady,
  onError,
  onStatusChange
}: OfficierEditorProps): ReactElement {
  const containerRef = useRef<HTMLDivElement | null>(null);
  const mountedRef = useRef<MountedOfficierWordEditor | null>(null);
  const [status, setStatus] = useState<OfficierEditorStatus>('idle');

  useEffect(() => {
    const container = containerRef.current;
    if (!container) return undefined;

    const controller = new AbortController();
    let disposed = false;

    const start = async () => {
      notifyStatus('loading', setStatus, onStatusChange);

      const options: LoadOfficierRuntimeOptions = {
        documentServerUrl,
        ...(manifestUrl !== undefined ? {manifestUrl} : {}),
        ...(adapter !== undefined ? {adapter} : {}),
        signal: controller.signal,
        document: container.ownerDocument
      };

      try {
        const runtime = await loadOfficierRuntime(options);
        const editor = await runtime.mountWord(container, session, {
          signal: controller.signal,
          onError: error => {
            notifyStatus('error', setStatus, onStatusChange);
            onError?.({code: error.code, error});
          }
        });
        if (disposed) {
          await editor.dispose();
          return;
        }
        mountedRef.current = editor;
        notifyStatus('ready', setStatus, onStatusChange);
        onReady?.({runtime, editor});
      } catch (error) {
        const officierError = toOfficierError(error);
        if (disposed && officierError.code === 'OFFICIER_ABORTED') return;
        notifyStatus('error', setStatus, onStatusChange);
        onError?.({code: officierError.code, error: officierError});
      }
    };

    void start();

    return () => {
      disposed = true;
      controller.abort();
      void mountedRef.current?.dispose();
      mountedRef.current = null;
    };
  }, [adapter, documentServerUrl, manifestUrl, onError, onReady, onStatusChange, session]);

  return createElement('div', {
    ref: containerRef,
    className,
    style,
    role: 'document',
    'aria-busy': status === 'loading' ? true : undefined,
    'data-officier-status': status,
    'data-officier-session': session.id
  });
}
