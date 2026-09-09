import type {ReactElement} from 'react';
import type {WordController, WordState} from './controller.js';
export * from './controller.js';
export function useWordState(controller: WordController): WordState;
export function OfficierToolbar(props: {
  controller: WordController;
  className?: string;
  labels?: Partial<Record<'toolbar' | 'bold' | 'italic' | 'undo' | 'redo' | 'save', string>>;
}): ReactElement;
