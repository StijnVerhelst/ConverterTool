/**
 * MonacoEditor Component
 *
 * A wrapper around @monaco-editor/react that provides:
 * - Local loading (no CDN) for Office Add-in compatibility
 * - Nunjucks syntax highlighting and autocomplete
 * - Fallback to textarea if Monaco fails to load
 * - Configurable read-only mode
 */

import React, { useState, useCallback, useRef, useEffect } from 'react';
import Editor, { loader } from '@monaco-editor/react';
import * as monaco from 'monaco-editor';
import { Textarea, Spinner, makeStyles, tokens } from '@fluentui/react-components';
import {
  defaultEditorOptions,
  configureMonacoForNunjucks,
  applyNunjucksDecorations,
  disableJsonValidation,
  applyNunjucksDiagnostics,
  subscribeToNunjucksDiagnostics,
} from '../utils/monacoConfig';

// Configure Monaco to load from local node_modules instead of CDN
// This is required for Office Add-ins due to Content Security Policy
loader.config({ monaco });

// Flag to ensure we only configure Monaco once
let monacoConfigured = false;

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    width: '100%',
  },
  loadingContainer: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    padding: tokens.spacingVerticalXXL,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  editorWrapper: {
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: tokens.borderRadiusMedium,
    overflow: 'hidden',
  },
  fallbackTextarea: {
    fontFamily: 'Consolas, "Courier New", monospace',
    fontSize: '12px',
  },
});

// Add CSS for Nunjucks syntax highlighting
const nunjucksStyles = document.createElement('style');
nunjucksStyles.textContent = `
  /* Nunjucks expression {{ }} - purple */
  .nunjucks-expression-inline {
    background-color: rgba(153, 50, 204, 0.15) !important;
    border-radius: 3px;
    padding: 1px 0;
  }
  .nunjucks-expression-line {
    background-color: rgba(153, 50, 204, 0.05);
  }

  /* Nunjucks statement {% %} - blue */
  .nunjucks-statement-inline {
    background-color: rgba(0, 100, 255, 0.15) !important;
    border-radius: 3px;
    padding: 1px 0;
  }
  .nunjucks-statement-line {
    background-color: rgba(0, 100, 255, 0.05);
  }

  /* Nunjucks comment {# #} - gray */
  .nunjucks-comment-inline {
    background-color: rgba(128, 128, 128, 0.2) !important;
    font-style: italic;
    border-radius: 3px;
    padding: 1px 0;
  }
  .nunjucks-comment-line {
    background-color: rgba(128, 128, 128, 0.05);
  }
`;
if (!document.head.querySelector('style[data-monaco-nunjucks]')) {
  nunjucksStyles.setAttribute('data-monaco-nunjucks', 'true');
  document.head.appendChild(nunjucksStyles);
}

export interface MonacoEditorProps {
  /**
   * The content to display in the editor
   */
  value: string;

  /**
   * Callback when content changes (not called in read-only mode)
   */
  onChange?: (value: string) => void;

  /**
   * Monaco language mode (e.g., 'json', 'html', 'xml', 'plaintext')
   */
  language?: string;

  /**
   * Whether the editor is read-only
   */
  readOnly?: boolean;

  /**
   * Height of the editor (CSS value)
   * @default '400px'
   */
  height?: string;

  /**
   * Additional Monaco editor options
   */
  options?: any;

  /**
   * Whether to enable Nunjucks syntax highlighting
   * @default true
   */
  enableNunjucks?: boolean;
}

/**
 * Monaco Editor wrapper component with Nunjucks support and fallback to textarea
 */
export const MonacoEditor: React.FC<MonacoEditorProps> = ({
  value,
  onChange,
  language = 'plaintext',
  readOnly = false,
  height = '400px',
  options = {},
  enableNunjucks = true,
}) => {
  const styles = useStyles();
  const [loadFailed, setLoadFailed] = useState(false);
  const editorRef = useRef<any>(null);
  const monacoRef = useRef<any>(null);
  const diagnosticsUnsubscribeRef = useRef<null | (() => void)>(null);
  const documentTabHandlerRef = useRef<null | ((ev: KeyboardEvent) => void)>(null);
  const tabHandlerActiveRef = useRef(false);
  const tabFocusDisposeRef = useRef<null | { dispose: () => void }>(null);
  const tabBlurDisposeRef = useRef<null | { dispose: () => void }>(null);
  const tabKeyDisposeRef = useRef<null | { dispose: () => void }>(null);

  /**
   * Handle Monaco editor mount
   */
  const handleEditorDidMount = useCallback((editor: any, monacoInstance: any) => {
    editorRef.current = editor;
    monacoRef.current = monacoInstance;

    if (tabKeyDisposeRef.current) {
      tabKeyDisposeRef.current.dispose();
    }
    tabKeyDisposeRef.current = editor.onKeyDown((ev: any) => {
      if (ev.keyCode !== monacoInstance.KeyCode.Tab) {
        return;
      }
      ev.preventDefault();
      ev.stopPropagation();
      if (ev.shiftKey) {
        editor.trigger('keyboard', 'outdent', {});
      } else {
        editor.trigger('keyboard', 'tab', {});
      }
    });
    editor.updateOptions({ tabFocusMode: false });

    const handler = (ev: KeyboardEvent) => {
      if (ev.key !== 'Tab') {
        return;
      }
      if (!editor.hasTextFocus?.()) {
        return;
      }
      const domNode = editor.getDomNode();
      if (domNode && ev.target instanceof Node && !domNode.contains(ev.target)) {
        return;
      }
      ev.preventDefault();
      ev.stopPropagation();
      if (ev.shiftKey) {
        editor.trigger('keyboard', 'outdent', {});
      } else {
        editor.trigger('keyboard', 'tab', {});
      }
    };
    documentTabHandlerRef.current = handler;
    const enableHandler = () => {
      if (tabHandlerActiveRef.current || !documentTabHandlerRef.current) {
        return;
      }
      document.addEventListener('keydown', documentTabHandlerRef.current, true);
      tabHandlerActiveRef.current = true;
    };
    const disableHandler = () => {
      if (!tabHandlerActiveRef.current || !documentTabHandlerRef.current) {
        return;
      }
      document.removeEventListener('keydown', documentTabHandlerRef.current, true);
      tabHandlerActiveRef.current = false;
    };
    if (tabFocusDisposeRef.current) {
      tabFocusDisposeRef.current.dispose();
    }
    if (tabBlurDisposeRef.current) {
      tabBlurDisposeRef.current.dispose();
    }
    tabFocusDisposeRef.current = editor.onDidFocusEditorText(enableHandler);
    tabBlurDisposeRef.current = editor.onDidBlurEditorText(disableHandler);

    // Configure Nunjucks support once
    if (enableNunjucks && !monacoConfigured) {
      configureMonacoForNunjucks(monacoInstance);
      monacoConfigured = true;
    }

    // Disable JSON validation for Nunjucks templates
    if (enableNunjucks) {
      disableJsonValidation(monacoInstance);
    }

    // Apply initial decorations
    if (enableNunjucks) {
      // Small delay to ensure editor is fully rendered
      setTimeout(() => {
        applyNunjucksDecorations(editor, monacoInstance);
        applyNunjucksDiagnostics(editor, monacoInstance);
      }, 100);
      const runDiagnostics = () => {
        if (editorRef.current && monacoRef.current) {
          applyNunjucksDiagnostics(editorRef.current, monacoRef.current);
        }
      };
      if (diagnosticsUnsubscribeRef.current) {
        diagnosticsUnsubscribeRef.current();
      }
      diagnosticsUnsubscribeRef.current = subscribeToNunjucksDiagnostics(runDiagnostics);
    }
  }, [enableNunjucks]);

  // Reapply decorations when value changes from props
  useEffect(() => {
    if (enableNunjucks && editorRef.current && monacoRef.current) {
      // Small delay to ensure the model is updated
      const timeoutId = setTimeout(() => {
        applyNunjucksDecorations(editorRef.current, monacoRef.current);
        applyNunjucksDiagnostics(editorRef.current, monacoRef.current);
      }, 50);
      return () => clearTimeout(timeoutId);
    }
  }, [value, enableNunjucks]);

  useEffect(() => {
    return () => {
      if (diagnosticsUnsubscribeRef.current) {
        diagnosticsUnsubscribeRef.current();
        diagnosticsUnsubscribeRef.current = null;
      }
      if (tabFocusDisposeRef.current) {
        tabFocusDisposeRef.current.dispose();
        tabFocusDisposeRef.current = null;
      }
      if (tabBlurDisposeRef.current) {
        tabBlurDisposeRef.current.dispose();
        tabBlurDisposeRef.current = null;
      }
      if (tabKeyDisposeRef.current) {
        tabKeyDisposeRef.current.dispose();
        tabKeyDisposeRef.current = null;
      }
      if (documentTabHandlerRef.current && tabHandlerActiveRef.current) {
        document.removeEventListener('keydown', documentTabHandlerRef.current, true);
        tabHandlerActiveRef.current = false;
      }
      documentTabHandlerRef.current = null;
    };
  }, []);

  /**
   * Handle Monaco editor change
   */
  const handleEditorChange = useCallback((newValue: string | undefined) => {
    if (onChange && newValue !== undefined) {
      onChange(newValue);
    }

    // Update decorations on change
    if (enableNunjucks && editorRef.current && monacoRef.current) {
      applyNunjucksDecorations(editorRef.current, monacoRef.current);
      applyNunjucksDiagnostics(editorRef.current, monacoRef.current);
    }
  }, [onChange, enableNunjucks]);

  /**
   * Handle textarea change (fallback mode)
   */
  const handleTextareaChange = useCallback((ev: React.ChangeEvent<HTMLTextAreaElement>) => {
    if (onChange) {
      onChange(ev.target.value);
    }
  }, [onChange]);

  // Merge editor options
  const editorOptions = {
    ...defaultEditorOptions,
    readOnly,
    ...options,
  };

  // Fallback to textarea if Monaco failed to load
  if (loadFailed) {
    return (
      <div className={styles.container}>
        <Textarea
          value={value}
          onChange={handleTextareaChange}
          readOnly={readOnly}
          className={styles.fallbackTextarea}
          resize="vertical"
          style={{ minHeight: height, fontFamily: 'Consolas, "Courier New", monospace' }}
          aria-label="Code editor (fallback)"
        />
      </div>
    );
  }

  // Render Monaco Editor
  return (
    <div className={styles.container}>
      <div className={styles.editorWrapper} style={{ height }}>
        <Editor
          height={height}
          language={language}
          value={value}
          onChange={handleEditorChange}
          onMount={handleEditorDidMount}
          options={editorOptions}
          loading={
            <div className={styles.loadingContainer} style={{ height }}>
              <Spinner label="Loading editor..." />
            </div>
          }
          theme="vs"
        />
      </div>
    </div>
  );
};

export default MonacoEditor;
