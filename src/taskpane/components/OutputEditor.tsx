/**
 * OutputEditor Component
 *
 * Edits a single output configuration with name, filename, and template.
 * Uses Monaco Editor for template editing with syntax highlighting.
 */

import React, { useCallback, useEffect, useState } from 'react';
import {
  Input,
  Label,
  Button,
  makeStyles,
  tokens,
  Field,
} from '@fluentui/react-components';
import { Delete24Regular } from '@fluentui/react-icons';
import { MonacoEditor } from './MonacoEditor';
import { detectLanguageFromFilename } from '../utils/monacoConfig';
import type { OutputConfig } from '@/types';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
    padding: tokens.spacingVerticalM,
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
  },
  fields: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
  },
  infoText: {
    fontSize: '11px',
    color: tokens.colorNeutralForeground3,
    marginTop: '2px',
  },
});

export interface OutputEditorProps {
  /**
   * The output key (e.g., "output1", "report")
   */
  outputKey: string;

  /**
   * The output configuration
   */
  output: OutputConfig;

  /**
   * Callback when the output key changes (rename)
   */
  onKeyChange: (oldKey: string, newKey: string) => void;

  /**
   * Callback when the output configuration changes
   */
  onChange: (key: string, output: OutputConfig) => void;

  /**
   * Callback when delete is requested
   */
  onDelete: (key: string) => void;

  /**
   * Whether this is the only output (disable delete)
   */
  isOnlyOutput?: boolean;

  /**
   * Existing output keys for validation
   */
  existingKeys?: string[];
}

/**
 * Single output editor with name, filename, and template fields
 */
export const OutputEditor: React.FC<OutputEditorProps> = ({
  outputKey,
  output,
  onKeyChange,
  onChange,
  onDelete,
  isOnlyOutput = false,
  existingKeys = [],
}) => {
  const styles = useStyles();
  const [draftKey, setDraftKey] = useState(outputKey);
  const [keyError, setKeyError] = useState<string | null>(null);

  useEffect(() => {
    setDraftKey(outputKey);
    setKeyError(null);
  }, [outputKey]);

  /**
   * Handle output key change
   */
  const handleKeyChange = useCallback((newKey: string) => {
    // Sanitize key: lowercase, alphanumeric and underscores only
    const sanitized = newKey.toLowerCase().replace(/[^a-z0-9_]/g, '_');
    if (!sanitized) {
      setKeyError('Output name is required.');
      return;
    }
    if (sanitized !== outputKey && existingKeys.includes(sanitized)) {
      setKeyError(`Output "${sanitized}" already exists.`);
      return;
    }
    setKeyError(null);
    if (sanitized !== outputKey) {
      onKeyChange(outputKey, sanitized);
    }
  }, [outputKey, onKeyChange, existingKeys]);

  const handleKeyInputChange = useCallback((newKey: string) => {
    setDraftKey(newKey);
    if (keyError) {
      setKeyError(null);
    }
  }, [keyError]);

  const handleKeyInputBlur = useCallback(() => {
    handleKeyChange(draftKey);
  }, [handleKeyChange, draftKey]);

  const handleKeyInputKeyDown = useCallback((ev: React.KeyboardEvent<HTMLInputElement>) => {
    if (ev.key === 'Enter') {
      handleKeyChange(draftKey);
      ev.currentTarget.blur();
    } else if (ev.key === 'Escape') {
      setDraftKey(outputKey);
      setKeyError(null);
      ev.currentTarget.blur();
    }
  }, [draftKey, handleKeyChange, outputKey]);

  /**
   * Handle filename change
   */
  const handleFilenameChange = useCallback((filename: string) => {
    onChange(outputKey, { ...output, filename });
  }, [outputKey, output, onChange]);

  /**
   * Handle template change
   */
  const handleTemplateChange = useCallback((template: string) => {
    onChange(outputKey, { ...output, template });
  }, [outputKey, output, onChange]);

  /**
   * Handle delete
   */
  const handleDelete = useCallback(() => {
    onDelete(outputKey);
  }, [outputKey, onDelete]);

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <Field
          label="Output Name"
          required
          style={{ flex: 1 }}
          validationMessage={keyError ?? undefined}
          validationState={keyError ? 'error' : 'none'}
        >
          <Input
            value={draftKey}
            onChange={(_, data) => handleKeyInputChange(data.value)}
            onBlur={handleKeyInputBlur}
            onKeyDown={handleKeyInputKeyDown}
            placeholder="e.g., report, export1"
          />
        </Field>
        <Button
          appearance="subtle"
          icon={<Delete24Regular />}
          onClick={handleDelete}
          disabled={isOnlyOutput}
          style={{ marginTop: '22px' }}
        >
          Delete
        </Button>
      </div>

      <div className={styles.fields}>
        <Field label="Filename" required>
          <Input
            value={output.filename}
            onChange={(_, data) => handleFilenameChange(data.value)}
            placeholder="e.g., output.json, report_{{ currentdatetime }}.csv"
          />
          <div className={styles.infoText}>
            Supports placeholders: {`{{ currentdatetime }}`} (YYYY-MM-DD_HH-mm-ss), {`{{ globals.variableName }}`}
          </div>
        </Field>

        <Field label="Template">
          <MonacoEditor
            value={output.template}
            onChange={handleTemplateChange}
            language={detectLanguageFromFilename(output.filename)}
            readOnly={false}
            height="300px"
          />
        </Field>
      </div>
    </div>
  );
};

export default OutputEditor;
