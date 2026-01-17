/**
 * OutputsTab Component
 *
 * Visual editor for managing output files with Monaco Editor.
 * Uses accordion-style list with expand/collapse (multiple expandable).
 * Acts as the master for outputs - changes auto-update Template JSON.
 */

import React, { useState, useEffect, useCallback, useMemo } from 'react';
import {
  Button,
  Text,
  makeStyles,
  tokens,
  Accordion,
  AccordionItem,
  AccordionHeader,
  AccordionPanel,
  MessageBar,
  MessageBarBody,
} from '@fluentui/react-components';
import { Add24Regular, DocumentText24Regular } from '@fluentui/react-icons';
import { OutputEditor } from './OutputEditor';
import { generateStarterTemplate } from '../utils/StarterTemplates';
import { updateAutocompleteData } from '../utils/monacoConfig';
import type { AddInTemplate, OutputConfig } from '@/types';
import type { TableInfo } from '@/services';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
  },
  title: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
  },
  accordion: {
    width: '100%',
  },
  emptyState: {
    padding: tokens.spacingVerticalXXL,
    textAlign: 'center',
    color: tokens.colorNeutralForeground3,
  },
  saveIndicator: {
    fontSize: '12px',
    color: tokens.colorNeutralForeground3,
  },
});

export interface OutputsTabProps {
  /**
   * The current template
   */
  template: AddInTemplate | null;

  /**
   * Callback when template changes
   */
  onTemplateChange: (template: AddInTemplate | null) => void;

  /**
   * Available tables for starter template generation
   */
  tables: TableInfo[];
}

/**
 * Outputs tab with accordion-style visual editor
 */
export const OutputsTab: React.FC<OutputsTabProps> = ({
  template,
  onTemplateChange,
  tables,
}) => {
  const styles = useStyles();
  const [localOutputs, setLocalOutputs] = useState<Record<string, OutputConfig>>({});
  const [saveStatus, setSaveStatus] = useState<'saved' | 'saving' | 'idle'>('idle');
  const [nextOutputNumber, setNextOutputNumber] = useState(1);
  const [openItems, setOpenItems] = useState<string[]>([]);
  const didInitOpenRef = React.useRef(false);
  const [pendingDeleteKey, setPendingDeleteKey] = useState<string | null>(null);
  const [deleteError, setDeleteError] = useState<string | null>(null);

  // Sync from template to local state
  useEffect(() => {
    if (template?.outputs) {
      setLocalOutputs(template.outputs);

      // Calculate next output number
      const outputKeys = Object.keys(template.outputs);
      const numbers = outputKeys
        .filter(key => key.startsWith('output'))
        .map(key => parseInt(key.replace('output', ''), 10))
        .filter(num => !isNaN(num));
      const maxNumber = numbers.length > 0 ? Math.max(...numbers) : 0;
      setNextOutputNumber(maxNumber + 1);
    }
  }, [template]);

  // Update autocomplete data when tables or template change
  useEffect(() => {
    // Build table autocomplete data from template.tables (these are the keys used in templates)
    const tableAutocomplete: { name: string; key: string; columns: string[] }[] = [];

    if (template?.tables) {
      template.tables.forEach((config) => {
        // Find the actual table to get columns
        const actualTable = tables.find(t => t.name === config.name);
        tableAutocomplete.push({
          name: config.name,
          key: config.name,
          columns: actualTable?.columns || [],
        });
      });
    }

    // Also add tables from Excel that aren't in template yet (for convenience)
    tables.forEach(table => {
      const alreadyAdded = tableAutocomplete.some(t => t.name === table.name);
      if (!alreadyAdded) {
        tableAutocomplete.push({
          name: table.name,
          key: table.name,
          columns: table.columns,
        });
      }
    });

    // Build globals autocomplete data from template
    const globalsAutocomplete = template?.globals
      ? template.globals.map((config) => ({
          name: config.name,
          key: config.name,
        }))
      : [];

    console.log('OutputsTab: Updating autocomplete data', {
      tables: tableAutocomplete,
      globals: globalsAutocomplete,
    });

    updateAutocompleteData(tableAutocomplete, globalsAutocomplete);
  }, [tables, template]);

  // Debounced update to parent (500ms)
  useEffect(() => {
    if (!template) return;

    const timeoutId = setTimeout(() => {
      if (JSON.stringify(localOutputs) !== JSON.stringify(template.outputs)) {
        setSaveStatus('saving');
        onTemplateChange({ ...template, outputs: localOutputs });

        // Show "saved" status briefly
        setTimeout(() => {
          setSaveStatus('saved');
          setTimeout(() => setSaveStatus('idle'), 1500);
        }, 200);
      }
    }, 500);

    return () => clearTimeout(timeoutId);
  }, [localOutputs, template, onTemplateChange]);

  /**
   * Add a new output
   */
  const handleAddOutput = useCallback(() => {
    const newKey = `output${nextOutputNumber}`;
    const defaultFilename = `output${nextOutputNumber}.json`;

    // Generate starter template based on filename extension
    const starterTemplate = generateStarterTemplate(defaultFilename, tables);

    const newOutput: OutputConfig = {
      filename: defaultFilename,
      template: starterTemplate,
      description: '',
    };

    setLocalOutputs(prev => ({
      ...prev,
      [newKey]: newOutput,
    }));
    setNextOutputNumber(prev => prev + 1);
  }, [nextOutputNumber, tables]);

  /**
   * Handle output key change (rename)
   */
  const handleKeyChange = useCallback((oldKey: string, newKey: string) => {
    // Validate new key
    if (!newKey || newKey === oldKey) return;

    // Check for duplicates
    if (localOutputs[newKey]) {
      return;
    }

    setLocalOutputs(prev => {
      const updated = { ...prev };
      updated[newKey] = updated[oldKey];
      delete updated[oldKey];
      return updated;
    });
  }, [localOutputs]);

  /**
   * Handle output configuration change
   */
  const handleOutputChange = useCallback((key: string, output: OutputConfig) => {
    setLocalOutputs(prev => ({
      ...prev,
      [key]: output,
    }));
  }, []);

  /**
   * Handle output deletion
   */
  const handleDelete = useCallback((key: string) => {
    if (Object.keys(localOutputs).length === 1) {
      setDeleteError('Cannot delete the last output. At least one output is required.');
      return;
    }

    setDeleteError(null);
    setPendingDeleteKey(key);
  }, [localOutputs]);

  const confirmDelete = useCallback(() => {
    if (!pendingDeleteKey) return;
    setLocalOutputs(prev => {
      const updated = { ...prev };
      delete updated[pendingDeleteKey];
      return updated;
    });
    setPendingDeleteKey(null);
    setDeleteError(null);
  }, [pendingDeleteKey]);

  const cancelDelete = useCallback(() => {
    setPendingDeleteKey(null);
  }, []);

  if (!template) {
    return (
      <div className={styles.container}>
        <MessageBar intent="info">
          <MessageBarBody>
            Please load or create a template first in the Template tab.
          </MessageBarBody>
        </MessageBar>
      </div>
    );
  }

  const outputKeys = useMemo(() => Object.keys(localOutputs), [localOutputs]);
  const hasOutputs = outputKeys.length > 0;

  useEffect(() => {
    if (!hasOutputs) {
      if (openItems.length !== 0) {
        setOpenItems([]);
      }
      didInitOpenRef.current = false;
      return;
    }

    setOpenItems(prev => {
      const filtered = prev.filter(key => outputKeys.includes(key));
      if (filtered.length !== prev.length) {
        return filtered;
      }
      if (!didInitOpenRef.current) {
        didInitOpenRef.current = true;
        return [outputKeys[0]];
      }
      return prev;
    });
  }, [hasOutputs, outputKeys, openItems.length]);

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <Text className={styles.title}>Output Files</Text>
        <div style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalM }}>
          {saveStatus === 'saving' && (
            <Text className={styles.saveIndicator}>Saving...</Text>
          )}
          {saveStatus === 'saved' && (
            <Text className={styles.saveIndicator} style={{ color: tokens.colorPaletteGreenForeground1 }}>
              ✓ Saved
            </Text>
          )}
          <Button
            appearance="primary"
            icon={<Add24Regular />}
            onClick={handleAddOutput}
          >
            Add Output
          </Button>
        </div>
      </div>

      {pendingDeleteKey && (
        <MessageBar intent="warning">
          <MessageBarBody>
            <div style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalM }}>
              <span>Delete output "{pendingDeleteKey}"?</span>
              <Button appearance="primary" size="small" onClick={confirmDelete}>
                Delete
              </Button>
              <Button appearance="secondary" size="small" onClick={cancelDelete}>
                Cancel
              </Button>
            </div>
          </MessageBarBody>
        </MessageBar>
      )}
      {deleteError && (
        <MessageBar intent="error">
          <MessageBarBody>{deleteError}</MessageBarBody>
        </MessageBar>
      )}

      {!hasOutputs && (
        <div className={styles.emptyState}>
          <DocumentText24Regular fontSize={48} />
          <Text size={400} style={{ display: 'block', marginTop: tokens.spacingVerticalM }}>
            No outputs defined
          </Text>
          <Text size={200} style={{ display: 'block', marginTop: tokens.spacingVerticalS }}>
            Click "Add Output" to create your first output file
          </Text>
        </div>
      )}

      {hasOutputs && (
        <Accordion
          className={styles.accordion}
          multiple
          collapsible
          openItems={openItems}
          onToggle={(_, data) => setOpenItems(data.openItems as string[])}
        >
          {outputKeys.map((key) => (
            <AccordionItem key={key} value={key}>
              <AccordionHeader>
                <div style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS }}>
                  <DocumentText24Regular fontSize={20} />
                  <Text weight="semibold">{key}</Text>
                  <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
                    → {localOutputs[key].filename}
                  </Text>
                </div>
              </AccordionHeader>
              <AccordionPanel>
                <OutputEditor
                  outputKey={key}
                  output={localOutputs[key]}
                  onKeyChange={handleKeyChange}
                  existingKeys={outputKeys}
                  onChange={handleOutputChange}
                  onDelete={handleDelete}
                  isOnlyOutput={outputKeys.length === 1}
                />
              </AccordionPanel>
            </AccordionItem>
          ))}
        </Accordion>
      )}
    </div>
  );
};

export default OutputsTab;
