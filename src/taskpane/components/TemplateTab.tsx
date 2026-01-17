/**
 * TemplateTab Component
 *
 * Allows importing, editing, and exporting template JSON files.
 */

import { useState, useCallback, useRef, useEffect } from "react";
import {
  Button,
  Text,
  Textarea,
  Card,
  makeStyles,
  tokens,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
} from "@fluentui/react-components";
import {
  ArrowUpload24Regular,
  ArrowDownload24Regular,
  DocumentAdd24Regular,
  Checkmark24Regular,
} from "@fluentui/react-icons";

import { dataOrchestrator } from "@/core";
import type { AddInTemplate } from "@/types";
import type { TableInfo } from "@/services";
import { MonacoEditor } from "./MonacoEditor";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
  },
  buttonRow: {
    display: "flex",
    gap: "8px",
    flexWrap: "wrap",
  },
  editorCard: {
    padding: "12px",
  },
  textarea: {
    width: "100%",
    minHeight: "300px",
    fontFamily: "Consolas, Monaco, 'Courier New', monospace",
    fontSize: "12px",
  },
  hiddenInput: {
    display: "none",
  },
  statusBar: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "8px 12px",
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: "4px",
  },
});

interface TemplateTabProps {
  template: AddInTemplate | null;
  tables: TableInfo[];
  onTemplateChange: (template: AddInTemplate | null) => void;
}

export function TemplateTab({ template, tables, onTemplateChange }: TemplateTabProps) {
  const styles = useStyles();
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [jsonText, setJsonText] = useState<string>(
    template ? JSON.stringify(template, null, 2) : ""
  );
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState<string | null>(null);

  // Sync jsonText when template changes from parent (e.g., from Outputs tab)
  useEffect(() => {
    if (template) {
      const newText = JSON.stringify(template, null, 2);
      // Only update if different to avoid cursor jumping
      if (newText !== jsonText) {
        setJsonText(newText);
      }
    }
  }, [template]);

  const handleImportClick = useCallback(() => {
    fileInputRef.current?.click();
  }, []);

  const handleFileChange = useCallback(
    (event: React.ChangeEvent<HTMLInputElement>) => {
      const file = event.target.files?.[0];
      if (!file) return;

      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const text = e.target?.result as string;
          const parsed = JSON.parse(text) as AddInTemplate;

          // Validate template structure
          const validation = dataOrchestrator.validateTemplate(parsed);
          if (!validation.valid) {
            setError(validation.errors.map((e) => e.message).join(", "));
            return;
          }

          setJsonText(text);
          onTemplateChange(parsed);
          setError(null);
          setSuccess("Template imported successfully");
          setTimeout(() => setSuccess(null), 3000);
        } catch (err) {
          setError(err instanceof Error ? err.message : "Invalid JSON file");
        }
      };
      reader.readAsText(file);

      // Reset input so same file can be selected again
      event.target.value = "";
    },
    [onTemplateChange]
  );

  const handleTextChange = useCallback(
    (text: string) => {
      setJsonText(text);
      setError(null);
      setSuccess(null);

      // Try to parse and validate on each change
      if (text.trim()) {
        try {
          const parsed = JSON.parse(text) as AddInTemplate;
          const validation = dataOrchestrator.validateTemplate(parsed);
          if (validation.valid) {
            onTemplateChange(parsed);
          } else {
            onTemplateChange(null);
          }
        } catch {
          onTemplateChange(null);
        }
      } else {
        onTemplateChange(null);
      }
    },
    [onTemplateChange]
  );

  const handleExport = useCallback(() => {
    if (!template) return;

    const blob = new Blob([JSON.stringify(template, null, 2)], {
      type: "application/json",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${template.metadata?.name || "template"}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, [template]);

  const handleCreateNew = useCallback(() => {
    // Create a starter template based on available tables
    const newTemplate: AddInTemplate = {
      metadata: {
        name: "New Template",
        version: "1.0.0",
        description: "Template description",
      },
      tables: [],
      globals: [],
      outputs: {
        output: {
          filename: "output.json",
          template: `{
  "items": [
    {% for item in data.myTable %}
    {
      "field": "{{ item.Column1 }}"
    }{{ "," if not loop.last }}
    {% endfor %}
  ]
}`,
          description: "Main output file",
        },
      },
    };

    // Add detected tables
    tables.forEach((table) => {
      newTemplate.tables.push({
        name: table.name,
      });
    });

    const text = JSON.stringify(newTemplate, null, 2);
    setJsonText(text);
    onTemplateChange(newTemplate);
    setError(null);
    setSuccess("New template created");
    setTimeout(() => setSuccess(null), 3000);
  }, [tables, onTemplateChange]);

  const isValid = template !== null;

  return (
    <div className={styles.container}>
      <input
        ref={fileInputRef}
        type="file"
        accept=".json"
        className={styles.hiddenInput}
        onChange={handleFileChange}
      />

      <div className={styles.buttonRow}>
        <Button
          appearance="primary"
          icon={<ArrowUpload24Regular />}
          onClick={handleImportClick}
        >
          Import
        </Button>
        <Button
          appearance="secondary"
          icon={<ArrowDownload24Regular />}
          onClick={handleExport}
          disabled={!template}
        >
          Export
        </Button>
        <Button
          appearance="subtle"
          icon={<DocumentAdd24Regular />}
          onClick={handleCreateNew}
        >
          New
        </Button>
      </div>

      <MessageBar intent="info">
        <MessageBarBody>
          <MessageBarTitle>Outputs Managed in Outputs Tab</MessageBarTitle>
          The "outputs" section of the template is managed in the Outputs tab for easier editing with visual editors. Changes made directly to the outputs section in this JSON will be overwritten by the Outputs tab.
        </MessageBarBody>
      </MessageBar>

      {error && (
        <MessageBar intent="error">
          <MessageBarBody>
            <MessageBarTitle>Error</MessageBarTitle>
            {error}
          </MessageBarBody>
        </MessageBar>
      )}

      {success && (
        <MessageBar intent="success">
          <MessageBarBody>
            <MessageBarTitle>Success</MessageBarTitle>
            {success}
          </MessageBarBody>
        </MessageBar>
      )}

      <Card className={styles.editorCard}>
        <Text weight="semibold" size={200} style={{ marginBottom: "8px", display: "block" }}>
          Template JSON
        </Text>
        <MonacoEditor
          value={jsonText}
          onChange={handleTextChange}
          language="json"
          readOnly={false}
          height="500px"
          enableNunjucks={false}
        />
      </Card>

      <div className={styles.statusBar}>
        <Text size={200}>
          {isValid ? (
            <>
              <Checkmark24Regular
                style={{ verticalAlign: "middle", marginRight: "4px", color: "green" }}
              />
              Valid template: {template?.metadata?.name} v{template?.metadata?.version}
            </>
          ) : (
            <span style={{ color: tokens.colorNeutralForeground3 }}>
              No valid template loaded
            </span>
          )}
        </Text>
        {isValid && (
          <Text size={200}>
            {template?.tables?.length || 0} table(s),{" "}
            {Object.keys(template?.outputs || {}).length} output(s)
          </Text>
        )}
      </div>
    </div>
  );
}
