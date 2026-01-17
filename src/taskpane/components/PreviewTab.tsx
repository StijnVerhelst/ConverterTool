/**
 * PreviewTab Component
 *
 * Shows a preview of rendered output before exporting.
 */

import { useState, useEffect, useCallback } from "react";
import {
  Button,
  Text,
  Dropdown,
  Option,
  Card,
  Spinner,
  makeStyles,
  tokens,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
} from "@fluentui/react-components";
import { ArrowSync24Regular, Copy24Regular, Checkmark24Regular } from "@fluentui/react-icons";

import { dataOrchestrator } from "@/core";
import { outputService } from "@/services";
import type { AddInTemplate, RenderedOutput } from "@/types";
import { MonacoEditor } from "./MonacoEditor";
import { detectLanguageFromFilename } from "../utils/monacoConfig";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
  },
  controls: {
    display: "flex",
    alignItems: "center",
    gap: "12px",
  },
  dropdown: {
    minWidth: "200px",
  },
  previewCard: {
    padding: "0",
    overflow: "hidden",
  },
  previewHeader: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "8px 12px",
    backgroundColor: tokens.colorNeutralBackground3,
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  previewContent: {
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground1,
  },
  previewEditor: {
    padding: "0",
  },
  emptyState: {
    padding: "32px",
    textAlign: "center",
    color: tokens.colorNeutralForeground3,
  },
});

interface PreviewTabProps {
  template: AddInTemplate | null;
}

export function PreviewTab({ template }: PreviewTabProps) {
  const styles = useStyles();
  const [selectedOutput, setSelectedOutput] = useState<string>("");
  const [preview, setPreview] = useState<RenderedOutput | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [copied, setCopied] = useState(false);

  const outputNames = template ? Object.keys(template.outputs) : [];

  // Auto-select first output when template changes
  useEffect(() => {
    if (outputNames.length > 0 && !outputNames.includes(selectedOutput)) {
      setSelectedOutput(outputNames[0]);
    }
  }, [outputNames, selectedOutput]);

  const loadPreview = useCallback(async () => {
    if (!template || !selectedOutput) return;

    setLoading(true);
    setError(null);

    try {
      const result = await dataOrchestrator.previewOutput(template, selectedOutput);
      setPreview(result);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to generate preview");
      setPreview(null);
    } finally {
      setLoading(false);
    }
  }, [template, selectedOutput]);

  // Load preview when selection changes
  useEffect(() => {
    if (selectedOutput) {
      loadPreview();
    }
  }, [selectedOutput, loadPreview]);

  const handleCopy = useCallback(async () => {
    if (!preview) return;

    try {
      await outputService.copyToClipboard(preview.content);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch (err) {
      setError("Failed to copy to clipboard");
    }
  }, [preview]);

  if (!template) {
    return (
      <div className={styles.container}>
        <Card className={styles.previewCard}>
          <div className={styles.emptyState}>
            <Text>No template loaded.</Text>
            <Text size={200} style={{ display: "block", marginTop: "8px" }}>
              Go to the Template tab to import or create a template.
            </Text>
          </div>
        </Card>
      </div>
    );
  }

  return (
    <div className={styles.container}>
      <div className={styles.controls}>
        <Text weight="semibold">Output:</Text>
        <Dropdown
          className={styles.dropdown}
          value={selectedOutput}
          onOptionSelect={(_, data) => setSelectedOutput(data.optionValue || "")}
          placeholder="Select output..."
        >
          {outputNames.map((name) => (
            <Option key={name} value={name}>
              {name}
            </Option>
          ))}
        </Dropdown>
        <Button
          appearance="subtle"
          icon={<ArrowSync24Regular />}
          onClick={loadPreview}
          disabled={loading || !selectedOutput}
        >
          Refresh
        </Button>
      </div>

      {error && (
        <MessageBar intent="error">
          <MessageBarBody>
            <MessageBarTitle>Error</MessageBarTitle>
            {error}
          </MessageBarBody>
        </MessageBar>
      )}

      <Card className={styles.previewCard}>
        <div className={styles.previewHeader}>
          <Text weight="semibold" size={200}>
            {preview?.filename || "Preview"}
          </Text>
          <Button
            appearance="subtle"
            size="small"
            icon={copied ? <Checkmark24Regular /> : <Copy24Regular />}
            onClick={handleCopy}
            disabled={!preview}
          >
            {copied ? "Copied!" : "Copy"}
          </Button>
        </div>

        {loading && (
          <div className={styles.previewContent}>
            <Spinner size="small" label="Generating preview..." />
          </div>
        )}

        {!loading && !preview && (
          <div className={styles.previewContent}>
            <Text className={styles.emptyState}>
              Select an output and click Refresh to generate preview.
            </Text>
          </div>
        )}

        {!loading && preview && (
          <div className={styles.previewEditor}>
            <MonacoEditor
              value={preview.content}
              language={detectLanguageFromFilename(preview.filename)}
              readOnly={true}
              height="400px"
            />
          </div>
        )}
      </Card>
    </div>
  );
}
