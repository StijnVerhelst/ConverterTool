/**
 * ExportTab Component
 *
 * Generate and download output files.
 */

import { useState, useCallback } from "react";
import {
  Button,
  Text,
  Checkbox,
  Card,
  Spinner,
  makeStyles,
  tokens,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Divider,
} from "@fluentui/react-components";
import { ArrowDownload24Regular, Checkmark24Regular } from "@fluentui/react-icons";

import { dataOrchestrator } from "@/core";
import { outputService } from "@/services";
import type { AddInTemplate, RenderedOutput } from "@/types";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
  },
  card: {
    padding: "16px",
  },
  outputList: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  outputItem: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  generateButton: {
    marginTop: "16px",
  },
  emptyState: {
    padding: "32px",
    textAlign: "center",
    color: tokens.colorNeutralForeground3,
  },
  successCard: {
    padding: "16px",
    backgroundColor: tokens.colorPaletteGreenBackground1,
  },
  resultItem: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "4px 0",
  },
});

interface ExportTabProps {
  template: AddInTemplate | null;
}

export function ExportTab({ template }: ExportTabProps) {
  const styles = useStyles();
  const [selectedOutputs, setSelectedOutputs] = useState<Set<string>>(new Set());
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [results, setResults] = useState<RenderedOutput[] | null>(null);

  const outputNames = template ? Object.keys(template.outputs) : [];

  // Initialize selected outputs to all when template changes
  useState(() => {
    if (outputNames.length > 0) {
      setSelectedOutputs(new Set(outputNames));
    }
  });

  const handleOutputToggle = useCallback((name: string, checked: boolean) => {
    setSelectedOutputs((prev) => {
      const next = new Set(prev);
      if (checked) {
        next.add(name);
      } else {
        next.delete(name);
      }
      return next;
    });
  }, []);

  const handleSelectAll = useCallback(() => {
    setSelectedOutputs(new Set(outputNames));
  }, [outputNames]);

  const handleSelectNone = useCallback(() => {
    setSelectedOutputs(new Set());
  }, []);

  const handleGenerate = useCallback(async () => {
    if (!template || selectedOutputs.size === 0) return;

    setLoading(true);
    setError(null);
    setResults(null);

    try {
      // Validate data availability first
      const validation = await dataOrchestrator.validateDataAvailability(template);
      if (!validation.valid) {
        const errorMessages = validation.errors
          .filter((e) => e.severity !== "warning")
          .map((e) => e.message);
        if (errorMessages.length > 0) {
          setError(errorMessages.join("; "));
          setLoading(false);
          return;
        }
      }

      // Generate all outputs
      const allOutputs = await dataOrchestrator.generateOutputs(template);

      // Filter to selected outputs
      const filteredOutputs = allOutputs.filter((o) => selectedOutputs.has(o.name));

      // Download
      await outputService.downloadOutputs(filteredOutputs, template.metadata?.name);

      setResults(filteredOutputs);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to generate outputs");
    } finally {
      setLoading(false);
    }
  }, [template, selectedOutputs]);

  if (!template) {
    return (
      <div className={styles.container}>
        <Card className={styles.card}>
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
      <Card className={styles.card}>
        <Text weight="semibold" size={300} style={{ marginBottom: "12px", display: "block" }}>
          Select Outputs to Generate
        </Text>

        <div style={{ marginBottom: "12px" }}>
          <Button appearance="subtle" size="small" onClick={handleSelectAll}>
            Select All
          </Button>
          <Button appearance="subtle" size="small" onClick={handleSelectNone}>
            Select None
          </Button>
        </div>

        <div className={styles.outputList}>
          {outputNames.map((name) => (
            <div key={name} className={styles.outputItem}>
              <Checkbox
                checked={selectedOutputs.has(name)}
                onChange={(_, data) => handleOutputToggle(name, data.checked as boolean)}
                label={
                  <span>
                    <Text weight="semibold">{name}</Text>
                    <Text size={200} style={{ marginLeft: "8px", color: tokens.colorNeutralForeground3 }}>
                      → {template.outputs[name].filename}
                    </Text>
                  </span>
                }
              />
            </div>
          ))}
        </div>

        <Divider style={{ margin: "16px 0" }} />

        <Button
          appearance="primary"
          icon={loading ? <Spinner size="tiny" /> : <ArrowDownload24Regular />}
          onClick={handleGenerate}
          disabled={loading || selectedOutputs.size === 0}
          className={styles.generateButton}
        >
          {loading ? "Generating..." : "Generate & Download"}
        </Button>
      </Card>

      {error && (
        <MessageBar intent="error">
          <MessageBarBody>
            <MessageBarTitle>Error</MessageBarTitle>
            {error}
          </MessageBarBody>
        </MessageBar>
      )}

      {results && (
        <Card className={styles.successCard}>
          <Text weight="semibold" style={{ marginBottom: "8px", display: "block" }}>
            <Checkmark24Regular style={{ verticalAlign: "middle", marginRight: "8px" }} />
            Successfully generated {results.length} file(s)
          </Text>
          {results.map((result) => (
            <div key={result.name} className={styles.resultItem}>
              <Text size={200}>• {result.filename}</Text>
            </div>
          ))}
        </Card>
      )}
    </div>
  );
}
