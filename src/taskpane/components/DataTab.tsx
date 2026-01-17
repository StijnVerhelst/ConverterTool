/**
 * DataTab Component
 *
 * Displays available Excel Tables and Named Ranges.
 * Auto-detects tables when mounted.
 * Provides checkboxes for auto-configuration and click-to-navigate functionality.
 */

import { useState, useEffect, useCallback } from "react";
import {
  Button,
  Text,
  Spinner,
  Card,
  CardHeader,
  makeStyles,
  tokens,
  Badge,
  Divider,
  Checkbox,
  Link,
} from "@fluentui/react-components";
import { ArrowSync24Regular, Table24Regular, Tag24Regular, Location24Regular } from "@fluentui/react-icons";

import { tableService, globalsService } from "@/services";
import type { TableInfo, NamedRangeInfo } from "@/services";
import type { AddInTemplate } from "@/types";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
  },
  section: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  sectionHeader: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
  },
  sectionTitle: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    fontWeight: 600,
  },
  card: {
    padding: "12px",
  },
  tableItem: {
    display: "flex",
    alignItems: "flex-start",
    gap: "8px",
    padding: "8px 0",
  },
  tableContent: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
    flex: 1,
  },
  tableName: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  nameLink: {
    cursor: "pointer",
    fontWeight: 600,
  },
  tableInfo: {
    display: "flex",
    flexDirection: "column",
    gap: "2px",
  },
  location: {
    display: "flex",
    alignItems: "center",
    gap: "4px",
    fontSize: "11px",
    color: tokens.colorNeutralForeground3,
  },
  columns: {
    fontSize: "12px",
    color: tokens.colorNeutralForeground3,
    marginTop: "2px",
  },
  emptyMessage: {
    padding: "16px",
    textAlign: "center",
    color: tokens.colorNeutralForeground3,
  },
  errorMessage: {
    padding: "12px",
    backgroundColor: tokens.colorPaletteRedBackground1,
    color: tokens.colorPaletteRedForeground1,
    borderRadius: "4px",
  },
});

interface DataTabProps {
  tables: TableInfo[];
  onTablesChange: (tables: TableInfo[]) => void;
  template: AddInTemplate | null;
  onTemplateChange: (template: AddInTemplate | null) => void;
}

export function DataTab({ tables, onTablesChange, template, onTemplateChange }: DataTabProps) {
  const styles = useStyles();
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [namedRanges, setNamedRanges] = useState<NamedRangeInfo[]>([]);
  const [selectedTables, setSelectedTables] = useState<Set<string>>(new Set());
  const [selectedRanges, setSelectedRanges] = useState<Set<string>>(new Set());

  const loadData = useCallback(async () => {
    setLoading(true);
    setError(null);

    try {
      const [tablesResult, rangesResult] = await Promise.all([
        tableService.getTables(),
        globalsService.getNamedRanges(),
      ]);

      onTablesChange(tablesResult);
      setNamedRanges(rangesResult);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to load data");
    } finally {
      setLoading(false);
    }
  }, [onTablesChange]);

  // Sync checkbox state with template on mount and template change
  // Also clean up orphaned table/range configs
  useEffect(() => {
    if (template && tables.length > 0) {
      // Get current table names from Excel
      const currentTableNames = new Set(tables.map(t => t.name));

      // Clean up orphaned table configs (tables that no longer exist in Excel)
      const updatedTables = template.tables.filter(config => currentTableNames.has(config.name));
      const tablesChanged = updatedTables.length !== template.tables.length;

      // Sync table checkboxes with cleaned config
      const tableKeys = new Set(updatedTables.map(config => config.name));
      setSelectedTables(tableKeys);

      const rangeKeys = new Set((template.globals || []).map(config => config.name));
      setSelectedRanges(rangeKeys);

      // Update template if any orphaned configs were found
      if (tablesChanged) {
        onTemplateChange({
          ...template,
          tables: updatedTables,
          globals: template.globals,
        });
      }
    } else if (template) {
      // No tables loaded yet, just sync with template
      const tableKeys = new Set(template.tables.map(config => config.name));
      setSelectedTables(tableKeys);

      const rangeKeys = new Set((template.globals || []).map(config => config.name));
      setSelectedRanges(rangeKeys);
    } else {
      setSelectedTables(new Set());
      setSelectedRanges(new Set());
    }
  }, [template, tables, namedRanges, onTemplateChange]);

  // Load data on mount
  useEffect(() => {
    loadData();
  }, [loadData]);

  /**
   * Handle table checkbox change
   * When checked: auto-add basic config with all columns
   * When unchecked: remove from template
   */
  const handleTableCheckboxChange = useCallback((table: TableInfo, checked: boolean) => {
    if (!template) return;

    const updatedTemplate = { ...template };
    updatedTemplate.tables = [...template.tables];

    if (checked) {
      // Add table to template with basic config (all columns)
      const exists = updatedTemplate.tables.some(config => config.name === table.name);
      if (!exists) {
        updatedTemplate.tables.push({
          name: table.name,
        });
      }
      setSelectedTables(prev => new Set(prev).add(table.name));
    } else {
      // Remove table from template
      updatedTemplate.tables = updatedTemplate.tables.filter(
        config => config.name !== table.name
      );
      setSelectedTables(prev => {
        const newSet = new Set(prev);
        newSet.delete(table.name);
        return newSet;
      });
    }

    onTemplateChange(updatedTemplate);
  }, [template, onTemplateChange]);

  /**
   * Handle named range checkbox change
   * When checked: add to globals
   * When unchecked: remove from globals
   */
  const handleRangeCheckboxChange = useCallback((range: NamedRangeInfo, checked: boolean) => {
    if (!template) return;

    const updatedTemplate = { ...template };
    updatedTemplate.globals = [...(template.globals || [])];

    if (checked) {
      // Add range to globals with basic config
      const exists = updatedTemplate.globals.some(config => config.name === range.name);
      if (!exists) {
        updatedTemplate.globals.push({
          name: range.name,
          type: "string",
        });
      }
      setSelectedRanges(prev => new Set(prev).add(range.name));
    } else {
      // Remove range from globals
      updatedTemplate.globals = updatedTemplate.globals.filter(
        config => config.name !== range.name
      );
      setSelectedRanges(prev => {
        const newSet = new Set(prev);
        newSet.delete(range.name);
        return newSet;
      });
    }

    onTemplateChange(updatedTemplate);
  }, [template, onTemplateChange]);

  /**
   * Navigate to a table in Excel
   */
  const handleNavigateToTable = useCallback(async (tableName: string) => {
    try {
      await tableService.navigateToTable(tableName);
    } catch (err) {
      setError(`Could not navigate to table. It may have been deleted.`);
      // Refresh data to reflect current state
      loadData();
    }
  }, [loadData]);

  /**
   * Navigate to a named range in Excel
   */
  const handleNavigateToRange = useCallback(async (rangeName: string) => {
    try {
      await globalsService.navigateToNamedRange(rangeName);
    } catch (err) {
      setError(`Could not navigate to named range. It may have been deleted.`);
      // Refresh data to reflect current state
      loadData();
    }
  }, [loadData]);

  return (
    <div className={styles.container}>
      {/* Tables Section */}
      <div className={styles.section}>
        <div className={styles.sectionHeader}>
          <Text className={styles.sectionTitle}>
            <Table24Regular />
            Excel Tables
          </Text>
          <Button
            appearance="subtle"
            icon={<ArrowSync24Regular />}
            onClick={loadData}
            disabled={loading}
          >
            Refresh
          </Button>
        </div>

        {loading && <Spinner size="small" label="Loading tables..." />}

        {error && <div className={styles.errorMessage}>{error}</div>}

        {!loading && !error && tables.length === 0 && (
          <Card className={styles.card}>
            <Text className={styles.emptyMessage}>
              No Excel Tables found in this workbook.
              <br />
              <br />
              To create a table: Select your data range, then go to Insert → Table.
            </Text>
          </Card>
        )}

        {!loading && tables.length > 0 && (
          <Card className={styles.card}>
            {tables.map((table, index) => (
              <div key={table.name}>
                {index > 0 && <Divider />}
                <div className={styles.tableItem}>
                  <Checkbox
                    checked={selectedTables.has(table.name)}
                    onChange={(_, data) => handleTableCheckboxChange(table, !!data.checked)}
                    disabled={!template}
                  />
                  <div className={styles.tableContent}>
                    <div className={styles.tableName}>
                      <Link
                        className={styles.nameLink}
                        onClick={() => handleNavigateToTable(table.name)}
                      >
                        {table.name}
                      </Link>
                      <Badge appearance="filled" color="informative">
                        {table.rowCount} rows
                      </Badge>
                    </div>
                    <div className={styles.location}>
                      <Location24Regular fontSize={12} />
                      <Text>{table.worksheetName} • {table.address}</Text>
                    </div>
                    <Text size={200} className={styles.columns}>
                      Columns: {table.columns.join(", ")}
                    </Text>
                  </div>
                </div>
              </div>
            ))}
          </Card>
        )}
      </div>

      {/* Named Ranges Section */}
      <div className={styles.section}>
        <Text className={styles.sectionTitle}>
          <Tag24Regular />
          Named Ranges (Globals)
        </Text>

        {!loading && namedRanges.length === 0 && (
          <Card className={styles.card}>
            <Text className={styles.emptyMessage}>
              No Named Ranges found.
              <br />
              <br />
              To create: Select a cell, then go to Formulas → Name Manager → New.
            </Text>
          </Card>
        )}

        {!loading && namedRanges.length > 0 && (
          <Card className={styles.card}>
            {namedRanges.map((range, index) => (
              <div key={range.name}>
                {index > 0 && <Divider />}
                <div className={styles.tableItem}>
                  <Checkbox
                    checked={selectedRanges.has(range.name)}
                    onChange={(_, data) => handleRangeCheckboxChange(range, !!data.checked)}
                    disabled={!template}
                  />
                  <div className={styles.tableContent}>
                    <div className={styles.tableName}>
                      <Link
                        className={styles.nameLink}
                        onClick={() => handleNavigateToRange(range.name)}
                      >
                        {range.name}
                      </Link>
                      <Badge appearance="outline">Available</Badge>
                    </div>
                    <div className={styles.location}>
                      <Location24Regular fontSize={12} />
                      <Text>{range.address}</Text>
                    </div>
                  </div>
                </div>
              </div>
            ))}
          </Card>
        )}
      </div>
    </div>
  );
}
