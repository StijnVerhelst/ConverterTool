/**
 * Excel Add-in Template Schema v1.0
 *
 * Simplified template format for the Excel Add-in version.
 * Leverages Excel's native capabilities (formulas, filtering, sorting)
 * instead of custom processing.
 */

/**
 * Root template structure
 */
export interface AddInTemplate {
  /** Template metadata */
  metadata: TemplateMetadata;

  /** List of Excel table configurations */
  tables: TableConfig[];

  /** List of global named range configurations */
  globals?: GlobalConfig[];

  /** Nunjucks output definitions */
  outputs: Record<string, OutputConfig>;
}

/**
 * Template metadata
 */
export interface TemplateMetadata {
  /** Template name */
  name: string;
  /** Template version (semantic versioning recommended) */
  version: string;
  /** Optional description */
  description?: string;
}

/**
 * Table configuration - maps an Excel Table to a template data key
 */
export interface TableConfig {
  /** Excel Table name (as shown in Table Design tab, case sensitive) */
  name: string;

  /**
   * Optional column renames: Excel header -> template field name
   * Use when Excel headers don't match desired template field names
   */
  columnMap?: Record<string, string>;

  /**
   * Optional value mappings for code-to-display conversions
   * Applied during rendering, not in Excel
   */
  valueMappings?: ValueMapping[];
}

/**
 * Value mapping configuration
 * Converts values from one column to new values in a new field
 */
export interface ValueMapping {
  /** Source column name */
  column: string;
  /** Target field name (new field added to row) */
  outputField: string;
  /** Lookup map: source value -> output value */
  map: Record<string, string | number | boolean>;
  /** Default value if source value not found in map */
  default: string | number | boolean;
}

/**
 * Global variable configuration - maps to a Named Range
 */
export interface GlobalConfig {
  /** Excel Named Range name (case sensitive) */
  name: string;
  /** Expected data type for conversion */
  type?: "string" | "number" | "boolean" | "date";
  /** Default value if Named Range is empty or not found */
  defaultValue?: string | number | boolean | null;
}

/**
 * Output configuration - defines a generated file
 */
export interface OutputConfig {
  /**
   * Output filename
   * Supports placeholders: {{ globals.project }}.json, {{ currentdatetime }}.csv
   */
  filename: string;

  /**
   * Nunjucks template string
   * Has access to:
   *   - globals: Record<string, any> - Global variables
   *   - data: Record<string, TableRow[]> - All table data
   */
  template: string;

  /** Optional description for documentation */
  description?: string;
}

/**
 * Represents a row from an Excel Table
 * Keys are column headers (or remapped names from columnMap)
 */
export type TableRow = Record<string, unknown>;

/**
 * Processed table data ready for rendering
 */
export interface ProcessedTableData {
  /** Template key (from tables config) */
  key: string;
  /** Excel Table name */
  tableName: string;
  /** Column headers */
  headers: string[];
  /** Row data */
  rows: TableRow[];
  /** Row count */
  rowCount: number;
}

/**
 * Nunjucks template context
 */
export interface TemplateContext {
  /** Global variables from Named Ranges */
  globals: Record<string, unknown>;
  /** Table data by template key */
  data: Record<string, TableRow[]>;
  /** Current datetime formatted for filenames (YYYY-MM-DD_HH-mm-ss) */
  currentdatetime: string;
}

/**
 * Rendered output result
 */
export interface RenderedOutput {
  /** Output name (from outputs config key) */
  name: string;
  /** Resolved filename */
  filename: string;
  /** Rendered content */
  content: string;
  /** MIME type for download */
  mimeType: string;
}

/**
 * Validation error
 */
export interface ValidationError {
  /** Field path that has the error */
  field: string;
  /** Error message */
  message: string;
  /** Severity level */
  severity?: "error" | "warning";
}

/**
 * Validation result
 */
export interface ValidationResult {
  /** Whether validation passed */
  valid: boolean;
  /** List of errors/warnings */
  errors: ValidationError[];
}
