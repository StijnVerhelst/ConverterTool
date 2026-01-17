/**
 * TableService
 *
 * Handles Excel Table detection and data reading via Office.js API.
 * Auto-detects tables when the Add-in opens and reads table data for rendering.
 */

import type { TableRow, ProcessedTableData } from "@/types";

/**
 * Information about an Excel Table
 */
export interface TableInfo {
  /** Table name (as shown in Table Design tab) */
  name: string;
  /** Column headers */
  columns: string[];
  /** Number of data rows (excluding header) */
  rowCount: number;
  /** Whether table has a header row */
  showHeaders: boolean;
  /** Whether table has a totals row */
  showTotals: boolean;
  /** Worksheet name where the table is located */
  worksheetName: string;
  /** Cell address range (e.g., "A1:D10") */
  address: string;
}

/**
 * Service for Excel Table operations via Office.js
 */
export class TableService {
  /**
   * Get all Excel Tables in the workbook
   * @returns Promise resolving to array of table information
   */
  async getTables(): Promise<TableInfo[]> {
    return Excel.run(async (context) => {
      const tables = context.workbook.tables;
      tables.load("items/name,items/showHeaders,items/showTotals,items/worksheet");
      await context.sync();

      const tableInfos: TableInfo[] = [];

      for (const table of tables.items) {
        // Get table columns
        const columns = table.columns;
        columns.load("items/name");

        // Get row count and address
        const bodyRange = table.getDataBodyRange();
        bodyRange.load("rowCount");

        // Get full table range for address
        const tableRange = table.getRange();
        tableRange.load("address");

        // Get worksheet name
        const worksheet = table.worksheet;
        worksheet.load("name");

        await context.sync();

        tableInfos.push({
          name: table.name,
          showHeaders: table.showHeaders,
          showTotals: table.showTotals,
          columns: columns.items.map((col) => col.name),
          rowCount: bodyRange.rowCount,
          worksheetName: worksheet.name,
          address: tableRange.address,
        });
      }

      return tableInfos;
    });
  }

  /**
   * Check if a specific table exists
   * @param tableName - Name of the table to check
   * @returns Promise resolving to boolean
   */
  async tableExists(tableName: string): Promise<boolean> {
    return Excel.run(async (context) => {
      try {
        const table = context.workbook.tables.getItem(tableName);
        table.load("name");
        await context.sync();
        return true;
      } catch {
        return false;
      }
    });
  }

  /**
   * Read data from an Excel Table
   * @param tableName - Name of the Excel Table
   * @param columnMap - Optional mapping of Excel headers to template field names
   * @returns Promise resolving to processed table data
   */
  async readTable(
    tableName: string,
    columnMap?: Record<string, string>
  ): Promise<ProcessedTableData> {
    return Excel.run(async (context) => {
      const table = context.workbook.tables.getItem(tableName);

      // Load header and body ranges
      const headerRange = table.getHeaderRowRange();
      const bodyRange = table.getDataBodyRange();

      headerRange.load("values");
      bodyRange.load("values");

      await context.sync();

      // Extract headers
      const rawHeaders = headerRange.values[0] as string[];

      // Apply column mapping if provided
      const headers = rawHeaders.map((header) => {
        if (columnMap && columnMap[header]) {
          return columnMap[header];
        }
        return header;
      });

      // Convert rows to objects
      const rows: TableRow[] = [];
      const bodyValues = bodyRange.values;

      for (const rowValues of bodyValues) {
        const row: TableRow = {};
        headers.forEach((header, index) => {
          row[header] = rowValues[index];
        });
        rows.push(row);
      }

      return {
        key: tableName, // Will be overridden by caller with template key
        tableName,
        headers,
        rows,
        rowCount: rows.length,
      };
    });
  }

  /**
   * Read multiple tables at once
   * @param tableConfigs - Array of table configurations
   * @returns Promise resolving to map of template key -> table data
   */
  async readTables(
    tableConfigs: Array<{
      templateKey: string;
      tableName: string;
      columnMap?: Record<string, string>;
    }>
  ): Promise<Record<string, TableRow[]>> {
    const result: Record<string, TableRow[]> = {};

    for (const config of tableConfigs) {
      try {
        const tableData = await this.readTable(config.tableName, config.columnMap);
        result[config.templateKey] = tableData.rows;
      } catch (error) {
        console.error(`Failed to read table "${config.tableName}":`, error);
        result[config.templateKey] = [];
      }
    }

    return result;
  }

  /**
   * Get column values for a specific column in a table
   * Useful for displaying unique values or validating data
   * @param tableName - Name of the Excel Table
   * @param columnName - Name of the column
   * @returns Promise resolving to array of column values
   */
  async getColumnValues(tableName: string, columnName: string): Promise<unknown[]> {
    return Excel.run(async (context) => {
      const table = context.workbook.tables.getItem(tableName);
      const column = table.columns.getItem(columnName);
      const bodyRange = column.getDataBodyRange();

      bodyRange.load("values");
      await context.sync();

      // Flatten 2D array to 1D
      return bodyRange.values.map((row: unknown[]) => row[0]);
    });
  }

  /**
   * Get distinct values from a column
   * @param tableName - Name of the Excel Table
   * @param columnName - Name of the column
   * @returns Promise resolving to array of unique values
   */
  async getDistinctColumnValues(tableName: string, columnName: string): Promise<unknown[]> {
    const values = await this.getColumnValues(tableName, columnName);
    return [...new Set(values)].filter((v) => v !== null && v !== undefined && v !== "");
  }

  /**
   * Navigate to a table in Excel
   * Activates the worksheet and selects the table range
   * @param tableName - Name of the table to navigate to
   * @returns Promise that resolves when navigation is complete
   */
  async navigateToTable(tableName: string): Promise<void> {
    return Excel.run(async (context) => {
      const table = context.workbook.tables.getItem(tableName);
      const range = table.getRange();
      const worksheet = table.worksheet;

      // Activate the worksheet
      worksheet.activate();

      // Select the table range
      range.select();

      await context.sync();
    });
  }
}

// Singleton instance
export const tableService = new TableService();
