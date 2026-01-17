/**
 * GlobalsService
 *
 * Handles reading Named Ranges from Excel for template-wide variables.
 * Named Ranges provide a way to reference specific cells by name.
 */

import type { GlobalConfig } from "@/types";

/**
 * Information about a Named Range
 */
export interface NamedRangeInfo {
  /** Name of the range */
  name: string;
  /** Whether the range is visible (not hidden) */
  visible: boolean;
  /** The range address (e.g., "Sheet1!$A$1") */
  address: string;
  /** Worksheet name where the range is located (if single-sheet range) */
  worksheetName?: string;
}

/**
 * Service for reading Named Ranges via Office.js
 */
export class GlobalsService {
  /**
   * Get all workbook-scoped Named Ranges
   * @returns Promise resolving to array of named range info
   */
  async getNamedRanges(): Promise<NamedRangeInfo[]> {
    return Excel.run(async (context) => {
      const names = context.workbook.names;
      names.load("items/name,items/visible");
      await context.sync();

      const rangeInfos: NamedRangeInfo[] = [];

      for (const name of names.items) {
        // Only include visible names
        if (name.visible) {
          try {
            // Get range to load address and worksheet
            const range = name.getRange();
            range.load("address");

            // Try to get worksheet (may fail for multi-sheet ranges)
            let worksheetName: string | undefined;
            try {
              const worksheet = range.worksheet;
              worksheet.load("name");
              await context.sync();
              worksheetName = worksheet.name;
            } catch {
              // Multi-sheet or non-standard range
              await context.sync();
              worksheetName = undefined;
            }

            rangeInfos.push({
              name: name.name,
              visible: name.visible,
              address: range.address,
              worksheetName,
            });
          } catch (error) {
            // If we can't get range info, skip this named range
            console.warn(`Could not load info for named range "${name.name}":`, error);
          }
        }
      }

      return rangeInfos;
    });
  }

  /**
   * Check if a Named Range exists
   * @param rangeName - Name to check
   * @returns Promise resolving to boolean
   */
  async rangeExists(rangeName: string): Promise<boolean> {
    return Excel.run(async (context) => {
      try {
        const namedItem = context.workbook.names.getItem(rangeName);
        namedItem.load("name");
        await context.sync();
        return true;
      } catch {
        return false;
      }
    });
  }

  /**
   * Read value from a Named Range
   * @param rangeName - Name of the Named Range
   * @param type - Expected data type for conversion
   * @param defaultValue - Default value if empty or not found
   * @returns Promise resolving to the value
   */
  async readNamedRange(
    rangeName: string,
    type: "string" | "number" | "boolean" | "date" = "string",
    defaultValue?: unknown
  ): Promise<unknown> {
    return Excel.run(async (context) => {
      try {
        const namedItem = context.workbook.names.getItem(rangeName);
        const range = namedItem.getRange();
        range.load("values");
        await context.sync();

        // Get the first cell value (Named Ranges can be single cells or ranges)
        const value = range.values[0]?.[0];

        // Handle empty values
        if (value === null || value === undefined || value === "") {
          return defaultValue ?? this.getDefaultForType(type);
        }

        // Type conversion
        return this.convertValue(value, type);
      } catch {
        // Named range not found
        return defaultValue ?? this.getDefaultForType(type);
      }
    });
  }

  /**
   * Read all globals defined in template configuration
   * @param globalsConfig - Global variable configurations from template
   * @returns Promise resolving to globals object
   */
  async readAllGlobals(
    globalsConfig: GlobalConfig[]
  ): Promise<Record<string, unknown>> {
    const globals: Record<string, unknown> = {};

    for (const config of globalsConfig) {
      globals[config.name] = await this.readNamedRange(
        config.name,
        config.type,
        config.defaultValue
      );
    }

    return globals;
  }

  /**
   * Convert a value to the specified type
   */
  private convertValue(
    value: unknown,
    type: "string" | "number" | "boolean" | "date"
  ): unknown {
    switch (type) {
      case "number":
        if (typeof value === "number") return value;
        const num = parseFloat(String(value));
        return isNaN(num) ? 0 : num;

      case "boolean":
        if (typeof value === "boolean") return value;
        const str = String(value).toLowerCase();
        return str === "true" || str === "1" || str === "yes";

      case "date":
        if (value instanceof Date) {
          return value.toISOString();
        }
        // Excel stores dates as serial numbers
        if (typeof value === "number") {
          return this.excelDateToISO(value);
        }
        return String(value);

      default:
        return String(value);
    }
  }

  /**
   * Get default value for a type
   */
  private getDefaultForType(type: "string" | "number" | "boolean" | "date"): unknown {
    switch (type) {
      case "number":
        return 0;
      case "boolean":
        return false;
      case "date":
        return new Date().toISOString();
      default:
        return "";
    }
  }

  /**
   * Convert Excel serial date to ISO string
   * Excel uses a serial date system where 1 = January 1, 1900
   */
  private excelDateToISO(serial: number): string {
    // Excel's epoch is December 30, 1899 (accounting for the 1900 leap year bug)
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const msPerDay = 24 * 60 * 60 * 1000;
    const date = new Date(excelEpoch.getTime() + serial * msPerDay);
    return date.toISOString();
  }

  /**
   * Navigate to a named range in Excel
   * Activates the worksheet and selects the range
   * @param rangeName - Name of the named range to navigate to
   * @returns Promise that resolves when navigation is complete
   */
  async navigateToNamedRange(rangeName: string): Promise<void> {
    return Excel.run(async (context) => {
      const namedItem = context.workbook.names.getItem(rangeName);
      const range = namedItem.getRange();

      // Try to activate the worksheet (may not work for multi-sheet ranges)
      try {
        const worksheet = range.worksheet;
        worksheet.activate();
      } catch {
        // Multi-sheet range or worksheet not accessible
      }

      // Select the range
      range.select();

      await context.sync();
    });
  }
}

// Singleton instance
export const globalsService = new GlobalsService();
