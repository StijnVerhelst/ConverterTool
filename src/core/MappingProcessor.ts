/**
 * MappingProcessor
 *
 * Applies value mappings to rows.
 * Converts values from one column to new values in a new field.
 * Adapted from CLI version - unchanged logic.
 */

import type { ValueMapping, TableRow } from "@/types";

/**
 * Apply value mappings to an array of rows
 *
 * @param rows - Array of data rows
 * @param mappings - Mapping definitions
 * @returns Rows with new mapped fields added
 */
export function applyMappings(rows: TableRow[], mappings: ValueMapping[]): TableRow[] {
  if (!mappings || mappings.length === 0) {
    return rows;
  }

  return rows.map((row) => {
    const mappedRow = { ...row };

    for (const mapping of mappings) {
      const { column, outputField, map, default: defaultValue } = mapping;

      // Get source value
      const sourceValue = row[column];

      if (sourceValue === null || sourceValue === undefined) {
        mappedRow[outputField] = defaultValue;
        continue;
      }

      // Convert to string for map lookup
      const key = String(sourceValue);

      // Look up mapped value
      const mappedValue = map[key];

      if (mappedValue !== undefined) {
        mappedRow[outputField] = mappedValue;
      } else {
        mappedRow[outputField] = defaultValue;
      }
    }

    return mappedRow;
  });
}

/**
 * Apply mappings to multiple tables
 *
 * @param tablesData - Map of table key to rows
 * @param tableMappings - Map of table key to mapping definitions
 * @returns Tables data with mappings applied
 */
export function applyMappingsToTables(
  tablesData: Record<string, TableRow[]>,
  tableMappings: Record<string, ValueMapping[] | undefined>
): Record<string, TableRow[]> {
  const result: Record<string, TableRow[]> = {};

  for (const [tableKey, rows] of Object.entries(tablesData)) {
    const mappings = tableMappings[tableKey];
    result[tableKey] = mappings ? applyMappings(rows, mappings) : rows;
  }

  return result;
}
