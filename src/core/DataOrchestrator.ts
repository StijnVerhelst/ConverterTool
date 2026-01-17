/**
 * DataOrchestrator
 *
 * Coordinates data reading, processing, and rendering.
 * Central orchestration point for the Add-in's conversion logic.
 */

import { tableService } from "@/services/TableService";
import { globalsService } from "@/services/GlobalsService";
import { nunjucksRenderer } from "./NunjucksRenderer";
import { applyMappingsToTables } from "./MappingProcessor";
import type {
  AddInTemplate,
  TemplateContext,
  RenderedOutput,
  ValidationResult,
  ValidationError,
} from "@/types";

/**
 * Orchestrates the full conversion process
 */
export class DataOrchestrator {
  /**
   * Validate a template structure
   * @param template - Template to validate
   * @returns Validation result
   */
  validateTemplate(template: AddInTemplate): ValidationResult {
    const errors: ValidationError[] = [];

    // Validate metadata
    if (!template.metadata?.name) {
      errors.push({ field: "metadata.name", message: "Template name is required" });
    }
    if (!template.metadata?.version) {
      errors.push({ field: "metadata.version", message: "Template version is required" });
    }

    // Validate tables
    if (!template.tables || template.tables.length === 0) {
      errors.push({ field: "tables", message: "At least one table mapping is required" });
    } else {
      template.tables.forEach((config, index) => {
        if (!config.name) {
          errors.push({
            field: `tables[${index}].name`,
            message: "Table name is required",
          });
        }
      });
    }

    // Validate outputs
    if (!template.outputs || Object.keys(template.outputs).length === 0) {
      errors.push({ field: "outputs", message: "At least one output is required" });
    } else {
      for (const [name, output] of Object.entries(template.outputs)) {
        if (!output.template) {
          errors.push({
            field: `outputs.${name}.template`,
            message: "Template content is required",
          });
        }
        if (!output.filename) {
          errors.push({
            field: `outputs.${name}.filename`,
            message: "Filename is required",
          });
        }
      }
    }

    return {
      valid: errors.length === 0,
      errors,
    };
  }

  /**
   * Validate data availability for a template
   * Checks that all referenced tables and named ranges exist
   * @param template - Template to validate against
   * @returns Validation result with warnings for missing items
   */
  async validateDataAvailability(template: AddInTemplate): Promise<ValidationResult> {
    const errors: ValidationError[] = [];

    // Check tables exist
    const availableTables = await tableService.getTables();
    const tableNames = availableTables.map((t) => t.name);

    for (const config of template.tables) {
      if (!tableNames.includes(config.name)) {
        errors.push({
          field: `tables.${config.name}`,
          message: `Excel Table "${config.name}" not found in workbook`,
          severity: "error",
        });
      }
    }

    // Check named ranges exist (warnings only)
    if (template.globals) {
      const availableRanges = await globalsService.getNamedRanges();
      const rangeNames = availableRanges.map((r) => r.name);

      for (const config of template.globals) {
        if (!rangeNames.includes(config.name)) {
          errors.push({
            field: `globals.${config.name}`,
            message: `Named Range "${config.name}" not found. Will use default value.`,
            severity: "warning",
          });
        }
      }
    }

    return {
      valid: errors.filter((e) => e.severity !== "warning").length === 0,
      errors,
    };
  }

  /**
   * Build template context from Excel data
   * Reads tables and globals according to template configuration
   * @param template - Template configuration
   * @returns Template context ready for rendering
   */
  async buildContext(template: AddInTemplate): Promise<TemplateContext> {
    // Read globals from Named Ranges
    const globals = template.globals
      ? await globalsService.readAllGlobals(template.globals)
      : {};

    // Read table data
    const tableConfigs = template.tables.map((config) => ({
      templateKey: config.name,
      tableName: config.name,
      columnMap: config.columnMap,
    }));

    let data = await tableService.readTables(tableConfigs);

    // Apply value mappings
    const mappingsConfig: Record<string, typeof template.tables[number]["valueMappings"]> = {};
    for (const config of template.tables) {
      mappingsConfig[config.name] = config.valueMappings;
    }
    data = applyMappingsToTables(data, mappingsConfig);

    const now = new Date();
    const pad = (value: number) => String(value).padStart(2, "0");
    const currentdatetime = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}_${pad(now.getHours())}-${pad(now.getMinutes())}-${pad(now.getSeconds())}`;

    return { globals, data, currentdatetime };
  }

  /**
   * Generate outputs from a template
   * Full pipeline: read data, apply mappings, render templates
   * @param template - Template configuration
   * @returns Array of rendered outputs
   */
  async generateOutputs(template: AddInTemplate): Promise<RenderedOutput[]> {
    // Build context
    const context = await this.buildContext(template);

    // Render all outputs
    return nunjucksRenderer.renderAllOutputs(template.outputs, context);
  }

  /**
   * Preview a single output
   * @param template - Template configuration
   * @param outputName - Name of the output to preview
   * @returns Rendered output or null if not found
   */
  async previewOutput(
    template: AddInTemplate,
    outputName: string
  ): Promise<RenderedOutput | null> {
    const output = template.outputs[outputName];
    if (!output) return null;

    const context = await this.buildContext(template);
    return nunjucksRenderer.renderOutput(output, context);
  }
}

// Singleton instance
export const dataOrchestrator = new DataOrchestrator();
