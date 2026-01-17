/**
 * NunjucksRenderer
 *
 * Renders Nunjucks templates with table data and global variables.
 * Adapted from CLI version - no filesystem operations, returns strings.
 */

import nunjucks from "nunjucks";
import type { OutputConfig, TemplateContext, RenderedOutput, TableRow } from "@/types";

/**
 * Renderer for Nunjucks templates
 */
export class NunjucksRenderer {
  private env: nunjucks.Environment;

  constructor() {
    this.env = this.configureNunjucks();
  }

  /**
   * Render a single template with context
   * @param template - Nunjucks template string
   * @param context - Data context (globals and data)
   * @returns Rendered content as string
   */
  render(template: string, context: TemplateContext): string {
    return this.env.renderString(template, context);
  }

  /**
   * Render a single output definition
   * @param output - Output configuration
   * @param context - Data context
   * @returns Rendered output with filename and content
   */
  renderOutput(output: OutputConfig, context: TemplateContext): RenderedOutput {
    // Resolve filename placeholders
    const filename = this.resolveFilename(output.filename, context);

    // Render template
    const content = this.render(output.template, context);

    // Determine MIME type from filename
    const mimeType = this.getMimeType(filename);

    return {
      name: filename,
      filename,
      content,
      mimeType,
    };
  }

  /**
   * Render all outputs from a template
   * @param outputs - Map of output configurations
   * @param context - Data context
   * @returns Array of rendered outputs
   */
  renderAllOutputs(
    outputs: Record<string, OutputConfig>,
    context: TemplateContext
  ): RenderedOutput[] {
    const results: RenderedOutput[] = [];

    for (const [name, output] of Object.entries(outputs)) {
      try {
        const rendered = this.renderOutput(output, context);
        rendered.name = name;
        results.push(rendered);
      } catch (error) {
        throw new Error(
          `Error rendering output "${name}": ${error instanceof Error ? error.message : String(error)}`
        );
      }
    }

    return results;
  }

  /**
   * Validate template syntax without rendering
   * @param template - Nunjucks template string
   * @returns Object with valid flag and optional error message
   */
  validateTemplate(template: string): { valid: boolean; error?: string } {
    try {
      // Try to compile the template (uses empty context)
      this.env.renderString(template, { data: {}, globals: {} });
      return { valid: true };
    } catch (error) {
      return {
        valid: false,
        error: error instanceof Error ? error.message : String(error),
      };
    }
  }

  /**
   * Configure Nunjucks environment with custom filters
   */
  private configureNunjucks(): nunjucks.Environment {
    const env = new nunjucks.Environment(null, {
      autoescape: false, // Don't auto-escape since we're generating various formats
      throwOnUndefined: false, // Don't throw on undefined variables
      trimBlocks: true, // Trim newlines from blocks
      lstripBlocks: true, // Strip leading whitespace from blocks
    });

    // Date formatting filter
    env.addFilter("date", (value: unknown, format: string = "YYYY-MM-DD") => {
      if (!value) return "";

      const date = value instanceof Date ? value : new Date(String(value));
      if (isNaN(date.getTime())) return String(value);

      return this.formatDate(date, format);
    });

    // JSON stringify filter
    env.addFilter("json", (value: unknown, indent: number = 2) => {
      return JSON.stringify(value, null, indent);
    });

    // JSON compact (no whitespace)
    env.addFilter("jsonCompact", (value: unknown) => {
      return JSON.stringify(value);
    });

    // CSV escape filter
    env.addFilter("csvEscape", (value: unknown) => {
      if (value === null || value === undefined) return "";
      const str = String(value);
      // Escape quotes and wrap if contains comma, quote, or newline
      if (str.includes(",") || str.includes('"') || str.includes("\n")) {
        return `"${str.replace(/"/g, '""')}"`;
      }
      return str;
    });

    // XML escape filter
    env.addFilter("xmlEscape", (value: unknown) => {
      if (value === null || value === undefined) return "";
      return String(value)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&apos;");
    });

    // Default filter (already built-in, but adding explicit version)
    env.addFilter("defaultValue", (value: unknown, defaultVal: unknown) => {
      return value === null || value === undefined || value === "" ? defaultVal : value;
    });

    return env;
  }

  /**
   * Resolve filename placeholders
   * Supports {{ currentdatetime }} and {{ globals.xxx }}
   */
  private resolveFilename(
    filename: string,
    context: TemplateContext
  ): string {
    return this.env.renderString(filename, {
      globals: context.globals,
      currentdatetime: context.currentdatetime,
    });
  }

  /**
   * Simple date formatting
   * Supports: YYYY, MM, DD, HH, mm, ss
   */
  private formatDate(date: Date, format: string): string {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    const hours = String(date.getHours()).padStart(2, "0");
    const minutes = String(date.getMinutes()).padStart(2, "0");
    const seconds = String(date.getSeconds()).padStart(2, "0");

    return format
      .replace("YYYY", String(year))
      .replace("MM", month)
      .replace("DD", day)
      .replace("HH", hours)
      .replace("mm", minutes)
      .replace("ss", seconds);
  }

  /**
   * Get MIME type from filename
   */
  private getMimeType(filename: string): string {
    const ext = filename.split(".").pop()?.toLowerCase();

    const mimeTypes: Record<string, string> = {
      json: "application/json",
      xml: "application/xml",
      csv: "text/csv",
      txt: "text/plain",
      html: "text/html",
      md: "text/markdown",
      yaml: "text/yaml",
      yml: "text/yaml",
    };

    return mimeTypes[ext || ""] || "text/plain";
  }
}

// Singleton instance
export const nunjucksRenderer = new NunjucksRenderer();
