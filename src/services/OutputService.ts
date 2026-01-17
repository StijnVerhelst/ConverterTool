/**
 * OutputService
 *
 * Handles file generation and download.
 * Uses browser APIs (Blob, anchor element) for downloads.
 * Uses JSZip for multiple file downloads.
 */

import JSZip from "jszip";
import type { RenderedOutput } from "@/types";

/**
 * Service for file output and downloads
 */
export class OutputService {
  /**
   * Download a single file
   * @param content - File content as string
   * @param filename - Filename for download
   * @param mimeType - MIME type (optional, auto-detected from extension)
   */
  downloadFile(content: string, filename: string, mimeType?: string): void {
    const type = mimeType || this.getMimeType(filename);
    const blob = new Blob([content], { type });
    this.downloadBlob(blob, filename);
  }

  /**
   * Download multiple files as a ZIP archive
   * @param files - Array of files with name and content
   * @param zipFilename - Name for the ZIP file
   */
  async downloadAsZip(
    files: Array<{ name: string; content: string }>,
    zipFilename: string
  ): Promise<void> {
    const zip = new JSZip();

    for (const file of files) {
      zip.file(file.name, file.content);
    }

    const blob = await zip.generateAsync({ type: "blob" });
    this.downloadBlob(blob, zipFilename);
  }

  /**
   * Download rendered outputs
   * Downloads as individual files if one, or as ZIP if multiple
   * @param outputs - Array of rendered outputs
   * @param templateName - Template name for ZIP filename
   */
  async downloadOutputs(outputs: RenderedOutput[], templateName?: string): Promise<void> {
    if (outputs.length === 0) {
      throw new Error("No outputs to download");
    }

    if (outputs.length === 1) {
      // Single file - download directly
      const output = outputs[0];
      this.downloadFile(output.content, output.filename, output.mimeType);
    } else {
      // Multiple files - download as ZIP
      const files = outputs.map((output) => ({
        name: output.filename,
        content: output.content,
      }));

      const zipName = templateName
        ? `${templateName}_outputs.zip`
        : `outputs_${this.getTimestamp()}.zip`;

      await this.downloadAsZip(files, zipName);
    }
  }

  /**
   * Copy content to clipboard
   * @param content - Content to copy
   */
  async copyToClipboard(content: string): Promise<void> {
    await navigator.clipboard.writeText(content);
  }

  /**
   * Get MIME type based on file extension
   */
  getMimeType(filename: string): string {
    const ext = filename.split(".").pop()?.toLowerCase();

    const mimeTypes: Record<string, string> = {
      json: "application/json",
      xml: "application/xml",
      csv: "text/csv",
      txt: "text/plain",
      html: "text/html",
      htm: "text/html",
      md: "text/markdown",
      yaml: "text/yaml",
      yml: "text/yaml",
      js: "text/javascript",
      ts: "text/typescript",
      css: "text/css",
    };

    return mimeTypes[ext || ""] || "text/plain";
  }

  /**
   * Download a Blob as a file
   */
  private downloadBlob(blob: Blob, filename: string): void {
    const url = URL.createObjectURL(blob);

    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = filename;
    anchor.style.display = "none";

    document.body.appendChild(anchor);
    anchor.click();
    document.body.removeChild(anchor);

    // Clean up the object URL after a short delay
    setTimeout(() => {
      URL.revokeObjectURL(url);
    }, 100);
  }

  /**
   * Get timestamp for filenames
   */
  private getTimestamp(): string {
    const now = new Date();
    return now.toISOString().replace(/[:.]/g, "-").slice(0, 19);
  }
}

// Singleton instance
export const outputService = new OutputService();
