# ConverterTool Excel Add-in

An Excel Add-in that converts Excel table data to structured output formats (JSON, XML, CSV) using Nunjucks templates.

## Overview

This Add-in simplifies the original CLI-based ConverterTool by leveraging Excel's native capabilities:

| Feature | CLI Version | Add-in Version |
|---------|-------------|----------------|
| Data filtering | Custom filter processor | Excel AutoFilter |
| Text transforms | Custom processor | Excel formulas (UPPER, TRIM, etc.) |
| Derived fields | JavaScript expressions | Excel formula columns |
| Relations | Custom join processor | Excel VLOOKUP/INDEX-MATCH |
| User prompts | CLI interactive prompts | Named Ranges |
| Value mappings | ✅ Kept | ✅ Kept |
| Nunjucks templates | ✅ Kept | ✅ Kept |

## Quick Start

### 1. Install Dependencies
```bash
cd excel-addin
npm install
```

### 2. Start Development Server
```bash
npm run dev-server
```

### 3. Sideload in Excel
```bash
npm run start
```

This opens Excel Desktop and loads the Add-in.

## Project Structure

```
excel-addin/
├── src/
│   ├── types/           # TypeScript interfaces
│   ├── services/        # Office.js integration
│   │   ├── TableService.ts      # Excel table operations
│   │   ├── GlobalsService.ts    # Named Range reading
│   │   └── OutputService.ts     # File download
│   ├── core/            # Processing logic
│   │   ├── NunjucksRenderer.ts  # Template rendering
│   │   ├── MappingProcessor.ts  # Value mappings
│   │   └── DataOrchestrator.ts  # Orchestration
│   └── taskpane/        # React UI
│       ├── App.tsx              # Main component
│       └── components/          # Tab components
├── manifest.xml         # Office Add-in manifest
├── example-template.json
└── package.json
```

## Template Format

```json
{
  "metadata": {
    "name": "My Template",
    "version": "1.0.0"
  },
  "tables": [
    {
      "name": "Equipment",
      "columnMap": { "Equipment Tag": "tag" },
      "valueMappings": [{
        "column": "status",
        "outputField": "statusCode",
        "map": { "Active": 1, "Inactive": 0 },
        "default": -1
      }]
    }
  ],
  "globals": [
    {
      "name": "ProjectName",
      "type": "string",
      "defaultValue": "Unnamed"
    }
  ],
  "outputs": {
    "json": {
      "filename": "{{ globals.ProjectName }}.json",
      "template": "{ \"items\": [...] }"
    }
  }
}
```

## UI Tabs

| Tab | Purpose |
|-----|---------|
| **Data** | View detected Excel Tables and Named Ranges |
| **Template** | Import/export/edit template JSON |
| **Preview** | Preview rendered output before downloading |
| **Export** | Select outputs and generate files |

## Nunjucks Filters

- `{{ value | date('YYYY-MM-DD') }}` - Format dates
- `{{ value | json }}` - JSON stringify
- `{{ value | csvEscape }}` - Escape for CSV
- `{{ value | xmlEscape }}` - Escape for XML
- `{{ value | default('N/A') }}` - Default value

## Excel Setup

1. **Create Tables**: Select data → Insert → Table
2. **Create Named Ranges**: Select cell → Formulas → Name Manager → New
3. **Use formulas** for calculated columns instead of derived fields

## Scripts

| Command | Description |
|---------|-------------|
| `npm run build` | Production build |
| `npm run dev-server` | Start dev server |
| `npm run start` | Sideload in Excel |
| `npm run stop` | Unload from Excel |

## Technology Stack

- **Office.js** - Excel integration
- **React 19** - UI framework
- **Fluent UI** - Microsoft design system
- **Nunjucks** - Template engine
- **JSZip** - Multi-file downloads
- **TypeScript** - Type safety
- **Webpack** - Bundling

## Migration from CLI

If you have existing CLI templates:

1. **Outputs**: Change `var.` to `globals.` in templates
2. **Tables**: Create Excel Tables, map names in template
3. **Globals**: Create Named Ranges for cell-based variables
4. **Filters/Transforms**: Recreate as Excel formulas or AutoFilter
5. **Derived fields**: Add formula columns in Excel

## License

MIT
