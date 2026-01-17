/**
 * Monaco Editor Configuration Utilities
 *
 * Provides language detection, Nunjucks syntax highlighting,
 * and autocomplete for template variables (globals, data.tablexx).
 */

import type { editor, languages, IRange } from "monaco-editor";

/**
 * Detects the Monaco language mode from a filename extension
 *
 * @param filename - The filename to analyze
 * @returns Monaco language identifier (e.g., 'json', 'html', 'xml')
 */
export function detectLanguageFromFilename(filename: string): string {
  const extension = filename.split('.').pop()?.toLowerCase() || '';

  const languageMap: Record<string, string> = {
    'json': 'json',
    'xml': 'xml',
    'html': 'html',
    'htm': 'html',
    'csv': 'plaintext',
    'txt': 'plaintext',
    'md': 'markdown',
    'js': 'javascript',
    'ts': 'typescript',
    'css': 'css',
    'sql': 'sql',
    'yaml': 'yaml',
    'yml': 'yaml',
  };

  return languageMap[extension] || 'plaintext';
}

/**
 * Default Monaco Editor options
 * Configured for use in the Excel Add-in
 */
export const defaultEditorOptions: editor.IStandaloneEditorConstructionOptions = {
  minimap: { enabled: false },
  lineNumbers: 'on',
  wordWrap: 'on',
  automaticLayout: true,
  scrollBeyondLastLine: false,
  fontSize: 12,
  fontFamily: 'Consolas, "Courier New", monospace',
  renderWhitespace: 'selection',
  bracketPairColorization: { enabled: true },
  occurrencesHighlight: 'off',
  selectionHighlight: false,
  suggestOnTriggerCharacters: true,
  tabFocusMode: false,
  quickSuggestions: {
    other: true,
    comments: false,
    strings: true,
  },
  tabSize: 2,
  formatOnPaste: true,
};

/**
 * Editor options for Nunjucks template editing (JSON validation disabled)
 */
export const nunjucksEditorOptions: editor.IStandaloneEditorConstructionOptions = {
  ...defaultEditorOptions,
  // These help with Nunjucks templates
  autoClosingBrackets: 'always',
  autoClosingQuotes: 'always',
};

/**
 * Disable JSON validation for a specific editor
 * Call this in onMount when enableNunjucks is true
 */
export function disableJsonValidation(monacoInstance: any): void {
  try {
    // Disable JSON diagnostics globally
    monacoInstance.languages.json?.jsonDefaults?.setDiagnosticsOptions({
      validate: false,
      allowComments: true,
      schemas: [],
      enableSchemaRequest: false,
    });
    console.log('JSON validation disabled');
  } catch (e) {
    console.warn('Could not disable JSON validation:', e);
  }
}

/**
 * Table info for autocomplete
 */
interface TableInfoForAutocomplete {
  name: string;
  key: string;
  columns: string[];
}

/**
 * Global info for autocomplete
 */
interface GlobalInfoForAutocomplete {
  name: string;
  key: string;
}

interface LoopScope {
  varName: string;
  tableKey: string;
  startIndex: number;
  endIndex: number;
}

// Store for autocomplete data
let autocompleteData: {
  tables: TableInfoForAutocomplete[];
  globals: GlobalInfoForAutocomplete[];
} = {
  tables: [],
  globals: [],
};

const diagnosticsListeners = new Set<() => void>();

function notifyDiagnosticsListeners(): void {
  diagnosticsListeners.forEach(listener => listener());
}

export function subscribeToNunjucksDiagnostics(listener: () => void): () => void {
  diagnosticsListeners.add(listener);
  return () => diagnosticsListeners.delete(listener);
}

/**
 * Update autocomplete data with available tables and globals
 */
export function updateAutocompleteData(
  tables: TableInfoForAutocomplete[],
  globals: GlobalInfoForAutocomplete[]
): void {
  autocompleteData = { tables, globals };
  console.log('Autocomplete data updated:', {
    tables: tables.map(t => ({ key: t.key, columns: t.columns.length })),
    globals: globals.map(g => g.key),
  });
  notifyDiagnosticsListeners();
}

/**
 * Configure Monaco editor with Nunjucks syntax highlighting and autocomplete
 * Call this in the onMount handler of the Editor component
 */
export function configureMonacoForNunjucks(monacoInstance: any): void {
  // Disable JSON validation to prevent errors on Nunjucks syntax
  try {
    // Try multiple ways to disable JSON validation
    if (monacoInstance.languages.json?.jsonDefaults) {
      monacoInstance.languages.json.jsonDefaults.setDiagnosticsOptions({
        validate: false,
        allowComments: true,
        schemaValidation: 'ignore',
      });
    }
    // Also try the direct method
    const jsonDefaults = monacoInstance.languages.json?.jsonDefaults;
    if (jsonDefaults?.setDiagnosticsOptions) {
      jsonDefaults.setDiagnosticsOptions({
        validate: false,
      });
    }
  } catch (e) {
    console.warn('Could not disable JSON validation:', e);
  }

  // Define Nunjucks token colors
  monacoInstance.editor.defineTheme('nunjucks-theme', {
    base: 'vs',
    inherit: true,
    rules: [
      { token: 'nunjucks-delimiter', foreground: '9932CC', fontStyle: 'bold' },
      { token: 'nunjucks-keyword', foreground: '0000FF', fontStyle: 'bold' },
      { token: 'nunjucks-variable', foreground: '008000' },
      { token: 'nunjucks-filter', foreground: 'FF8C00' },
      { token: 'nunjucks-globals', foreground: 'DC143C', fontStyle: 'bold' },
      { token: 'nunjucks-data', foreground: '4169E1', fontStyle: 'bold' },
      { token: 'nunjucks-comment', foreground: '808080', fontStyle: 'italic' },
    ],
    colors: {},
  });

  // Register completion provider for all languages
  const languages = ['json', 'html', 'xml', 'plaintext', 'javascript', 'typescript', 'css', 'markdown'];

  languages.forEach(lang => {
    monacoInstance.languages.registerCompletionItemProvider(lang, {
      triggerCharacters: ['{', '.', '[', '"', "'", ' ', 'g', 'd', 'r'],
      provideCompletionItems: (model: any, position: any) => {
        const completions = getNunjucksCompletions(monacoInstance, model, position);
        console.log('Autocomplete triggered, suggestions:', completions.suggestions.length);
        return completions;
      },
    });
  });

  // Register hover provider for Nunjucks syntax
  languages.forEach(lang => {
    monacoInstance.languages.registerHoverProvider(lang, {
      provideHover: (model: any, position: any) => {
        return getNunjucksHover(model, position);
      },
    });
  });

  console.log('Monaco Nunjucks configuration complete');
}

/**
 * Get Nunjucks completions based on context
 */
function getNunjucksCompletions(
  monacoInstance: any,
  model: any,
  position: any
): { suggestions: any[] } {
  const textUntilPosition = model.getValueInRange({
    startLineNumber: position.lineNumber,
    startColumn: 1,
    endLineNumber: position.lineNumber,
    endColumn: position.column,
  });

  const propertyAccess = getPropertyAccessAtCursor(model, position);
  const word = model.getWordUntilPosition(position);
  const defaultRange: IRange = {
    startLineNumber: position.lineNumber,
    endLineNumber: position.lineNumber,
    startColumn: word.startColumn,
    endColumn: word.endColumn,
  };

  const suggestions: any[] = [];

  // Check context
  const insideExpression = /\{\{[^}]*$/.test(textUntilPosition) || /\{%[^%]*$/.test(textUntilPosition);
  const accessBase = propertyAccess?.base;

  // Find which table we're iterating over (for item./row. completions)
  const fullText = model.getValue();
  const cursorOffset = model.getOffsetAt(position);
  const loopScopes = getLoopScopes(fullText);
  let currentTableKey: string | null = null;
  if (accessBase) {
    currentTableKey = findLoopTableKey(loopScopes, cursorOffset, accessBase);
  }

  console.log('Autocomplete context:', {
    insideExpression,
    accessBase,
    accessProperty: propertyAccess?.property ?? '',
    currentTableKey,
    availableTables: autocompleteData.tables.map(t => t.key),
    availableGlobals: autocompleteData.globals.map(g => g.key),
  });

  if (accessBase === 'globals') {
    // Suggest global variable names
    autocompleteData.globals.forEach(global => {
      suggestions.push({
        label: global.key,
        kind: monacoInstance.languages.CompletionItemKind.Variable,
        detail: `Global: ${global.name}`,
        insertText: global.key,
        range: propertyAccess?.range ?? defaultRange,
      });
    });
    // If no globals, add a hint
    if (autocompleteData.globals.length === 0) {
      suggestions.push({
        label: '(no globals defined)',
        kind: monacoInstance.languages.CompletionItemKind.Text,
        detail: 'Add Named Ranges in Data tab',
        insertText: '',
        range: propertyAccess?.range ?? defaultRange,
      });
    }
  } else if (accessBase === 'data') {
    // Suggest table keys
    autocompleteData.tables.forEach(table => {
      if (propertyAccess?.accessType === 'dot' && !isValidIdentifier(table.key)) {
        return;
      }
      suggestions.push({
        label: table.key,
        kind: monacoInstance.languages.CompletionItemKind.Variable,
        detail: `Table: ${table.name} (${table.columns.length} columns)`,
        insertText: table.key,
        range: propertyAccess?.range ?? defaultRange,
      });
    });
    // If no tables, add a hint
    if (autocompleteData.tables.length === 0) {
      suggestions.push({
        label: '(no tables configured)',
        kind: monacoInstance.languages.CompletionItemKind.Text,
        detail: 'Check tables in Data tab',
        insertText: '',
        range: propertyAccess?.range ?? defaultRange,
      });
    }
  } else if (accessBase && currentTableKey) {
    // Suggest column names
    const table = autocompleteData.tables.find(t => t.key === currentTableKey);
    if (table) {
      console.log('Suggesting columns for table:', table.key, '; columns:', table.columns);
      table.columns.forEach(col => {
        suggestions.push({
          label: col,
          kind: monacoInstance.languages.CompletionItemKind.Field,
          detail: `Column from ${table.name}`,
          insertText: col,
          range: propertyAccess?.range ?? defaultRange,
        });
      });
    }
  } else if (accessBase && (accessBase === 'row' || accessBase === 'item')) {
    // Fallback for legacy row/item variable names
    autocompleteData.tables.forEach(table => {
      table.columns.forEach(col => {
        suggestions.push({
          label: col,
          kind: monacoInstance.languages.CompletionItemKind.Field,
          detail: `Column from ${table.name}`,
          insertText: col,
          range: propertyAccess?.range ?? defaultRange,
        });
      });
    });
  } else if (insideExpression) {
    // Inside {{ }} or {% %} - suggest top-level variables
    suggestions.push({
      label: 'globals',
      kind: monacoInstance.languages.CompletionItemKind.Module,
      detail: 'Global variables from Named Ranges',
      insertText: 'globals.',
      range: defaultRange,
    });
    suggestions.push({
      label: 'data',
      kind: monacoInstance.languages.CompletionItemKind.Module,
      detail: 'Table data',
      insertText: 'data.',
      range: defaultRange,
    });
    const activeLoopVars = getActiveLoopVariables(loopScopes, cursorOffset);
    activeLoopVars.forEach(varName => {
      suggestions.push({
        label: varName,
        kind: monacoInstance.languages.CompletionItemKind.Variable,
        detail: 'Current item in for loop',
        insertText: `${varName}.`,
        range: defaultRange,
      });
    });
    suggestions.push({
      label: 'loop',
      kind: monacoInstance.languages.CompletionItemKind.Variable,
      detail: 'Loop info: loop.index, loop.first, loop.last',
      insertText: 'loop.',
      range: defaultRange,
    });
    suggestions.push({
      label: 'loop.index',
      kind: monacoInstance.languages.CompletionItemKind.Property,
      detail: '1-based loop index',
      insertText: 'loop.index',
      range: defaultRange,
    });
    suggestions.push({
      label: 'loop.first',
      kind: monacoInstance.languages.CompletionItemKind.Property,
      detail: 'True if first iteration',
      insertText: 'loop.first',
      range: defaultRange,
    });
    suggestions.push({
      label: 'loop.last',
      kind: monacoInstance.languages.CompletionItemKind.Property,
      detail: 'True if last iteration',
      insertText: 'loop.last',
      range: defaultRange,
    });
    suggestions.push({
      label: 'currentdatetime',
      kind: monacoInstance.languages.CompletionItemKind.Variable,
      detail: 'Current datetime (filename-safe)',
      insertText: 'currentdatetime',
      range: defaultRange,
    });

    // Add direct table suggestions
    autocompleteData.tables.forEach(table => {
      suggestions.push({
        label: formatDataAccess(table.key),
        kind: monacoInstance.languages.CompletionItemKind.Variable,
        detail: `Table: ${table.name}`,
        insertText: formatDataAccess(table.key),
        range: defaultRange,
      });
    });

    // Add direct global suggestions
    autocompleteData.globals.forEach(global => {
      suggestions.push({
        label: formatGlobalsAccess(global.key),
        kind: monacoInstance.languages.CompletionItemKind.Variable,
        detail: `Global: ${global.name}`,
        insertText: formatGlobalsAccess(global.key),
        range: defaultRange,
      });
    });

    // Nunjucks filters (only suggest after a pipe or when typing)
    if (textUntilPosition.includes('|')) {
      const filters = ['safe', 'escape', 'upper', 'lower', 'capitalize', 'title', 'trim', 'default', 'first', 'last', 'length', 'join', 'sort', 'reverse', 'round', 'int', 'float', 'string'];
      filters.forEach(filter => {
        suggestions.push({
          label: filter,
          kind: monacoInstance.languages.CompletionItemKind.Function,
          detail: `Filter: | ${filter}`,
          insertText: filter,
          range: defaultRange,
        });
      });
    }
  } else {
    // Outside Nunjucks blocks - suggest opening a block
    suggestions.push({
      label: '{{',
      kind: monacoInstance.languages.CompletionItemKind.Snippet,
      detail: 'Nunjucks expression',
      insertText: '{{ $0 }}',
      insertTextRules: monacoInstance.languages.CompletionItemInsertTextRule.InsertAsSnippet,
      range: defaultRange,
    });
    suggestions.push({
      label: '{%',
      kind: monacoInstance.languages.CompletionItemKind.Snippet,
      detail: 'Nunjucks statement',
      insertText: '{% $0 %}',
      insertTextRules: monacoInstance.languages.CompletionItemInsertTextRule.InsertAsSnippet,
      range: defaultRange,
    });
    suggestions.push({
      label: '{% for %}',
      kind: monacoInstance.languages.CompletionItemKind.Snippet,
      detail: 'For loop over table',
      insertText: "{% for item in data['${1:tableName}'] %}\n$0\n{% endfor %}",
      insertTextRules: monacoInstance.languages.CompletionItemInsertTextRule.InsertAsSnippet,
      range: defaultRange,
    });
    suggestions.push({
      label: '{% if %}',
      kind: monacoInstance.languages.CompletionItemKind.Snippet,
      detail: 'If statement',
      insertText: '{% if ${1:condition} %}\n$0\n{% endif %}',
      insertTextRules: monacoInstance.languages.CompletionItemInsertTextRule.InsertAsSnippet,
      range: defaultRange,
    });
  }

  return { suggestions };
}

/**
 * Get hover information for Nunjucks syntax
 */
function getNunjucksHover(model: any, position: any): any {
  const word = model.getWordAtPosition(position);
  if (!word) return null;

  const wordText = word.word;

  // Check for globals
  if (wordText === 'globals') {
    return {
      contents: [
        { value: '**globals**' },
        { value: 'Access global variables from Excel Named Ranges.\n\nExample: `{{ globals.projectName }}`' },
      ],
    };
  }

  // Check for data
  if (wordText === 'data') {
    return {
      contents: [
        { value: '**data**' },
        { value: 'Access table data.\n\nExample: `{% for row in data.products %}`' },
      ],
    };
  }

  // Check for row
  if (wordText === 'row') {
    return {
      contents: [
        { value: '**row**' },
        { value: 'Current row in a for loop.\n\nExample: `{{ row.ProductName }}`' },
      ],
    };
  }

  // Check for loop
  if (wordText === 'loop') {
    return {
      contents: [
        { value: '**loop**' },
        { value: 'Loop information object:\n- `loop.index`: 1-based index\n- `loop.index0`: 0-based index\n- `loop.first`: true if first iteration\n- `loop.last`: true if last iteration\n- `loop.length`: total items' },
      ],
    };
  }

  // Check for currentdatetime
  if (wordText === 'currentdatetime') {
    return {
      contents: [
        { value: '**currentdatetime**' },
        { value: 'Current datetime formatted for filenames (YYYY-MM-DD_HH-mm-ss).\n\nExample: `{{ currentdatetime }}`' },
      ],
    };
  }

  return null;
}

function formatDataAccess(key: string): string {
  return isValidIdentifier(key) ? `data.${key}` : `data['${key}']`;
}

function formatGlobalsAccess(key: string): string {
  return isValidIdentifier(key) ? `globals.${key}` : `globals['${key}']`;
}

function isValidIdentifier(value: string): boolean {
  return /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(value);
}

function getPropertyAccessAtCursor(
  model: any,
  position: any
): { base: string; property: string; range: IRange; accessType: 'dot' | 'bracket' } | null {
  const lineTextUntil = model.getLineContent(position.lineNumber).slice(0, position.column - 1);
  const dotMatch = lineTextUntil.match(/\b(\w+)\.(\w*)$/);
  if (dotMatch) {
    const property = dotMatch[2] ?? '';
    return {
      base: dotMatch[1],
      property,
      range: {
        startLineNumber: position.lineNumber,
        endLineNumber: position.lineNumber,
        startColumn: position.column - property.length,
        endColumn: position.column,
      },
      accessType: 'dot',
    };
  }

  const bracketMatch = lineTextUntil.match(/\b(\w+)\[['"]([^'"]*)$/);
  if (bracketMatch) {
    const property = bracketMatch[2] ?? '';
    return {
      base: bracketMatch[1],
      property,
      range: {
        startLineNumber: position.lineNumber,
        endLineNumber: position.lineNumber,
        startColumn: position.column - property.length,
        endColumn: position.column,
      },
      accessType: 'bracket',
    };
  }

  return null;
}

function getLoopScopes(text: string): LoopScope[] {
  const scopes: LoopScope[] = [];
  const stack: LoopScope[] = [];
  const tagRegex = /\{%\s*(for|endfor)\b[^%]*%\}/g;
  let match: RegExpExecArray | null;

  while ((match = tagRegex.exec(text)) !== null) {
    const tagText = match[0];
    if (match[1] === 'for') {
      const forMatch = tagText.match(
        /for\s+(\w+)\s+in\s+(?:data\.(\w+)|data\[['"]([^'"]+)['"]\]|(\w+))/
      );
      if (forMatch) {
        const tableKey = forMatch[2] || forMatch[3] || forMatch[4];
        if (!tableKey) {
          continue;
        }
        stack.push({
          varName: forMatch[1],
          tableKey,
          startIndex: match.index,
          endIndex: text.length,
        });
      }
      continue;
    }

    const scope = stack.pop();
    if (scope) {
      scope.endIndex = match.index + tagText.length;
      scopes.push(scope);
    }
  }

  while (stack.length > 0) {
    const scope = stack.pop();
    if (scope) {
      scopes.push(scope);
    }
  }

  return scopes;
}

function getActiveLoopVariables(scopes: LoopScope[], offset: number): string[] {
  const active = scopes.filter(scope => offset >= scope.startIndex && offset <= scope.endIndex);
  const uniqueVars = new Set<string>();
  active.forEach(scope => uniqueVars.add(scope.varName));
  return Array.from(uniqueVars);
}

function findLoopTableKey(scopes: LoopScope[], offset: number, varName?: string): string | null {
  const candidates = scopes.filter(scope => {
    if (offset < scope.startIndex || offset > scope.endIndex) {
      return false;
    }
    if (varName && scope.varName !== varName) {
      return false;
    }
    return true;
  });

  if (candidates.length === 0) {
    return null;
  }

  let best = candidates[0];
  for (const scope of candidates) {
    if (scope.startIndex > best.startIndex) {
      best = scope;
    }
  }

  return best.tableKey;
}

/**
 * Apply Nunjucks diagnostics for missing globals/tables/columns
 */
export function applyNunjucksDiagnostics(
  editor: any,
  monacoInstance: any
): void {
  const model = editor.getModel();
  if (!model) return;

  const text = model.getValue();
  const markers: any[] = [];
  const tablesByKey = new Map(autocompleteData.tables.map(table => [table.key, table]));
  const globalsSet = new Set(autocompleteData.globals.map(global => global.key));
  const loopScopes = getLoopScopes(text);

  const pushMarker = (startIndex: number, endIndex: number, message: string) => {
    const startPos = model.getPositionAt(startIndex);
    const endPos = model.getPositionAt(endIndex);
    markers.push({
      severity: monacoInstance.MarkerSeverity.Error,
      message,
      startLineNumber: startPos.lineNumber,
      startColumn: startPos.column,
      endLineNumber: endPos.lineNumber,
      endColumn: endPos.column,
    });
  };

  const globalsRegex = /\bglobals\.(\w+)\b/g;
  let match: RegExpExecArray | null;
  while ((match = globalsRegex.exec(text)) !== null) {
    const key = match[1];
    if (!globalsSet.has(key)) {
      const startIndex = match.index + 'globals.'.length;
      pushMarker(startIndex, startIndex + key.length, `Unknown global "${key}"`);
    }
  }
  const globalsBracketRegex = /\bglobals\[['"]([^'"]+)['"]\]/g;
  while ((match = globalsBracketRegex.exec(text)) !== null) {
    const key = match[1];
    if (!globalsSet.has(key)) {
      const startIndex = match.index + "globals['".length;
      pushMarker(startIndex, startIndex + key.length, `Unknown global "${key}"`);
    }
  }

  const dataRegex = /\bdata\.(\w+)\b/g;
  while ((match = dataRegex.exec(text)) !== null) {
    const tableKey = match[1];
    if (!tablesByKey.has(tableKey)) {
      const startIndex = match.index + 'data.'.length;
      pushMarker(startIndex, startIndex + tableKey.length, `Unknown table "${tableKey}"`);
    }
  }
  const dataBracketRegex = /\bdata\[['"]([^'"]+)['"]\]/g;
  while ((match = dataBracketRegex.exec(text)) !== null) {
    const tableKey = match[1];
    if (!tablesByKey.has(tableKey)) {
      const startIndex = match.index + "data['".length;
      pushMarker(startIndex, startIndex + tableKey.length, `Unknown table "${tableKey}"`);
    }
  }

  const columnDotRegex = /\b(\w+)\.(\w+)\b/g;
  while ((match = columnDotRegex.exec(text)) !== null) {
    const varName = match[1];
    if (varName === 'data' || varName === 'globals') {
      continue;
    }
    const column = match[2];
    const tableKey = findLoopTableKey(loopScopes, match.index, varName);
    if (!tableKey) {
      continue;
    }
    const table = tablesByKey.get(tableKey);
    if (!table || !table.columns.includes(column)) {
      const startIndex = match.index + `${varName}.`.length;
      pushMarker(startIndex, startIndex + column.length, `Unknown column "${column}"`);
    }
  }
  const columnBracketRegex = /\b(\w+)\[['"]([^'"]+)['"]\]/g;
  while ((match = columnBracketRegex.exec(text)) !== null) {
    const varName = match[1];
    if (varName === 'data' || varName === 'globals') {
      continue;
    }
    const column = match[2];
    const tableKey = findLoopTableKey(loopScopes, match.index, varName);
    if (!tableKey) {
      continue;
    }
    const table = tablesByKey.get(tableKey);
    if (!table || !table.columns.includes(column)) {
      const startIndex = match.index + `${varName}['`.length;
      pushMarker(startIndex, startIndex + column.length, `Unknown column "${column}"`);
    }
  }

  loopScopes.forEach(scope => {
    if (tablesByKey.has(scope.tableKey)) {
      return;
    }
    const loopText = text.slice(scope.startIndex, scope.endIndex);
    const tableMatch = loopText.match(/in\s+(?:data\.(\w+)|data\[['"]([^'"]+)['"]\]|(\w+))/);
    if (!tableMatch) return;
    const tableKey = tableMatch[1] || tableMatch[2] || tableMatch[3];
    const tableOffset = loopText.indexOf(tableKey);
    if (tableOffset < 0) return;
    const startIndex = scope.startIndex + tableOffset;
    pushMarker(startIndex, startIndex + tableKey.length, `Unknown table "${tableKey}"`);
  });

  monacoInstance.editor.setModelMarkers(model, 'nunjucks', markers);
}

// Store decoration IDs per editor instance
const editorDecorations = new WeakMap<any, string[]>();

/**
 * Apply Nunjucks syntax decorations to the editor
 * Call this after content changes to highlight Nunjucks syntax
 */
export function applyNunjucksDecorations(
  editor: any,
  monacoInstance: any
): void {
  const model = editor.getModel();
  if (!model) return;

  const text = model.getValue();
  const decorations: any[] = [];

  // Match {{ expression }}
  const expressionRegex = /\{\{[\s\S]*?\}\}/g;
  let match;
  while ((match = expressionRegex.exec(text)) !== null) {
    const startPos = model.getPositionAt(match.index);
    const endPos = model.getPositionAt(match.index + match[0].length);
    decorations.push({
      range: new monacoInstance.Range(
        startPos.lineNumber,
        startPos.column,
        endPos.lineNumber,
        endPos.column
      ),
      options: {
        isWholeLine: false,
        className: 'nunjucks-expression-line',
        inlineClassName: 'nunjucks-expression-inline',
        overviewRuler: {
          color: '#9932CC',
          position: monacoInstance.editor.OverviewRulerLane.Right,
        },
        minimap: {
          color: '#9932CC',
          position: monacoInstance.editor.MinimapPosition.Inline,
        },
        hoverMessage: { value: '**Nunjucks Expression** `{{ }}`' },
      },
    });
  }

  // Match {% statement %}
  const statementRegex = /\{%[\s\S]*?%\}/g;
  while ((match = statementRegex.exec(text)) !== null) {
    const startPos = model.getPositionAt(match.index);
    const endPos = model.getPositionAt(match.index + match[0].length);
    decorations.push({
      range: new monacoInstance.Range(
        startPos.lineNumber,
        startPos.column,
        endPos.lineNumber,
        endPos.column
      ),
      options: {
        isWholeLine: false,
        className: 'nunjucks-statement-line',
        inlineClassName: 'nunjucks-statement-inline',
        overviewRuler: {
          color: '#0000FF',
          position: monacoInstance.editor.OverviewRulerLane.Right,
        },
        hoverMessage: { value: '**Nunjucks Statement** `{% %}`' },
      },
    });
  }

  // Match {# comment #}
  const commentRegex = /\{#[\s\S]*?#\}/g;
  while ((match = commentRegex.exec(text)) !== null) {
    const startPos = model.getPositionAt(match.index);
    const endPos = model.getPositionAt(match.index + match[0].length);
    decorations.push({
      range: new monacoInstance.Range(
        startPos.lineNumber,
        startPos.column,
        endPos.lineNumber,
        endPos.column
      ),
      options: {
        isWholeLine: false,
        className: 'nunjucks-comment-line',
        inlineClassName: 'nunjucks-comment-inline',
        hoverMessage: { value: '**Nunjucks Comment** `{# #}`' },
      },
    });
  }

  // Get old decorations and apply new ones
  const oldDecorations = editorDecorations.get(editor) || [];
  const newDecorations = editor.deltaDecorations(oldDecorations, decorations);
  editorDecorations.set(editor, newDecorations);
}
