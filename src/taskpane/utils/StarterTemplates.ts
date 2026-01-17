/**
 * Starter Templates Utility
 *
 * Auto-generates starter templates for different file formats based on
 * available Excel tables. Supports JSON, CSV, XML, HTML, and plain text.
 */

import type { TableInfo } from '@/services/TableService';

/**
 * Generates a starter template based on file extension
 *
 * @param filename - Output filename (extension determines format)
 * @param tables - Available Excel tables
 * @returns Generated template string
 */
export function generateStarterTemplate(filename: string, tables: TableInfo[]): string {
  const extension = filename.split('.').pop()?.toLowerCase() || 'txt';
  const firstTable = tables.length > 0 ? tables[0] : null;

  switch (extension) {
    case 'json':
      return generateJsonTemplate(firstTable);
    case 'csv':
      return generateCsvTemplate(firstTable);
    case 'xml':
      return generateXmlTemplate(firstTable);
    case 'html':
    case 'htm':
      return generateHtmlTemplate(firstTable);
    case 'txt':
    default:
      return generateTextTemplate(firstTable);
  }
}

/**
 * Generates a JSON array template
 *
 * Example output:
 * [
 *   {% for row in products %}
 *   {
 *     "name": "{{ row.ProductName }}",
 *     "price": {{ row.Price }}
 *   }{{ "," if not loop.last }}
 *   {% endfor %}
 * ]
 */
function generateJsonTemplate(table: TableInfo | null): string {
  if (!table) {
    return `[\n  {}\n]`;
  }

  const tableKey = table.name;
  const tableAccess = formatDataAccess(tableKey);
  const columns = table.columns;

  // Build JSON object properties
  const properties = columns.map((col, index) => {
    const isLast = index === columns.length - 1;
    const value = isNumericColumn(col) ? `{{ row.${col} }}` : `"{{ row.${col} }}"`;
    return `    "${col}": ${value}${isLast ? '' : ','}`;
  }).join('\n');

  return `[\n  {% for row in ${tableAccess} %}\n  {\n${properties}\n  }{{ "," if not loop.last }}\n  {% endfor %}\n]`;
}

/**
 * Generates a CSV template
 *
 * Example output:
 * ProductName,Price,Quantity
 * {% for row in products %}
 * {{ row.ProductName }},{{ row.Price }},{{ row.Quantity }}
 * {% endfor %}
 */
function generateCsvTemplate(table: TableInfo | null): string {
  if (!table) {
    return `Column1,Column2,Column3\nValue1,Value2,Value3`;
  }

  const tableKey = table.name;
  const tableAccess = formatDataAccess(tableKey);
  const columns = table.columns;

  // Header row
  const header = columns.join(',');

  // Data rows with Nunjucks loop
  const dataRow = columns.map(col => `{{ row.${col} }}`).join(',');

  return `${header}\n{% for row in ${tableAccess} %}\n${dataRow}\n{% endfor %}`;
}

/**
 * Generates an XML template
 *
 * Example output:
 * <?xml version="1.0" encoding="UTF-8"?>
 * <products>
 *   {% for row in products %}
 *   <product>
 *     <ProductName>{{ row.ProductName }}</ProductName>
 *     <Price>{{ row.Price }}</Price>
 *   </product>
 *   {% endfor %}
 * </products>
 */
function generateXmlTemplate(table: TableInfo | null): string {
  if (!table) {
    return `<?xml version="1.0" encoding="UTF-8"?>\n<data>\n  <item>\n    <field>value</field>\n  </item>\n</data>`;
  }

  const tableKey = table.name;
  const tableAccess = formatDataAccess(tableKey);
  const safeElementBase = generateTableKey(tableKey) || "data";
  const rootElement = pluralize(safeElementBase);
  const itemElement = singularize(safeElementBase);
  const columns = table.columns;

  // Build XML elements for each column
  const elements = columns.map(col => {
    const elementName = col.replace(/[^a-zA-Z0-9]/g, '_');
    return `    <${elementName}>{{ row.${col} }}</${elementName}>`;
  }).join('\n');

  return `<?xml version="1.0" encoding="UTF-8"?>\n<${rootElement}>\n  {% for row in ${tableAccess} %}\n  <${itemElement}>\n${elements}\n  </${itemElement}>\n  {% endfor %}\n</${rootElement}>`;
}

/**
 * Generates an HTML table template
 *
 * Example output:
 * <!DOCTYPE html>
 * <html>
 * <body>
 *   <table border="1">
 *     <thead>
 *       <tr><th>ProductName</th><th>Price</th></tr>
 *     </thead>
 *     <tbody>
 *       {% for row in products %}
 *       <tr><td>{{ row.ProductName }}</td><td>{{ row.Price }}</td></tr>
 *       {% endfor %}
 *     </tbody>
 *   </table>
 * </body>
 * </html>
 */
function generateHtmlTemplate(table: TableInfo | null): string {
  if (!table) {
    return `<!DOCTYPE html>\n<html>\n<body>\n  <h1>Data Export</h1>\n  <p>No tables available.</p>\n</body>\n</html>`;
  }

  const tableKey = table.name;
  const tableAccess = formatDataAccess(tableKey);
  const columns = table.columns;

  // Header cells
  const headerCells = columns.map(col => `<th>${col}</th>`).join('');

  // Data cells
  const dataCells = columns.map(col => `<td>{{ row.${col} }}</td>`).join('');

  return `<!DOCTYPE html>
<html>
<head>
  <title>${table.name} Export</title>
  <style>
    table { border-collapse: collapse; width: 100%; }
    th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
    th { background-color: #f2f2f2; font-weight: bold; }
  </style>
</head>
<body>
  <h1>${table.name}</h1>
  <table>
    <thead>
      <tr>${headerCells}</tr>
    </thead>
    <tbody>
      {% for row in ${tableAccess} %}
      <tr>${dataCells}</tr>
      {% endfor %}
    </tbody>
  </table>
</body>
</html>`;
}

/**
 * Generates a plain text template
 *
 * Example output:
 * Products List
 * =============
 *
 * {% for row in products %}
 * ProductName: {{ row.ProductName }}
 * Price: {{ row.Price }}
 * ---
 * {% endfor %}
 */
function generateTextTemplate(table: TableInfo | null): string {
  if (!table) {
    return `Data Export\n===========\n\nNo tables available.`;
  }

  const tableKey = table.name;
  const tableAccess = formatDataAccess(tableKey);
  const columns = table.columns;

  // Build field list
  const fields = columns.map(col => `${col}: {{ row.${col} }}`).join('\n');

  const title = table.name;
  const underline = '='.repeat(title.length);

  return `${title}\n${underline}\n\n{% for row in ${tableAccess} %}\n${fields}\n---\n{% endfor %}`;
}

/**
 * Generates a template key from a table name
 * Converts to lowercase and replaces non-alphanumeric characters with underscores
 *
 * @param tableName - Excel table name
 * @returns Template key (e.g., "ProductList" -> "productlist" or "product_list")
 */
export function generateTableKey(tableName: string): string {
  return tableName.toLowerCase().replace(/[^a-z0-9]/g, '_').replace(/_+/g, '_').replace(/^_|_$/g, '');
}

function formatDataAccess(tableName: string): string {
  return isValidIdentifier(tableName) ? `data.${tableName}` : `data['${tableName}']`;
}

function isValidIdentifier(value: string): boolean {
  return /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(value);
}

/**
 * Attempts to pluralize a word (simple English rules)
 */
function pluralize(word: string): string {
  if (word.endsWith('s')) return word;
  if (word.endsWith('y')) return word.slice(0, -1) + 'ies';
  return word + 's';
}

/**
 * Attempts to singularize a word (simple English rules)
 */
function singularize(word: string): string {
  if (word.endsWith('ies')) return word.slice(0, -3) + 'y';
  if (word.endsWith('s')) return word.slice(0, -1);
  return word;
}

/**
 * Heuristic to detect if a column likely contains numeric data
 * Based on column name patterns
 */
function isNumericColumn(columnName: string): boolean {
  const numericPatterns = [
    /price/i,
    /cost/i,
    /amount/i,
    /total/i,
    /quantity/i,
    /qty/i,
    /count/i,
    /number/i,
    /num/i,
    /id/i,
    /age/i,
    /year/i,
    /rate/i,
    /percent/i,
  ];

  return numericPatterns.some(pattern => pattern.test(columnName));
}
