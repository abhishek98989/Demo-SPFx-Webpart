import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import {
  Table,
  TableHeader,
  TableRow,
  TableHeaderCell,
  TableBody,
  TableCell,
  TableCellLayout,
  useTableFeatures,
  TableColumnDefinition,
  TableColumnId,
  useTableSort,
  createTableColumn,
} from '@fluentui/react-components';

export interface ISortableDetailsListProps {
  items: any[];
  columns: Array<{
    key: string;
    name: string;
    fieldName: string;
    minWidth?: number;
    maxWidth?: number;
    isResizable?: boolean;
    isSorted?: boolean;
    isSortedDescending?: boolean;
    onRender?: (item: any) => React.ReactNode;
  }>;
  isLoading?: boolean;
  constrainMode?: number;
  layoutMode?: number;
  maxHeight?: string;
  rowHeight?: number;
  headerHeight?: number;
}

// Compute a definitive pixel width for each column
const resolveColumnWidth = (col: ISortableDetailsListProps['columns'][number]): number => {
  if (col.maxWidth) return col.maxWidth;
  if (col.minWidth) return col.minWidth;
  return 150;
};

export const SortableDetailsList: React.FunctionComponent<ISortableDetailsListProps> = (props) => {
  const [calculatedHeight, setCalculatedHeight] = useState<string>('400px');
  const containerRef = useRef<HTMLDivElement>(null);

  const rowHeight = props.rowHeight || 42;
  const headerHeight = props.headerHeight || 48;

  // Pre-compute column widths once so header and body share identical values
  const columnWidths = React.useMemo(
    () => props.columns.map(resolveColumnWidth),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [props.columns.map(c => `${c.key}:${c.minWidth}:${c.maxWidth}`).join(',')]
  );

  const totalTableWidth = columnWidths.reduce((sum, w) => sum + w, 0);

  useEffect(() => {
    if (props.items.length > 0 && !props.maxHeight) {
      const contentHeight = props.items.length * rowHeight + headerHeight + 20;
      const maxAllowedHeight = window.innerHeight - 200;
      const finalHeight = Math.max(Math.min(contentHeight, maxAllowedHeight), 300);
      setCalculatedHeight(`${finalHeight}px`);
    }
  }, [props.items.length, props.maxHeight, rowHeight, headerHeight]);

  useEffect(() => {
    const styleId = 'fluent-table-styles';
    if (document.getElementById(styleId)) return;

    const style = document.createElement('style');
    style.id = styleId;
    style.innerHTML = `
      .scrollable-table-wrapper {
        border: 1px solid var(--colorNeutralStroke2);
        border-radius: var(--borderRadiusMedium);
        background-color: var(--colorNeutralBackground1);
        position: relative;
        overflow: hidden;
        display: flex;
        flex-direction: column;
      }

      /* ── Sticky-header scroll trick ─────────────────────────────────────
         The outer wrapper clips overflow.
         An inner scroll div handles both axes.
         The header row is sticky-top inside that scroll div.
      ─────────────────────────────────────────────────────────────────── */
      .scrollable-table-inner {
        overflow: auto;
        flex: 1;
      }

      /* Prevent Fluent from overriding our explicit cell widths */
      .scrollable-table-wrapper .fui-Table {
        border-collapse: collapse;
        /* width is set inline to match sum of columns */
      }

      .scrollable-table-wrapper .fui-TableHeader {
        position: sticky;
        top: 0;
        z-index: 100;
        background-color: var(--colorNeutralBackground2);
      }

      .scrollable-table-wrapper .fui-TableHeaderCell {
        box-sizing: border-box;
        overflow: hidden;
        white-space: nowrap;
        text-overflow: ellipsis;
        flex-shrink: 0;
        flex-grow: 0;
        align-items: center;
      }

      .scrollable-table-wrapper .fui-TableCell {
        box-sizing: border-box;
        overflow: visible;
        white-space: normal;
        word-break: break-word;
        overflow-wrap: break-word;
        flex-shrink: 0;
        flex-grow: 0;
        /* Let cell height grow with content */
        height: auto !important;
        align-items: flex-start;
      }

      /* Allow the cell layout inner div to wrap too */
      .scrollable-table-wrapper .fui-TableCell .fui-TableCellLayout {
        white-space: normal;
        word-break: break-word;
        overflow-wrap: break-word;
        width: 100%;
      }

      .scrollable-table-wrapper .fui-TableRow {
        border-bottom: 1px solid var(--colorNeutralStroke3);
        display: flex; /* keep cells in a row */
      }

      .scrollable-table-wrapper .fui-TableRow:hover {
        background-color: var(--colorSubtleBackgroundHover);
      }

      /* Loading overlay */
      .table-loading-overlay {
        position: absolute;
        inset: 0;
        background-color: rgba(255, 255, 255, 0.8);
        display: flex;
        align-items: center;
        justify-content: center;
        z-index: 200;
        font-size: 16px;
        color: var(--colorNeutralForeground1);
      }

      /* Scrollbar */
      .scrollable-table-inner::-webkit-scrollbar { width: 8px; height: 8px; }
      .scrollable-table-inner::-webkit-scrollbar-track { background: var(--colorNeutralBackground3); border-radius: 4px; }
      .scrollable-table-inner::-webkit-scrollbar-thumb { background: var(--colorNeutralStroke1); border-radius: 4px; }
      .scrollable-table-inner::-webkit-scrollbar-thumb:hover { background: var(--colorNeutralStroke2); }
    `;
    document.head.appendChild(style);

    return () => { document.getElementById(styleId)?.remove(); };
  }, []);

  const tableColumns: TableColumnDefinition<any>[] = props.columns.map(col =>
    createTableColumn<any>({
      columnId: col.key,
      compare: (a, b) => {
        const aVal = a[col.fieldName];
        const bVal = b[col.fieldName];
        if (aVal == null && bVal == null) return 0;
        if (aVal == null) return -1;
        if (bVal == null) return 1;
        return String(aVal).toLowerCase().localeCompare(String(bVal).toLowerCase());
      },
    })
  );

  const {
    getRows,
    sort: { getSortDirection, toggleColumnSort, sort },
  } = useTableFeatures(
    { columns: tableColumns, items: props.items },
    [
      useTableSort({
        defaultSortState: (() => {
          const sortedCol = props.columns.find(c => c.isSorted);
          return sortedCol
            ? { sortColumn: sortedCol.key, sortDirection: sortedCol.isSortedDescending ? 'descending' : 'ascending' }
            : undefined;
        })(),
      }),
    ]
  );

  const headerSortProps = (columnId: TableColumnId) => ({
    onClick: (e: React.MouseEvent) => toggleColumnSort(e, columnId),
    sortDirection: getSortDirection(columnId),
  });

  const rows = sort(getRows());
  const containerHeight = props.maxHeight || calculatedHeight;

  const renderCellContent = (item: any, col: ISortableDetailsListProps['columns'][number]) => {
    if (col.onRender) return col.onRender(item);
    const value = item[col.fieldName];
    if (React.isValidElement(value)) return value;
    if (value == null) return '';
    if (typeof value === 'boolean') return value ? 'Yes' : 'No';
    if (typeof value === 'number') return value.toLocaleString();
    return String(value);
  };

  // Shared style factory — same logic for header and body cells
  const cellStyle = (colIndex: number): React.CSSProperties => ({
    width: columnWidths[colIndex],
    minWidth: columnWidths[colIndex],
    maxWidth: columnWidths[colIndex],
    boxSizing: 'border-box',
    overflow: 'hidden',
    flexShrink: 0,
    flexGrow: 0,
  });

  return (
    <div
      ref={containerRef}
      className="scrollable-table-wrapper"
      style={{ height: containerHeight, width: '100%' }}
    >
      <div className="scrollable-table-inner">
        <Table
          sortable
          aria-label="Sortable data table"
          style={{
            width: totalTableWidth,
            minWidth: totalTableWidth,
            tableLayout: 'fixed',
          }}
        >
          <TableHeader>
            <TableRow style={{ display: 'flex', width: totalTableWidth }}>
              {props.columns.map((column, colIndex) => (
                <TableHeaderCell
                  key={column.key}
                  {...headerSortProps(column.key)}
                  style={cellStyle(colIndex)}
                >
                  {column.name}
                </TableHeaderCell>
              ))}
            </TableRow>
          </TableHeader>

          <TableBody>
            {rows.map(({ item }, rowIndex) => (
              <TableRow key={`row-${rowIndex}`} style={{ display: 'flex', width: totalTableWidth }}>
                {props.columns.map((column, colIndex) => (
                  <TableCell
                    key={`${rowIndex}-${column.key}`}
                    style={cellStyle(colIndex)}
                  >
                    <TableCellLayout>
                      {renderCellContent(item, column)}
                    </TableCellLayout>
                  </TableCell>
                ))}
              </TableRow>
            ))}
          </TableBody>
        </Table>
      </div>

      {props.isLoading && (
        <div className="table-loading-overlay">Loading…</div>
      )}
    </div>
  );
};