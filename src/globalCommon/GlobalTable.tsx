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

// Define the props for our component
export interface ISortableDetailsListProps {
  items: any[]; // The data to display
  columns: Array<{
    key: string;
    name: string;
    fieldName: string;
    minWidth?: number;
    maxWidth?: number;
    isResizable?: boolean;
    isSorted?: boolean;
    isSortedDescending?: boolean;
  }>; // Column configuration
  isLoading?: boolean; // Optional flag for showing a loading shimmer
  maxHeight?: string; // Optional max height, defaults to calculated height
  rowHeight?: number; // Height of each row (default: 42px)
  headerHeight?: number; // Height of header (default: 32px)
}

export const SortableDetailsList: React.FunctionComponent<ISortableDetailsListProps> = (props) => {
  const [calculatedHeight, setCalculatedHeight] = useState<string>('400px');
  const containerRef = useRef<HTMLDivElement>(null);

  const rowHeight = props.rowHeight || 42;
  const headerHeight = props.headerHeight || 48; // Fluent UI v9 header is taller

  // Calculate optimal table height based on content
  useEffect(() => {
    if (props.items.length > 0 && !props.maxHeight) {
      const contentHeight = (props.items.length * rowHeight) + headerHeight + 20;
      const availableHeight = window.innerHeight;
      const reservedSpace = 200;
      const maxAllowedHeight = availableHeight - reservedSpace;
      const optimalHeight = Math.min(contentHeight, maxAllowedHeight);
      const finalHeight = Math.max(optimalHeight, 300);
      
      setCalculatedHeight(`${finalHeight}px`);
    }
  }, [props.items.length, props.maxHeight, rowHeight, headerHeight]);

  // Add CSS for sticky header and scrolling
  useEffect(() => {
    const styleId = 'fluent-table-styles';
    let existingStyle = document.getElementById(styleId);
    
    if (!existingStyle) {
      const style = document.createElement('style');
      style.id = styleId;
      style.innerHTML = `
        .scrollable-table-container {
          border: 1px solid var(--colorNeutralStroke2);
          border-radius: var(--borderRadiusMedium);
          background-color: var(--colorNeutralBackground1);
          position: relative;
          overflow: hidden;
          display: flex;
          flex-direction: column;
        }

        .scrollable-table-container .fui-Table {
          height: 100%;
          display: flex;
          flex-direction: column;
        }

        .scrollable-table-container .fui-TableHeader {
          flex-shrink: 0;
          position: sticky;
          top: 0;
          z-index: 100;
          background-color: var(--colorNeutralBackground2);
          border-bottom: 1px solid var(--colorNeutralStroke2);
          box-shadow: 0 1px 2px rgba(0,0,0,0.1);
        }

        .scrollable-table-container .fui-TableBody {
          flex: 1;
          overflow-y: auto;
          overflow-x: hidden;
        }

        .scrollable-table-container .fui-TableRow {
          border-bottom: 1px solid var(--colorNeutralStroke3);
        }

        .scrollable-table-container .fui-TableRow:hover {
          background-color: var(--colorSubtleBackgroundHover);
        }

        /* Loading overlay */
        .table-loading-overlay {
          position: absolute;
          top: 0;
          left: 0;
          right: 0;
          bottom: 0;
          background-color: rgba(255, 255, 255, 0.8);
          display: flex;
          align-items: center;
          justify-content: center;
          z-index: 200;
          font-size: 16px;
          color: var(--colorNeutralForeground1);
        }

        /* Custom scrollbar */
        .scrollable-table-container .fui-TableBody::-webkit-scrollbar {
          width: 8px;
        }

        .scrollable-table-container .fui-TableBody::-webkit-scrollbar-track {
          background: var(--colorNeutralBackground3);
          border-radius: 4px;
        }

        .scrollable-table-container .fui-TableBody::-webkit-scrollbar-thumb {
          background: var(--colorNeutralStroke1);
          border-radius: 4px;
        }

        .scrollable-table-container .fui-TableBody::-webkit-scrollbar-thumb:hover {
          background: var(--colorNeutralStroke2);
        }
      `;
      document.head.appendChild(style);
    }

    return () => {
      const styleElement = document.getElementById(styleId);
      if (styleElement) {
        styleElement.remove();
      }
    };
  }, []);

  // Convert props.columns to Fluent UI v9 column definitions
  const tableColumns: TableColumnDefinition<any>[] = props.columns.map(col => 
    createTableColumn<any>({
      columnId: col.key,
      compare: (a, b) => {
        const aVal = a[col.fieldName];
        const bVal = b[col.fieldName];
        
        // Handle null/undefined values
        if (aVal == null && bVal == null) return 0;
        if (aVal == null) return -1;
        if (bVal == null) return 1;
        
        // Convert to strings for comparison
        const aStr = String(aVal).toLowerCase();
        const bStr = String(bVal).toLowerCase();
        
        return aStr.localeCompare(bStr);
      },
    })
  );

  const {
    getRows,
    sort: { getSortDirection, toggleColumnSort, sort },
  } = useTableFeatures(
    {
      columns: tableColumns,
      items: props.items,
    },
    [
      useTableSort({
        defaultSortState: props.columns.find(col => col.isSorted) 
          ? { 
              sortColumn: props.columns.find(col => col.isSorted)!.key, 
              sortDirection: props.columns.find(col => col.isSorted)!.isSortedDescending ? "descending" : "ascending" 
            }
          : undefined,
      }),
    ]
  );

  const headerSortProps = (columnId: TableColumnId) => ({
    onClick: (e: React.MouseEvent) => {
      toggleColumnSort(e, columnId);
    },
    sortDirection: getSortDirection(columnId),
  });

  const rows = sort(getRows());
  const containerHeight = props.maxHeight || calculatedHeight;

const renderCellContent = (item: any, fieldName: string, column: any) => {
  // If column has custom renderer, use it
  if (column.onRender) {
    return column.onRender(item);
  }

  const value = item[fieldName];

  if (React.isValidElement(value)) return value;
  if (value == null) return '';
  if (typeof value === 'boolean') return value ? 'Yes' : 'No';
  if (typeof value === 'number') return value.toLocaleString();

  return String(value);
};



  return (
    <div 
      ref={containerRef} 
      className="scrollable-table-container"
      style={{ 
        height: containerHeight,
        width: '100%'
      }}
    >
      <Table 
        sortable 
        aria-label="Sortable table with scrolling"
        style={{ height: '100%' }}
      >
        <TableHeader>
          <TableRow>
            {props.columns.map((column) => (
              <TableHeaderCell 
                key={column.key}
                {...headerSortProps(column.key)}
                style={{
                  minWidth: column.minWidth,
                  maxWidth: column.maxWidth,
                  width: column.maxWidth || column.minWidth
                }}
              >
                {column.name}
              </TableHeaderCell>
            ))}
          </TableRow>
        </TableHeader>
        <TableBody>
          {rows.map(({ item }, index) => (
            <TableRow key={`row-${index}`}>
             {props.columns.map((column) => (
  <TableCell key={`${index}-${column.key}`} style={{ minWidth: column.minWidth, maxWidth: column.maxWidth ,width: column.maxWidth || column.minWidth}}>
    <TableCellLayout>
      {renderCellContent(item, column.fieldName, column)}
    </TableCellLayout>
  </TableCell>
))}

            </TableRow>
          ))}
        </TableBody>
      </Table>
      
      {props.isLoading && (
        <div className="table-loading-overlay">
          Loading...
        </div>
      )}
    </div>
  );
};