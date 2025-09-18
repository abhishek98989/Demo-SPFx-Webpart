import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import { DetailsList, IColumn, DetailsListLayoutMode, SelectionMode } from '@fluentui/react/lib/DetailsList';
import { ShimmeredDetailsList } from '@fluentui/react/lib/ShimmeredDetailsList';

// Define the props for our component
export interface ISortableDetailsListProps {
  items: any[]; // The data to display
  columns: IColumn[]; // The column configuration
  isLoading?: boolean; // Optional flag for showing a loading shimmer
  maxHeight?: string; // Optional max height, defaults to calculated height
  rowHeight?: number; // Height of each row (default: 42px for Fluent UI)
  headerHeight?: number; // Height of header (default: 32px for Fluent UI)
}

export const SortableDetailsList: React.FunctionComponent<ISortableDetailsListProps> = (props) => {
  // State to hold the items and columns
  const [items, setItems] = useState<any[]>(props.items);
  const [columns, setColumns] = useState<IColumn[]>(props.columns);
  const [calculatedHeight, setCalculatedHeight] = useState<string>('400px');
  const containerRef = useRef<HTMLDivElement>(null);

  const rowHeight = props.rowHeight || 42; // Default Fluent UI row height
  const headerHeight = props.headerHeight || 32; // Default Fluent UI header height

  // This effect updates the items when the props change
  useEffect(() => {
    setItems(props.items);
  }, [props.items]);

  // Calculate optimal table height based on content
  useEffect(() => {
    if (items.length > 0 && !props.maxHeight) {
      // Calculate height needed for all rows plus header
      const contentHeight = (items.length * rowHeight) + headerHeight;
      
      // Get available viewport height
      const availableHeight = window.innerHeight;
      
      // Reserve space for other page elements (adjust as needed)
      const reservedSpace = 200;
      const maxAllowedHeight = availableHeight - reservedSpace;
      
      // Use the smaller of content height or max allowed height
      const optimalHeight = Math.min(contentHeight, maxAllowedHeight);
      
      // Ensure minimum height
      const finalHeight = Math.max(optimalHeight, 200);
      
      setCalculatedHeight(`${finalHeight}px`);
    }
  }, [items.length, props.maxHeight, rowHeight, headerHeight]);

  // Add CSS for sticky header globally
  useEffect(() => {
    const styleId = 'sticky-header-fix';
    let existingStyle = document.getElementById(styleId);
    
    if (!existingStyle) {
      const style = document.createElement('style');
      style.id = styleId;
      style.innerHTML = `
        .sticky-table-container .ms-DetailsList-headerWrapper {
          position: sticky !important;
          top: 0 !important;
          z-index: 200 !important;
          background-color: #faf9f8 !important;
          border-bottom: 1px solid #edebe9 !important;
          box-shadow: 0 1px 2px rgba(0,0,0,0.1) !important;
        }
        
        .sticky-table-container .ms-DetailsList-contentWrapper {
          overflow-x: hidden !important;
          overflow-y: auto !important;
        }
        
        .sticky-table-container .ms-DetailsHeader {
          background-color: #faf9f8 !important;
        }
        
        .sticky-table-container .ms-DetailsHeader-cell {
          background-color: #faf9f8 !important;
        }
      `;
      document.head.appendChild(style);
    }

    return () => {
      // Cleanup on unmount
      const styleElement = document.getElementById(styleId);
      if (styleElement) {
        styleElement.remove();
      }
    };
  }, []);

  // The core sorting logic
  const onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const newColumns: IColumn[] = columns.slice();
    const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];
    
    // Toggle the sort direction or set it to ascending if it's a new column
    const newSortDirection = currColumn.isSortedDescending ? false : true;

    // Update the sorted column properties
    const sortedColumns = newColumns.map(col => {
      col.isSorted = (col.key === column.key);
      if (col.isSorted) {
        col.isSortedDescending = newSortDirection;
      }
      return col;
    });

    // Sort the items array
    const sortedItems = [...items].sort((a, b) => {
      const firstValue = a[column.fieldName!];
      const secondValue = b[column.fieldName!];
      
      // Handle null/undefined values
      if (firstValue == null && secondValue == null) return 0;
      if (firstValue == null) return newSortDirection ? 1 : -1;
      if (secondValue == null) return newSortDirection ? -1 : 1;
      
      // Convert to strings for comparison to handle mixed data types
      const firstStr = String(firstValue).toLowerCase();
      const secondStr = String(secondValue).toLowerCase();
      
      if (newSortDirection) { // Ascending
        return firstStr > secondStr ? 1 : firstStr < secondStr ? -1 : 0;
      } else { // Descending
        return firstStr < secondStr ? 1 : firstStr > secondStr ? -1 : 0;
      }
    });

    // Update the state with the new sorted columns and items
    setColumns(sortedColumns);
    setItems(sortedItems);
  };

  const tableContainerStyle: React.CSSProperties = {
    height: props.maxHeight || calculatedHeight,
    border: '1px solid #edebe9',
    borderRadius: '2px',
    backgroundColor: '#ffffff',
    overflow: 'hidden',
    position: 'relative'
  };

  return (
    <div 
      ref={containerRef} 
      className="sticky-table-container"
      style={tableContainerStyle}
    >
      <ShimmeredDetailsList
        items={items}
        columns={columns.map(col => ({ ...col, onColumnClick: onColumnClick }))}
        setKey="set"
        layoutMode={DetailsListLayoutMode.justified}
        selectionMode={SelectionMode.none}
        enableShimmer={props.isLoading}
        compact={false}
        detailsListStyles={{
          root: {
            height: '100%',
            overflow: 'hidden'
          },
          contentWrapper: {
            height: '100%',
            overflowX: 'hidden',
            overflowY: 'auto'
          }
        }}
      />
    </div>
  );
};