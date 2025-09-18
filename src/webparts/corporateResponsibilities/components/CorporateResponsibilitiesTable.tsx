import * as React from 'react';
import { IColumn } from '@fluentui/react/lib/DetailsList';
import { SortableDetailsList } from '../../../globalCommon/GlobalTable';
import { useState, useEffect } from 'react';
import { spfi, SPFx } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export interface IContactItem {
    key: number;
    what: string;
    who: string;
}

const CorporateResponsibilitiesTable = (props: any) => {
    const [data, setData] = useState<any[]>([]);
    const [isLoading, setIsLoading] = useState<boolean>(true);

    useEffect(() => {
        loadData();
    }, []);

    const loadData = async () => {
        try {
            setIsLoading(true);
            // Create sp instance pointing to the specific site
            const sp = spfi("https://vaughnconstruction.sharepoint.com").using(SPFx(props.Context));
            
            const items = await sp.web.lists
                .getById('8c97c9e1-ecaf-4fcb-93a9-4016509d185c')
                .items
                .select('Id,Title,Who')
                .top(5000)();

            setData(items);
        } catch (error) {
            console.error('Error loading data:', error);
        } finally {
            setIsLoading(false);
        }
    };

    const columns: IColumn[] = [
        {
            key: 'column1',
            name: 'What',
            fieldName: 'Title',
            onRender: (item: any) => (
                <a
                    target="_blank"
                    href={`https://vaughnconstruction.sharepoint.com/Lists/WhoDoesWhat/DispForm.aspx?ID=${item.Id}`}
                    data-interception="off"
                    style={{ 
                        color: '#0078d4', 
                        textDecoration: 'none',
                        fontSize: '14px'
                    }}
                    onMouseOver={(e) => (e.target as HTMLElement).style.textDecoration = 'underline'}
                    onMouseOut={(e) => (e.target as HTMLElement).style.textDecoration = 'none'}
                >
                    {item.Title}
                </a>
            ),
            minWidth: 400,
            maxWidth: 600,
            isResizable: true,
            isSorted: false,
            isSortedDescending: false,
        },
        {
            key: 'column2',
            name: 'Who',
            fieldName: 'Who',
            minWidth: 200,
            maxWidth: 300,
            isResizable: true,
            isSorted: false,
            isSortedDescending: false,
            onRender: (item: any) => (
                <span style={{ fontSize: '14px' }}>
                    {item.Who || 'Not assigned'}
                </span>
            ),
        },
    ];

    const sectionStyle: React.CSSProperties = {
        padding: '20px',
        backgroundColor: '#ffffff'
    };

    const titleStyle: React.CSSProperties = {
        fontSize: '24px',
        fontWeight: '600',
        color: '#323130',
        marginBottom: '16px',
        fontFamily: 'Segoe UI, system-ui, sans-serif'
    };

    return (
        <section style={sectionStyle}>
            <h2 style={titleStyle}>Corporate Responsibilities</h2>
            <SortableDetailsList
                items={data}
                columns={columns}
                isLoading={isLoading}
                // Let the table calculate its own height based on content
                // maxHeight="600px" // Remove this to allow automatic height calculation
                rowHeight={48} // Slightly larger row height for better readability
                headerHeight={40} // Slightly larger header height
            />
            {data.length > 0 && (
                <div style={{ 
                    marginTop: '12px', 
                    fontSize: '12px', 
                    color: '#605e5c',
                    textAlign: 'right'
                }}>
                    Showing {data.length} {data.length === 1 ? 'item' : 'items'}
                </div>
            )}
        </section>
    );
};

export default CorporateResponsibilitiesTable;