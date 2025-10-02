import * as React from 'react';
import { SortableDetailsList } from '../../../globalCommon/GlobalTable';
import { useState, useEffect } from 'react';
import { spfi, SPFx } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/security";
import { PermissionKind } from "@pnp/sp/security";
const CorporateResponsibilitiesTable = (props: any) => {
    const [data, setData] = useState<any[]>([]);
    const [isLoading, setIsLoading] = useState<boolean>(true);
    const [canEdit, setCanEdit] = useState<boolean>(false);
    useEffect(() => {
        loadData();
        checkPermissions();
    }, []);

    const loadData = async () => {
        try {
            setIsLoading(true);
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
 const checkPermissions = async () => {
            try {
                const sp = spfi("https://vaughnconstruction.sharepoint.com").using(SPFx(props.Context));
                const list = sp.web.lists.getById('8c97c9e1-ecaf-4fcb-93a9-4016509d185c');
                const perms = await list.getCurrentUserEffectivePermissions();
                if (sp.web.hasPermissions(perms, PermissionKind.EditListItems)) {
                    setCanEdit(true);
                }
            } catch (error) {
                console.error("Error checking permissions:", error);
            }
        };
    // Columns adapted to SortableDetailsList format
  const columns = [
  {
    key: 'what',
    name: 'What',
    fieldName: 'Title',
    minWidth: 400,
    isResizable: true,
    isSorted: false,
    isSortedDescending: false,
    // add custom renderer
    onRender: (item: any) => (
      <>{canEdit? <a
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
      </a> : <span style={{fontSize: '14px'}}>{item.Title}</span>}
     
      </>
    )
  },
  {
    key: 'who',
    name: 'Who',
    fieldName: 'Who',
    minWidth: 200,
    isResizable: true,
    isSorted: false,
    isSortedDescending: false,
  },
];


    const sectionStyle: React.CSSProperties = {
        padding: '20px',
        backgroundColor: '#ffffff'
    };

    const titleStyle: React.CSSProperties = {
        fontSize: '24pt',
        fontWeight: '600',
        color: '#323130',
        marginBottom: '16px',
        fontFamily: 'Segoe UI, system-ui, sans-serif'
    };

    return (
        <section style={sectionStyle}>
             {canEdit ? (
                  <a href="https://vaughnconstruction.sharepoint.com/Lists/WhoDoesWhat/AllItems.aspx" target="_blank" data-interception="off" style={{ textDecoration: 'none' }}>
                      <h2 style={titleStyle}>Corporate Responsibilities</h2>
                  </a>
              ) : (
                  <h2 style={titleStyle}>Corporate Responsibilities</h2>
              )}
            <SortableDetailsList
                items={data}
                columns={columns}
                isLoading={isLoading}
                rowHeight={48}
                headerHeight={40}
            />
        </section>
    );
};

export default CorporateResponsibilitiesTable;
