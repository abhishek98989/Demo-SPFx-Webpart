import React from 'react'
import { SortableDetailsList } from '../../../globalCommon/GlobalTable';
import { useState, useEffect } from 'react';
import { spfi, SPFx } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/security";
import { PermissionKind } from "@pnp/sp/security";

export const LocationsTable = (props:any) => {
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
                    .getById('450dda32-a7e7-4439-8779-e1cc1523e7fd')
                    .items
                    .select('Id,Title,Project_x0020_Location,Phone,Fax,Location')
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
                const list = sp.web.lists.getById('450dda32-a7e7-4439-8779-e1cc1523e7fd');
                const perms = await list.getCurrentUserEffectivePermissions();
                if (sp.web.hasPermissions(perms, PermissionKind.EditListItems)) {
                    setCanEdit(true);
                }
            } catch (error) {
                console.error("Error checking permissions:", error);
            }
        };
          const columns = [
          {
            key: 'OfficeProject',
            name: 'Office/Project',
            fieldName: 'Title',
            minWidth: 300,
            isResizable: true,
            isSorted: false,
            isSortedDescending: false,
          },
          {
            key: 'Address',
            name: 'Address',
            fieldName: 'Project_x0020_Location',
            minWidth: 250,
            isResizable: true,
            isSorted: false,
            isSortedDescending: false,
          },
           {
            key: 'Phone',
            name: 'Phone',
            fieldName: 'Phone',
            minWidth: 150,
            isResizable: true,
            isSorted: false,
            isSortedDescending: false,
          },
           {
            key: 'Fax',
            name: 'Fax',
            fieldName: 'Fax',
            minWidth: 150,
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
                  <a href="https://vaughnconstruction.sharepoint.com/Lists/Locations" target="_blank" data-interception="off" style={{ textDecoration: 'none' }}>
                      <h2 style={titleStyle}>Locations​​​​</h2>
                  </a>
              ) : (
                  <h2 style={titleStyle}>Locations​​​​</h2>
              )}
              <SortableDetailsList
                  items={data}
                  columns={columns}
                  isLoading={isLoading}
                  rowHeight={48}
                  headerHeight={40}
              />
          </section>
  )
}
