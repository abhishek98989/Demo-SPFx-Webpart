import * as React from 'react';
import { SortableDetailsList } from '../../../globalCommon/GlobalTable';
import { useState, useEffect } from 'react';
import { spfi, SPFx } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/security";
import { PermissionKind } from "@pnp/sp/security";
import { Persona, PersonaSize } from '@fluentui/react/lib/Persona';
import { LivePersona } from '@pnp/spfx-controls-react/lib/LivePersona';

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
        .select(
          'Id',
          'Title',
          'Who',
          'Policy_x002f_Procedure',
          'Contact/Id',
          'Contact/Title',
          'Contact/EMail'
        )
        .expand('Contact')
        .top(5000)();

      console.log('Raw items:', items);

      // Default sort by Title (What) ascending
      const sorted = [...(items || [])].sort((a, b) =>
        (a.Title || '').toLowerCase().localeCompare((b.Title || '').toLowerCase())
      );

      setData(sorted);
    } catch (error) {
      console.error('Error loading data:', error);
      setData([]);
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

  const renderPolicyLink = (value: any) => {
    const url: string | undefined = value?.Url || value?.url || (typeof value === 'string' ? value : undefined);
    const desc: string | undefined = value?.Description || value?.description;
    const safeUrl = typeof url === 'string' ? url.trim() : '';
    const safeText = typeof desc === 'string' ? desc.trim() : '';

    if (!safeUrl && !safeText) return <span style={{ fontSize: 14, color: '#605e5c' }}>—</span>;
    if (!safeUrl && safeText) return <span style={{ fontSize: 14 }}>{safeText}</span>;

    return (
      <a
        target="_blank"
        rel="noreferrer"
        href={safeUrl}
        data-interception="off"
        style={{ color: '#0078d4', textDecoration: 'none', fontSize: '14px' }}
        onMouseOver={(e) => (e.currentTarget as HTMLElement).style.textDecoration = 'underline'}
        onMouseOut={(e) => (e.currentTarget as HTMLElement).style.textDecoration = 'none'}
      >
        {safeText || safeUrl}
      </a>
    );
  };

  // ✅ Render Contact using PnP LivePersona directly — no Graph/photo fetch needed
  const renderContact = (contactValue: any) => {
    if (!contactValue) return <span style={{ fontSize: 14, color: '#605e5c' }}>—</span>;

    const contacts: any[] = Array.isArray(contactValue) ? contactValue : [contactValue];
    if (contacts.length === 0) return <span style={{ fontSize: 14, color: '#605e5c' }}>—</span>;

    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', paddingTop: '4px', paddingBottom: '4px' }}>
        {contacts.map((person: any, index: number) => {
          const name: string = person?.Title || '';
          const email: string = person?.EMail || '';

          if (!email) {
            // Fallback: render plain Persona if no email available
            return (
              <Persona
                key={index}
                text={name}
                size={PersonaSize.size32}
                coinSize={32}
                styles={{
                  root: { maxWidth: 220 },
                  primaryText: { fontSize: '13px', color: '#323130' },
                }}
              />
            );
          }

          // ✅ LivePersona wraps a Persona coin — hover shows the full PnP profile card
          const photoUrl = `https://vaughnconstruction.sharepoint.com/_layouts/15/userphoto.aspx?size=S&accountname=${encodeURIComponent(email)}`;

          return (
            <LivePersona
              key={email}
              serviceScope={props.Context?.serviceScope}
              upn={email}
              template={
                <Persona
                  text={name}
                  size={PersonaSize.size32}
                  coinSize={32}
                  imageUrl={photoUrl}
                  styles={{
                    root: { maxWidth: 220 },
                    primaryText: { fontSize: '13px', color: '#323130' },
                  }}
                />
              }
            />
          );
        })}
      </div>
    );
  };

  // ✅ Column order: What, Policy/Procedure, Notes, Contact
  const columns = [
    {
      key: 'what',
      name: 'What',
      fieldName: 'Title',
      minWidth: 400,
      isResizable: true,
      onRender: (item: any) => (
        <>
          {canEdit ? (
            <a
              target="_blank"
              rel="noreferrer"
              href={`https://vaughnconstruction.sharepoint.com/Lists/WhoDoesWhat/DispForm.aspx?ID=${item.Id}`}
              data-interception="off"
              style={{ color: '#0078d4', textDecoration: 'none', fontSize: '14px' }}
              onMouseOver={(e) => (e.currentTarget as HTMLElement).style.textDecoration = 'underline'}
              onMouseOut={(e) => (e.currentTarget as HTMLElement).style.textDecoration = 'none'}
            >
              {item?.Title ?? ''}
            </a>
          ) : (
            <span style={{ fontSize: '14px' }}>{item?.Title ?? ''}</span>
          )}
        </>
      )
    },
    {
      key: 'policyProcedure',
      name: 'Policy / Procedure',
      fieldName: 'Policy_x002f_Procedure',
      minWidth: 220,
      isResizable: true,
      onRender: (item: any) => renderPolicyLink(item?.Policy_x002f_Procedure)
    },
    {
      key: 'notes',
      name: 'Notes',       // ✅ Renamed from "Who"
      fieldName: 'Who',
      minWidth: 200,
      isResizable: true,
    },
    {
      key: 'contact',
      name: 'Contact',     // ✅ PnP LivePersona hover card
      fieldName: 'Contact',
      minWidth: 220,
      isResizable: true,
      onRender: (item: any) => renderContact(item?.Contact)
    },
  ];

  const sectionStyle: React.CSSProperties = {
    padding: '20px',
    backgroundColor: '#ffffff',
    overflowX: 'auto',
    width: '100%'
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
        <a
          href="https://vaughnconstruction.sharepoint.com/Lists/WhoDoesWhat/AllItems.aspx"
          target="_blank"
          rel="noreferrer"
          data-interception="off"
          style={{ textDecoration: 'none' }}
        >
          <h2 style={titleStyle}>Corporate Responsibilities</h2>
        </a>
      ) : (
        <h2 style={titleStyle}>Corporate Responsibilities</h2>
      )}

      <div style={{ overflowX: 'auto', width: '100%' }}>
        <SortableDetailsList
          items={data}
          columns={columns}
          isLoading={isLoading}
          rowHeight={48}
          headerHeight={40}
          layoutMode={1}    // 1 = fixedColumns
          constrainMode={0} // 0 = unconstrained
        />
      </div>
    </section>
  );
};

export default CorporateResponsibilitiesTable;