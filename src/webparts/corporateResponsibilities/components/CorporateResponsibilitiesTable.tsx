import * as React from 'react';
import { SortableDetailsList } from '../../../globalCommon/GlobalTable';
import { useState, useEffect, useCallback } from 'react';
import { spfi, SPFx } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/security";
import { PermissionKind } from "@pnp/sp/security";
import { Persona, PersonaSize } from '@fluentui/react/lib/Persona';
import { LivePersona } from '@pnp/spfx-controls-react/lib/LivePersona';
import { IColumn } from '@fluentui/react/lib/DetailsList';

// ─── Types ────────────────────────────────────────────────────────────────────

interface IListItem {
  Id: number;
  Title: string;
  Who?: string;
  Policy_x002f_Procedure?: { Url?: string; Description?: string } | string;
  Contact?: IContactPerson | IContactPerson[];

  // Used for sorting Contact column
  _contactSortKey?: string;
}

interface IContactPerson {
  Id?: number;
  Title?: string;
  EMail?: string;
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

const getPolicySortValue = (
  value: IListItem['Policy_x002f_Procedure']
): string => {
  if (!value) return '';

  if (typeof value === 'string') {
    return value.toLowerCase();
  }

  return ((value.Description || value.Url) ?? '').toLowerCase();
};

const compareStrings = (
  a: string,
  b: string,
  desc: boolean
): number => {
  const result = a.localeCompare(b, undefined, {
    sensitivity: 'base',
  });

  return desc ? -result : result;
};

// ─── Component ────────────────────────────────────────────────────────────────

const CorporateResponsibilitiesTable = (props: any) => {

  const [rawData, setRawData] = useState<IListItem[]>([]);
  const [data, setData] = useState<IListItem[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [canEdit, setCanEdit] = useState<boolean>(false);

  // Sorting state
  const [sortColumn, setSortColumn] = useState<string>('what');
  const [sortDescending, setSortDescending] = useState<boolean>(false);

  // ─── Load Data ──────────────────────────────────────────────────────────────

  const loadData = async () => {

    try {

      setIsLoading(true);

      const sp = spfi(
        'https://vaughnconstruction.sharepoint.com'
      ).using(SPFx(props.Context));

      const items: IListItem[] = await sp.web.lists
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

      console.log('Raw Items:', items);

      // Normalize Contact sorting
      const normalized: IListItem[] = (items || []).map(
        (item: IListItem) => {

          const contacts: IContactPerson[] =
            !item.Contact
              ? []
              : Array.isArray(item.Contact)
              ? item.Contact
              : [item.Contact];

          const contactSortKey = contacts
            .map(c => (c?.Title || '').trim().toLowerCase())
            .filter(Boolean)
            .join(';');

          return {
            ...item,
            _contactSortKey: contactSortKey,
          };
        }
      );

      console.log('Normalized:', normalized);

      setRawData(normalized);

    } catch (error) {

      console.error('Error loading data:', error);
      setRawData([]);

    } finally {

      setIsLoading(false);

    }
  };

  // ─── Permissions ────────────────────────────────────────────────────────────

  const checkPermissions = async () => {

    try {

      const sp = spfi(
        'https://vaughnconstruction.sharepoint.com'
      ).using(SPFx(props.Context));

      const list = sp.web.lists.getById(
        '8c97c9e1-ecaf-4fcb-93a9-4016509d185c'
      );

      const perms = await list.getCurrentUserEffectivePermissions();

      if (
        sp.web.hasPermissions(
          perms,
          PermissionKind.EditListItems
        )
      ) {
        setCanEdit(true);
      }

    } catch (error) {

      console.error('Permission Error:', error);

    }
  };

  // ─── Initial Load ───────────────────────────────────────────────────────────

  useEffect(() => {

    loadData();
    checkPermissions();

  }, []);

  // ─── Sorting ────────────────────────────────────────────────────────────────

  useEffect(() => {

    if (!rawData.length) {
      setData([]);
      return;
    }

    const sorted = [...rawData].sort((a, b) => {

      const desc = sortDescending;

      switch (sortColumn) {

        case 'what':

          return compareStrings(
            (a.Title ?? '').toLowerCase(),
            (b.Title ?? '').toLowerCase(),
            desc
          );

        case 'policyProcedure':

          return compareStrings(
            getPolicySortValue(a.Policy_x002f_Procedure),
            getPolicySortValue(b.Policy_x002f_Procedure),
            desc
          );

        case 'notes':

          return compareStrings(
            (a.Who ?? '').toLowerCase(),
            (b.Who ?? '').toLowerCase(),
            desc
          );

        case 'contact':

          return compareStrings(
            a._contactSortKey ?? '',
            b._contactSortKey ?? '',
            desc
          );

        default:
          return 0;
      }
    });

    setData(sorted);

  }, [rawData, sortColumn, sortDescending]);

  // ─── Column Click ───────────────────────────────────────────────────────────

  const onColumnHeaderClick = useCallback(

    (
      _ev: React.MouseEvent<HTMLElement> | undefined,
      column?: IColumn
    ) => {

      if (!column) return;

      if (column.key === sortColumn) {

        setSortDescending(prev => !prev);

      } else {

        setSortColumn(column.key);
        setSortDescending(false);

      }
    },

    [sortColumn]
  );

  // ─── Render Policy Link ─────────────────────────────────────────────────────

  const renderPolicyLink = (
    value: IListItem['Policy_x002f_Procedure']
  ) => {

    const url: string =
      typeof value === 'object' && value
        ? (value.Url ?? '').trim()
        : typeof value === 'string'
        ? value.trim()
        : '';

    const desc: string =
      typeof value === 'object' && value
        ? (value.Description ?? '').trim()
        : '';

    if (!url && !desc) {
      return (
        <span style={{ fontSize: 14, color: '#605e5c' }}>
          —
        </span>
      );
    }

    if (!url && desc) {
      return <span style={{ fontSize: 14 }}>{desc}</span>;
    }

    return (
      <a
        target="_blank"
        rel="noreferrer"
        href={url}
        data-interception="off"
        style={{
          color: '#0078d4',
          textDecoration: 'none',
          fontSize: '14px',
        }}
        onMouseOver={e =>
          (e.currentTarget.style.textDecoration = 'underline')
        }
        onMouseOut={e =>
          (e.currentTarget.style.textDecoration = 'none')
        }
      >
        {desc || url}
      </a>
    );
  };

  // ─── Render Contact ─────────────────────────────────────────────────────────

  const renderContact = (
    contactValue: IListItem['Contact']
  ) => {

    if (!contactValue) {
      return (
        <span style={{ fontSize: 14, color: '#605e5c' }}>
          —
        </span>
      );
    }

    const contacts = Array.isArray(contactValue)
      ? contactValue
      : [contactValue];

    if (!contacts.length) {
      return (
        <span style={{ fontSize: 14, color: '#605e5c' }}>
          —
        </span>
      );
    }

    return (
      <div
        style={{
          display: 'flex',
          flexDirection: 'column',
          gap: '8px',
          paddingTop: '4px',
          paddingBottom: '4px',
        }}
      >
        {contacts.map((person, index) => {

          const name = person?.Title ?? '';
          const email = person?.EMail ?? '';

          if (!email) {

            return (
              <Persona
                key={index}
                text={name}
                size={PersonaSize.size32}
                coinSize={32}
                styles={{
                  root: { maxWidth: 220 },
                  primaryText: {
                    fontSize: '13px',
                    color: '#323130',
                  },
                }}
              />
            );
          }

          const photoUrl =
            `https://vaughnconstruction.sharepoint.com/_layouts/15/userphoto.aspx?size=S&accountname=${encodeURIComponent(email)}`;

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
                    primaryText: {
                      fontSize: '13px',
                      color: '#323130',
                    },
                  }}
                />
              }
            />
          );
        })}
      </div>
    );
  };

  // ─── Columns ────────────────────────────────────────────────────────────────

  const columns: any[] = [

    {
      key: 'what',
      name: 'What',
      fieldName: 'Title',
      minWidth: 400,
      isResizable: true,
      isSorted: sortColumn === 'what',
      isSortedDescending:
        sortColumn === 'what'
          ? sortDescending
          : false,
      onColumnClick: onColumnHeaderClick,

      onRender: (item: IListItem) => (

        canEdit ? (

          <a
            target="_blank"
            rel="noreferrer"
            href={`https://vaughnconstruction.sharepoint.com/Lists/WhoDoesWhat/DispForm.aspx?ID=${item.Id}`}
            data-interception="off"
            style={{
              color: '#0078d4',
              textDecoration: 'none',
              fontSize: '14px',
            }}
            onMouseOver={e =>
              (e.currentTarget.style.textDecoration = 'underline')
            }
            onMouseOut={e =>
              (e.currentTarget.style.textDecoration = 'none')
            }
          >
            {item.Title ?? ''}
          </a>

        ) : (

          <span style={{ fontSize: '14px' }}>
            {item.Title ?? ''}
          </span>

        )
      ),
    },

    {
      key: 'policyProcedure',
      name: 'Policy / Procedure',
      fieldName: 'Policy_x002f_Procedure',
      minWidth: 220,
      isResizable: true,
      isSorted: sortColumn === 'policyProcedure',
      isSortedDescending:
        sortColumn === 'policyProcedure'
          ? sortDescending
          : false,
      onColumnClick: onColumnHeaderClick,

      onRender: (item: IListItem) =>
        renderPolicyLink(item.Policy_x002f_Procedure),
    },

    {
      key: 'notes',
      name: 'Notes',
      fieldName: 'Who',
      minWidth: 200,
      isResizable: true,
      isSorted: sortColumn === 'notes',
      isSortedDescending:
        sortColumn === 'notes'
          ? sortDescending
          : false,
      onColumnClick: onColumnHeaderClick,
    },

    {
      key: 'contact',
      name: 'Contact',

      // IMPORTANT FIX
      fieldName: '_contactSortKey',

      minWidth: 220,
      isResizable: true,
      isSorted: sortColumn === 'contact',
      isSortedDescending:
        sortColumn === 'contact'
          ? sortDescending
          : false,
      onColumnClick: onColumnHeaderClick,

      onRender: (item: IListItem) =>
        renderContact(item.Contact),
    },
  ];

  // ─── Styles ─────────────────────────────────────────────────────────────────

  const sectionStyle: React.CSSProperties = {
    padding: '20px',
    backgroundColor: '#ffffff',
    overflowX: 'auto',
    width: '100%',
  };

  const titleStyle: React.CSSProperties = {
    fontSize: '24pt',
    fontWeight: '600',
    color: '#323130',
    marginBottom: '16px',
    fontFamily: 'Segoe UI, system-ui, sans-serif',
  };

  // ─── Render ─────────────────────────────────────────────────────────────────

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
          <h2 style={titleStyle}>
            Corporate Responsibilities
          </h2>
        </a>

      ) : (

        <h2 style={titleStyle}>
          Corporate Responsibilities
        </h2>

      )}

      <div
        style={{
          overflowX: 'auto',
          width: '100%',
        }}
      >

        <SortableDetailsList
          items={data}
          columns={columns}
          isLoading={isLoading}
          rowHeight={48}
          headerHeight={40}
          layoutMode={1}
          constrainMode={0}
        />

      </div>

    </section>
  );
};

export default CorporateResponsibilitiesTable;