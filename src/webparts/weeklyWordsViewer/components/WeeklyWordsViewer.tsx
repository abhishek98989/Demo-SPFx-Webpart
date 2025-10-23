import * as React from 'react';
import { useEffect, useMemo, useState } from 'react';
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

export type ViewerView = 'Post' | 'Tiles';
export type ViewerSource = 'Weekly Words' | 'Department Specific';
export type ViewerDepartment = 'HR' | 'IT' | 'Marketing' | 'OPS' | 'Safety' | 'VDC';

export interface IWeeklyWordsViewerProps {
  context: any;
  siteUrl: string;
  listId: string;
  view: ViewerView;
  source: ViewerSource;
  department?: ViewerDepartment;
}

interface IListItem {
  Id: number;
  Title: string;
  Created?: string;
  VaughnContent?: string;
  Abstract?: string;
  PublishingRollupImage?: { Url?: string } | string;
  PublishingRollupImageUrl?: string;
  PublishingSource?: string;
  Department?: string;
  InCaseYouMissed?: Array<{ Id: number; Title: string }>;
}

const TilesList: React.FC<{ items: IListItem[] }> = ({ items }) => {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
      {items.map(it => (
        <div key={it.Id} style={{ display: 'flex', gap: 16, alignItems: 'flex-start', border: '1px solid #eee', borderRadius: 8, padding: 12 }}>
          <div style={{ width: 160, height: 100, overflow: 'hidden', borderRadius: 6, flexShrink: 0 }}>
            {it.PublishingRollupImageUrl ? (
              <img src={it.PublishingRollupImageUrl} alt="" style={{ width: '100%', height: '100%', objectFit: 'cover' }} />
            ) : (
              <div style={{ width: '100%', height: '100%', background: '#f3f2f1' }} />
            )}
          </div>
          <div style={{ flex: 1 }}>
            <div style={{ fontSize: 18, fontWeight: 600, marginBottom: 6 }}>{it.Title}</div>
            <div style={{ color: '#605e5c', fontSize: 14 }}>{it.Abstract || ''}</div>
          </div>
        </div>
      ))}
    </div>
  );
};

const ICYMGrid: React.FC<{ items: IListItem[] }> = ({ items }) => {
  if (!items?.length) return null;
  return (
    <div style={{ marginTop: 32, paddingTop: 24, borderTop: '1px solid #edebe9' }}>
      <div style={{ fontSize: 20, fontWeight: 700, marginBottom: 16 }}>In case you missed it</div>
      <div style={{ display: 'flex', gap: 16, justifyContent: 'flex-start' }}>
        {items.slice(0, 4).map(it => (
          <div key={it.Id} style={{ flex: '0 0 calc(25% - 12px)', border: '1px solid #eee', borderRadius: 8, padding: 12, textAlign: 'center' }}>
            <div style={{ width: '100%', aspectRatio: '16/9', overflow: 'hidden', borderRadius: 6, marginBottom: 12, background: '#faf9f8' }}>
              {it.PublishingRollupImageUrl ? (
                <img src={it.PublishingRollupImageUrl} alt="" style={{ width: '100%', height: '100%', objectFit: 'cover' }} />
              ) : null}
            </div>
            <div style={{ fontWeight: 600, fontSize: 14 }}>{it.Title}</div>
          </div>
        ))}
      </div>
    </div>
  );
};

const WeeklyWordsViewer: React.FC<IWeeklyWordsViewerProps> = ({ context, siteUrl, listId, view, source, department }) => {
  const sp = useMemo(() => (siteUrl ? spfi(siteUrl).using(SPFx(context)) : spfi().using(SPFx(context))), [siteUrl, context]);

  const [items, setItems] = useState<IListItem[]>([]);
  const [main, setMain] = useState<IListItem | null>(null);
  const [icym, setIcym] = useState<IListItem[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | null>(null);

  const isDept = source === 'Department Specific';

  const buildFilter = React.useCallback(() => {
    const clauses: string[] = [];
    
    // Content approval: 0 = Approved
    clauses.push("OData__ModerationStatus eq 0");

    if (isDept) {
      // Department Specific: PublishingSource is 'Department Specific' OR 'Both'
      clauses.push("(PublishingSource eq 'Department Specific' or PublishingSource eq 'Both')");
      if (department) {
        clauses.push(`Department eq '${department}'`);
      }
    } else {
      // Weekly Words: PublishingSource is 'Weekly Words' OR 'Both'
      clauses.push("(PublishingSource eq 'Weekly Words' or PublishingSource eq 'Both')");
    }

    return clauses.join(' and ');
  }, [isDept, department]);

  async function fetchICYMDetails(ids: number[]): Promise<IListItem[]> {
    if (!ids?.length) return [];
    try {
      const idFilter = ids.map(id => `Id eq ${id}`).join(' or ');
      const rows: any[] = await sp.web.lists.getById(listId).items
        .select('Id', 'Title', 'PublishingRollupImage')
        .filter(idFilter)
        .top(4)();

      return rows.map(r => ({
        Id: r.Id,
        Title: r.Title,
        PublishingRollupImageUrl: r.PublishingRollupImage?.Url || r.PublishingRollupImage || ''
      }));
    } catch (e) {
      console.error('Error fetching ICYM details:', e);
      return [];
    }
  }

  async function loadTiles() {
    const filter = buildFilter();
    const rows: any[] = await sp.web.lists.getById(listId).items
      .select('Id', 'Title', 'Abstract', 'PublishingRollupImage', 'PublishingSource', 'Department', 'OData__ModerationStatus', 'Created')
      .filter(filter)
      .orderBy('Created', false)
      .top(4)();

    const mapped: IListItem[] = rows.map(r => ({
      Id: r.Id,
      Title: r.Title,
      Abstract: r.Abstract,
      PublishingRollupImageUrl: r.PublishingRollupImage?.Url || r.PublishingRollupImage || '',
      PublishingSource: r.PublishingSource,
      Department: r.Department,
      Created: r.Created
    }));

    setItems(mapped);
    setMain(null);
    setIcym([]);
  }

  async function loadPost() {
    const filter = buildFilter();
    const rows: any[] = await sp.web.lists.getById(listId).items
      .select(
        'Id',
        'Title',
        'VaughnContent',
        'Abstract',
        'PublishingRollupImage',
        'PublishingSource',
        'Department',
        'OData__ModerationStatus',
        'Created',
        'InCaseYouMissed/Id',
        'InCaseYouMissed/Title'
      )
      .expand('InCaseYouMissed')
      .filter(filter)
      .orderBy('Created', false)
      .top(1)();

    const first = rows?.[0];
    if (!first) {
      setMain(null);
      setIcym([]);
      return;
    }

    const mapped: IListItem = {
      Id: first.Id,
      Title: first.Title,
      VaughnContent: first.VaughnContent,
      Abstract: first.Abstract,
      PublishingRollupImageUrl: first.PublishingRollupImage?.Url || first.PublishingRollupImage || '',
      PublishingSource: first.PublishingSource,
      Department: first.Department,
      Created: first.Created,
      InCaseYouMissed: Array.isArray(first.InCaseYouMissed)
        ? first.InCaseYouMissed.map((x: any) => ({ Id: x.Id, Title: x.Title }))
        : []
    };

    setMain(mapped);

    // Load ICYM details with images
    const icymIds = (mapped.InCaseYouMissed || []).map(x => x.Id).slice(0, 4);
    if (icymIds.length > 0) {
      const icymItems = await fetchICYMDetails(icymIds);
      setIcym(icymItems);
    } else {
      setIcym([]);
    }
  }

  useEffect(() => {
    (async () => {
      try {
        setLoading(true);
        setError(null);
        
        if (!listId) {
          setError('List ID is required.');
          setLoading(false);
          return;
        }

        if (isDept && !department) {
          setError('Department is required when source is "Department Specific".');
          setLoading(false);
          return;
        }

        if (view === 'Tiles') {
          await loadTiles();
        } else {
          await loadPost();
        }
      } catch (e: any) {
        console.error('Load error:', e);
        setError(e?.message || 'Failed to load items');
      } finally {
        setLoading(false);
      }
    })();
  }, [view, source, department, listId, siteUrl]);

  if (loading) {
    return (
      <div style={{ padding: 24, textAlign: 'center', color: '#605e5c' }}>
        Loading…
      </div>
    );
  }

  if (error) {
    return (
      <div style={{ padding: 16, border: '1px solid #d13438', borderRadius: 8, color: '#d13438', backgroundColor: '#fef0f0' }}>
        <strong>Error:</strong> {error}
      </div>
    );
  }

  if (view === 'Tiles') {
    if (!items.length) return <div style={{ padding: 16, color: '#605e5c' }}>No posts found.</div>;
    return <TilesList items={items} />;
  }

  // Post view
  if (!main) return <div style={{ padding: 16, color: '#605e5c' }}>No post found.</div>;

  return (
    <div style={{ maxWidth: 1200, margin: '0 auto' }}>
      <h2 style={{ margin: '0 0 16px', fontSize: 28, fontWeight: 700 }}>{main.Title}</h2>
      
      {main.PublishingRollupImageUrl ? (
        <div style={{ width: '100%', maxHeight: 400, overflow: 'hidden', borderRadius: 8, marginBottom: 24 }}>
          <img src={main.PublishingRollupImageUrl} alt="" style={{ width: '100%', height: '100%', objectFit: 'cover' }} />
        </div>
      ) : null}

      {main.VaughnContent ? (
        <div
          style={{ lineHeight: 1.6, fontSize: 16, color: '#323130' }}
          dangerouslySetInnerHTML={{ __html: main.VaughnContent }}
        />
      ) : (
        <div style={{ padding: 16, background: '#faf9f8', border: '1px solid #edebe9', borderRadius: 8, color: '#605e5c' }}>
          No content available.
        </div>
      )}

      <ICYMGrid items={icym} />
    </div>
  );
};

export default WeeklyWordsViewer;