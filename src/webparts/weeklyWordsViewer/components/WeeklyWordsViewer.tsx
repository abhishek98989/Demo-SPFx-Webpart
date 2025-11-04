import * as React from 'react';
import { useEffect, useMemo, useState } from 'react';
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/items/list';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import './style.css';
import { ArrowLeft, Calendar, User, Eye } from 'lucide-react';
export interface IWeeklyWordsViewerProps {
  context: any;
  siteUrl: string;            // your web absolute URL
  listId: string;
  view: 'Tiles' | 'Post' | 'Redirected';  // ⬅️ NEW
  publishingSource: 'Weekly Words' | 'Department Specific';
  department?: 'HR' | 'IT' | 'Marketing' | 'OPS' | 'Safety' | 'VDC';
}
// Shape mapped from your SharePoint list (based on NewsManagement.tsx)
interface ISpItem {
  Id: number;
  Title: string;
  VaughnContent?: string; // HTML body
  Department?: string;
  PublishingSource?: string; // 'Weekly Words' | 'Department Specific' | 'Both'
  PublishingRollupImage?: { Url?: string; Description?: string } | string | null;
  Abstract?: string;
  ArticleDate?: string; // ISO date string
  Created?: string; // fallback if ArticleDate missing
  PublishingContact?: { Id: number; Title: string; EMail?: string } | null;
  InCaseYouMissed?: Array<{ Id: number; Title: string }>; // lookup IDs only
}

interface INormalizedItem {
  id: number;
  title: string;
  vaughn_content?: string;
  department?: string;
  publishing_source?: string;
  publishing_rollup_image_url?: string;
  abstract?: string;
  article_date?: string; // ISO
  created?: string; // ISO
  publishing_contact?: string; // display name
  publishing_contact_email?: string;
  in_case_you_missed_ids?: number[];
}

const WeeklyWordsViewer: React.FC<IWeeklyWordsViewerProps> = (props) => {
  const { context, siteUrl, listId, view, publishingSource, department } = props;
  function buildViewerUrl(siteUrl: string, id: number) {
    const base = context.pageContext.web.absoluteUrl;
    return `${base}/SitePages/Weekly-Words-Viewer.aspx?WeeklyWordsId=${id}`;
  }
  const sp = useMemo(() => spfi(siteUrl).using(SPFx(context)), [siteUrl, context]);

  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  // Tiles view state (top 4)
  const [tiles, setTiles] = useState<INormalizedItem[]>([]);

  // Post view state (top most + related up to 4)
  const [post, setPost] = useState<INormalizedItem | null>(null);
  const [related, setRelated] = useState<INormalizedItem[]>([]);

  useEffect(() => {
    setError(null);
    setLoading(true);
    (async () => {
      if (view === 'Tiles') {
        const items = await fetchItems(sp, listId, publishingSource, department, 4);
        setTiles(items);
        return;
      }

      if (view === 'Redirected') {
        const idStr = getQueryParam('WeeklyWordsId');
        const id = idStr ? parseInt(idStr, 10) : NaN;
        if (!id || Number.isNaN(id)) {
          setPost(null);
          setRelated([]);
          setError('Missing or invalid WeeklyWordsId in the URL.');
          return;
        }

        const full = await fetchOneByIdFull(sp, listId, id);
        if (!full) {
          setPost(null);
          setRelated([]);
          setError(`No item found for WeeklyWordsId=${id}.`);
          return;
        }

      

        setPost(full);

        if (full.in_case_you_missed_ids?.length) {
          const rel = await fetchByIds(sp, listId, full.in_case_you_missed_ids.slice(0, 4));
          setRelated(rel);
        } else {
          setRelated([]);
        }
        return;
      }

      // Default Post path (existing behavior)
      const items = await fetchItems(sp, listId, publishingSource, department, 1);
      const top = items[0] ?? null;
      setPost(top || null);
      if (top?.in_case_you_missed_ids?.length) {
        const rel = await fetchByIds(sp, listId, top.in_case_you_missed_ids.slice(0, 4));
        setRelated(rel);
      } else {
        setRelated([]);
      }
    })()
      .catch((e) => setError(parseError(e)))
      .finally(() => setLoading(false));
  }, [view, publishingSource, department, sp, listId]);

  if (loading) {
    return (
      <div className="flex items-center justify-center min-h-[40vh]">
        <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600"></div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="max-w-3xl mx-auto px-4 py-8 text-red-600">
        {error}
      </div>
    );
  }

  if (view === 'Tiles') {
    return (
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-12">
        <h1 className="text-4xl font-bold text-gray-900 mb-12">Latest Articles</h1>

        <div className="space-y-8">
          {tiles.map((article) => (
            <article
              key={article.id}
              className="bg-white rounded-xl shadow-md hover:shadow-xl transition-all duration-300 overflow-hidden group"
            >
              <div className="flex flex-col sm:flex-row">
                <div className="sm:w-80 h-56 sm:h-auto flex-shrink-0 overflow-hidden">
                  <img
                    src={article.publishing_rollup_image_url || 'https://images.pexels.com/photos/3861969/pexels-photo-3861969.jpeg?auto=compress&cs=tinysrgb&w=800'}
                    alt={article.title}
                    className="w-full h-full object-cover group-hover:scale-105 transition-transform duration-300"
                  />
                </div>

                <div className="flex-1 p-6 sm:p-8 flex flex-col justify-between">
                  <div>
                    <div className="flex items-center gap-3 mb-3">
                      {article.department && (
                        <span className="inline-block px-3 py-1 text-xs font-semibold text-blue-700 bg-blue-100 rounded-full">
                          {article.department}
                        </span>
                      )}
                      {(article.article_date || article.created) && (
                        <span className="text-sm text-gray-500">
                          {new Date(article.article_date || article.created!).toLocaleDateString('en-US', {
                            year: 'numeric', month: 'long', day: 'numeric'
                          })}
                        </span>
                      )}
                    </div>

                    <h2 className="text-2xl font-bold text-gray-900 mb-3 group-hover:text-blue-600 transition-colors">
                      {article.title}
                    </h2>

                    {article.abstract && (
                      <p className="text-gray-600 leading-relaxed line-clamp-3">
                        {article.abstract}
                      </p>
                    )}
                  </div>

                  <div className="flex items-center justify-between mt-4">
                    {article.publishing_contact && (
                      <span className="text-sm text-gray-500">
                        By {article.publishing_contact}
                      </span>
                    )}
                    <a
                      href={buildViewerUrl(siteUrl, article.id)}
                      className="flex items-center gap-2 text-blue-600 font-semibold"
                    >
                      <span>Read More</span>
                      <Eye className="w-4 h-4" />
                    </a>

                  </div>
                </div>
              </div>
            </article>
          ))}
        </div>
      </div>
    );
  }

  // Post view
  if (!post) {
    return (
      <div className="max-w-4xl mx-auto px-4 py-12">
        <p className="text-gray-600">No post found for the selected filters.</p>
      </div>
    );
  }

  return (
    <article className="bg-white rounded-xl shadow-lg overflow-hidden">


      <div className="pb-6 pr-6 pl-6 pt-4">
        <h3 style={{ color: '#991426' ,backgroundColor:'#bbbcbc'}} className="text-4xl sm:text-4xl font-bold text-gray-900 mb-2 leading-tight">
          Weekly Words from Westpark
        </h3>

        <div className="border-b border-gray-200">
          <h3 className="text-2xl sm:text-2xl font-bold text-gray-900 mb-2 leading-tight">
            {post.title}
          </h3>

          <div className="flex flex-wrap gap-6 pt-4 mb-2 text-gray-600">
            {post.publishing_contact && (
              <div className="flex items-center gap-2">
                <User className="w-5 h-5 text-blue-600" />
                <span>
                  <span className="font-semibold">Posted by:</span> {post.publishing_contact}
                </span>
              </div>
            )}
            {(post.article_date || post.created) && (
              <div className="flex items-center gap-2">
                <Calendar className="w-5 h-5 text-blue-600" />
                <span>
                  <span className="font-semibold">Posted on:</span>{' '}
                  {new Date(post.article_date || post.created!).toLocaleDateString('en-US', {
                    year: 'numeric', month: 'long', day: 'numeric'
                  })}
                </span>
              </div>
            )}
            {post.department && (
              <span className="inline-block px-4 py-1 text-sm font-semibold text-blue-700 bg-blue-100 rounded-full">
                {post.department}
              </span>
            )}
          </div>

        </div>






        {post.vaughn_content && (
          <div
            className="ck-content prose max-w-none mb-12"
            dangerouslySetInnerHTML={{ __html: post.vaughn_content }}
          />
        )}

        {related.length > 0 && (
          <div className="related-section">
            <h3 className="text-4xl p-2 mb-4 sm:text-2xl font-bold" style={{ background: "gainsboro" }}>In Case You Missed It</h3>
            <div className="related-grid">
              {related.map((r) => (
                <div key={r.id} className="related-card">
                  <div className="related-card-content">
                    <h4 className="related-card-title">
                      <a href={buildViewerUrl(siteUrl, r.id)} className="hover:underline">
                        {r.title}
                      </a>
                    </h4>

                    <div className="related-card-image">
                      <img
                        src={r.publishing_rollup_image_url || 'https://images.pexels.com/photos/3861969/pexels-photo-3861969.jpeg?auto=compress&cs=tinysrgb&w=400'}
                        alt={r.title}
                      />
                    </div>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    </article>
  );
};

export default WeeklyWordsViewer;

/** ------------------------------
 * Data helpers
 * -------------------------------*/
async function fetchItems(
  sp: ReturnType<typeof spfi>,
  listId: string,
  publishingSource: 'Weekly Words' | 'Department Specific',
  department: string | undefined,
  top: number
): Promise<INormalizedItem[]> {
  const today = new Date();
  today.setHours(23, 59, 59, 999);
  const todayISO = today.toISOString();
  // Build filter: include items where PublishingSource == requested OR 'Both'
  const srcFilter = `(PublishingSource eq '${escapeOData(publishingSource)}' or PublishingSource eq 'Both') and OData__ModerationStatus eq 0 and (ArticleDate le datetime'${todayISO}')`;
  const deptFilter = publishingSource === 'Department Specific' && department
    ? ` and Department eq '${escapeOData(department)}'`
    : '';

  const filter = `${srcFilter}${deptFilter}`;

  // Pull the fields we need. Order by ArticleDate desc, fallback by Created.
  const rows: ISpItem[] = await sp.web.lists
    .getById(listId)
    .items
    .select(
      'Id',
      'Title',
      'VaughnContent',
      'Department',
      'PublishingSource',
      'PublishingRollupImage',
      'Abstract',
      'ArticleDate',
      'Created',
      'InCaseYouMissed/Id',
      'InCaseYouMissed/Title',
      'PublishingContact/Id',
      'PublishingContact/Title',
      'PublishingContact/EMail'
    )
    .expand('InCaseYouMissed', 'PublishingContact')
    .filter(filter)
    .orderBy("ArticleDate", false)
    .top(top) // over-fetch a bit; we'll sort in JS by ArticleDate
    () as any;

  // Normalize and sort by ArticleDate desc then Created desc
  const norm = rows.map(normalizeItem)
    .sort((a, b) => {
      const adA = a.article_date ? Date.parse(a.article_date) : 0;
      const adB = b.article_date ? Date.parse(b.article_date) : 0;
      return adB - adA;
    })
    .slice(0, top);

  return norm;
}

async function fetchByIds(
  sp: ReturnType<typeof spfi>,
  listId: string,
  ids: number[]
): Promise<INormalizedItem[]> {
  if (!ids.length) return [];
  const filter = ids.map((id) => `Id eq ${id}`).join(' or ');
  const rows: ISpItem[] = await sp.web.lists
    .getById(listId)
    .items
    .select('Id', 'Title', 'PublishingRollupImage', 'ArticleDate', 'Created')
    .filter(filter)
    .top(4)();
  return rows.map(normalizeItem);
}

function normalizeItem(r: ISpItem): INormalizedItem {
  return {
    id: r.Id,
    title: r.Title,
    vaughn_content: r.VaughnContent || undefined,
    department: r.Department || undefined,
    publishing_source: r.PublishingSource || undefined,
    publishing_rollup_image_url: r.PublishingRollupImage
      ? (typeof r.PublishingRollupImage === 'string'
        ? r.PublishingRollupImage
        : r.PublishingRollupImage?.Url || '')
      : undefined,
    abstract: r.Abstract || undefined,
    article_date: r.ArticleDate || undefined,
    created: r.Created || undefined,
    publishing_contact: r.PublishingContact?.Title || undefined,
    publishing_contact_email: r.PublishingContact?.EMail || undefined,
    in_case_you_missed_ids: Array.isArray(r.InCaseYouMissed)
      ? r.InCaseYouMissed.map((x: any) => x.Id)
      : undefined,
  };
}

function escapeOData(val: string) {
  return val.replace(/'/g, "''");
}

function parseError(e: any): string {
  if (!e) return 'Unknown error';
  if (typeof e === 'string') return e;
  if (e?.message) return e.message;
  try { return JSON.stringify(e); } catch { return String(e); }
}


function getQueryParam(name: string) {
  if (typeof window === 'undefined') return null;
  return new URLSearchParams(window.location.search).get(name);
}
async function fetchOneByIdFull(
  sp: ReturnType<typeof spfi>,
  listId: string,
  id: number
): Promise<INormalizedItem | null> {
  const rows: ISpItem[] = await sp.web.lists
    .getById(listId)
    .items
    .select(
      'Id',
      'Title',
      'VaughnContent',
      'Department',
      'PublishingSource',
      'PublishingRollupImage',
      'Abstract',
      'ArticleDate',
      'Created',
      'InCaseYouMissed/Id',
      'InCaseYouMissed/Title',
      'PublishingContact/Id',
      'PublishingContact/Title',
      'PublishingContact/EMail'
    )
    .expand('InCaseYouMissed', 'PublishingContact')
    .filter(`Id eq ${id}`)
    .top(1)();

  if (!rows?.length) return null;
  return normalizeItem(rows[0]);
}
