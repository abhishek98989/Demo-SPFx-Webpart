// WeeklyWordsViewer.tsx
import * as React from 'react';
import { useEffect, useMemo, useState } from 'react';
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/items/list';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import './style.css';
import { Calendar, User, ArrowRight } from 'lucide-react';

export interface IWeeklyWordsViewerProps {
  context: any;
  siteUrl: string;
  listId: string;
  view: 'Tiles' | 'Post' | 'Redirected';
  publishingSource: 'Weekly Words' | 'Department Specific';
  department?: 'HR' | 'IT' | 'Marketing' | 'OPS' | 'Safety' | 'VDC';
}

interface ISpItem {
  Id: number;
  Title: string;
  VaughnContent?: string;
  Department?: string;
  PublishingSource?: string;
  PublishingRollupImage?: { Url?: string } | string | null;
  Abstract?: string;
  ArticleDate?: string;
  Created?: string;
  PublishingContact?: { Id: number; Title: string; EMail?: string } | null;
  InCaseYouMissed?: Array<{ Id: number; Title: string }>;
}

interface INormalizedItem {
  id: number;
  title: string;
  vaughn_content?: string;
  department?: string;
  publishing_source?: string;
  publishing_rollup_image_url?: string;
  abstract?: string;
  article_date?: string;
  created?: string;
  publishing_contact?: string;
  publishing_contact_email?: string;
  in_case_you_missed_ids?: number[];
}

const WeeklyWordsViewer: React.FC<IWeeklyWordsViewerProps> = (props) => {
  const {
    context,
    siteUrl,
    listId,
    view,
    publishingSource,
    department
  } = props;

  const sp = useMemo(
    () => spfi(siteUrl).using(SPFx(context)),
    [siteUrl, context]
  );

  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const [tiles, setTiles] = useState<INormalizedItem[]>([]);
  const [post, setPost] = useState<INormalizedItem | null>(null);
  const [related, setRelated] = useState<INormalizedItem[]>([]);

  function buildViewerUrl(id: number) {
    const base = context.pageContext.web.absoluteUrl;
    return `${base}/SitePages/Weekly-Words-Viewer.aspx?WeeklyWordsId=${id}`;
  }

  useEffect(() => {
    setError(null);
    setLoading(true);

    (async () => {
      if (view === 'Tiles') {
        const items = await fetchItems(
          sp,
          listId,
          publishingSource,
          department,
          4
        );

        setTiles(items);
        return;
      }

      if (view === 'Redirected') {
        const idStr = getQueryParam('WeeklyWordsId');
        const id = idStr ? parseInt(idStr, 10) : NaN;

        if (!id || Number.isNaN(id)) {
          setError('Missing or invalid WeeklyWordsId.');
          return;
        }

        const full = await fetchOneByIdFull(sp, listId, id);

        if (!full) {
          setError(`No item found for WeeklyWordsId=${id}`);
          return;
        }

        setPost(full);

        if (full.in_case_you_missed_ids?.length) {
          const rel = await fetchByIds(
            sp,
            listId,
            full.in_case_you_missed_ids.slice(0, 4)
          );

          setRelated(rel);
        }

        return;
      }

      const items = await fetchItems(
        sp,
        listId,
        publishingSource,
        department,
        1
      );

      const top = items[0] ?? null;

      setPost(top);

      if (top?.in_case_you_missed_ids?.length) {
        const rel = await fetchByIds(
          sp,
          listId,
          top.in_case_you_missed_ids.slice(0, 4)
        );

        setRelated(rel);
      }
    })()
      .catch((e) => setError(parseError(e)))
      .finally(() => setLoading(false));

  }, [view, publishingSource, department, sp, listId]);

  if (loading) {
    return (
      <div className="ww-loading">
        <div className="ww-loader"></div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="ww-error">
        {error}
      </div>
    );
  }

  // =========================================
  // TILES VIEW
  // =========================================

  if (view === 'Tiles') {
    return (
      <div className="ww-wrapper">

        <div className="ww-heading-row">
          <h1 className="ww-page-title">
            Weekly Words
          </h1>
        </div>

        <div className="ww-grid">
          {tiles.map((article) => (
            <article
              key={article.id}
              className="ww-tile"
            >
              <div className="ww-tile-image-wrap">
                <img
                  src={
                    article.publishing_rollup_image_url ||
                    'https://images.pexels.com/photos/3861969/pexels-photo-3861969.jpeg?auto=compress&cs=tinysrgb&w=1200'
                  }
                  alt={article.title}
                  className="ww-tile-image"
                />
              </div>

              <div className="ww-tile-body">

                <div className="ww-meta-row">
                  {article.department && (
                    <span className="ww-tag">
                      {article.department}
                    </span>
                  )}

                  {(article.article_date || article.created) && (
                    <span className="ww-date">
                      {new Date(
                        article.article_date || article.created!
                      ).toLocaleDateString('en-US', {
                        year: 'numeric',
                        month: 'long',
                        day: 'numeric'
                      })}
                    </span>
                  )}
                </div>

                <h2 className="ww-tile-title">
                  {article.title}
                </h2>

                {article.abstract && (
                  <p className="ww-tile-description">
                    {article.abstract}
                  </p>
                )}

                <div className="ww-footer">
                  <span className="ww-author">
                    {article.publishing_contact}
                  </span>

                  <a
                    href={buildViewerUrl(article.id)}
                    className="ww-read-link"
                  >
                    Read Article
                    <ArrowRight size={16} />
                  </a>
                </div>

              </div>
            </article>
          ))}
        </div>
      </div>
    );
  }

  // =========================================
  // POST VIEW
  // =========================================

  if (!post) {
    return (
      <div className="ww-empty">
        No article found.
      </div>
    );
  }

  return (
    <article className="ww-article">

      <div className="ww-article-container">

        <div className="ww-brand-header">
          Weekly Words from Westpark
        </div>

        <div className="ww-post-header">

          <div className="ww-meta-row">
            {post.department && (
              <span className="ww-tag">
                {post.department}
              </span>
            )}

            {(post.article_date || post.created) && (
              <span className="ww-date">
                {new Date(
                  post.article_date || post.created!
                ).toLocaleDateString('en-US', {
                  year: 'numeric',
                  month: 'long',
                  day: 'numeric'
                })}
              </span>
            )}
              {post.publishing_contact && (
              <div className="ww-info-item">
                <User size={16} />
                <span>{post?.publishing_contact}</span>
              </div>
            )}
          </div>

          <h1 className="ww-post-title">
            {post.title}
          </h1>

      

        </div>

        {post.publishing_rollup_image_url && (
          <div className="ww-hero-wrap">
            <img
              src={post.publishing_rollup_image_url}
              alt={post.title}
              className="ww-hero-image"
            />
          </div>
        )}

        {post.vaughn_content && (
          <div
            className="ww-content"
            dangerouslySetInnerHTML={{
              __html: post.vaughn_content
            }}
          />
        )}

        {related.length > 0 && (
          <div className="ww-related">

            <h2 className="ww-related-title">
              In Case You Missed It
            </h2>

            <div className="ww-related-grid">

              {related.map((r) => (
                <a
                  key={r.id}
                  href={buildViewerUrl(r.id)}
                  className="ww-related-card"
                >
                  <div className="ww-related-image-wrap">
                    <img
                      src={
                        r.publishing_rollup_image_url ||
                        'https://images.pexels.com/photos/3861969/pexels-photo-3861969.jpeg?auto=compress&cs=tinysrgb&w=800'
                      }
                      alt={r.title}
                      className="ww-related-image"
                    />
                  </div>

                  <div className="ww-related-body">
                    <h3 className="ww-related-heading">
                      {r.title}
                    </h3>
                  </div>
                </a>
              ))}

            </div>

          </div>
        )}

      </div>
    </article>
  );
};

export default WeeklyWordsViewer;

/* =========================================
DATA HELPERS
========================================= */

async function fetchItems(
  sp: ReturnType<typeof spfi>,
  listId: string,
  publishingSource: 'Weekly Words' | 'Department Specific',
  department: string | undefined,
  top: number
): Promise<INormalizedItem[]> {

  const today = new Date();
  today.setHours(23, 59, 59, 999);

  const filter = `
    (PublishingSource eq '${escapeOData(publishingSource)}'
    or PublishingSource eq 'Both')
    and OData__ModerationStatus eq 0
    and (ArticleDate le datetime'${today.toISOString()}')
    ${publishingSource === 'Department Specific' && department
      ? `and Department eq '${escapeOData(department)}'`
      : ''}
  `;

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
      'PublishingContact/Title',
      'PublishingContact/EMail'
    )
    .expand('InCaseYouMissed', 'PublishingContact')
    .filter(filter)
    .orderBy('ArticleDate', false)
    .top(top)();

  return rows.map(normalizeItem);
}

async function fetchByIds(
  sp: ReturnType<typeof spfi>,
  listId: string,
  ids: number[]
): Promise<INormalizedItem[]> {

  if (!ids.length) return [];

  const filter = ids
    .map((id) => `Id eq ${id}`)
    .join(' or ');

  const rows: ISpItem[] = await sp.web.lists
    .getById(listId)
    .items
    .select(
      'Id',
      'Title',
      'PublishingRollupImage'
    )
    .filter(filter)
    .top(4)();

  return rows.map(normalizeItem);
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
      'PublishingContact/Title',
      'PublishingContact/EMail'
    )
    .expand('InCaseYouMissed', 'PublishingContact')
    .filter(`Id eq ${id}`)
    .top(1)();

  if (!rows?.length) return null;

  return normalizeItem(rows[0]);
}

function normalizeItem(r: ISpItem): INormalizedItem {
  return {
    id: r.Id,
    title: r.Title,
    vaughn_content: r.VaughnContent,
    department: r.Department,
    publishing_source: r.PublishingSource,
    publishing_rollup_image_url:
      typeof r.PublishingRollupImage === 'string'
        ? r.PublishingRollupImage
        : r.PublishingRollupImage?.Url || '',
    abstract: r.Abstract,
    article_date: r.ArticleDate,
    created: r.Created,
    publishing_contact: r.PublishingContact?.Title,
    publishing_contact_email: r.PublishingContact?.EMail,
    in_case_you_missed_ids:
      r.InCaseYouMissed?.map((x) => x.Id)
  };
}

function escapeOData(val: string) {
  return val.replace(/'/g, "''");
}

function parseError(e: any): string {
  if (!e) return 'Unknown error';
  if (typeof e === 'string') return e;
  if (e?.message) return e.message;

  try {
    return JSON.stringify(e);
  } catch {
    return String(e);
  }
}

function getQueryParam(name: string) {
  if (typeof window === 'undefined') return null;

  return new URLSearchParams(
    window.location.search
  ).get(name);
}