// SearchDocuments.tsx
import * as React from "react";
import { useState, useMemo, useEffect, useRef } from "react";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/search";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { TooltipHost, IconButton } from "@fluentui/react";
import DOMPurify from "dompurify";
import {
  TextField,
  PrimaryButton,
  ChoiceGroup,
  IChoiceGroupOption,
  Checkbox,
  Stack,
  Spinner,
  Persona,
  PersonaSize,
} from "@fluentui/react";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { ResponseType } from "@microsoft/microsoft-graph-client";
import { LivePersona } from "@pnp/spfx-controls-react/lib/LivePersona";

interface ISearchDocumentsProps {
  context: any;
  msGraphClientFactory: any; // this.context.msGraphClientFactory
  /** UI variant controlled from property pane via parent webpart */
  variant?: "SearchBar" | "SearchWithResult";
}

/** 🔧 CONFIG – change these if needed */
const SITE_URL = "https://vaughnconstruction.sharepoint.com";
const DOCUMENT_LIBRARY_ID = "8177736d-9faa-49cc-82b3-4ff5a93fa02e"; // Resource Library doc library Id

/** Other sources */
const WEEKLY_WORDS_LIST_ID = "b2b9bc28-8e97-4716-9b88-3c456f8d2b6a";
// ✅ IMPORTANT: use the working Lessons Learned list id
const LESSONS_LEARNED_LIST_ID = "42F972A2-3D9D-4DBE-86C8-32EBAF4ACFDB";

const WEEKLY_WORDS_SITE_URL = `${SITE_URL}/news`;
const LESSONS_LEARNED_SITE_URL = `${SITE_URL}/ll`;

// ✅ IT Blogs (IT SitePages library)
const IT_SITE_URL = `${SITE_URL}/sites/ITSite`;
const IT_BLOGS_LIBRARY_ID = "7dd1ac79-1372-41a6-972f-0e6eca0d9f0e";
const IT_BLOGS_PATH_CONTAINS = "/sites/ITSite/SitePages";

// ✅ NEW: Company Calendar (Events list) — Graph-only
const COMPANY_CAL_SITE_URL = `${SITE_URL}/sites/OPSSite`;
const COMPANY_CAL_LIST_ID = "3c14f133-7cb6-4582-ac0a-4e5c401a5952";

// ✅ NEW: Corporate Responsibilities (list) — Graph-only
const CORP_RESP_SITE_URL = `${SITE_URL}/sites/opsite`;
const CORP_RESP_LIST_ID = "8c97c9e1-ecaf-4fcb-93a9-4016509d185c";

// ✅ NEW: Marketing Calendar (Events list) — Graph-only
const MARKETING_CAL_SITE_URL = `${SITE_URL}/Sites/Marketing/`;
const MARKETING_CAL_LIST_ID = "9be28777-0981-4627-a366-779f280711ef";

// ✅ NEW: Locations (list) — Graph-only
const LOCATIONS_SITE_URL = `${SITE_URL}/`;
const LOCATIONS_LIST_ID = "450dda32-a7e7-4439-8779-e1cc1523e7fd";

// People directory result source
const PEOPLE_SOURCE_ID = "B09A7990-05EA-4AF9-81EF-EDFAB16C4E31";

/** Target page that will show full results */
const RESULTS_PAGE_URL =
  "https://vaughnconstruction.sharepoint.com/sitepages/Custom-Search-Results.aspx";

type Mode = "exact" | "all" | "any";

/** Searchable fields for KQL (Resource Library scoring) */
const fields = ["Title", "Filename", "Content", "Author", "Tags"];

// Escape special chars for KQL
const escapeKql = (text: string) => {
  return text.replace(/(["\\()])/g, "\\$1");
};

// Build KQL for exact / all / any across multiple fields
const buildKql = (text: string, mode: Mode, searchFields: string[]) => {
  const escapedText = escapeKql(text.trim());
  const words = escapedText.split(/\s+/);

  if (mode === "exact") {
    const phrase = `"${escapedText}"`;
    return searchFields.map((f) => `${f}:${phrase}`).join(" OR ");
  }

  if (mode === "all") {
    const perWord = words.map((w) =>
      searchFields.map((f) => `${f}:${w}`).join(" OR ")
    );
    return perWord.join(" AND ");
  }

  // mode === "any"
  const clauses: string[] = [];
  for (const w of words) {
    for (const f of searchFields) {
      clauses.push(`${f}:${w}`);
    }
  }
  return clauses.join(" OR ");
};

/** Build the queryString for Graph search */
const buildGraphQueryString = (text: string, mode: Mode) => {
  const trimmed = text.trim();
  const words = trimmed.split(/\s+/).filter(Boolean);
  if (!trimmed) return "";

  if (mode === "exact") return `"${trimmed}"`;
  if (mode === "all") return trimmed;

  // any
  if (words.length === 0) return trimmed;
  return words.join(" OR ");
};

// Highlight keyword function
const highlight = (text: string, words: string[]) => {
  if (!text) return "";
  let newText = text;
  words.forEach((w) => {
    if (!w) return;
    const reg = new RegExp(`(${w})`, "gi");
    newText = newText.replace(reg, `<mark style="background:yellow">$1</mark>`);
  });
  return newText;
};

// Normalize GUIDs (Graph sometimes returns {GUID} / different casing)
const normalizeGuid = (g?: string) =>
  (g || "")
    .toLowerCase()
    .replace(/[{}]/g, "")
    .replace(/-/g, "");

// Normalize UPN / account formats for photo+LivePersona
const normalizeUpn = (value?: string) => {
  const v = (value || "").trim();
  if (!v) return "";
  // handle claims: i:0#.f|membership|user@domain.com
  const pipe = v.lastIndexOf("|");
  if (pipe >= 0 && pipe < v.length - 1) return v.substring(pipe + 1).toLowerCase();
  return v.toLowerCase();
};


// ==========================
// ✅ People filtering helpers
// ==========================

// Exclude obvious external/guest accounts returned by People Search
// Common patterns:
// - Azure AD B2B guests often contain "#EXT#" in UPN/email or AccountName
// - SharePoint claims may include "urn:spo:guest"
// - Some guests have no WorkEmail at all (keep conservative)
const isGuestPerson = (p: Partial<ISearchResult>) => {
  const account = (p.AccountName || "").toLowerCase();
  const sip = (p.SipAddress || "").toLowerCase();
  const email = (p.WorkEmail || "").toLowerCase();

  // If we literally have no identity signals, ignore it
  if (!account && !sip && !email) return true;

  const hay = `${account} ${sip} ${email}`;

  if (hay.includes("#ext#")) return true;
  if (hay.includes("urn:spo:guest")) return true;

  return false;
};

// Local, strict "double-check" match so unexpected People results don't show.
// Only matches against the allowed People fields.
const peopleMatchesQuery = (p: Partial<ISearchResult>, queryText: string, mode: Mode) => {
  const q = (queryText || "").trim().toLowerCase();
  if (!q) return true;

  const hay = [
    p.PreferredName,
    p.Title,
    p.JobTitle,
    p.Department,
    p.WorkEmail,
    p.SipAddress,
    p.AccountName,
    p.OfficeNumber,
    p.CellPhone,
  ]
    .filter(Boolean)
    .join(" ")
    .toLowerCase();

  const words = q.split(/\s+/).filter(Boolean);

  if (mode === "exact") return hay.includes(q);
  if (mode === "all") return words.every((w) => hay.includes(w));
  return words.some((w) => hay.includes(w)); // any
};

type SourceKey =
  | "resource"
  | "weekly"
  | "lessons"
  | "people"
  | "itblogs"
  | "marketingcal"
  | "companycal"
  | "locations"
  | "corpresp";

const ALL_SOURCES: SourceKey[] = [
  "resource",
  "weekly",
  "lessons",
  "itblogs",
  "people",
  "marketingcal",
  "companycal",
  "locations",
  "corpresp",
];

type ResultSource =
  | "ResourceLibrary"
  | "WeeklyWords"
  | "LessonsLearned"
  | "People"
  | "ITBlogs"
  | "MarketingCalendar"
  | "CompanyCalendar"
  | "Locations"
  | "CorporateResponsibilities";

interface ISearchResult {
  Title: string;
  Filename?: string;
  Path?: string;
  OriginalPath?: string;
  HitHighlightedSummary?: string; // used ONLY for tooltip preview on specific sources
  FileType?: string;
  ListItemId?: number;
  LastModifiedTime?: string;
  Author?: string;

  /** Which source this came from */
  Source: ResultSource;

  /** People-specific fields (only for Source === "People") */
  PreferredName?: string;
  JobTitle?: string;
  Department?: string;
  AccountName?: string;
  SipAddress?: string;
  WorkEmail?: string;
  CellPhone?: string;
  OfficeNumber?: string;

  /** Common list id mapping for view pages */
  OriginalID?: number;

  /** Event list fields (Marketing/Company calendars) */
  EventDate?: string;
  EndDate?: string;
  EventLocation?: string;
  AllDayEvent?: boolean;
  Category?: string;
  RecurrenceData?: string;
  fRecurrence?: boolean;
  FreeBusy?: string;
  Overbook?: string;

  /** Locations */
  ProjectLocation?: string; // Project_x0020_Location
  Location?: string;
  Phone?: string;
  Fax?: string;

  /** Corporate responsibilities */
  Who?: string;
}

// Simple mapping for badges
const sourceLabels: Record<ResultSource, string> = {
  ResourceLibrary: "Resource Library",
  WeeklyWords: "Weekly Words",
  LessonsLearned: "Lessons Learned",
  People: "People Directory",
  ITBlogs: "IT Blogs",
  MarketingCalendar: "Marketing Calendar",
  CompanyCalendar: "Company Calendar",
  Locations: "Locations",
  CorporateResponsibilities: "Corporate Responsibilities",
};

// Photo cache like UsersTable
const peoplePhotoCache = new Map<
  string,
  { url: string | null; loading: boolean; error: boolean }
>();

const PeoplePersonaWithLive: React.FC<{
  context: any;
  graphClient: MSGraphClientV3 | null;
  displayName: string;
  upnOrEmail: string;
  jobTitle?: string;
  department?: string;
}> = ({ context, graphClient, displayName, upnOrEmail, jobTitle, department }) => {
  const normalized = normalizeUpn(upnOrEmail);

  const [photoState, setPhotoState] = useState<{
    url: string | null;
    loading: boolean;
    error: boolean;
  }>({ url: null, loading: false, error: false });

  const mountedRef = useRef(true);
  const loadingRef = useRef(false);

  const getInitials = (name: string): string => {
    if (!name) return "??";
    return name
      .trim()
      .split(/\s+/)
      .map((w) => w?.charAt(0).toUpperCase())
      .join("")
      .substring(0, 2);
  };

  // hydrate from cache
  useEffect(() => {
    if (!normalized) return;
    const cached = peoplePhotoCache.get(normalized);
    if (cached) setPhotoState(cached);
  }, [normalized]);

  // load photo
  useEffect(() => {
    const loadPhoto = async () => {
      if (!graphClient || !normalized || loadingRef.current) return;

      const cached = peoplePhotoCache.get(normalized);
      if (cached && (cached.url || cached.error)) {
        setPhotoState(cached);
        return;
      }
      if (cached?.loading) return;

      loadingRef.current = true;
      const loadingState = { url: null, loading: true, error: false };
      peoplePhotoCache.set(normalized, loadingState);
      setPhotoState(loadingState);

      try {
        const blob = await graphClient
          .api(`/users/${normalized}/photo/$value`)
          .responseType(ResponseType.BLOB)
          .get();

        if (!mountedRef.current) return;

        const url = URL.createObjectURL(blob as Blob);
        const ok = { url, loading: false, error: false };
        peoplePhotoCache.set(normalized, ok);
        setPhotoState(ok);
      } catch {
        const fail = { url: null, loading: false, error: true };
        peoplePhotoCache.set(normalized, fail);
        if (mountedRef.current) setPhotoState(fail);
      } finally {
        loadingRef.current = false;
      }
    };

    loadPhoto();
  }, [graphClient, normalized]);

  useEffect(() => {
    return () => {
      mountedRef.current = false;
    };
  }, []);

  const personaElement = (
    <Persona
      text={displayName}
      secondaryText={jobTitle}
      tertiaryText={department}
      size={PersonaSize.size40}
      imageUrl={photoState.url || undefined}
      imageInitials={photoState.url ? undefined : getInitials(displayName)}
      coinSize={40}
    />
  );

  // LivePersona hover overlay
  if (context?.serviceScope && normalized && !photoState.error) {
    return (
      <div style={{ position: "relative" }}>
        {personaElement}
        <div style={{ position: "absolute", inset: 0, opacity: 0, pointerEvents: "all" }}>
          <LivePersona serviceScope={context.serviceScope} upn={normalized} template={personaElement} />
        </div>
      </div>
    );
  }

  return personaElement;
};

// ✅ SharePoint Search HitHighlightedSummary uses tags like <c0> and <ddd/>
const normalizeSearchSummaryHtml = (html: string) => {
  if (!html) return "";

  return (
    html
      .replace(/<ddd\s*\/>/gi, "… ")
      .replace(/<c0>/gi, '<mark style="background:yellow">')
      .replace(/<\/c0>/gi, "</mark>")
      .replace(/<c\d+>/gi, '<mark style="background:yellow">')
      .replace(/<\/c\d+>/gi, "</mark>")
  );
};

const sanitizeHtml = (html: string) => {
  return DOMPurify.sanitize(html, {
    USE_PROFILES: { html: true },
    ADD_TAGS: ["mark"],
    ADD_ATTR: ["style"],
  });
};

const getSafeHtmlForTooltip = (r: ISearchResult, words: string[]) => {
  if (r.Source === "WeeklyWords" || r.Source === "LessonsLearned") {
    return sanitizeHtml(r.HitHighlightedSummary || "");
  }
  if (r.Source === "ITBlogs") {
    return sanitizeHtml(normalizeSearchSummaryHtml(r.HitHighlightedSummary || ""));
  }
  return sanitizeHtml(highlight(r.HitHighlightedSummary || "", words));
};

const SearchDocuments: React.FC<ISearchDocumentsProps> = ({
  context,
  msGraphClientFactory,
  variant = "SearchWithResult",
}) => {
  const spRoot = useMemo(() => spfi(SITE_URL).using(SPFx(context)), [context]);

  const [graphClient, setGraphClient] = useState<MSGraphClientV3 | null>(null);

  useEffect(() => {
    const initGraph = async () => {
      if (!msGraphClientFactory) return;
      const client = await msGraphClientFactory.getClient("3");
      setGraphClient(client);
    };
    initGraph();
  }, [msGraphClientFactory]);

  /** STATE **/
  const [query, setQuery] = useState("");
  const [mode, setMode] = useState<Mode>("exact");
  const [results, setResults] = useState<ISearchResult[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string>("");

  // Result-type style file type filter – applies only to Resource Library
  const [fileTypeFilters, setFileTypeFilters] = useState<string[]>([]);
  const [availableFileTypes, setAvailableFileTypes] = useState<string[]>([
    "docx",
    "xlsx",
    "pptx",
    "pdf",
    "msg",
    "eml",
  ]);

  /** Source filters */
  const [sourceFilters, setSourceFilters] = useState<SourceKey[]>(ALL_SOURCES);

  /** Pagination **/
  const pageSize = 10;
  const [page, setPage] = useState(1);
  const totalPages = Math.ceil(results.length / pageSize);
  const paginatedResults = results.slice((page - 1) * pageSize, page * pageSize);

  const words = query.trim() ? query.trim().split(/\s+/) : [];

  const modeOptions: IChoiceGroupOption[] = [
    { key: "exact", text: "Exact phrase" },
    { key: "all", text: "All words" },
    { key: "any", text: "Any word" },
  ];

  const toggleFileType = (ext: string, checked?: boolean) => {
    setPage(1);
    setFileTypeFilters((prev) => {
      const exists = prev.includes(ext);
      if (checked && !exists) return [...prev, ext];
      if (!checked && exists) return prev.filter((e) => e !== ext);
      return prev;
    });
  };

  const toggleSource = (key: SourceKey, checked?: boolean) => {
    setPage(1);
    setSourceFilters((prev) => {
      const exists = prev.includes(key);
      if (checked && !exists) return [...prev, key];
      if (!checked && exists) return prev.filter((k) => k !== key);
      return prev;
    });
  };

  // ✅ All checkbox should select/unselect ALL at once
  const handleAllSourcesToggle = (_e: any, checked?: boolean) => {
    setPage(1);
    setSourceFilters(checked ? ALL_SOURCES : []);
  };

  // Dashboard stats: count by source
  const sourceCounts = useMemo(() => {
    const map: Record<string, number> = {};
    results.forEach((r) => {
      map[r.Source] = (map[r.Source] || 0) + 1;
    });
    return map;
  }, [results]);

  // ==========================
  // ✅ Graph helpers for new 4 sources
  // ==========================
  const siteIdCache = useRef(new Map<string, string>());

  const getGraphSiteId = async (absoluteSiteUrl: string): Promise<string> => {
    if (!graphClient) throw new Error("Graph client not initialized.");
    const u = new URL(absoluteSiteUrl);
    const hostname = u.hostname;
    const path = u.pathname.replace(/\/$/, ""); // no trailing slash
    const key = `${hostname}:${path || "/"}`;

    const cached = siteIdCache.current.get(key);
    if (cached) return cached;

    // Graph format: /sites/{hostname}:{server-relative-path}
    const resp: any = await graphClient.api(`/sites/${hostname}:${path || "/"}`).get();
    const id = resp?.id;
    if (!id) throw new Error(`Unable to resolve site id for ${absoluteSiteUrl}`);
    siteIdCache.current.set(key, id);
    return id;
  };

  const graphSearchListItems = async (queryString: string, listId: string): Promise<any[]> => {
    if (!graphClient) return [];
    const body = {
      requests: [
        {
          entityTypes: ["listItem"],
          query: { queryString },
          from: 0,
          size: 50,
          fields: ["Title", "webUrl", "lastModifiedDateTime", "createdBy", "sharepointIds"],
        },
      ],
    };

    const res: any = await graphClient.api("/search/query").post(body);
    const hits = res?.value?.[0]?.hitsContainers?.[0]?.hits ?? [];
    const target = normalizeGuid(listId);

    return hits
      .map((h: any) => h.resource)
      .filter((r: any) => normalizeGuid(r?.sharepointIds?.listId) === target);
  };

  // Get item fields via Graph (Graph-only requirement)
  const graphGetListItemFields = async (
    siteUrl: string,
    listId: string,
    listItemId: number,
    fieldsSelect: string[]
  ) => {
    if (!graphClient) throw new Error("Graph client not initialized.");

    const siteId = await getGraphSiteId(siteUrl);

    // Expand fields and select internal names
    const selectPart = fieldsSelect.join(",");
    const item: any = await graphClient
      .api(`/sites/${siteId}/lists/${listId}/items/${listItemId}?expand=fields($select=${selectPart})`)
      .get();

    const fieldsObj = item?.fields || {};
    const createdBy = item?.createdByUser?.displayName || item?.createdBy?.user?.displayName || "";
    const lastModifiedBy = item?.lastModifiedByUser?.displayName || item?.lastModifiedBy?.user?.displayName || "";

    return { fields: fieldsObj, createdBy, lastModifiedBy };
  };

  /** ===== Core search engine (merged across selected sources) ===== */
  const runSearch = async (overrideText?: string, overrideMode?: Mode) => {
    const textToUse = overrideText !== undefined ? overrideText : query;
    const modeToUse = overrideMode !== undefined ? overrideMode : mode;

    const trimmed = textToUse.trim();
    if (!trimmed) {
      setResults([]);
      setError("");
      return;
    }

    // 🔁 When in SearchBar-only mode: just redirect to results page
    if (variant === "SearchBar") {
      const url = `${RESULTS_PAGE_URL}?q=${encodeURIComponent(trimmed)}`;
      window.open(url, "_blank");
      return;
    }

    setLoading(true);
    setError("");
    setPage(1);

    try {
      const kqlBody = buildKql(trimmed, modeToUse, fields);
      const graphQueryString = buildGraphQueryString(trimmed, modeToUse);

      // Determine which sources are active (if user cleared all, treat as ALL)
      const activeSources: SourceKey[] =
        sourceFilters.length > 0 ? sourceFilters : ALL_SOURCES;

      // Graph is needed for weekly/lessons/itblogs + NEW 4 sources + photo
      const needGraph = activeSources.some((s) =>
        ["weekly", "lessons", "itblogs", "marketingcal", "companycal", "locations", "corpresp"].includes(s)
      );

      if (needGraph && !graphClient) {
        setLoading(false);
        setError("Graph client is not initialized yet. Please try again.");
        return;
      }

      const searchPromises: Promise<ISearchResult[]>[] = [];

      // ============================================================
      // RESOURCE LIBRARY – SharePoint Search (stable)
      // ============================================================
      if (activeSources.includes("resource")) {
        const docsPromise = (async (): Promise<ISearchResult[]> => {
          let filter = `(${kqlBody}) AND (contentclass:STS_ListItem_DocumentLibrary AND IsDocument:"true" AND ListId:"${DOCUMENT_LIBRARY_ID}")`;

          if (fileTypeFilters.length > 0) {
            const ftClause = fileTypeFilters.map((ft) => `FileType:${ft}`).join(" OR ");
            filter += ` AND (${ftClause})`;
          }

          const res: any = await spRoot.search({
            Querytext: filter,
            RowLimit: 500,
            TrimDuplicates: false,
            SelectProperties: [
              "Title",
              "Path",
              "OriginalPath",
              "FileType",
              "Filename",
              "ListItemId",
              "LastModifiedTime",
              "Author",
              "HitHighlightedSummary",
            ],
            EnableQueryRules: true,
            SortList: [{ Property: "Rank", Direction: 1 }],
          });

          const items: ISearchResult[] = (res.PrimarySearchResults || []).map((i: any) => ({
            Title: i.Title || "",
            Filename: i.Filename || "",
            Path: i.Path,
            OriginalPath: i.OriginalPath,
            HitHighlightedSummary: i.HitHighlightedSummary,
            FileType: (i.FileType || "").toLowerCase(),
            ListItemId: Number(i.ListItemId),
            LastModifiedTime: i.LastModifiedTime,
            Author: i.Author,
            Source: "ResourceLibrary",
          }));

          const dynamicTypes: string[] = Array.from(
            new Set(items.map((it) => it.FileType).filter((x): x is string => !!x && x.trim().length > 0))
          );
          if (dynamicTypes.length > 0) setAvailableFileTypes(dynamicTypes);

          return items;
        })();

        searchPromises.push(docsPromise);
      }

      // ============================================================
      // IT BLOGS – Graph Search (listItem) + PnP full item load (keep as before)
      // ============================================================
      if (activeSources.includes("itblogs") && graphClient) {
        const itBlogsPromise = (async (): Promise<ISearchResult[]> => {
          const body = {
            requests: [
              {
                entityTypes: ["listItem"],
                query: { queryString: graphQueryString },
                from: 0,
                size: 50,
                fields: ["Title", "webUrl", "lastModifiedDateTime", "createdBy", "sharepointIds"],
              },
            ],
          };

          const res: any = await graphClient.api("/search/query").post(body);
          const hits = res?.value?.[0]?.hitsContainers?.[0]?.hits ?? [];

          const targetListId = normalizeGuid(IT_BLOGS_LIBRARY_ID);

          const filtered: any[] = hits
            .map((h: any) => h.resource)
            .filter((r: any) => {
              const listId = normalizeGuid(r?.sharepointIds?.listId);
              if (listId && listId === targetListId) return true;
              return (r?.webUrl || "").toLowerCase().includes(IT_BLOGS_PATH_CONTAINS.toLowerCase());
            });

          if (filtered.length === 0) return [];

          const spIT = spfi(IT_SITE_URL).using(SPFx(context));

          const itemsWithDetails = await Promise.all(
            filtered.map(async (r: any) => {
              const id = parseInt(r?.sharepointIds?.listItemId || "0", 10);
              if (!id) return null;

              try {
                const item: any = await spIT.web.lists
                  .getById(IT_BLOGS_LIBRARY_ID)
                  .items.getById(id)
                  .select("Id", "Title", "CanvasContent1", "Modified", "Author/Title")
                  .expand("Author")();

                const authorName = item.Author?.Title || r.createdBy?.user?.displayName || "";

                return {
                  Title: item.Title || r?.fields?.Title || "(no title)",
                  Filename: "",
                  Path: r.webUrl,
                  OriginalPath: "",
                  HitHighlightedSummary: item.CanvasContent1 || "",
                  FileType: "aspx",
                  ListItemId: item.Id,
                  LastModifiedTime: r.lastModifiedDateTime || item.Modified || "",
                  Author: authorName,
                  Source: "ITBlogs" as const,
                } as ISearchResult;
              } catch {
                return {
                  Title: r?.fields?.Title || "(no title)",
                  Filename: "",
                  Path: r.webUrl,
                  OriginalPath: "",
                  HitHighlightedSummary: "",
                  FileType: "aspx",
                  ListItemId: 0,
                  LastModifiedTime: r.lastModifiedDateTime || "",
                  Author: r.createdBy?.user?.displayName || "",
                  Source: "ITBlogs" as const,
                } as ISearchResult;
              }
            })
          );

          return itemsWithDetails.filter((x): x is ISearchResult => x !== null);
        })();

        searchPromises.push(itBlogsPromise);
      }

      // ============================================================
      // WEEKLY WORDS – Graph Search (listItem) + PnP full item load
      // ============================================================
      if (activeSources.includes("weekly") && graphClient) {
        const weeklyPromise = (async (): Promise<ISearchResult[]> => {
          const body = {
            requests: [
              {
                entityTypes: ["listItem"],
                query: { queryString: graphQueryString },
                from: 0,
                size: 50,
                fields: [
                  "Title",
                  "ArticleDate",
                  "VaughnContent",
                  "Abstract",
                  "PublishingContact",
                  "webUrl",
                  "lastModifiedDateTime",
                  "createdBy",
                  "sharepointIds",
                  "OriginalID",
                ],
              },
            ],
          };

          const res: any = await graphClient.api("/search/query").post(body);
          const hits = res?.value?.[0]?.hitsContainers?.[0]?.hits ?? [];

          const filtered: any[] = hits
            .map((h: any) => h.resource)
            .filter((r: any) => normalizeGuid(r?.sharepointIds?.listId) === normalizeGuid(WEEKLY_WORDS_LIST_ID));

          if (filtered.length === 0) return [];

          const spWeekly = spfi(WEEKLY_WORDS_SITE_URL).using(SPFx(context));

          const itemsWithDetails = await Promise.all(
            filtered.map(async (r: any) => {
              const id = parseInt(r.sharepointIds?.listItemId || "0", 10);
              if (!id) return null;

              try {
                const item: any = await spWeekly.web.lists
                  .getById(WEEKLY_WORDS_LIST_ID)
                  .items.getById(id)
                  .select(
                    "Id",
                    "Title",
                    "ArticleDate",
                    "VaughnContent",
                    "Abstract",
                    "PublishingContact/Title",
                    "PublishingContact/Id",
                    "OriginalID",
                    "Modified"
                  )
                  .expand("PublishingContact")();

                const fullBodyHtml = item.VaughnContent || item.Abstract || "";
                const authorName = item.PublishingContact?.Title || r.createdBy?.user?.displayName || "";
                const originalId = item.OriginalID ? Number(item.OriginalID) : undefined;

                return {
                  Title: item.Title || r.fields?.Title || "(no title)",
                  Filename: "",
                  Path: r.webUrl,
                  OriginalPath: "",
                  HitHighlightedSummary: fullBodyHtml, // tooltip only
                  FileType: "item",
                  ListItemId: item.Id,
                  OriginalID: originalId,
                  LastModifiedTime: item.Modified || r.lastModifiedDateTime || "",
                  Author: authorName,
                  Source: "WeeklyWords" as const,
                } as ISearchResult;
              } catch (err) {
                console.error("WeeklyWords item fetch failed", err);
                return null;
              }
            })
          );

          return itemsWithDetails.filter((x): x is ISearchResult => x !== null);
        })();

        searchPromises.push(weeklyPromise);
      }

      // ============================================================
      // LESSONS LEARNED – Graph Search (listItem) + PnP full item load
      // ============================================================
      if (activeSources.includes("lessons") && graphClient) {
        const lessonsPromise = (async (): Promise<ISearchResult[]> => {
          const body = {
            requests: [
              {
                entityTypes: ["listItem"],
                query: { queryString: graphQueryString },
                from: 0,
                size: 50,
                fields: [
                  "Title",
                  "Body",
                  "Contact",
                  "webUrl",
                  "lastModifiedDateTime",
                  "createdBy",
                  "sharepointIds",
                  "OriginalID",
                ],
              },
            ],
          };

          const res: any = await graphClient.api("/search/query").post(body);
          const hits = res?.value?.[0]?.hitsContainers?.[0]?.hits ?? [];

          const targetListId = normalizeGuid(LESSONS_LEARNED_LIST_ID);

          const filtered: any[] = hits
            .map((h: any) => h.resource)
            .filter((r: any) => {
              const listId = normalizeGuid(r?.sharepointIds?.listId);
              if (listId && listId === targetListId) return true;
              return (r?.webUrl || "").toLowerCase().includes("/ll/");
            });

          if (filtered.length === 0) return [];

          const spLessons = spfi(LESSONS_LEARNED_SITE_URL).using(SPFx(context));

          const itemsWithDetails = await Promise.all(
            filtered.map(async (r: any) => {
              const id = parseInt(r.sharepointIds?.listItemId || "0", 10);
              if (!id) return null;

              try {
                const item: any = await spLessons.web.lists
                  .getById(LESSONS_LEARNED_LIST_ID)
                  .items.getById(id)
                  .select(
                    "Id",
                    "Title",
                    "PublishedDate",
                    "Body",
                    "Contact/Title",
                    "OriginalID",
                    "Modified"
                  )
                  .expand("Contact")();

                const authorName = item.Contact?.Title || r.createdBy?.user?.displayName || "";
                const originalId = item.OriginalID ? Number(item.OriginalID) : undefined;

                return {
                  Title: item.Title || r.fields?.Title || "(no title)",
                  Filename: "",
                  Path: r.webUrl,
                  OriginalPath: "",
                  HitHighlightedSummary: item.Body || "", // tooltip only
                  FileType: "item",
                  ListItemId: item.Id,
                  OriginalID: originalId,
                  LastModifiedTime: item.Modified || item.PublishedDate || r.lastModifiedDateTime || "",
                  Author: authorName,
                  Source: "LessonsLearned" as const,
                } as ISearchResult;
              } catch (err) {
                console.error("LessonsLearned item fetch failed", err);
                return null;
              }
            })
          );

          return itemsWithDetails.filter((x): x is ISearchResult => x !== null);
        })();

        searchPromises.push(lessonsPromise);
      }

      // ============================================================
      // PEOPLE DIRECTORY – SharePoint Search with People Source
      // ============================================================
      if (activeSources.includes("people")) {
        const peoplePromise = (async (): Promise<ISearchResult[]> => {
          const peopleQuery = {
            Querytext: trimmed,
            SourceId: PEOPLE_SOURCE_ID,
            RowLimit: 500,
            RowsPerPage: 500,
            TrimDuplicates: true,
            SelectProperties: [
              "SipAddress",
              "PreferredName",
              "JobTitle",
              "Department",
              "WorkEmail",
              "CellPhone",
              "OfficeNumber",
              "AccountName",
            ],
          };

          const res: any = await spRoot.search(peopleQuery);

          const items: ISearchResult[] = (res.PrimarySearchResults || []).map((i: any) => ({
            Title: i.PreferredName || "(no name)",
            Filename: "",
            Path: i.WorkEmail || i.SipAddress || "",
            OriginalPath: "",
            FileType: "person",
            ListItemId: 0,
            LastModifiedTime: "",
            Author: "",
            Source: "People",
            PreferredName: i.PreferredName,
            JobTitle: i.JobTitle,
            Department: i.Department,
            AccountName: i.AccountName,
            SipAddress: i.SipAddress,
            WorkEmail: i.WorkEmail,
            CellPhone: i.CellPhone,
            OfficeNumber: i.OfficeNumber,
          }));

          return items
            .filter((p) => !isGuestPerson(p))
            .filter((p) => peopleMatchesQuery(p, trimmed, modeToUse));
        })();

        searchPromises.push(peoplePromise);
      }

      // ============================================================
      // ✅ NEW 4 SOURCES — Graph-only (search + field read)
      // ============================================================

      // Calendars: use columns shared by you
      const calendarFieldSelect = [
        "Id",
        "Title",
        "Location",
        "EventDate",
        "RecurrenceData",
        "fRecurrence",
        "fAllDayEvent",
        "EndDate",
        "Description",
        "Category",
        "FreeBusy",
        "Overbook",
        "Modified",
        "Created",
        // NOTE: Author/Editor/ParticipantsPicker are complex; Graph fields may return string/lookup ids.
        "ParticipantsPicker",
      ];

      if (activeSources.includes("companycal") && graphClient) {
        const p = (async (): Promise<ISearchResult[]> => {
          const hits = await graphSearchListItems(graphQueryString, COMPANY_CAL_LIST_ID);
          if (!hits.length) return [];

          const items = await Promise.all(
            hits.map(async (r: any) => {
              const id = parseInt(r?.sharepointIds?.listItemId || "0", 10);
              if (!id) return null;

              try {
                const { fields: f, createdBy, lastModifiedBy } = await graphGetListItemFields(
                  COMPANY_CAL_SITE_URL,
                  COMPANY_CAL_LIST_ID,
                  id,
                  calendarFieldSelect
                );

                return {
                  Title: f?.Title || r?.fields?.Title || "(no title)",
                  Path: r?.webUrl,
                  FileType: "event",
                  ListItemId: id,
                  LastModifiedTime: f?.Modified || r?.lastModifiedDateTime || "",
                  Author: lastModifiedBy || createdBy || r?.createdBy?.user?.displayName || "",
                  Source: "CompanyCalendar" as const,
                  EventDate: f?.EventDate,
                  EndDate: f?.EndDate,
                  EventLocation: f?.Location,
                  AllDayEvent: f?.fAllDayEvent,
                } as any;
              } catch (e) {
                // fallback: minimal
                return {
                  Title: r?.fields?.Title || "(no title)",
                  Path: r?.webUrl,
                  FileType: "event",
                  ListItemId: id,
                  LastModifiedTime: r?.lastModifiedDateTime || "",
                  Author: r?.createdBy?.user?.displayName || "",
                  Source: "CompanyCalendar" as const,
                } as ISearchResult;
              }
            })
          );

          // fix boolean mapping in TS (we will rewrite after python)
          return items.filter(Boolean) as any;
        })();

        searchPromises.push(p as any);
      }

      if (activeSources.includes("marketingcal") && graphClient) {
        const p = (async (): Promise<ISearchResult[]> => {
          const hits = await graphSearchListItems(graphQueryString, MARKETING_CAL_LIST_ID);
          if (!hits.length) return [];

          const items = await Promise.all(
            hits.map(async (r: any) => {
              const id = parseInt(r?.sharepointIds?.listItemId || "0", 10);
              if (!id) return null;

              try {
                const { fields: f, createdBy, lastModifiedBy } = await graphGetListItemFields(
                  MARKETING_CAL_SITE_URL,
                  MARKETING_CAL_LIST_ID,
                  id,
                  calendarFieldSelect
                );

                return {
                  Title: f?.Title || r?.fields?.Title || "(no title)",
                  Path: r?.webUrl,
                  FileType: "event",
                  ListItemId: id,
                  LastModifiedTime: f?.Modified || r?.lastModifiedDateTime || "",
                  Author: lastModifiedBy || createdBy || r?.createdBy?.user?.displayName || "",
                  Source: "MarketingCalendar" as const,
                  EventDate: f?.EventDate,
                  EndDate: f?.EndDate,
                  EventLocation: f?.Location,
                  AllDayEvent: f?.fAllDayEvent,
                  Category: f?.Category,
                  RecurrenceData: f?.RecurrenceData,
                  fRecurrence: f?.fRecurrence,
                  FreeBusy: f?.FreeBusy,
                  Overbook: f?.Overbook,
                } as ISearchResult;
              } catch {
                return {
                  Title: r?.fields?.Title || "(no title)",
                  Path: r?.webUrl,
                  FileType: "event",
                  ListItemId: id,
                  LastModifiedTime: r?.lastModifiedDateTime || "",
                  Author: r?.createdBy?.user?.displayName || "",
                  Source: "MarketingCalendar" as const,
                } as ISearchResult;
              }
            })
          );

          return items.filter((x): x is ISearchResult => x !== null);
        })();

        searchPromises.push(p);
      }

      if (activeSources.includes("locations") && graphClient) {
        const p = (async (): Promise<ISearchResult[]> => {
          const hits = await graphSearchListItems(graphQueryString, LOCATIONS_LIST_ID);
          if (!hits.length) return [];

          const fieldsSelect = ["Id", "Title", "Project_x0020_Location", "Phone", "Fax", "Location"];

          const items = await Promise.all(
            hits.map(async (r: any) => {
              const id = parseInt(r?.sharepointIds?.listItemId || "0", 10);
              if (!id) return null;

              try {
                const { fields: f, createdBy, lastModifiedBy } = await graphGetListItemFields(
                  LOCATIONS_SITE_URL,
                  LOCATIONS_LIST_ID,
                  id,
                  fieldsSelect
                );

                return {
                  Title: f?.Title || r?.fields?.Title || "(no title)",
                  Path: r?.webUrl,
                  FileType: "item",
                  ListItemId: id,
                  LastModifiedTime: "",
                  Author: lastModifiedBy || createdBy || "",
                  Source: "Locations" as const,
                  ProjectLocation: f?.Project_x0020_Location,
                  Phone: f?.Phone,
                  Fax: f?.Fax,
                  Location: f?.Location,
                } as ISearchResult;
              } catch {
                return {
                  Title: r?.fields?.Title || "(no title)",
                  Path: r?.webUrl,
                  FileType: "item",
                  ListItemId: id,
                  Source: "Locations" as const,
                } as ISearchResult;
              }
            })
          );

          return items.filter((x): x is ISearchResult => x !== null);
        })();

        searchPromises.push(p);
      }

      if (activeSources.includes("corpresp") && graphClient) {
        const p = (async (): Promise<ISearchResult[]> => {
          const hits = await graphSearchListItems(graphQueryString, CORP_RESP_LIST_ID);
          if (!hits.length) return [];

          const fieldsSelect = ["Id", "Title", "Who"];

          const items = await Promise.all(
            hits.map(async (r: any) => {
              const id = parseInt(r?.sharepointIds?.listItemId || "0", 10);
              if (!id) return null;

              try {
                const { fields: f, createdBy, lastModifiedBy } = await graphGetListItemFields(
                  CORP_RESP_SITE_URL,
                  CORP_RESP_LIST_ID,
                  id,
                  fieldsSelect
                );

                return {
                  Title: f?.Title || r?.fields?.Title || "(no title)",
                  Path: r?.webUrl,
                  FileType: "item",
                  ListItemId: id,
                  Source: "CorporateResponsibilities" as const,
                  Who: f?.Who,
                  Author: lastModifiedBy || createdBy || "",
                } as ISearchResult;
              } catch {
                return {
                  Title: r?.fields?.Title || "(no title)",
                  Path: r?.webUrl,
                  FileType: "item",
                  ListItemId: id,
                  Source: "CorporateResponsibilities" as const,
                } as ISearchResult;
              }
            })
          );

          return items.filter((x): x is ISearchResult => x !== null);
        })();

        searchPromises.push(p);
      }

      // Run all active source searches in parallel
      const resultArrays = await Promise.all(searchPromises);

      // NOTE: we injected an invalid python placeholder above for companycal bool; we'll fix below in file post-process
      const merged: ISearchResult[] = ([] as ISearchResult[]).concat(...resultArrays);

      // simple rank: title first
      const lowerQuery = trimmed.toLowerCase();
      const queryWords = lowerQuery.split(/\s+/).filter(Boolean);

      const scored = merged
        .map((it, index) => {
          const title = (it.Title || "").toLowerCase();
          const secondary = (it.Filename || it.Department || it.JobTitle || it.EventLocation || it.Category || it.ProjectLocation || it.Location || it.Who || "").toLowerCase();
          let score = 0;

          if (modeToUse === "exact") {
            if (title.includes(lowerQuery)) score += 100;
            if (secondary.includes(lowerQuery)) score += 50;
          } else {
            let t = 0;
            let s = 0;
            queryWords.forEach((w) => {
              if (w && title.includes(w)) t++;
              if (w && secondary.includes(w)) s++;
            });
            if (modeToUse === "all") {
              score += t * 10 + s * 5;
              if (t === queryWords.length && queryWords.length > 0) score += 40;
            } else {
              score += t * 5 + s * 2;
            }
          }
          return { it, score, index };
        })
        .sort((a, b) => (b.score !== a.score ? b.score - a.score : a.index - b.index))
        .map((x) => x.it);

      setResults(scored);
    } catch (e: any) {
      console.error(e);
      setError(e?.message || "Search failed.");
    }

    setLoading(false);
  };

  // ✅ Auto-run exact search on results page if ?q=... is present,
  // but WAIT for graphClient if required sources are enabled.
  const autoRunRef = useRef<{ q: string; ran: boolean }>({ q: "", ran: false });

  useEffect(() => {
    if (variant !== "SearchWithResult") return;

    let qParam = "";
    try {
      const params = new URLSearchParams(window.location.search);
      qParam = params.get("q") || "";
    } catch {
      qParam = "";
    }

    const trimmed = qParam.trim();
    if (!trimmed) return;

    const activeSources: SourceKey[] = sourceFilters.length > 0 ? sourceFilters : ALL_SOURCES;
    const needGraph = activeSources.some((s) =>
      ["weekly", "lessons", "itblogs", "marketingcal", "companycal", "locations", "corpresp"].includes(s)
    );

    if (needGraph && !graphClient) return;

    if (autoRunRef.current.ran && autoRunRef.current.q === trimmed) return;
    autoRunRef.current = { q: trimmed, ran: true };

    setQuery(trimmed);
    setMode("exact");
    runSearch(trimmed, "exact");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [variant, graphClient]);

  const openResult = (item: ISearchResult) => {
    if (item.Source === "People") {
      if (item.WorkEmail) window.location.href = `mailto:${item.WorkEmail}`;
      else if (item.SipAddress) window.open(`sip:${item.SipAddress}`, "_blank");
      return;
    }

    if (item.Source === "LessonsLearned" && item.OriginalID) {
      const url = `${LESSONS_LEARNED_SITE_URL}/SitePages/LessonsLearnedView.aspx?LessonsLearnedId=${item.OriginalID}`;
      window.open(url, "_blank");
      return;
    }

    const basePath = item.OriginalPath || item.Path;
    if (!basePath) return;

    const isPage =
      (item.FileType || "").toLowerCase() === "aspx" || basePath.toLowerCase().includes(".aspx");

    const url = isPage ? basePath : `${basePath}?web=1`;
    window.open(url, "_blank");
  };

  // ✅ Compact card styles
  const cardStyle: React.CSSProperties = {
    background: "#ffffff",
    borderRadius: 10,
    boxShadow: "0 1px 3px rgba(0,0,0,0.08)",
    padding: 12,
  };

  const pillStyle: React.CSSProperties = {
    display: "inline-block",
    padding: "2px 8px",
    borderRadius: 999,
    fontSize: 10,
    fontWeight: 600,
    textTransform: "uppercase",
    background: "#f3f2f1",
    color: "#605e5c",
    lineHeight: 1.6,
  };

  const metaRowStyle: React.CSSProperties = {
    fontSize: 12,
    color: "#605e5c",
    marginTop: 6,
    display: "flex",
    flexWrap: "wrap",
    gap: 8,
  };

  // 🔹 Simple SearchBar variant
  if (variant === "SearchBar") {
    return (
      <div style={{ padding: 20 }}>
        <Stack tokens={{ childrenGap: 8 }} styles={{ root: { maxWidth: 600 } }}>
          <TextField
            label="Search"
            placeholder="Type search text…"
            value={query}
            onChange={(_e, v) => setQuery(v || "")}
            onKeyDown={(e) => {
              if (e.key === "Enter") {
                e.preventDefault();
                runSearch();
              }
            }}
          />
          <PrimaryButton text="Search" onClick={() => runSearch()} />
        </Stack>
      </div>
    );
  }

  const allSourcesSelected = sourceFilters.length === ALL_SOURCES.length;

  return (
    <div style={{ padding: 20, background: "#faf9f8" }}>
      <div style={{ maxWidth: 1200, margin: "0 auto" }}>
        {/* Header */}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 16 }}>
          <div>
            <h2 style={{ margin: 0 }}>Advanced Search</h2>
            <p style={{ margin: "4px 0 0", color: "#605e5c" }}>
              Search across documents, pages, people, lists, and calendars in one place.
            </p>
          </div>

          {results.length > 0 && (
            <div style={cardStyle}>
              <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 6 }}>Result summary</div>
              <div style={{ fontSize: 12 }}>
                <div style={{ marginBottom: 4 }}>
                  Total: <strong>{results.length} item{results.length !== 1 ? "s" : ""}</strong>
                </div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                  {Object.entries(sourceCounts).map(([key, count]) => (
                    <span key={key} style={pillStyle}>
                      {sourceLabels[key as ResultSource]}: {count}
                    </span>
                  ))}
                </div>
              </div>
            </div>
          )}
        </div>

        {/* Search controls */}
        <div style={{ display: "grid", gridTemplateColumns: "minmax(0, 2fr) minmax(0, 1.4fr)", gap: 16, alignItems: "flex-start" }}>
          <div style={cardStyle}>
            <Stack tokens={{ childrenGap: 8 }}>
              <TextField
                label="Search text"
                placeholder="Type keywords…"
                value={query}
                onChange={(_e, v) => setQuery(v || "")}
                onKeyDown={(e) => {
                  if (e.key === "Enter") {
                    e.preventDefault();
                    runSearch();
                  }
                }}
              />
              <ChoiceGroup label="Match mode" selectedKey={mode} options={modeOptions} onChange={(_e, o) => setMode(o?.key as Mode)} />
              <div style={{ display: "flex", justifyContent: "flex-end" }}>
                <PrimaryButton text="Search" onClick={() => runSearch()} disabled={loading} />
              </div>
            </Stack>
          </div>

          <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
            <div style={cardStyle}>
              <div style={{ fontWeight: 600, marginBottom: 6 }}>Source</div>
              <div style={{ fontSize: 12, color: "#605e5c", marginBottom: 8 }}>
                Select/unselect sources. “All” toggles everything.
              </div>

              <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                <Checkbox label="All" checked={allSourcesSelected} onChange={handleAllSourcesToggle} />

                <Checkbox label="Resource Library" checked={sourceFilters.includes("resource")} onChange={(_e, checked) => toggleSource("resource", !!checked)} />
                <Checkbox label="Weekly Words" checked={sourceFilters.includes("weekly")} onChange={(_e, checked) => toggleSource("weekly", !!checked)} />
                <Checkbox label="Lessons Learned" checked={sourceFilters.includes("lessons")} onChange={(_e, checked) => toggleSource("lessons", !!checked)} />
                <Checkbox label="IT Blogs" checked={sourceFilters.includes("itblogs")} onChange={(_e, checked) => toggleSource("itblogs", !!checked)} />
                <Checkbox label="People Directory" checked={sourceFilters.includes("people")} onChange={(_e, checked) => toggleSource("people", !!checked)} />

                <Checkbox label="Marketing Calendar" checked={sourceFilters.includes("marketingcal")} onChange={(_e, checked) => toggleSource("marketingcal", !!checked)} />
                <Checkbox label="Company Calendar" checked={sourceFilters.includes("companycal")} onChange={(_e, checked) => toggleSource("companycal", !!checked)} />
                <Checkbox label="Locations" checked={sourceFilters.includes("locations")} onChange={(_e, checked) => toggleSource("locations", !!checked)} />
                <Checkbox label="Corporate Responsibilities" checked={sourceFilters.includes("corpresp")} onChange={(_e, checked) => toggleSource("corpresp", !!checked)} />
              </div>
            </div>

            <div style={cardStyle}>
              <div style={{ fontWeight: 600, marginBottom: 6 }}>File types (Resource Library)</div>
              <div style={{ fontSize: 12, color: "#605e5c", marginBottom: 8 }}>
                Only applies to documents in the Resource Library.
              </div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                {availableFileTypes.map((ext) => (
                  <Checkbox key={ext} label={ext.toUpperCase()} checked={fileTypeFilters.includes(ext)} onChange={(_e, checked) => toggleFileType(ext, !!checked)} />
                ))}
              </div>
            </div>
          </div>
        </div>

        {loading && (
          <div style={{ marginTop: 20 }}>
            <Spinner label="Searching…" />
          </div>
        )}

        {error && !loading && <p style={{ color: "red", marginTop: 10 }}>{error}</p>}

        {!loading && paginatedResults.length > 0 && (
          <div style={{ marginTop: 20 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
              <h3 style={{ margin: 0 }}>Results</h3>
              <div style={{ fontSize: 12, color: "#605e5c" }}>
                Showing <strong>{(page - 1) * pageSize + 1}–{Math.min(page * pageSize, results.length)}</strong> of <strong>{results.length}</strong>
              </div>
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "minmax(0, 1fr) minmax(0, 1fr)", gap: 12 }}>
              {paginatedResults.map((r, i) => {
                const displayTitle = r.Title || r.Filename || "(no title)";
                const onCardClick = () => openResult(r);

                if (r.Source === "People") {
                  const upnOrEmail = r.WorkEmail || r.SipAddress || r.AccountName || "";
                  return (
                    <div
                      key={(r.ListItemId || 0) + "-" + i + "-" + r.Source}
                      style={{ ...cardStyle, cursor: r.WorkEmail || r.SipAddress ? "pointer" : "default", borderLeft: "4px solid #0078d4" }}
                      onClick={onCardClick}
                    >
                      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}>
                        <span style={pillStyle}>{sourceLabels[r.Source]}</span>
                      </div>

                      <PeoplePersonaWithLive
                        context={context}
                        graphClient={graphClient}
                        displayName={r.PreferredName || displayTitle}
                        upnOrEmail={upnOrEmail}
                        jobTitle={r.JobTitle}
                        department={r.Department}
                      />

                      <div style={metaRowStyle}>
                        {r.WorkEmail && <span><strong>Email:</strong> {r.WorkEmail}</span>}
                        {r.OfficeNumber && <span><strong>Office:</strong> {r.OfficeNumber}</span>}
                        {r.CellPhone && <span><strong>Mobile:</strong> {r.CellPhone}</span>}
                      </div>
                    </div>
                  );
                }

                const isTooltipSource =
                  r.Source === "WeeklyWords" || r.Source === "LessonsLearned" || r.Source === "ITBlogs";

                const borderLeft =
                  r.Source === "WeeklyWords"
                    ? "4px solid #107c10"
                    : r.Source === "LessonsLearned"
                    ? "4px solid #ffb900"
                    : r.Source === "ITBlogs"
                    ? "4px solid #5c2d91"
                    : r.Source === "MarketingCalendar"
                    ? "4px solid #d83b01"
                    : r.Source === "CompanyCalendar"
                    ? "4px solid #038387"
                    : r.Source === "CorporateResponsibilities"
                    ? "4px solid #8764b8"
                    : r.Source === "Locations"
                    ? "4px solid #ca5010"
                    : "4px solid #605e5c";

                return (
                  <div
                    key={(r.ListItemId || 0) + "-" + i + "-" + r.Source}
                    style={{ ...cardStyle, cursor: r.Path ? "pointer" : "default", borderLeft }}
                    onClick={onCardClick}
                  >
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6, gap: 8 }}>
                      <span style={pillStyle}>{sourceLabels[r.Source]}</span>
                      {r.FileType && (
                        <span style={{ fontSize: 10, fontWeight: 600, textTransform: "uppercase", color: "#605e5c" }}>
                          {r.FileType}
                        </span>
                      )}
                    </div>

                    <div style={{ display: "flex", alignItems: "flex-start", gap: 6 }}>
                      <div
                        style={{ fontWeight: 600, fontSize: 14, lineHeight: 1.25, flex: 1, minWidth: 0 }}
                        dangerouslySetInnerHTML={{ __html: highlight(displayTitle, words) }}
                      />

                      {isTooltipSource && (
                        <TooltipHost
                          styles={{ root: { display: "inline-flex" } }}
                          content={
                            <div style={{ maxHeight: 250, overflowY: "auto" }}>
                              <div dangerouslySetInnerHTML={{ __html: getSafeHtmlForTooltip(r, words) }} />
                            </div>
                          }
                        >
                          <IconButton
                            iconProps={{ iconName: "Info" }}
                            title="Preview"
                            ariaLabel="Preview"
                            styles={{ root: { padding: 0, height: 20, width: 20, marginTop: 1 } }}
                            onClick={(e) => e.stopPropagation()}
                          />
                        </TooltipHost>
                      )}
                    </div>

                    {r.Source === "ResourceLibrary" && r.Filename && r.Filename !== r.Title && (
                      <div style={{ fontSize: 12, color: "#605e5c", marginTop: 4 }}>File: {r.Filename}</div>
                    )}

                    {(r.Source === "MarketingCalendar" || r.Source === "CompanyCalendar") && (
                      <div style={metaRowStyle}>
                        {r.EventDate && <span><strong>Start:</strong> {new Date(r.EventDate).toLocaleString()}</span>}
                        {r.EndDate && <span><strong>End:</strong> {new Date(r.EndDate).toLocaleString()}</span>}
                        {typeof r.AllDayEvent === "boolean" && <span><strong>All day:</strong> {r.AllDayEvent ? "Yes" : "No"}</span>}
                        {r.EventLocation && <span><strong>Location:</strong> {r.EventLocation}</span>}
                        {r.Category && <span><strong>Category:</strong> {r.Category}</span>}
                        {typeof r.fRecurrence === "boolean" && <span><strong>Recurs:</strong> {r.fRecurrence ? "Yes" : "No"}</span>}
                      </div>
                    )}

                    {r.Source === "Locations" && (
                      <div style={metaRowStyle}>
                        {r.ProjectLocation && <span><strong>Project Location:</strong> {r.ProjectLocation}</span>}
                        {r.Location && <span><strong>Location:</strong> {r.Location}</span>}
                        {r.Phone && <span><strong>Phone:</strong> {r.Phone}</span>}
                        {r.Fax && <span><strong>Fax:</strong> {r.Fax}</span>}
                      </div>
                    )}

                    {r.Source === "CorporateResponsibilities" && (
                      <div style={metaRowStyle}>
                        {r.Who && <span><strong>Who:</strong> {r.Who}</span>}
                      </div>
                    )}

                    {(r.Source === "WeeklyWords" || r.Source === "LessonsLearned" || r.Source === "ITBlogs") && (
                      <div style={metaRowStyle}>
                        {r.Author && <span><strong>Author:</strong> {r.Author}</span>}
                        {r.LastModifiedTime && <span><strong>Modified:</strong> {new Date(r.LastModifiedTime).toLocaleString()}</span>}
                      </div>
                    )}

                    {r.Source === "ResourceLibrary" && (
                      <div style={metaRowStyle}>
                        {r.Author && <span><strong>Author:</strong> {r.Author}</span>}
                        {r.LastModifiedTime && <span><strong>Modified:</strong> {new Date(r.LastModifiedTime).toLocaleString()}</span>}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>

            <div style={{ marginTop: 16, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <div>
                <PrimaryButton text="Previous" disabled={page == 1} onClick={() => setPage(page - 1)} style={{ marginRight: 8 }} />
                <PrimaryButton text="Next" disabled={page === totalPages} onClick={() => setPage(page + 1)} />
              </div>
              <div style={{ fontSize: 12, color: "#605e5c" }}>Page {page} of {totalPages || 1}</div>
            </div>
          </div>
        )}

        {!loading && results.length === 0 && query.trim() && !error && <p style={{ marginTop: 16 }}>No results found.</p>}
      </div>
    </div>
  );
};

export default SearchDocuments;
