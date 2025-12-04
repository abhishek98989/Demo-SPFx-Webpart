// SearchDocuments.tsx
import * as React from "react";
import { useState, useMemo, useEffect } from "react";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/search";
import { SearchQueryBuilder } from "@pnp/sp/search";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import {
  TextField,
  PrimaryButton,
  ChoiceGroup,
  IChoiceGroupOption,
  Checkbox,
  Stack,
  Spinner,
} from "@fluentui/react";

interface ISearchDocumentsProps {
  context: any;
  /** UI variant controlled from property pane via parent webpart */
  variant?: "SearchBar" | "SearchWithResult";
}

/** 🔧 CONFIG – change these if needed */
const SITE_URL = "https://vaughnconstruction.sharepoint.com";
const DOCUMENT_LIBRARY_ID = "8177736d-9faa-49cc-82b3-4ff5a93fa02e"; // Resource Library doc library Id

/** Other sources */
const WEEKLY_WORDS_LIST_ID = "216B6A79-4A9F-4E15-962E-23133411D96B";
const LESSONS_LEARNED_LIST_ID = "6991af0c-d310-4590-bbd2-2e25cd59eebd";

const WEEKLY_WORDS_SITE_URL = "https://vaughnconstruction.sharepoint.com/news";
const LESSONS_LEARNED_SITE_URL = "https://vaughnconstruction.sharepoint.com/ll";

// People directory result source
const PEOPLE_SOURCE_ID = "B09A7990-05EA-4AF9-81EF-EDFAB16C4E31";

/** Target page that will show full results */
const RESULTS_PAGE_URL =
  "https://vaughnconstruction.sharepoint.com/sitepages/Custom-Search-Results.aspx";

type Mode = "exact" | "all" | "any";

/** Searchable fields for content KQL */
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

// Highlight keyword function
const highlight = (text: string, words: string[]) => {
  if (!text) return "";
  let newText = text;
  words.forEach((w) => {
    if (!w) return;
    const reg = new RegExp(`(${w})`, "gi");
    newText = newText.replace(
      reg,
      `<mark style="background:yellow">$1</mark>`
    );
  });
  return newText;
};

type SourceKey = "resource" | "weekly" | "lessons" | "people";
const ALL_SOURCES: SourceKey[] = ["resource", "weekly", "lessons", "people"];

interface ISearchResult {
  Title: string;
  Filename?: string;
  Path?: string;
  OriginalPath?: string;
  HitHighlightedSummary?: string;
  FileType?: string;
  ListItemId?: number;
  LastModifiedTime?: string;
  Author?: string;

  /** Which source this came from */
  Source: "ResourceLibrary" | "WeeklyWords" | "LessonsLearned" | "People";

  /** People-specific fields (only for Source === "People") */
  PreferredName?: string;
  JobTitle?: string;
  Department?: string;
  AccountName?: string;
  SipAddress?: string;
  WorkEmail?: string;
  CellPhone?: string;
  OfficeNumber?: string;
}

const SearchDocuments: React.FC<ISearchDocumentsProps> = ({
  context,
  variant = "SearchWithResult",
}) => {
  const spRoot = useMemo(() => spfi(SITE_URL).using(SPFx(context)), [context]);

  /** STATE **/
  const [query, setQuery] = useState("");
  const [mode, setMode] = useState<Mode>("exact");
  const [results, setResults] = useState<ISearchResult[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string>("");

  // Result-type style file type filter (simplified) – applies only to doc library
  const [fileTypeFilters, setFileTypeFilters] = useState<string[]>([]);
  const [availableFileTypes, setAvailableFileTypes] = useState<string[]>([
    "docx",
    "xlsx",
    "pptx",
    "pdf",
    "msg",
    "eml",
  ]);

  /** NEW: Source filters */
  const [sourceFilters, setSourceFilters] = useState<SourceKey[]>(ALL_SOURCES);

  /** Pagination **/
  const pageSize = 10;
  const [page, setPage] = useState(1);
  const totalPages = Math.ceil(results.length / pageSize);
  const paginatedResults = results.slice(
    (page - 1) * pageSize,
    page * pageSize
  );

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

  const handleAllSourcesToggle = (_e: any, checked?: boolean) => {
    setPage(1);
    if (checked) {
      // Select all sources
      setSourceFilters(ALL_SOURCES);
    } else {
      // When unchecking All, don't clear others – user can uncheck individually
    }
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
      window.open(url, "_blank"); // Open in new tab
      return;
    }

    setLoading(true);
    setError("");
    setPage(1);

    try {
      const kqlBody = buildKql(trimmed, modeToUse, fields);

      // Determine which sources are active (if user cleared all, treat as ALL)
      const activeSources: SourceKey[] =
        sourceFilters.length > 0 ? sourceFilters : ALL_SOURCES;

      const searchPromises: Promise<ISearchResult[]>[] = [];

      /** Helper: docs from Resource Library (document library only) */
      if (activeSources.includes("resource")) {
        const docsPromise = (async (): Promise<ISearchResult[]> => {
          let filter = `(${kqlBody}) AND (contentclass:STS_ListItem_DocumentLibrary AND IsDocument:"true" AND ListId:"${DOCUMENT_LIBRARY_ID}")`;

          // FileType refiners (only for doc library)
          if (fileTypeFilters.length > 0) {
            const ftClause = fileTypeFilters
              .map((ft) => `FileType:${ft}`)
              .join(" OR ");
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
              "HitHighlightedSummary",
              "FileType",
              "Filename",
              "ListItemId",
              "LastModifiedTime",
              "Author",
            ],
            EnableQueryRules: true,
            SortList: [
              {
                Property: "Rank",
                Direction: 1, // Desc
              },
            ],
          });

          const items: ISearchResult[] = (res.PrimarySearchResults || []).map(
            (i: any) => ({
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
            })
          );

          // Build available file types dynamically *from doc results only*
          const dynamicTypes :any= Array.from(
            new Set(
              items
                .map((it) => it.FileType)
                .filter((x) => x && x.trim().length > 0)
            )
          );
          if (dynamicTypes.length > 0) {
            setAvailableFileTypes(dynamicTypes);
          }

          return items;
        })();

        searchPromises.push(docsPromise);
      }

  if (activeSources.includes("weekly")) {
  const weeklyPromise = (async (): Promise<ISearchResult[]> => {
    const spWeekly = spfi(WEEKLY_WORDS_SITE_URL).using(SPFx(context));

    // Search across Weekly Words page fields
    const weeklyFields = [
      "Title",
      "ArticleDate",
      "VaughnContent",
      "PublishingContact",
      "CanvasContent1", // Modern page content
    ];
    
    // Build KQL specifically for Weekly Words pages
    const weeklyKql = buildKql(trimmed, modeToUse, weeklyFields);
    
    // FIX: Use STS_ListItem_DocumentLibrary for Site Pages library + ListId for safety
    const listFilter = `(${weeklyKql}) AND (contentclass:STS_ListItem_DocumentLibrary AND IsDocument:"true" AND ListId:"${WEEKLY_WORDS_LIST_ID}")`;

    const res: any = await spWeekly.search({
      Querytext: listFilter,
      RowLimit: 500,
      TrimDuplicates: false,
      SelectProperties: [
        "Title",
        "Path",
        "OriginalPath",
        "HitHighlightedSummary",
        "ListItemId",
        "LastModifiedTime",
        "Author",
        "ArticleDate",
        "VaughnContent",
        "PublishingContact",
        "FileType",
        "Filename",
        "CanvasContent1",
      ],
      EnableQueryRules: true,
      SortList: [
        {
          Property: "Rank",
          Direction: 1, // Desc
        },
      ],
    });

    const items: ISearchResult[] = (res.PrimarySearchResults || []).map(
      (i: any) => ({
        Title: i.Title || "",
        Filename: i.Filename || "",
        Path: i.Path,
        OriginalPath: i.OriginalPath,
        HitHighlightedSummary: i.HitHighlightedSummary,
        FileType: (i.FileType || "aspx").toLowerCase(),
        ListItemId: Number(i.ListItemId),
        LastModifiedTime: i.LastModifiedTime,
        Author: i.Author,
        Source: "WeeklyWords",
      })
    );

    return items;
  })();

  searchPromises.push(weeklyPromise);
}


// ============================================================
// PATCH 2: Replace Lessons Learned Search Section
// FOR SITE PAGES (stored in Site Pages document library)
// ============================================================

/** Helper: Lessons Learned (site-scoped search) */
if (activeSources.includes("lessons")) {
  const lessonsPromise = (async (): Promise<ISearchResult[]> => {
    const spLessons = spfi(LESSONS_LEARNED_SITE_URL).using(SPFx(context));

    // Search across Lessons Learned page fields
    const lessonsFields = [
      "Title",
      "Body",
      "Contact",
      "Key_x0020_Words",
      "Job_x0020__x0023_",
      "Location",
      "Project_x0020_Type",
      "Construction_x002f_System_x0020_",
      "CSI_x0020_Div",
      "CanvasContent1", // Modern page content
    ];
    
    // Build KQL specifically for Lessons Learned pages
    const lessonsKql = buildKql(trimmed, modeToUse, lessonsFields);
    
    // FIX: Use STS_ListItem_DocumentLibrary for Site Pages library + ListId for safety
    const listFilter = `(${lessonsKql}) AND (contentclass:STS_ListItem_DocumentLibrary AND IsDocument:"true" AND ListId:"${LESSONS_LEARNED_LIST_ID}")`;

    const res: any = await spLessons.search({
      Querytext: listFilter,
      RowLimit: 500,
      TrimDuplicates: false,
      SelectProperties: [
        "Title",
        "Path",
        "OriginalPath",
        "HitHighlightedSummary",
        "ListItemId",
        "LastModifiedTime",
        "Author",
        "PublishedDate",
        "Body",
        "Contact",
        "Key_x0020_Words",
        "Job_x0020__x0023_",
        "Location",
        "Project_x0020_Type",
        "Construction_x002f_System_x0020_",
        "CSI_x0020_Div",
        "FileType",
        "Filename",
        "CanvasContent1",
      ],
      EnableQueryRules: true,
      SortList: [
        {
          Property: "Rank",
          Direction: 1, // Desc
        },
      ],
    });

    const items: ISearchResult[] = (res.PrimarySearchResults || []).map(
      (i: any) => ({
        Title: i.Title || "",
        Filename: i.Filename || "",
        Path: i.Path,
        OriginalPath: i.OriginalPath,
        HitHighlightedSummary: i.HitHighlightedSummary,
        FileType: (i.FileType || "aspx").toLowerCase(),
        ListItemId: Number(i.ListItemId),
        LastModifiedTime: i.LastModifiedTime,
        Author: i.Author,
        Source: "LessonsLearned",
      })
    );

    return items;
  })();

  searchPromises.push(lessonsPromise);
}

      /** Helper: People Directory (using people search source on root) */
      if (activeSources.includes("people")) {
        const peoplePromise = (async (): Promise<ISearchResult[]> => {
          // Build a plain SearchQuery object (not the builder) so spRoot.search accepts it
          const peopleQuery = {
            Querytext: trimmed,
            SourceId: PEOPLE_SOURCE_ID,
            RowLimit: 500,
            RowsPerPage: 500,
            TrimDuplicates: true,
            SelectProperties: [
              "SipAddress",
              "PreferredName",
              "AccountName",
              "Department",
              "JobTitle",
              "WorkEmail",
              "CellPhone",
              "OfficeNumber",
            ],
          };

          const res: any = await spRoot.search(peopleQuery);

          const items: ISearchResult[] = (res.PrimarySearchResults || []).map(
            (i: any) => {
              const preferredName = i.PreferredName || "";
              const dept = i.Department || "";
              const job = i.JobTitle || "";
              const summaryParts = [];
              if (job) summaryParts.push(job);
              if (dept) summaryParts.push(dept);

              return {
                Title: preferredName,
                Filename: "",
                Path: i.WorkEmail || i.SipAddress || "",
                OriginalPath: "",
                HitHighlightedSummary: summaryParts.join(" • "),
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
              };
            }
          );

          return items;
        })();

        searchPromises.push(peoplePromise);
      }

      // Run all active source searches in parallel
      const resultArrays = await Promise.all(searchPromises);
      const merged: ISearchResult[] = ([] as ISearchResult[]).concat(
        ...resultArrays
      );

      // 🔥 TITLE-FIRST SORTING (client-side) across all sources
      const lowerQuery = trimmed.toLowerCase();
      const queryWords = lowerQuery.split(/\s+/).filter(Boolean);

      const scored = merged
        .map((it, index) => {
          const title = (it.Title || "").toLowerCase();
          // secondary text for scoring: filename or dept/job
          const secondary = (
            it.Filename ||
            it.Department ||
            it.JobTitle ||
            ""
          ).toLowerCase();

          let score = 0;

          if (modeToUse === "exact") {
            if (title.includes(lowerQuery)) score += 100;
            if (secondary.includes(lowerQuery)) score += 50;
          } else {
            let titleMatches = 0;
            let secondaryMatches = 0;
            queryWords.forEach((w) => {
              if (w && title.includes(w)) titleMatches++;
              if (w && secondary.includes(w)) secondaryMatches++;
            });

            if (modeToUse === "all") {
              score += titleMatches * 10 + secondaryMatches * 5;
              if (titleMatches === queryWords.length && queryWords.length > 0) {
                score += 40; // bonus for full title match
              }
            } else {
              // any
              score += titleMatches * 5 + secondaryMatches * 2;
            }
          }

          return { it, score, index };
        })
        .sort((a, b) => {
          if (b.score !== a.score) return b.score - a.score;
          return a.index - b.index;
        })
        .map((x) => x.it);

      setResults(scored);
    } catch (e: any) {
      console.error(e);
      setError(e?.message || "Search failed.");
    }

    setLoading(false);
  };

  /** Auto-run exact search on results page if ?q=... is present */
  useEffect(() => {
    if (variant !== "SearchWithResult") return;

    try {
      const params = new URLSearchParams(window.location.search);
      const qParam = params.get("q");
      if (qParam && qParam.trim()) {
        setQuery(qParam);
        setMode("exact");
        // run with override text + mode for exact match
        runSearch(qParam, "exact");
      }
    } catch {
      // ignore URL parsing errors
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [variant]);

  const openResult = (item: ISearchResult) => {
    // Handle People results specially – open mailto if available
    if (item.Source === "People") {
      if (item.WorkEmail) {
        window.location.href = `mailto:${item.WorkEmail}`;
      } else if (item.SipAddress) {
        window.open(`sip:${item.SipAddress}`, "_blank");
      }
      return;
    }

    const basePath = item.OriginalPath || item.Path;
    if (!basePath) return;

    // For pages (aspx) just open the page; for docs, use ?web=1
    const isPage =
      (item.FileType || "").toLowerCase() === "aspx" ||
      (basePath || "").toLowerCase().includes(".aspx");

    const url = isPage ? basePath : `${basePath}?web=1`;
    window.open(url, "_blank");
  };

  // =======================
  // RENDER
  // =======================

  // 🔹 Simple SearchBar variant: only a single textbox + button,
  // and Enter key will redirect to results page.
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

  // 🔹 Full SearchWithResult variant (advanced UI)
  const allSourcesSelected = sourceFilters.length === ALL_SOURCES.length;

  return (
    <div style={{ padding: 20 }}>
      <h2>Advanced Search</h2>

      {/* Query + Mode */}
      <Stack tokens={{ childrenGap: 8 }} styles={{ root: { maxWidth: 600 } }}>
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

        <ChoiceGroup
          label="Match mode"
          selectedKey={mode}
          options={modeOptions}
          onChange={(_e, o) => setMode(o?.key as Mode)}
        />
      </Stack>

      {/* File type filters */}
      <div style={{ marginTop: 10 }}>
        <span style={{ fontWeight: 600 }}>File types (Resource Library):</span>
        <div style={{ display: "flex", flexWrap: "wrap", marginTop: 4 }}>
          {availableFileTypes.map((ext) => (
            <Checkbox
              key={ext}
              label={ext.toUpperCase()}
              checked={fileTypeFilters.includes(ext)}
              onChange={(_e, checked) => toggleFileType(ext, !!checked)}
              styles={{ root: { marginRight: 8 } }}
            />
          ))}
        </div>
      </div>

      {/* Source filter */}
      <div style={{ marginTop: 10 }}>
        <span style={{ fontWeight: 600 }}>Source:</span>
        <div style={{ display: "flex", flexWrap: "wrap", marginTop: 4 }}>
          <Checkbox
            label="All"
            checked={allSourcesSelected}
            onChange={handleAllSourcesToggle}
            styles={{ root: { marginRight: 12 } }}
          />
          <Checkbox
            label="Weekly Words"
            checked={sourceFilters.includes("weekly")}
            onChange={(_e, checked) => toggleSource("weekly", !!checked)}
            styles={{ root: { marginRight: 12 } }}
          />
          <Checkbox
            label="Lessons Learned"
            checked={sourceFilters.includes("lessons")}
            onChange={(_e, checked) => toggleSource("lessons", !!checked)}
            styles={{ root: { marginRight: 12 } }}
          />
          <Checkbox
            label="Resource Library"
            checked={sourceFilters.includes("resource")}
            onChange={(_e, checked) => toggleSource("resource", !!checked)}
            styles={{ root: { marginRight: 12 } }}
          />
          <Checkbox
            label="People Directory"
            checked={sourceFilters.includes("people")}
            onChange={(_e, checked) => toggleSource("people", !!checked)}
            styles={{ root: { marginRight: 12 } }}
          />
        </div>
      </div>

      <PrimaryButton
        text="Search"
        onClick={() => runSearch()}
        style={{ marginTop: 12 }}
      />

      {loading && (
        <div style={{ marginTop: 20 }}>
          <Spinner label="Searching…" />
        </div>
      )}

      {error && !loading && (
        <p style={{ color: "red", marginTop: 10 }}>{error}</p>
      )}

      {/* RESULTS */}
      {!loading && paginatedResults.length > 0 && (
        <div style={{ marginTop: 20 }}>
          <h3>Results ({results.length})</h3>

          {paginatedResults.map((r, i) => (
            <div
              key={(r.ListItemId || 0) + "-" + i + "-" + r.Source}
              style={{
                borderBottom: "1px solid #ddd",
                padding: "10px 0",
                cursor: "pointer",
              }}
              onClick={() => openResult(r)}
            >
              {/* Source badge */}
              <div
                style={{
                  fontSize: 11,
                  color: "#444",
                  marginBottom: 2,
                  textTransform: "uppercase",
                }}
              >
                [
                {r.Source === "ResourceLibrary"
                  ? "Resource Library"
                  : r.Source === "WeeklyWords"
                  ? "Weekly Words"
                  : r.Source === "LessonsLearned"
                  ? "Lessons Learned"
                  : "People Directory"}
                ]
              </div>

              {/* Title */}
              <div
                style={{ fontWeight: 600, fontSize: 16 }}
                dangerouslySetInnerHTML={{
                  __html: highlight(
                    r.Title || r.Filename || "(no title)",
                    words
                  ),
                }}
              />

              {/* People subtitle */}
              {r.Source === "People" && (
                <div style={{ fontSize: 13, color: "#555", marginTop: 2 }}>
                  {r.JobTitle && <span>{r.JobTitle}</span>}
                  {r.JobTitle && r.Department && <span> • </span>}
                  {r.Department && <span>{r.Department}</span>}
                  {r.WorkEmail && (
                    <>
                      {" "}
                      • <span>{r.WorkEmail}</span>
                    </>
                  )}
                </div>
              )}

              {/* File name (non-people) */}
              {r.Filename && r.Source !== "People" && (
                <div style={{ fontSize: 12, color: "#555", marginTop: 2 }}>
                  File name: {r.Filename}
                </div>
              )}

              {/* Path + type (non-people) */}
              {r.Source !== "People" && r.Path && (
                <div style={{ fontSize: 12, color: "#666", marginTop: 2 }}>
                  {r.FileType && (
                    <span
                      style={{
                        fontWeight: 600,
                        marginRight: 6,
                        textTransform: "uppercase",
                      }}
                    >
                      [{r.FileType}]
                    </span>
                  )}
                  <span>{r.Path}</span>
                </div>
              )}

              {/* Snippet */}
              {r.HitHighlightedSummary && (
                <div
                  style={{
                    fontSize: 13,
                    marginTop: 4,
                    color: "#333",
                    lineHeight: 1.4,
                  }}
                  dangerouslySetInnerHTML={{
                    __html: r.HitHighlightedSummary,
                  }}
                />
              )}

              {/* Meta */}
              {r.Source !== "People" && (
                <div style={{ fontSize: 12, color: "#777", marginTop: 4 }}>
                  {r.Author && <span>Author: {r.Author} • </span>}
                  {r.LastModifiedTime && (
                    <span>
                      Modified: {new Date(r.LastModifiedTime).toLocaleString()}
                    </span>
                  )}
                </div>
              )}
            </div>
          ))}

          {/* PAGINATION */}
          <div style={{ marginTop: 20 }}>
            <PrimaryButton
              text="Previous"
              disabled={page === 1}
              onClick={() => setPage(page - 1)}
              style={{ marginRight: 10 }}
            />
            <PrimaryButton
              text="Next"
              disabled={page === totalPages}
              onClick={() => setPage(page + 1)}
            />
            <div style={{ marginTop: 5 }}>
              Page {page} of {totalPages || 1}
            </div>
          </div>
        </div>
      )}

      {!loading && results.length === 0 && query.trim() && !error && (
        <p style={{ marginTop: 16 }}>No results found.</p>
      )}
    </div>
  );
};

export default SearchDocuments;
