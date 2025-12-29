// SearchDocuments.tsx
import * as React from "react";
import { useState, useMemo, useEffect } from "react";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/search";
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
  Persona,
  PersonaSize,
} from "@fluentui/react";
import { MSGraphClientV3 } from "@microsoft/sp-http";

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
const LESSONS_LEARNED_LIST_ID = "42F972A2-3D9D-4DBE-86C8-32EBAF4ACFDB";

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

  if (mode === "exact") {
    return `"${trimmed}"`;
  }

  if (mode === "all") {
    // default behavior = all words
    return trimmed;
  }

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

// Simple mapping for badges
const sourceLabels: Record<ISearchResult["Source"], string> = {
  ResourceLibrary: "Resource Library",
  WeeklyWords: "Weekly Words",
  LessonsLearned: "Lessons Learned",
  People: "People Directory",
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

  // Dashboard stats: count by source
  const sourceCounts = useMemo(() => {
    const map: Record<string, number> = {};
    results.forEach((r) => {
      map[r.Source] = (map[r.Source] || 0) + 1;
    });
    return map;
  }, [results]);

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
      const graphQueryString = buildGraphQueryString(trimmed, modeToUse);

      // Determine which sources are active (if user cleared all, treat as ALL)
      const activeSources: SourceKey[] =
        sourceFilters.length > 0 ? sourceFilters : ALL_SOURCES;

      // Graph is needed only for Weekly Words + Lessons Learned
      const needGraph = activeSources.some((s) =>
        ["weekly", "lessons"].includes(s)
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
          const dynamicTypes: string[] = Array.from(
            new Set(
              items
                .map((it) => it.FileType)
                .filter((x): x is string => !!x && x.trim().length > 0)
            )
          );
          if (dynamicTypes.length > 0) {
            setAvailableFileTypes(dynamicTypes);
          }

          return items;
        })();

        searchPromises.push(docsPromise);
      }

      // ============================================================
      // WEEKLY WORDS – Graph Search (listItem)
      // ➜ Returns Title + full VaughnContent/CanvasContent1/Abstract HTML
      // ============================================================
      if (activeSources.includes("weekly") && graphClient) {
        const weeklyPromise = (async (): Promise<ISearchResult[]> => {
          const body = {
            requests: [
              {
                entityTypes: ["listItem"],
                query: {
                  queryString: graphQueryString,
                },
                from: 0,
                size: 50,
                fields: [
                  "Title",
                  "Abstract",
                  "VaughnContent",
                  "CanvasContent1",
                  "ArticleDate",
                  "PublishingContact",
                  "webUrl",
                  "lastModifiedDateTime",
                  "createdBy",
                  "sharepointIds",
                ],
              },
            ],
          };

          const res: any = await graphClient.api("/search/query").post(body);

          const hits = res?.value?.[0]?.hitsContainers?.[0]?.hits ?? [];

          const items: ISearchResult[] = hits
            .map((h: any) => h.resource)
            // keep only items from Weekly Words list
            .filter(
              (r: any) =>
                r?.sharepointIds?.listId?.toLowerCase() ===
                WEEKLY_WORDS_LIST_ID.toLowerCase()
            )
            .map((r: any) => {
              const f = r.fields || {};

              const title =
                f.Title ||
                f.title ||
                f.FileLeafRef ||
                r.name ||
                "(no title)";

              // FULL BODY HTML
              const bodyHtml =
                f.VaughnContent || f.CanvasContent1 || f.Abstract || "";

              return {
                Title: title,
                Filename: "",
                Path: r.webUrl,
                OriginalPath: "",
                HitHighlightedSummary: bodyHtml,
                FileType: "item",
                ListItemId: parseInt(
                  r.sharepointIds?.listItemId || "0",
                  10
                ),
                LastModifiedTime: r.lastModifiedDateTime,
                Author:
                  f.PublishingContact ||
                  f.Author ||
                  r.createdBy?.user?.displayName ||
                  "",
                Source: "WeeklyWords",
              };
            });

          return items;
        })();

        searchPromises.push(weeklyPromise);
      }

      // ============================================================
      // LESSONS LEARNED – Graph Search (listItem)
      // ➜ Returns Title + full Body/CanvasContent1 HTML
      // ============================================================
      if (activeSources.includes("lessons") && graphClient) {
        const lessonsPromise = (async (): Promise<ISearchResult[]> => {
          const body = {
            requests: [
              {
                entityTypes: ["listItem"],
                query: {
                  queryString: graphQueryString,
                },
                from: 0,
                size: 50,
                fields: [
                  "Title",
                  "Body",
                  "Key_x0020_Words",
                  "Job_x0020__x0023_",
                  "Location",
                  "Project_x0020_Type",
                  "Construction_x002f_System_x0020_",
                  "CSI_x0020_Div",
                  "CanvasContent1",
                  "Contact",
                  "webUrl",
                  "lastModifiedDateTime",
                  "createdBy",
                  "sharepointIds",
                ],
              },
            ],
          };

          const res: any = await graphClient.api("/search/query").post(body);

          const hits = res?.value?.[0]?.hitsContainers?.[0]?.hits ?? [];

          const items: ISearchResult[] = hits
            .map((h: any) => h.resource)
            // keep only items from Lessons Learned list
            .filter(
              (r: any) =>
                r?.sharepointIds?.listId?.toLowerCase() ===
                LESSONS_LEARNED_LIST_ID.toLowerCase()
            )
            .map((r: any) => {
              const f = r.fields || {};

              const title =
                f.Title ||
                f.title ||
                f.FileLeafRef ||
                r.name ||
                "(no title)";

              // FULL BODY HTML
              const bodyHtml = f.Body || f.CanvasContent1 || "";

              return {
                Title: title,
                Filename: "",
                Path: r.webUrl,
                OriginalPath: "",
                HitHighlightedSummary: bodyHtml,
                FileType: "item",
                ListItemId: parseInt(
                  r.sharepointIds?.listItemId || "0",
                  10
                ),
                LastModifiedTime: r.lastModifiedDateTime,
                Author:
                  f.Contact ||
                  f.Author ||
                  r.createdBy?.user?.displayName ||
                  "",
                Source: "LessonsLearned",
              };
            });

          return items;
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

  // Shared card styles for dashboard feel
  const cardStyle: React.CSSProperties = {
    background: "#ffffff",
    borderRadius: 8,
    boxShadow: "0 1px 3px rgba(0,0,0,0.08)",
    padding: 16,
  };

  const pillStyle: React.CSSProperties = {
    display: "inline-block",
    padding: "2px 8px",
    borderRadius: 999,
    fontSize: 11,
    fontWeight: 600,
    textTransform: "uppercase",
    background: "#f3f2f1",
    color: "#605e5c",
  };

  // =======================
  // RENDER
  // =======================

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

  // 🔹 Full SearchWithResult variant (advanced UI)
  const allSourcesSelected = sourceFilters.length === ALL_SOURCES.length;

  return (
    <div style={{ padding: 20, background: "#faf9f8" }}>
      <div style={{ maxWidth: 1200, margin: "0 auto" }}>
        {/* Header */}
        <div
          style={{
            display: "flex",
            justifyContent: "space-between",
            alignItems: "flex-start",
            marginBottom: 16,
          }}
        >
          <div>
            <h2 style={{ margin: 0 }}>Advanced Search</h2>
            <p style={{ margin: "4px 0 0", color: "#605e5c" }}>
              Search across Resource Library, Weekly Words, Lessons Learned, and
              the People Directory in one place.
            </p>
          </div>

          {/* Summary chip */}
          {results.length > 0 && (
            <div style={cardStyle}>
              <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 6 }}>
                Result summary
              </div>
              <div style={{ fontSize: 12 }}>
                <div style={{ marginBottom: 4 }}>
                  Total:{" "}
                  <strong>
                    {results.length} item{results.length !== 1 ? "s" : ""}
                  </strong>
                </div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                  {Object.entries(sourceCounts).map(([key, count]) => (
                    <span key={key} style={pillStyle}>
                      {sourceLabels[key as ISearchResult["Source"]]}: {count}
                    </span>
                  ))}
                </div>
              </div>
            </div>
          )}
        </div>

        {/* Search controls area as dashboard cards */}
        <div
          style={{
            display: "grid",
            gridTemplateColumns: "minmax(0, 2fr) minmax(0, 1.4fr)",
            gap: 16,
            alignItems: "flex-start",
          }}
        >
          {/* Query + mode */}
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

              <ChoiceGroup
                label="Match mode"
                selectedKey={mode}
                options={modeOptions}
                onChange={(_e, o) => setMode(o?.key as Mode)}
              />

              <div style={{ display: "flex", justifyContent: "flex-end" }}>
                <PrimaryButton
                  text="Search"
                  onClick={() => runSearch()}
                  disabled={loading}
                />
              </div>
            </Stack>
          </div>

          {/* Filters column */}
          <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
            {/* Source filter */}
            <div style={cardStyle}>
              <div style={{ fontWeight: 600, marginBottom: 6 }}>Source</div>
              <div style={{ fontSize: 12, color: "#605e5c", marginBottom: 8 }}>
                Choose which sources to include in your search.
              </div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                <Checkbox
                  label="All"
                  checked={allSourcesSelected}
                  onChange={handleAllSourcesToggle}
                  styles={{ root: { marginRight: 0 } }}
                />
                <Checkbox
                  label="Weekly Words"
                  checked={sourceFilters.includes("weekly")}
                  onChange={(_e, checked) => toggleSource("weekly", !!checked)}
                  styles={{ root: { marginRight: 0 } }}
                />
                <Checkbox
                  label="Lessons Learned"
                  checked={sourceFilters.includes("lessons")}
                  onChange={(_e, checked) => toggleSource("lessons", !!checked)}
                  styles={{ root: { marginRight: 0 } }}
                />
                <Checkbox
                  label="Resource Library"
                  checked={sourceFilters.includes("resource")}
                  onChange={(_e, checked) => toggleSource("resource", !!checked)}
                  styles={{ root: { marginRight: 0 } }}
                />
                <Checkbox
                  label="People Directory"
                  checked={sourceFilters.includes("people")}
                  onChange={(_e, checked) => toggleSource("people", !!checked)}
                  styles={{ root: { marginRight: 0 } }}
                />
              </div>
            </div>

            {/* File type filters */}
            <div style={cardStyle}>
              <div style={{ fontWeight: 600, marginBottom: 6 }}>
                File types (Resource Library)
              </div>
              <div style={{ fontSize: 12, color: "#605e5c", marginBottom: 8 }}>
                Only applies to documents in the Resource Library.
              </div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                {availableFileTypes.map((ext) => (
                  <Checkbox
                    key={ext}
                    label={ext.toUpperCase()}
                    checked={fileTypeFilters.includes(ext)}
                    onChange={(_e, checked) => toggleFileType(ext, !!checked)}
                    styles={{ root: { marginRight: 0 } }}
                  />
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

        {error && !loading && (
          <p style={{ color: "red", marginTop: 10 }}>{error}</p>
        )}

        {/* RESULTS */}
        {!loading && paginatedResults.length > 0 && (
          <div style={{ marginTop: 20 }}>
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                marginBottom: 8,
              }}
            >
              <h3 style={{ margin: 0 }}>Results</h3>
              <div style={{ fontSize: 12, color: "#605e5c" }}>
                Showing{" "}
                <strong>
                  {(page - 1) * pageSize + 1}–
                  {Math.min(page * pageSize, results.length)}
                </strong>{" "}
                of <strong>{results.length}</strong>
              </div>
            </div>

            <div
              style={{
                display: "grid",
                gridTemplateColumns: "minmax(0, 1fr) minmax(0, 1fr)",
                gap: 16,
              }}
            >
              {paginatedResults.map((r, i) => {
                const displayTitle =
                  r.Title || r.Filename || "(no title)";

                const onCardClick = () => openResult(r);

                // PEOPLE CARD
                if (r.Source === "People") {
                  return (
                    <div
                      key={(r.ListItemId || 0) + "-" + i + "-" + r.Source}
                      style={{
                        ...cardStyle,
                        cursor:
                          r.WorkEmail || r.SipAddress ? "pointer" : "default",
                        borderLeft: "4px solid #0078d4",
                      }}
                      onClick={onCardClick}
                    >
                      <div
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          marginBottom: 8,
                        }}
                      >
                        <span style={pillStyle}>
                          {sourceLabels[r.Source] || r.Source}
                        </span>
                      </div>

                      <div
                        style={{
                          display: "flex",
                          gap: 12,
                          alignItems: "center",
                        }}
                      >
                        <Persona
                          text={r.PreferredName || displayTitle}
                          secondaryText={r.JobTitle}
                          tertiaryText={r.Department}
                          size={PersonaSize.size40}
                          imageInitials={(
                            (r.PreferredName || displayTitle)
                              .split(" ")
                              .map((p) => p[0])
                              .join("") || "?"
                          ).toUpperCase()}
                        />
                      </div>

                      <div style={{ marginTop: 10, fontSize: 13 }}>
                        {r.WorkEmail && (
                          <div>
                            <strong>Email:</strong> {r.WorkEmail}
                          </div>
                        )}
                        {r.CellPhone && (
                          <div>
                            <strong>Mobile:</strong> {r.CellPhone}
                          </div>
                        )}
                        {r.OfficeNumber && (
                          <div>
                            <strong>Office:</strong> {r.OfficeNumber}
                          </div>
                        )}
                        {r.SipAddress && (
                          <div>
                            <strong>Teams:</strong> {r.SipAddress}
                          </div>
                        )}
                      </div>

                      {r.HitHighlightedSummary && (
                        <div
                          style={{
                            marginTop: 10,
                            fontSize: 12,
                            color: "#605e5c",
                          }}
                          dangerouslySetInnerHTML={{
                            __html: highlight(r.HitHighlightedSummary, words),
                          }}
                        />
                      )}
                    </div>
                  );
                }

                // NON-PEOPLE CARD (docs, Weekly Words, Lessons Learned)
                return (
                  <div
                    key={(r.ListItemId || 0) + "-" + i + "-" + r.Source}
                    style={{
                      ...cardStyle,
                      cursor: r.Path ? "pointer" : "default",
                      borderLeft:
                        r.Source === "WeeklyWords"
                          ? "4px solid #107c10"
                          : r.Source === "LessonsLearned"
                          ? "4px solid #ffb900"
                          : "4px solid #605e5c",
                    }}
                    onClick={onCardClick}
                  >
                    {/* Source + meta header */}
                    <div
                      style={{
                        display: "flex",
                        justifyContent: "space-between",
                        alignItems: "center",
                        marginBottom: 8,
                      }}
                    >
                      <span style={pillStyle}>
                        {sourceLabels[r.Source] || r.Source}
                      </span>

                      {r.FileType && (
                        <span
                          style={{
                            fontSize: 11,
                            fontWeight: 600,
                            textTransform: "uppercase",
                            color: "#605e5c",
                          }}
                        >
                          {r.FileType}
                        </span>
                      )}
                    </div>

                    {/* Title */}
                    <div
                      style={{ fontWeight: 600, fontSize: 15, lineHeight: 1.3 }}
                      dangerouslySetInnerHTML={{
                        __html: highlight(displayTitle, words),
                      }}
                    />

                    {/* File name (if exists and different from title) */}
                    {r.Filename &&
                      r.Filename !== r.Title && (
                        <div
                          style={{
                            fontSize: 12,
                            color: "#605e5c",
                            marginTop: 2,
                          }}
                        >
                          File name: {r.Filename}
                        </div>
                      )}

                    {/* Path */}
                    {r.Path && (
                      <div
                        style={{
                          fontSize: 12,
                          color: "#605e5c",
                          marginTop: 4,
                          wordBreak: "break-all",
                        }}
                      >
                        {r.Path}
                      </div>
                    )}

                    {/* Snippet */}
                    {r.HitHighlightedSummary && (
                      <div
                        style={{
                          fontSize: 13,
                          marginTop: 8,
                          color: "#323130",
                          lineHeight: 1.4,
                        }}
                        dangerouslySetInnerHTML={{
                          __html:
                            r.Source === "WeeklyWords" ||
                            r.Source === "LessonsLearned"
                              ? // For Weekly/Lessons: use full HTML body as-is
                                r.HitHighlightedSummary
                              : // For docs: apply keyword highlighting
                                highlight(r.HitHighlightedSummary, words),
                        }}
                      />
                    )}

                    {/* Meta */}
                    <div
                      style={{
                        fontSize: 12,
                        color: "#605e5c",
                        marginTop: 8,
                      }}
                    >
                      {r.Author && <span>Author: {r.Author}</span>}
                      {r.Author && r.LastModifiedTime && <span> • </span>}
                      {r.LastModifiedTime && (
                        <span>
                          Modified:{" "}
                          {new Date(r.LastModifiedTime).toLocaleString()}
                        </span>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>

            {/* PAGINATION */}
            <div
              style={{
                marginTop: 20,
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
              }}
            >
              <div>
                <PrimaryButton
                  text="Previous"
                  disabled={page === 1}
                  onClick={() => setPage(page - 1)}
                  style={{ marginRight: 8 }}
                />
                <PrimaryButton
                  text="Next"
                  disabled={page === totalPages}
                  onClick={() => setPage(page + 1)}
                />
              </div>
              <div style={{ fontSize: 12, color: "#605e5c" }}>
                Page {page} of {totalPages || 1}
              </div>
            </div>
          </div>
        )}

        {!loading && results.length === 0 && query.trim() && !error && (
          <p style={{ marginTop: 16 }}>No results found.</p>
        )}
      </div>
    </div>
  );
};

export default SearchDocuments;
