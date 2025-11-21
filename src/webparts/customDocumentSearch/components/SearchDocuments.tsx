import * as React from "react";
import { useState, useEffect } from "react";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/search";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import {
  TextField,
  PrimaryButton,
  Dropdown,
  IDropdownOption,
  DatePicker,
} from "@fluentui/react";

interface ISearchDocumentsProps {
  context: any;
}

// Highlight keyword function
const highlight = (text: string, words: string[]) => {
  if (!text) return "";
  let newText = text;
  words.forEach((w) => {
    const reg = new RegExp(`(${w})`, "gi");
    newText = newText.replace(reg, `<mark style="background:yellow">$1</mark>`);
  });
  return newText;
};

const SearchDocuments: React.FC<ISearchDocumentsProps> = ({ context }) => {
  const sp = spfi("https://vaughnconstruction.sharepoint.com").using(SPFx(context));

  /** STATE **/
  const [query, setQuery] = useState("");
  const [results, setResults] = useState<any[]>([]);
  const [loading, setLoading] = useState(false);
  const [searchMode, setSearchMode] = useState("exact");
  const [fileType, setFileType] = useState<string | null>(null);
  const [author, setAuthor] = useState<string | null>(null);
  const [startDate, setStartDate] = useState<Date | undefined>();
  const [endDate, setEndDate] = useState<Date | undefined>();

  const [sortMode, setSortMode] = useState("relevance");

  /** Pagination **/
  const pageSize = 10;
  const [page, setPage] = useState(1);

  const searchModes: IDropdownOption[] = [
    { key: "exact", text: "Exact Phrase" },
    { key: "any", text: "Any Words" },
    { key: "all", text: "All Words" },
  ];

  const sortModes: IDropdownOption[] = [
    { key: "relevance", text: "Relevance" },
    { key: "modified", text: "Modified Date" },
    { key: "name", text: "File Name" },
  ];

  const fileFilters: IDropdownOption[] = [
    { key: '', text: "All File Types" },
    { key: "pdf", text: "PDF" },
    { key: "docx", text: "Word" },
    { key: "xlsx", text: "Excel" },
    { key: "pptx", text: "PowerPoint" },
  ];

  const libraryId = "8177736d-9faa-49cc-82b3-4ff5a93fa02e";

  /** ===== KQL Builder ===== **/
const buildKQL = (): string => {
  if (!query) return "";

  const words = query.trim().split(" ");

  let base = "";

  if (searchMode === "exact") {
    base = `Title:"${query}"`;
  }

  if (searchMode === "any") {
    base = words.map(w => `Title:${w}`).join(" OR ");
  }

  if (searchMode === "all") {
    base = words.map(w => `Title:${w}`).join(" AND ");
  }

  // Always limit search to your library
  let kql = `(${base}) AND ListId:${libraryId}`;

  // Apply filters
  if (fileType) kql += ` AND FileType:${fileType}`;
  if (author) kql += ` AND Author:${author}`;

  if (startDate)
    kql += ` AND LastModifiedTime>=${startDate.toISOString()}`;

  if (endDate)
    kql += ` AND LastModifiedTime<=${endDate.toISOString()}`;

  return kql;
};


  /** ===== Search Engine ===== **/
 const runSearch = async () => {
  if (!query) return;

  setLoading(true);
  setPage(1);

  const kql = buildKQL();

  const sortClause =
    sortMode === "modified"
      ? "LastModifiedTime:descending"
      : sortMode === "name"
      ? "FileName:ascending"
      : "Rank:descending";

  try {
    const res = await sp.search({
      Querytext: kql,
      RowLimit: 100,
      EnableQueryRules: false,
      SelectProperties: [
        "Title",
        "Path",
        "Author",
        "LastModifiedTime",
        "Size",
        "FileType"
      ],
      SortList: [
        {
          Property: sortClause.split(":")[0],
          Direction: sortClause.endsWith("descending") ? 1 : 0,
        },
      ],
    });

    const items = res.PrimarySearchResults.map((i: any) => ({
      Title: i.Title,
      Path: i.Path,
      Author: i.Author,
      Size: i.Size,
      FileType: i.FileType,
      Modified: i.LastModifiedTime,
    }));

    setResults(items);
  } catch (err) {
    console.error(err);
  }

  setLoading(false);
};


  /** ===== Pagination Logic ===== **/
  const paginatedResults = results.slice((page - 1) * pageSize, page * pageSize);
  const totalPages = Math.ceil(results.length / pageSize);

  const words = query.split(" ");

  return (
    <div style={{ padding: 20 }}>
      <h2>🔍 Advanced Document Search</h2>

      <Dropdown
        label="Search Mode"
        selectedKey={searchMode}
        onChange={(e, o:any) => setSearchMode(o.key.toString())}
        options={searchModes}
        styles={{ root: { width: 250 } }}
      />

      <TextField
        label="Search Text"
        value={query}
        onChange={(e, v:any) => setQuery(v)}
        styles={{ root: { width: 400 } }}
      />

      <Dropdown
        label="File Type"
        options={fileFilters}
        selectedKey={fileType}
        onChange={(e, o:any) => setFileType(o.key as any)}
        styles={{ root: { width: 250, marginTop: 10 } }}
      />

      <TextField
        label="Author (optional)"
        value={author || ""}
        onChange={(e, v:any) => setAuthor(v)}
        styles={{ root: { width: 250 } }}
      />

      <DatePicker
        label="Start Modified Date"
        value={startDate}
        onSelectDate={(date) => setStartDate(date ?? undefined)}
      />

      <DatePicker
        label="End Modified Date"
        value={endDate}
        onSelectDate={(date) => setEndDate(date ?? undefined)}
      />

      <Dropdown
        label="Sort By"
        selectedKey={sortMode}
        options={sortModes}
        onChange={(e, o:any) => setSortMode(o?.key?.toString())}
        styles={{ root: { width: 250, marginTop: 10 } }}
      />

      <PrimaryButton
        text="Search"
        onClick={runSearch}
        style={{ marginTop: 10 }}
      />

      {loading && <p>Searching...</p>}

      {!loading && paginatedResults.length > 0 && (
        <div style={{ marginTop: 20 }}>
          <h3>Results ({results.length})</h3>

          {paginatedResults.map((r, i) => (
            <div
              key={i}
              style={{
                borderBottom: "1px solid #ddd",
                padding: "10px 0",
              }}
            >
              <div
                dangerouslySetInnerHTML={{
                  __html: highlight(r.Title, words),
                }}
              />

              <a href={r.Path} target="_blank">Open Document</a>

              <div><strong>Author:</strong> {r.Author}</div>
              <div><strong>Modified:</strong> {r.Modified}</div>
              <div><strong>Type:</strong> {r.FileType}</div>

              {/* PDF/Office snippet */}
              {r.Snippet && (
                <div
                  dangerouslySetInnerHTML={{
                    __html: highlight(r.Snippet, words),
                  }}
                />
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
              Page {page} of {totalPages}
            </div>
          </div>
        </div>
      )}

      {!loading && results.length === 0 && query && (
        <p>No results found.</p>
      )}
    </div>
  );
};

export default SearchDocuments;
