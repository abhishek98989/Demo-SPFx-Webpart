import * as React from "react";
import { useState } from "react";
import { spfi, SPFx } from "@pnp/sp";
import { SearchQueryBuilder, SearchQueryInit } from "@pnp/sp/search";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import { TextField, PrimaryButton } from "@fluentui/react";

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
 
  /** STATE **/
  const [query, setQuery] = useState("");
  const [results, setResults] = useState<any[]>([]);
  const [loading, setLoading] = useState(false);
  

  /** Pagination **/
  const pageSize = 10;
  const [page, setPage] = useState(1);

  /** ===== Search Engine ===== **/
 const runSearch = async () => {
  const trimmedQuery = query.trim();
  if (!trimmedQuery) {
    setResults([]);
    return;
  }

  setLoading(true);
  setPage(1);
 const sp = spfi("https://vaughnconstruction.sharepoint.com").using(SPFx(context));
  const peopleSourceId = "B09A7990-05EA-4AF9-81EF-EDFAB16C4E31";
  try {
    const peopleQuery: SearchQueryInit = SearchQueryBuilder(trimmedQuery)
      .sourceId(peopleSourceId)
      .rowLimit(500)
      .rowsPerPage(500)
      .selectProperties(
        "SipAddress",
        "PreferredName",
        "AccountName",
        "Department",
        "JobTitle",
        "WorkEmail",
        "CellPhone",
        "OfficeNumber"
      )
      .trimDuplicates;

    const res = await sp.search(peopleQuery);

    const items = res.PrimarySearchResults.map((i: any) => ({
      PreferredName: i.PreferredName,
      JobTitle: i.JobTitle,
      Department: i.Department,
      AccountName: i.AccountName,
      SipAddress: i.SipAddress,
      WorkEmail: i.WorkEmail,
      CellPhone: i.CellPhone,
      OfficeNumber: i.OfficeNumber,
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

  const words = query.trim() ? query.trim().split(/\s+/) : [];

  return (
    <div style={{ padding: 20 }}>
      <h2>👥 People Search</h2>

      <TextField
        label="Enter a name or keyword"
        value={query}
        onChange={(e, v: any) => setQuery(v || "")}
        styles={{ root: { width: 400 } }}
      />

      <PrimaryButton
        text="Search People"
        onClick={runSearch}
        style={{ marginTop: 10 }}
      />

      {loading && <p>Searching directory...</p>}

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
                style={{ fontWeight: 600, fontSize: 16 }}
                dangerouslySetInnerHTML={{
                  __html: highlight(r.PreferredName || r.AccountName, words),
                }}
              />

              {r.JobTitle && (
                <div>
                  <strong>Title:</strong> {r.JobTitle}
                </div>
              )}

              {r.Department && (
                <div>
                  <strong>Department:</strong> {r.Department}
                </div>
              )}

              {r.WorkEmail && (
                <div>
                  <strong>Email:</strong>{" "}
                  <a href={`mailto:${r.WorkEmail}`}>{r.WorkEmail}</a>
                </div>
              )}

              {r.SipAddress && (
                <div>
                  <strong>SIP:</strong> {r.SipAddress}
                </div>
              )}

              {r.CellPhone && (
                <div>
                  <strong>Cell:</strong> {r.CellPhone}
                </div>
              )}

              {r.OfficeNumber && (
                <div>
                  <strong>Office:</strong> {r.OfficeNumber}
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
              Page {page} of {totalPages}
            </div>
          </div>
        </div>
      )}

      {!loading && results.length === 0 && query.trim() && (
        <p>No people found.</p>
      )}
    </div>
  );
};

export default SearchDocuments;
