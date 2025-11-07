import * as React from "react";
import "./style.css";
import { useEffect, useState, useRef, useCallback } from "react";
import {
  TextField,
  Stack,
  IconButton,
  Spinner,
  Text,
  TooltipHost,
  PrimaryButton,
} from "@fluentui/react";
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
} from "@fluentui/react/lib/DetailsList";
import { spfi, SPFx } from "@pnp/sp/presets/all";
import DOMPurify from "dompurify";
import "../../../globalCommon/style.css";
import "@pnp/sp/security";
import { PermissionKind } from "@pnp/sp/security";
interface IBlogPost {
  Id: number;
  Title: string;
  Author: string;
  PublishedDate: string;
  PublishedYear: number;
  PublishedDateAndMonth: string;
  PublishedTime: string;
  Url: string;
  Body: string;
  // FIX: contact is multi-lookup
  Contact?: { Id: number; Title: string }[];
  CSIDivisions?: { Id: number; Title: string }[];
}

interface ICSIDivision {
  Id: number;
  Title: string;
}

let siteUrl: any = "";
let ListIds: any = {};

const BlogsTable = (props: any) => {
  const [allPosts, setAllPosts] = useState<IBlogPost[]>([]);
  const [visiblePosts, setVisiblePosts] = useState<IBlogPost[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [canEdit, setCanEdit] = useState<boolean>(false);
  const [searchQuery, setSearchQuery] = useState<string>("");
  const [searchResults, setSearchResults] = useState<number[]>([]);
  const [currentResultIndex, setCurrentResultIndex] = useState<number>(0);

  // pagination JUST for cards
  const [page, setPage] = useState<number>(1);
  const postsPerPage = 6;

  // CSI Divisions
  const [csiDivisions, setCsiDivisions] = useState<ICSIDivision[]>([]);
  const [selectedCsiDivisions, setSelectedCsiDivisions] = useState<
    ICSIDivision[]
  >([]);

  // view toggle
  const [activeView, setActiveView] = useState<"cards" | "table">("table");

  // table view
  const [tableItems, setTableItems] = useState<IBlogPost[]>([]);
  const [columns, setColumns] = useState<IColumn[]>([]);
  const [sortConfig, setSortConfig] = useState<{
    key: string;
    isSortedDescending: boolean;
  }>({
    key: "PublishedDate",
    isSortedDescending: true,
  });

  const observerRef = useRef<IntersectionObserver | null>(null);
  const loaderRef = useRef<HTMLDivElement | null>(null);
  const contentRef = useRef<HTMLDivElement | null>(null);

  const sortByDate = (array: IBlogPost[]) => {
    return [...array].sort(
      (a, b) =>
        new Date(b?.PublishedDate).getTime() -
        new Date(a?.PublishedDate).getTime()
    );
  };

  // fetch posts (CSI-based)
  const fetchPosts = async (sp: any) => {
    try {
        const list = sp.web.lists.getById(ListIds?.posts);
                 try{
                         const perms = await list.getCurrentUserEffectivePermissions();
                        if (sp.web.hasPermissions(perms, PermissionKind.EditListItems)) {
                            setCanEdit(true);
                        }
                 }catch(err){
console.error("Error checking permissions:", err);
                 }
      const items = await sp.web.lists
        .getById(ListIds?.posts)
        .items.select(
          "Id",
          "Title",
          "Author/Title",
          "Created",
          "Body",
          "EncodedAbsUrl",
          "PublishedDate",
          "Contact/Title",
          "Contact/Id",
          "CSI_x0020_Div/Id",
          "CSI_x0020_Div/Title"
        )
        .expand("Author,Contact,CSI_x0020_Div")();

      const mapped: IBlogPost[] = items.map((item: any) => {
        const dateObj = item?.PublishedDate
          ? new Date(item?.PublishedDate)
          : new Date(item?.Created);
        const year = dateObj.getFullYear();
        const monthDate = `${dateObj.getMonth() + 1}/${dateObj.getDate()}`;
        let hours = dateObj.getHours();
        const minutes = dateObj.getMinutes().toString().padStart(2, "0");
        const ampm = hours >= 12 ? "PM" : "AM";
        hours = hours % 12 || 12;
        const time = `${hours}:${minutes} ${ampm}`;

        // FIX: normalize multi-lookup Contact
        let contacts: { Id: number; Title: string }[] = [];
        if (item.Contact) {
          if (Array.isArray(item.Contact)) {
            contacts = item.Contact.map((c: any) => ({
              Id: c.Id,
              Title: c.Title,
            }));
          } else {
            contacts = [{ Id: item.Contact.Id, Title: item.Contact.Title }];
          }
        }

        return {
          Id: item.Id,
          Title: item.Title || "",
          Author: item.Author?.Title || "",
          PublishedDate: item.PublishedDate || item.Created,
          PublishedYear: year,
          PublishedDateAndMonth: monthDate,
          PublishedTime: time,
          Url: item.EncodedAbsUrl || "",
          Body: item.Body || "",
          Contact: contacts,
          CSIDivisions: item["CSI_x0020_Div"]
            ? Array.isArray(item["CSI_x0020_Div"])
              ? item["CSI_x0020_Div"]
              : [item["CSI_x0020_Div"]]
            : [],
        };
      });

      const sorted = sortByDate(mapped);
      setAllPosts(sorted);
      setLoading(false);
    } catch (err) {
      console.error("Error fetching posts:", err);
      setLoading(false);
    }
  };

  // fetch CSI divisions
  const fetchCsiDivisions = async (sp: any) => {
    try {
      const items = await sp.web.lists
        .getById("90df1643-9bbd-4dfa-bc72-327e6263048f")
        .items.select("Id", "Title")();
      setCsiDivisions(items);
    } catch (err) {
      console.error("Error fetching CSI divisions:", err);
    }
  };

  useEffect(() => {
    siteUrl = props?.Context?.pageContext?.web?.absoluteUrl;
    ListIds = {
      posts: props?.postsListId,
    };
    const sp = spfi().using(SPFx(props.Context));
    fetchCsiDivisions(sp);
    fetchPosts(sp);
  }, [props.Context]);

  // CSI selection toggle
  const handleCsiDivisionSelect = (division: ICSIDivision) => {
    setSelectedCsiDivisions((prev) => {
      const exists = prev.some((d) => d.Id === division.Id);
      return exists
        ? prev.filter((d) => d.Id !== division.Id)
        : [...prev, division];
    });
    setPage(1);
  };

  const removeCsiDivision = (id: number) => {
    setSelectedCsiDivisions((prev) => prev.filter((d) => d.Id !== id));
    setPage(1);
  };

  const clearAllCsiDivisions = () => {
    setSelectedCsiDivisions([]);
    setPage(1);
  };

  const isCsiSelected = (id: number) =>
    selectedCsiDivisions.some((d) => d.Id === id);

  // filtered by CSI only
  const getFilteredPostsByCsi = useCallback(() => {
    let filtered = [...allPosts];
    if (selectedCsiDivisions.length > 0) {
      const selectedIds = selectedCsiDivisions.map((d) => d.Id);
      filtered = filtered.filter((p) =>
        (p.CSIDivisions || []).some((d) => selectedIds.includes(d.Id))
      );
    }
    return sortByDate(filtered);
  }, [allPosts, selectedCsiDivisions]);

  // global filter (search + CSI)
  const getFilteredPosts = useCallback(() => {
    let filtered = getFilteredPostsByCsi();

    if (searchQuery.trim()) {
      filtered = filtered.filter((post) => {
        const tempDiv = document.createElement("div");
        tempDiv.innerHTML = DOMPurify.sanitize(post.Body);
        const textContent = tempDiv.textContent || tempDiv.innerText || "";

        const csiTitles =
          post.CSIDivisions?.map((c: any) => c.Title || "").join(" ") || "";
        const contactTitle =
          post.Contact && post.Contact.length > 0
            ? post.Contact.map((c) => c.Title).join(" ")
            : "";

        const contentToSearch = `${post.Title} ${post.Author} ${
          post.PublishedDate
        } ${post.PublishedYear} ${post.PublishedDateAndMonth} ${
          post.PublishedTime
        } ${csiTitles} ${contactTitle} ${textContent}`.toLowerCase();

        return contentToSearch.includes(searchQuery.toLowerCase());
      });
    }

    return sortByDate(filtered);
  }, [getFilteredPostsByCsi, searchQuery]);

  // update search result indices (over ALL filtered, not just current page)
  const updateSearchResultIndices = useCallback(() => {
    if (!searchQuery.trim()) {
      setSearchResults([]);
      setCurrentResultIndex(0);
      return;
    }

    const filteredByCsi = getFilteredPostsByCsi();
    const matches: number[] = [];

    filteredByCsi.forEach((post, index) => {
      const tempDiv = document.createElement("div");
      tempDiv.innerHTML = DOMPurify.sanitize(post.Body);
      const textContent = tempDiv.textContent || tempDiv.innerText || "";

      const csiTitles =
        post.CSIDivisions?.map((c: any) => c.Title || "").join(" ") || "";
      const contactTitle =
        post.Contact && post.Contact.length > 0
          ? post.Contact.map((c) => c.Title).join(" ")
          : "";

      const contentToSearch = `${post.Title} ${post.Author} ${
        post.PublishedDate
      } ${post.PublishedYear} ${post.PublishedDateAndMonth} ${
        post.PublishedTime
      } ${csiTitles} ${contactTitle} ${textContent}`.toLowerCase();

      if (contentToSearch.includes(searchQuery.toLowerCase())) {
        matches.push(index);
      }
    });

    setSearchResults(matches);
    setCurrentResultIndex(matches.length > 0 ? 0 : -1);
  }, [searchQuery, getFilteredPostsByCsi]);

  // apply filters to cards + table
  const applyFiltersAndUpdatePosts = useCallback(() => {
    const filtered = getFilteredPosts();

    if (searchQuery.trim()) {
      updateSearchResultIndices();
    }

    // TABLE: always show everything filtered
    setTableItems(filtered);

    // CARDS: show only till current page
    const endIndex = page * postsPerPage;
    setVisiblePosts(filtered.slice(0, endIndex));
  }, [getFilteredPosts, updateSearchResultIndices, page, searchQuery]);

  useEffect(() => {
    if (allPosts.length > 0) {
      applyFiltersAndUpdatePosts();
    }
  }, [
    selectedCsiDivisions,
    searchQuery,
    allPosts.length,
    page,
    applyFiltersAndUpdatePosts,
  ]);

  // infinite scroll for cards view
  useEffect(() => {
    if (activeView !== "cards") return;

    const handleObserver = (entries: IntersectionObserverEntry[]) => {
      const [entry] = entries;
      if (entry.isIntersecting && !loading) {
        const filtered = getFilteredPosts();
        if (visiblePosts.length < filtered.length) {
          setPage((prev) => prev + 1);
        }
      }
    };

    const opts = {
      root: null,
      rootMargin: "100px",
      threshold: 0.1,
    };

    if (observerRef.current) observerRef.current.disconnect();
    observerRef.current = new IntersectionObserver(handleObserver, opts);
    if (loaderRef.current) observerRef.current.observe(loaderRef.current);

    return () => {
      if (observerRef.current) observerRef.current.disconnect();
    };
  }, [activeView, loading, visiblePosts.length, getFilteredPosts]);

  const handleSearch = (text: string) => {
    setSearchQuery(text);
    setPage(1);
    if (!text.trim()) {
      setSearchResults([]);
      setCurrentResultIndex(0);
    }
  };

  const highlightText = (text: any) => {
    if (!searchQuery || !text) return text;
    const regex = new RegExp(
      `(${searchQuery.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")})`,
      "gi"
    );
    const parts = text.toString().split(regex);
    return parts.map((part: any, index: any) =>
      regex.test(part) ? (
        <mark key={index} style={{ backgroundColor: "#ffff00" }}>
          {part}
        </mark>
      ) : (
        part
      )
    );
  };

  const highlightHtmlContent = useCallback(
    (html: string) => {
      if (!searchQuery || !html) return html;

      const parser = new DOMParser();
      const doc = parser.parseFromString(html, "text/html");

      const processNode = (node: Node) => {
        if (node.nodeType === Node.TEXT_NODE) {
          const text = node.textContent || "";
          const sanitizedQuery = searchQuery.replace(
            /[.*+?^${}()|[\]\\]/g,
            "\\$&"
          );
          const regex = new RegExp(`(${sanitizedQuery})`, "gi");

          if (regex.test(text)) {
            const span = doc.createElement("span");
            const parts = text.split(regex);

            parts.forEach((part, i) => {
              if (i % 2 === 1) {
                const mark = doc.createElement("mark");
                mark.style.backgroundColor = "#ffff00";
                mark.textContent = part;
                span.appendChild(mark);
              } else if (part) {
                span.appendChild(doc.createTextNode(part));
              }
            });

            if (node.parentNode) {
              node.parentNode.replaceChild(span, node);
            }
          }
        } else if (node.nodeType === Node.ELEMENT_NODE) {
          Array.from(node.childNodes).forEach(processNode);
        }
      };

      Array.from(doc.body.childNodes).forEach(processNode);
      return doc.body.innerHTML;
    },
    [searchQuery]
  );

  // navigation through search results (cards)
  const moveToNext = () => {
    if (
      searchResults.length === 0 ||
      currentResultIndex === searchResults.length - 1
    )
      return;
    const nextIndex = currentResultIndex + 1;
    setCurrentResultIndex(nextIndex);

    const filteredByCsi = getFilteredPostsByCsi();
    const postIndex = searchResults[nextIndex];
    const postId = filteredByCsi[postIndex]?.Id;

    const visibleIndex = visiblePosts.findIndex((p) => p.Id === postId);
    if (visibleIndex === -1) {
      const pageForNext = Math.ceil((postIndex + 1) / postsPerPage);
      setPage((p) => Math.max(p, pageForNext));
    }

    setTimeout(() => {
      const el = document.querySelector(
        `[data-post-id="${postId}"]`
      ) as HTMLDivElement;
      if (el) el.scrollIntoView({ behavior: "smooth", block: "center" });
    }, 100);
  };

  const moveToPrev = () => {
    if (searchResults.length === 0 || currentResultIndex === 0) return;
    const prevIndex = currentResultIndex - 1;
    setCurrentResultIndex(prevIndex);

    const filteredByCsi = getFilteredPostsByCsi();
    const postIndex = searchResults[prevIndex];
    const postId = filteredByCsi[postIndex]?.Id;

    const visibleIndex = visiblePosts.findIndex((p) => p.Id === postId);
    if (visibleIndex === -1) {
      const pageForPrev = Math.ceil((postIndex + 1) / postsPerPage);
      setPage(pageForPrev);
    }

    setTimeout(() => {
      const el = document.querySelector(
        `[data-post-id="${postId}"]`
      ) as HTMLDivElement;
      if (el) el.scrollIntoView({ behavior: "smooth", block: "center" });
    }, 100);
  };

  const getResultCountText = () => {
    const filtered = getFilteredPosts();
    if (searchQuery.trim()) {
      return searchResults.length === 0
        ? "No matches found"
        : `Match ${currentResultIndex + 1} of ${searchResults.length}`;
    }
    return `Showing ${visiblePosts.length} of ${filtered.length} posts`;
  };

  const shouldShowLoader = () => {
    const filtered = getFilteredPosts();
    return activeView === "cards" && visiblePosts.length < filtered.length;
  };

  // column click
  const onColumnClick = (
    ev?: React.MouseEvent<HTMLElement>,
    column?: IColumn
  ) => {
    if (!column) return;
    const desc =
      sortConfig.key === column.key ? !sortConfig.isSortedDescending : false;
    setSortConfig({ key: column.key, isSortedDescending: desc });

    const sorted = sortTable(tableItems, column.key, desc);
    setTableItems(sorted);
  };

  // define columns
  useEffect(() => {
    const cols: IColumn[] = [
      {
        key: "Title",
        name: "Title",
        fieldName: "Title",
        minWidth: 160,
        isResizable: true,
        isSorted: sortConfig.key === "Title",
        isSortedDescending:
          sortConfig.key === "Title" && sortConfig.isSortedDescending,
        onColumnClick: onColumnClick,
        onRender: (item: IBlogPost) => (
          <Stack horizontal tokens={{ childrenGap: 6 }} verticalAlign="center">
            <a
              href={`${siteUrl}/SitePages/LessonsLearnedView.aspx?LessonsLearnedId=${item.Id}`}
            >
              {item.Title}
            </a>
            <TooltipHost
              content={
                <div
                  style={{ maxHeight: 250, overflowY: "auto" }}
                  dangerouslySetInnerHTML={{
                    __html: DOMPurify.sanitize(item.Body || ""),
                  }}
                />
              }
            >
              <IconButton iconProps={{ iconName: "Info" }} />
            </TooltipHost>
          </Stack>
        ),
      },
      {
        key: "PublishedDate",
        name: "Published",
        fieldName: "PublishedDate",
        minWidth: 120,
        isResizable: true,
        isSorted: sortConfig.key === "PublishedDate",
        isSortedDescending:
          sortConfig.key === "PublishedDate" && sortConfig.isSortedDescending,
        onColumnClick: onColumnClick,
        onRender: (item: IBlogPost) =>
          item.PublishedDate
            ? new Date(item.PublishedDate).toLocaleString()
            : "",
      },
      {
        key: "Contact",
        name: "Contact",
        fieldName: "Contact",
        minWidth: 160,
        isResizable: true,
        isSorted: sortConfig.key === "Contact",
        isSortedDescending:
          sortConfig.key === "Contact" && sortConfig.isSortedDescending,
        onColumnClick: onColumnClick,
        onRender: (item: IBlogPost) => {
          const contacts = item.Contact || [];
          if (!contacts || contacts.length === 0) return "—";
          return (
            <div>
              {contacts.map((c, idx) => (
                <div key={idx}>{c?.Title || "Contact"}</div>
              ))}
            </div>
          );
        },
      },
      {
        key: "CSI",
        name: "CSI Division",
        fieldName: "CSI",
        minWidth: 160,
        isResizable: true,
        isSorted: sortConfig.key === "CSI",
        isSortedDescending:
          sortConfig.key === "CSI" && sortConfig.isSortedDescending,
        onColumnClick: onColumnClick,
        onRender: (item: IBlogPost) => (
          <div>
            {(item.CSIDivisions || []).map((d) => (
              <div key={d?.Id}>{d?.Title}</div>
            ))}
          </div>
        ),
      },
    ];
    setColumns(cols);
  }, [siteUrl, sortConfig]);

  const sortTable = (
    items: IBlogPost[],
    key: string,
    desc: boolean
  ): IBlogPost[] => {
    const sorted = [...items].sort((a, b) => {
      let aVal: any = "";
      let bVal: any = "";

      switch (key) {
        case "Title":
          aVal = a.Title || "";
          bVal = b.Title || "";
          break;
        case "PublishedDate":
          aVal = a.PublishedDate ? new Date(a.PublishedDate).getTime() : 0;
          bVal = b.PublishedDate ? new Date(b.PublishedDate).getTime() : 0;
          break;
        case "Contact":
          aVal = a.Contact && a.Contact.length > 0 ? a.Contact[0].Title : "";
          bVal = b.Contact && b.Contact.length > 0 ? b.Contact[0].Title : "";
          break;
        case "CSI":
          aVal = (a.CSIDivisions || [])
            .map((d: any) => d.Title)
            .sort()
            .join(",");
          bVal = (b.CSIDivisions || [])
            .map((d: any) => d.Title)
            .sort()
            .join(",");
          break;
      }

      if (aVal < bVal) return desc ? 1 : -1;
      if (aVal > bVal) return desc ? -1 : 1;
      return 0;
    });

    return sorted;
  };

  return (
    <>
     {canEdit ? (
                  <a href="https://vaughnconstruction.sharepoint.com/ll/Lists/Posts/AllPosts.aspx" target="_blank" data-interception="off" style={{ textDecoration: 'none' }}>
                      <h2 style={{paddingBottom:'10px'}}>Lessons Learned </h2>
                  </a>
              ) : (
                  <h2 style={{paddingBottom:'10px'}}>Lessons Learned </h2>
              )}
                <div className="blogs-table-layout">
        
      {/* LEFT SIDE: CSI ONLY */}
      <div className="csi-panel">
        <div className="ms-core-listMenu-verticalBox ms-noList">
          <ul className="static ms-blog-listMenu-root ms-core-listMenu-root root" style={{width: '100%'}}>
            <li>
              <span className="ms-core-listMenu-item ms-blog-quickLinksTitle">
                CSI Division
              </span>
              <ul className="static">
                {csiDivisions.map((div) => (
                  <li className="static" key={div.Id}>
                    <a
                      className="ms-link"
                      style={{
                        cursor: "pointer",
                        fontWeight: isCsiSelected(div.Id) ? "bold" : "normal",
                        textDecoration: isCsiSelected(div.Id)
                          ? "underline"
                          : "none",
                      }}
                      onClick={() => handleCsiDivisionSelect(div)}
                    >
                      {div.Title}
                    </a>
                  </li>
                ))}
              </ul>
            </li>
          </ul>
        </div>
      </div>

      {/* RIGHT SIDE */}
      <div className="content-area">
        <div className="sticky-header" role="search">
          {/* selected CSI chips */}
          {selectedCsiDivisions.length > 0 && (
            <div
              style={{
                marginBottom: "10px",
                padding: "5px 10px",
                backgroundColor: "#ffffff", // was #f0f0f0
                borderRadius: "4px",
                border: "1px solid #e5e5e5",
              }}
            >
              <div
                style={{
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "center",
                  marginBottom: "5px",
                }}
              >
                <Text variant="medium">Selected CSI Divisions:</Text>
                <IconButton
                  iconProps={{ iconName: "Clear" }}
                  onClick={clearAllCsiDivisions}
                  title="Clear all CSI division filters"
                  styles={{ root: { height: "28px" } }}
                />
              </div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: "5px" }}>
                {selectedCsiDivisions.map((div) => (
                  <div
                    key={div.Id}
                    style={{
                      display: "flex",
                      alignItems: "center",
                      padding: "3px 8px",
                      backgroundColor: "#ffffff",
                      border: "1px solid #d0d0d0",
                      borderRadius: "3px",
                      marginBottom: "3px",
                    }}
                  >
                    <Text variant="smallPlus">{div.Title}</Text>
                    <IconButton
                      iconProps={{ iconName: "Cancel" }}
                      onClick={() => removeCsiDivision(div.Id)}
                      title={`Remove ${div.Title}`}
                      styles={{
                        root: {
                          height: "24px",
                          width: "24px",
                          marginLeft: "4px",
                          padding: "0",
                        },
                        icon: { fontSize: "10px" },
                      }}
                    />
                  </div>
                ))}
              </div>
            </div>
          )}

          <Stack horizontal tokens={{ childrenGap: 10 }}>
            <TextField
              label="Search"
              value={searchQuery}
              onChange={(_, v) => handleSearch(v || "")}
              styles={{ root: { width: "100%" } }}
            />
            <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 5 }}>
              <IconButton
                iconProps={{ iconName: "Up" }}
                onClick={moveToPrev}
                disabled={searchResults.length === 0 || currentResultIndex === 0}
                title="Previous result"
              />
              <IconButton
                iconProps={{ iconName: "Down" }}
                onClick={moveToNext}
                disabled={
                  searchResults.length === 0 ||
                  currentResultIndex === searchResults.length - 1
                }
                title="Next result"
              />
            </Stack>
          </Stack>

          <Stack horizontal tokens={{ childrenGap: 8 }} style={{ marginTop: 8 }}>
            {activeView !== "table" && <LabelLike text={getResultCountText()} />}
            <Stack horizontal tokens={{ childrenGap: 4 }}>
              <PrimaryButton
                text="Table View"
                onClick={() => setActiveView("table")}
                styles={{
                  root: {
                    backgroundColor:
                      activeView === "table" ? "#005a9e" : "#ffffff",
                    color: activeView === "table" ? "#ffffff" : "#005a9e",
                    border: "1px solid #005a9e",
                  },
                }}
              />
              <PrimaryButton
                text="Cards View"
                onClick={() => setActiveView("cards")}
                styles={{
                  root: {
                    backgroundColor:
                      activeView === "cards" ? "#005a9e" : "#ffffff",
                    color: activeView === "cards" ? "#ffffff" : "#005a9e",
                    border: "1px solid #005a9e",
                  },
                }}
              />
            </Stack>
          </Stack>
        </div>
        <div ref={contentRef} className="scrollable-content">
          {loading ? (
            <Spinner label="Loading blog posts..." />
          ) : activeView === "table" ? (
            <DetailsList
              items={tableItems}
              columns={columns}
              selectionMode={SelectionMode.none}
              layoutMode={DetailsListLayoutMode.justified}
              setKey="set"
            />
          ) : visiblePosts.length === 0 ? (
            <div className="no-results">
              <div className="no-results-title">
                No matching blog posts found
              </div>
              <div className="no-results-subtitle">
                Try adjusting your search criteria or CSI filters
              </div>
            </div>
          ) : (
            <>
              {visiblePosts.map((post) => {
                const filteredByCsi = getFilteredPostsByCsi();
                const indexInFiltered = filteredByCsi.findIndex(
                  (p) => p.Id === post.Id
                );
                const isHighlighted =
                  searchQuery.trim() && searchResults.includes(indexInFiltered);
                const isCurrent =
                  searchQuery.trim() &&
                  searchResults[currentResultIndex] === indexInFiltered;

                return (
                  <div
                    className={`blog-post ms-blog-postBox ms-shadow ${
                      isCurrent ? "current" : isHighlighted ? "highlighted" : ""
                    }`}
                    key={post.Id}
                    data-post-id={post.Id}
                  >
                    <div className="ms-blog-postBoxDate">
                      <div className="ms-textSmall">
                        {highlightText(post.PublishedYear)}
                      </div>
                      <div className="ms-textXLarge ms-blog-dateText">
                        {highlightText(post.PublishedDateAndMonth)}
                      </div>
                    </div>
                    <div className="ms-blog-postBoxMargin">
                      <h2>
                        <a
                          href={`${siteUrl}/SitePages/DemoLessonsLearned.aspx?PostId=${post.Id}`}
                        >
                          {highlightText(post.Title)}
                        </a>
                      </h2>
                      <div className="ms-metadata ms-textSmall">
                        <span>
                          by{" "}
                          <span className="ms-noWrap ms-imnSpan">
                            <a className="ms-subtleLink">
                              {highlightText(post.Author)}
                            </a>
                          </span>{" "}
                          at {highlightText(post.PublishedTime)} in{" "}
                          {(post.CSIDivisions || []).map((div, idx) => (
                            <span key={div.Id}>
                              {idx > 0 && ", "}
                              <a
                                style={{
                                  cursor: "pointer",
                                  fontWeight: isCsiSelected(div.Id)
                                    ? "bold"
                                    : "normal",
                                  textDecoration: isCsiSelected(div.Id)
                                    ? "underline"
                                    : "none",
                                }}
                                onClick={() => handleCsiDivisionSelect(div)}
                                className="ms-link"
                              >
                                {highlightText(div.Title)}
                              </a>
                            </span>
                          ))}
                        </span>
                      </div>
                      <div className="ms-blog-postBody">
                        <div
                          className="ms-rtestate-field"
                          dangerouslySetInnerHTML={{
                            __html: DOMPurify.sanitize(
                              highlightHtmlContent(post.Body)
                            ),
                          }}
                        />
                      </div>
                    </div>
                  </div>
                );
              })}

              {shouldShowLoader() && (
                <div ref={loaderRef} className="loader">
                  <Spinner label="Loading more posts..." />
                </div>
              )}
            </>
          )}
        </div>
      </div>
    </div>
    </>
  
  );
};

// tiny helper so we don't import Label again
const LabelLike = ({ text }: { text: string }) => (
  <span style={{ paddingTop: 4 }}>{text}</span>
);

export default BlogsTable;
