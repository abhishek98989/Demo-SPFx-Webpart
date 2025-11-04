import * as React from "react";
import './style.css'
import { useEffect, useState, useRef, useCallback } from "react";
import {
  TextField,
  Stack,
  Label,
  IconButton,
  Spinner,
  Text,
  Panel,
  PanelType,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  DefaultButton,
  TooltipHost,
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
  Contact?: { Id: number; Title: string } | null;
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

  // edit panel
  const [isEditOpen, setIsEditOpen] = useState<boolean>(false);
  const [editItem, setEditItem] = useState<IBlogPost | null>(null);
  const [editTitle, setEditTitle] = useState<string>("");
  const [editCsiDivisions, setEditCsiDivisions] = useState<number[]>([]);

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
      const params = new URLSearchParams(window.location.search);
      const postFilter = params.get("PostId");

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
          Contact: item.Contact
            ? { Id: item.Contact.Id, Title: item.Contact.Title }
            : null,
          CSIDivisions: item["CSI_x0020_Div"]
            ? Array.isArray(item["CSI_x0020_Div"])
              ? item["CSI_x0020_Div"]
              : [item["CSI_x0020_Div"]]
            : [],
        };
      });

      const sorted = sortByDate(mapped);
      setAllPosts(sorted);

      // FIXED: Don't set visiblePosts or tableItems here
      // Let the filter effect handle it

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
      return exists ? prev.filter((d) => d.Id !== division.Id) : [...prev, division];
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
        const contactTitle = post.Contact?.Title || "";

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
      const contactTitle = post.Contact?.Title || "";

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

  // FIXED: apply filters to cards + table - removed dependencies causing issues
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
  }, [getFilteredPosts, updateSearchResultIndices, page]);

  // FIXED: Added allPosts.length check to trigger on data load
  useEffect(() => {
    if (allPosts.length > 0) {
      applyFiltersAndUpdatePosts();
    }
  }, [selectedCsiDivisions, searchQuery, allPosts.length, page, applyFiltersAndUpdatePosts]);

  // FIXED: infinite scroll for cards view - added getFilteredPosts dependency
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

  // FIXED: TABLE columns - moved onColumnClick definition before columns array
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

  useEffect(() => {
    const cols: IColumn[] = [
      {
        key: "Title",
        name: "Title",
        fieldName: "Title",
        minWidth: 200,
        isResizable: true,
        isSorted: sortConfig.key === "Title",
        isSortedDescending:
          sortConfig.key === "Title" && sortConfig.isSortedDescending,
        onColumnClick: onColumnClick,
        onRender: (item: IBlogPost) => (
          <Stack horizontal tokens={{ childrenGap: 6 }} verticalAlign="center">
            <a
              href={`${siteUrl}/SitePages/DemoLessonsLearned.aspx?PostId=${item.Id}`}
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
        minWidth: 120,
        isResizable: true,
        isSorted: sortConfig.key === "Contact",
        isSortedDescending:
          sortConfig.key === "Contact" && sortConfig.isSortedDescending,
        onColumnClick: onColumnClick,
        onRender: (item: IBlogPost) => item.Contact?.Title || "",
      },
      {
        key: "CSI",
        name: "CSI Division",
        fieldName: "CSI",
        minWidth: 140,
        isResizable: true,
        isSorted: sortConfig.key === "CSI",
        isSortedDescending:
          sortConfig.key === "CSI" && sortConfig.isSortedDescending,
        onColumnClick: onColumnClick,
        onRender: (item: IBlogPost) =>
          (item.CSIDivisions || [])
            .map((d: any) => d.Title)
            .filter(Boolean)
            .join(", "),
      },
      {
        key: "Edit",
        name: "Edit",
        fieldName: "edit",
        minWidth: 50,
        onRender: (item: IBlogPost) => (
          <IconButton
            iconProps={{ iconName: "Edit" }}
            onClick={() => openEditPanel(item)}
            title="Edit"
          />
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
          aVal = a.Contact?.Title || "";
          bVal = b.Contact?.Title || "";
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

  // edit panel
  const openEditPanel = (item: IBlogPost) => {
    setEditItem(item);
    setEditTitle(item.Title);
    setEditCsiDivisions((item.CSIDivisions || []).map((d: any) => d.Id));
    setIsEditOpen(true);
  };
  const closeEditPanel = () => setIsEditOpen(false);

  const onSaveEdit = async () => {
    if (!editItem) return;
    try {
      const sp = spfi().using(SPFx(props.Context));
      await sp.web.lists.getById(ListIds.posts).items.getById(editItem.Id).update({
        Title: editTitle,
        "CSI_x0020_DivId": editCsiDivisions,
      });

      const updated = allPosts.map((p) =>
        p.Id === editItem.Id
          ? {
              ...p,
              Title: editTitle,
              CSIDivisions: csiDivisions.filter((d) =>
                editCsiDivisions.includes(d.Id)
              ),
            }
          : p
      );
      setAllPosts(sortByDate(updated));
      setIsEditOpen(false);
    } catch (err) {
      console.error("Error updating post", err);
    }
  };

  const csiDropdownOptions: IDropdownOption[] = csiDivisions.map((d) => ({
    key: d.Id,
    text: d.Title,
  }));

  return (
    <div className="blogs-table-layout">
      {/* LEFT SIDE: CSI ONLY */}
      <div className="csi-panel">
        <div className="ms-core-listMenu-verticalBox ms-noList">
          <ul className="static ms-blog-listMenu-root ms-core-listMenu-root root">
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
                backgroundColor: "#f0f0f0",
                borderRadius: "4px",
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
                      backgroundColor: "#e0e0e0",
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
            {activeView !== "table" && <Label>{getResultCountText()}</Label>}
            <Stack horizontal tokens={{ childrenGap: 4 }}>
              <PrimaryButton
                text="Table View"
                onClick={() => setActiveView("table")}
                styles={{
                  root: {
                    backgroundColor:
                      activeView === "table" ? "#005a9e" : undefined,
                  },
                }}
              />
              <DefaultButton
                text="Cards View"
                onClick={() => setActiveView("cards")}
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

      {/* EDIT PANEL */}
      <Panel
        isOpen={isEditOpen}
        onDismiss={closeEditPanel}
        headerText="Edit Post"
        closeButtonAriaLabel="Close"
        type={PanelType.medium}
      >
        <TextField
          label="Title"
          value={editTitle}
          onChange={(_, v) => setEditTitle(v || "")}
        />
        <Dropdown
          label="CSI Division"
          multiSelect
          selectedKeys={editCsiDivisions}
          onChange={(_, option) => {
            if (!option) return;
            const id = option.key as number;
            setEditCsiDivisions((prev) => {
              if (option.selected) return [...prev, id];
              return prev.filter((p) => p !== id);
            });
          }}
          options={csiDropdownOptions}
        />
        <Stack horizontal tokens={{ childrenGap: 8 }} style={{ marginTop: 16 }}>
          <PrimaryButton text="Save" onClick={onSaveEdit} />
          <DefaultButton text="Cancel" onClick={closeEditPanel} />
        </Stack>
      </Panel>
    </div>
  );
};

export default BlogsTable;
