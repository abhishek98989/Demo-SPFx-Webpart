import * as React from "react";
import { useEffect, useRef, useState, useMemo } from "react";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/items/list";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/folders";
import "@pnp/sp/files";

import { DatePicker } from "@fluentui/react/lib/DatePicker";
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
} from "@fluentui/react/lib/DetailsList";
// ADD this import
import { SPHttpClient } from "@microsoft/sp-http";
import { PeoplePicker, PrincipalType, IPeoplePickerUserItem } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { Dropdown, IDropdownOption } from "@fluentui/react/lib/Dropdown";
import { TextField } from "@fluentui/react/lib/TextField";
import { PrimaryButton, DefaultButton, IconButton } from "@fluentui/react/lib/Button";
import { Stack } from "@fluentui/react/lib/Stack";
import { Panel, PanelType } from "@fluentui/react/lib/Panel";
import { Dialog, DialogType, DialogFooter } from "@fluentui/react/lib/Dialog";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import { Breadcrumb, IBreadcrumbItem } from "@fluentui/react/lib/Breadcrumb";
import { Checkbox } from "@fluentui/react/lib/Checkbox";
import JoditEditor from "jodit-react";

// NOTE: class kept but not used by paste anymore
class SPImageUploadAdapter {
  private loader: any;
  private sp: any;
  private folderUrl: string;

  constructor(loader: any, context: any, folderUrl: string) {
    this.loader = loader;
    this.sp = spfi('https://vaughnconstruction.sharepoint.com/news').using(SPFx(context));
    this.folderUrl = folderUrl;
  }

  async upload() {
    const file = await this.loader.file;
    const fileName = (file?.name || `image-${Date.now()}.png`).trim();

    const folder = this.sp.web.getFolderByServerRelativePath(this.folderUrl);
    const result = await folder.files.addUsingPath(fileName, file, { Overwrite: true });

    const url =
      (result as any)?.data?.ServerRelativeUrl ||
      (result as any)?.file?.serverRelativeUrl ||
      (result as any)?.ServerRelativeUrl;

    if (!url) throw new Error("Upload succeeded but no URL returned from SharePoint.");

    return { default: url };
  }

  abort() { }
}

export interface INewsManagerProps {
  context: any;
  listId: string;
  siteUrl: any;
  imagesWebServerRelativeUrl?: string; // default: "/news"
  imagesLibraryTitle?: string; // default: "Publishing Images"
  tinymceApiKey?: string;
}

interface IListItem {
  Id: number;
  Title: string;
  Author?: { Id: number; Title: string; EMail?: string };
  Created: string;
  Published?: boolean;
  Description?: string;                // HTML (VaughnContent)
  // DepartmentSpecific removed
  Department?: string;                 // NEW (Choice)
  PublishingSource?: string;           // NEW (Choice)
  PublishingRollupImageUrl?: string;   // hyperlink URL
  Abstract?: string;                   // multiline
  InCaseYouMissed?: Array<{ Id: number; Title: string }>;
  ArticleDate?: string;
}

const NewsManager: React.FC<INewsManagerProps> = (props) => {
  const {
    context,
    listId,
    siteUrl,
    imagesWebServerRelativeUrl = "/news",
    imagesLibraryTitle = "Publishing Images",
    tinymceApiKey = "no-api-key",
  } = props;

  const sp = siteUrl != undefined ? spfi(siteUrl).using(SPFx(context)) : spfi().using(SPFx(context));
  const [searchRaw, setSearchRaw] = useState("");
  const [search, setSearch] = useState("");
  const [sortKey, setSortKey] = useState<string>("col_artdate");
  const [sortDesc, setSortDesc] = useState<boolean>(true);

  const [items, setItems] = useState<IListItem[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | null>(null);

  // Panel state (Add/Edit)
  const [isPanelOpen, setIsPanelOpen] = useState(false);
  const [editingItem, setEditingItem] = useState<IListItem | null>(null);
  const [title, setTitle] = useState("");
  const [descriptionHtml, setDescriptionHtml] = useState("");
  const [published, setPublished] = useState<boolean>(false);

  // NEW: Department & PublishingSource
  const [department, setDepartment] = useState<string | undefined>(undefined);
  const [publishingSource, setPublishingSource] = useState<string | undefined>(undefined);

  // Existing
  const [rollupImageUrl, setRollupImageUrl] = useState<string>("");
  const [abstractText, setAbstractText] = useState<string>("");

  // People picker
  const [publishingContactId, setPublishingContactId] = useState<number | null>(null);
  const [publishingContactEmail, setPublishingContactEmail] = useState<string | undefined>(undefined);

  // ICYM
  const [icymOptions, setIcymOptions] = useState<Array<{ key: number; text: string }>>([]);
  const [icymSelected, setIcymSelected] = useState<number[]>([]);

  // Article Date 
  const [articleDate, setArticleDate] = useState<Date | null>(null);

  // Image picker
  const [imagePickerTarget, setImagePickerTarget] = useState<"editor" | "rollup">("editor");
  const [isImagePickerOpen, setIsImagePickerOpen] = useState(false);
  const [pickerError, setPickerError] = useState<string | null>(null);
  const [currentFolder, setCurrentFolder] = useState<string>("");
  const [folderTrail, setFolderTrail] = useState<IBreadcrumbItem[]>([]);
  const [folderSubfolders, setFolderSubfolders] = useState<any[]>([]);
  const [folderFiles, setFolderFiles] = useState<any[]>([]);
  const [uploading, setUploading] = useState(false);
  const [urlToInsert, setUrlToInsert] = useState("");

  // NEW: processing pasted images state
  const [processingImages, setProcessingImages] = useState(false);

  const editorRef = useRef<any>(null);

  // SIMPLE JODIT CONFIG – paste is default, no custom events.afterPaste
  const joditConfig = React.useMemo(
    () => ({
      readonly: false,
      height: 400,
      toolbarSticky: false,
      buttons: [
        "source", "|",
        "bold", "italic", "underline", "strikethrough", "|",
        "ul", "ol", "indent", "outdent", "|",
        "font", "fontsize", "paragraph", "|",
        "left", "center", "right", "justify", "|",
        "image", "table", "link", "|",
        "hr", "eraser", "fullsize", "selectall", "undo", "redo"
      ],
      allowResizeX: true,
      tableAllowCellResize: true,
      allowResizeY: true,
      imageDefaultWidth: 300,
        allowResizeTags: new Set(['img', 'iframe', 'table']),
      enableDragAndDropFileToEditor: true,
      // let Jodit handle paste normally (images will be base64/blob)
    }),
    []
  );

  // NEW: dropdown options
  const departmentOptions: IDropdownOption[] = [
    { key: "__NONE__", text: "None" },
    { key: "HR", text: "HR" },
    { key: "IT", text: "IT" },
    { key: "Marketing", text: "Marketing" },
    { key: "OPS", text: "OPS" },
    { key: "VDC", text: "VDC" },
    { key: "Safety", text: "Safety" },
  ];

  const publishingSourceOptions: IDropdownOption[] = [
    { key: "Weekly Words", text: "Weekly Words" },
    { key: "Department Specific", text: "Department Specific" },
    { key: "Both", text: "Both" },
  ];

  const baseColumns: IColumn[] = [
    { key: "col_title", name: "Title", fieldName: "Title", minWidth: 160, isResizable: true },
    { key: "col_author", name: "Author", minWidth: 120, onRender: (i: IListItem) => i.Author?.Title ?? "—" },
    { key: "col_pub", name: "Published", minWidth: 110, onRender: (i) => i.Published ? "Approved" : "Pending" },
    { key: "col_dept", name: "Department", minWidth: 110, onRender: (i) => i.Department || "None" },
    { key: "col_src", name: "Publishing Source", minWidth: 140, onRender: (i) => i.PublishingSource || "—" },
    {
      key: "col_artdate",
      name: "Article Date",
      minWidth: 140,
      onRender: (i: IListItem) => {
        const raw = i.ArticleDate || ""; // fallback for safety
        if (!raw) return "—";
        const d = new Date(raw);
        return Number.isFinite(d.getTime()) ? d.toLocaleDateString() : "—";
      },
    }, {
      key: "col_rollup", name: "Preview Image", minWidth: 100, onRender: (i) =>
        i.PublishingRollupImageUrl ? <img src={i.PublishingRollupImageUrl} alt="" style={{ width: 72, height: 48, objectFit: "cover", borderRadius: 4 }} /> : "—"
    },
    {
      key: "col_actions", name: "Actions", minWidth: 140, onRender: (i: IListItem) => (
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <IconButton iconProps={{ iconName: "Edit" }} title="Edit" ariaLabel="Edit" onClick={() => openEditPanel(i)} />
          <IconButton iconProps={{ iconName: "Delete" }} title="Delete" ariaLabel="Delete" onClick={() => confirmDelete(i)} />
        </Stack>
      )
    }
  ];

  // decorate with sort props and click handler (avoid sorting on actions column)
  const columns: IColumn[] = baseColumns.map((c) => {
    const sortable = c.key !== "col_actions" && c.key !== "col_rollup";
    return {
      ...c,
      isSorted: sortable && c.key === sortKey,
      isSortedDescending: sortable && c.key === sortKey ? sortDesc : undefined,
      onColumnClick: sortable ? onColumnClick : undefined,
    };
  });

  // Initial data load
  useEffect(() => {
    (async () => {
      await loadItems();
      const rootFolder = `${imagesWebServerRelativeUrl}/PublishingImages`;
      setCurrentFolder(rootFolder);
      setFolderTrail([
        { key: "root", text: imagesLibraryTitle, onClick: () => setCurrentFolder(rootFolder) },
      ]);
    })().catch((e) => setError(parseError(e)));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [imagesWebServerRelativeUrl, imagesLibraryTitle]);

  useEffect(() => {
    const t = setTimeout(() => setSearch(searchRaw.trim().toLowerCase()), 250);
    return () => clearTimeout(t);
  }, [searchRaw]);

  async function loadItems() {
    try {
      setLoading(true);
      setError(null);

      const rows: any[] = await sp.web.lists
        .getById(listId)
        .items
        .select(
          "Id",
          "Title",
          "Created",
          "OData__ModerationStatus",     // publish flag (your existing usage)
          "VaughnContent",               // HTML
          "Department",                  // NEW
          "PublishingSource",            // NEW
          "PublishingRollupImage",       // Hyperlink
          "Abstract",
          "InCaseYouMissed/Id",
          "InCaseYouMissed/Title",
          "PublishingContact/Id",
          "PublishingContact/Title",
          "PublishingContact/EMail",
          "ArticleDate"
        )
        .expand("InCaseYouMissed", "PublishingContact")
        .orderBy("Created", false)();

      const mapped: IListItem[] = rows.map(r => ({
        Id: r.Id,
        Title: r.Title,
        Author: r.PublishingContact ? {
          Id: r.PublishingContact.Id,
          Title: r.PublishingContact.Title,
          EMail: r.PublishingContact.EMail
        } : undefined,
        Created: r.Created,
        // CHANGED: 0=Approved, 2=Pending, 1=Rejected, 3=Draft, 4=Scheduled
        Published: r.OData__ModerationStatus === 0,
        Description: r.VaughnContent as string,
        Department: r.Department || undefined,
        PublishingSource: r.PublishingSource || undefined,
        PublishingRollupImageUrl: r.PublishingRollupImage ? (r.PublishingRollupImage.Url || r.PublishingRollupImage) : "",
        Abstract: r.Abstract || "",
        ArticleDate: r.ArticleDate || undefined,
        InCaseYouMissed: Array.isArray(r.InCaseYouMissed) ? r.InCaseYouMissed.map((x: any) => ({ Id: x.Id, Title: x.Title })) : []
      }));

      setItems(mapped);

      const options = mapped
        .filter(m => m?.Published && !!m?.Title && m?.Title.trim().length > 0)
        .map(m => ({ key: m.Id, text: m.Title }));

      setIcymOptions(options);

    } catch (e) {
      setError(parseError(e));
    } finally {
      setLoading(false);
    }
  }

  function openAddPanel() {
    setEditingItem(null);
    setTitle("");
    setDescriptionHtml("");
    setPublished(false);

    // NEW fields reset
    setDepartment(undefined);
    setPublishingSource(undefined);

    setRollupImageUrl("");
    setAbstractText("");
    setPublishingContactId(null);
    setPublishingContactEmail(undefined);
    setIcymSelected([]);
    setArticleDate(null);
    setIsPanelOpen(true);
  }

  function openEditPanel(item: IListItem) {
    setEditingItem(item);
    setTitle(item.Title);
    setDescriptionHtml(item.Description || "");
    setPublished(!!item.Published);
    setArticleDate(item.ArticleDate ? new Date(item.ArticleDate) : null);
    // NEW fields populate
    setDepartment(item.Department ?? undefined);
    setPublishingSource(item.PublishingSource);

    setRollupImageUrl(item.PublishingRollupImageUrl || "");
    setAbstractText(item.Abstract || "");
    setPublishingContactId(item.Author?.Id ?? null);
    setPublishingContactEmail(item.Author?.EMail);
    setIcymSelected((item.InCaseYouMissed || []).map(x => x.Id));

    setIsPanelOpen(true);
  }

  // HELPER: insert image at cursor into Jodit
  function insertImageAtCursor(src: string, alt = "") {
    const imgHtml = `<img src="${src}" alt="${alt}" />`;

    const maybeEditor =
      (editorRef.current as any)?.editor ||
      (editorRef.current as any) ||
      (editorRef.current as any)?.getEditor?.();

    if (maybeEditor) {
      try {
        if (maybeEditor?.s?.focus) {
          maybeEditor.s.focus();
        }

        if (maybeEditor?.selection?.insertHTML) {
          maybeEditor.selection.insertHTML(imgHtml);
        } else if (maybeEditor?.value !== undefined) {
          maybeEditor.value = (maybeEditor.value || "") + `<p>${imgHtml}</p>`;
        }

        const currentHtml =
          maybeEditor?.value !== undefined
            ? maybeEditor.value
            : (descriptionHtml || "") + `<p>${imgHtml}</p>`;

        setDescriptionHtml(currentHtml);
        return;
      } catch (e) {
        // fall through to React-only fallback
      }
    }

    setDescriptionHtml((h) => (h || "") + `<p>${imgHtml}</p>`);
  }

  // NEW: PROCESS PASTED IMAGES (button + used in save)
 // NEW: PROCESS PASTED IMAGES (button + used in save)
// NEW: PROCESS PASTED IMAGES (button + used in save)
// NEW: PROCESS PASTED IMAGES (button + used in save)
const processEditorImages = async (): Promise<string | null> => {
  const maybeEditor =
    (editorRef.current as any)?.editor ||
    (editorRef.current as any)?.getEditor?.() ||
    (editorRef.current as any);

  if (!maybeEditor) {
    console.warn("No editor instance found");
    return descriptionHtml;
  }

  // Check if maybeEditor itself IS the root element (DIV)
  let root: HTMLElement | null = null;
  
  if (maybeEditor instanceof HTMLElement) {
    // maybeEditor is directly the div
    root = maybeEditor.classList?.contains('jodit-wysiwyg') 
      ? maybeEditor 
      : maybeEditor.querySelector?.('.jodit-wysiwyg') || maybeEditor;
  } else {
    // Try the usual properties
    root =
      maybeEditor?.editor ||
      maybeEditor?.container?.querySelector?.(".jodit-wysiwyg") ||
      maybeEditor?.workplace ||
      document.querySelector?.(".jodit-wysiwyg") ||
      null;
  }

  if (!root) {
    console.warn("Could not find editor root element");
    return descriptionHtml;
  }

  // Query ALL images deeply nested (querySelectorAll searches all descendants)
  const imgs = root.querySelectorAll("img[src^='data:image'], img[src^='blob:']");
  
  if (!imgs.length) {
    // Get HTML from either the div or editor value
    const htmlNow = (maybeEditor instanceof HTMLElement) 
      ? root.innerHTML 
      : (maybeEditor?.value ?? root.innerHTML ?? descriptionHtml);
    return htmlNow;
  }

  try {
    setProcessingImages(true);

    const spImg = spfi("https://vaughnconstruction.sharepoint.com/news").using(SPFx(context));
    const folderUrl = `${imagesWebServerRelativeUrl}/PublishingImages`;
    const folder = spImg.web.getFolderByServerRelativePath(folderUrl);

    for (const node of Array.from(imgs)) {
      const img = node as HTMLImageElement;
      const src = img.getAttribute("src");
      if (!src) continue;

      // Double-check (though selector should have filtered)
      if (!src.startsWith("data:image") && !src.startsWith("blob:")) continue;

      const blob = await fetch(src).then((r) => r.blob());
      const fileName = `paste-${Date.now()}-${Math.floor(Math.random() * 1000)}.png`;

      const addResult = await folder.files.addUsingPath(fileName, blob, { Overwrite: true });

      const url =
        (addResult as any)?.data?.ServerRelativeUrl ||
        (addResult as any)?.file?.serverRelativeUrl ||
        (addResult as any)?.ServerRelativeUrl;

      if (url) {
        // Preserve original Word dimensions if available
        const width = img.getAttribute("width");
        const height = img.getAttribute("height");
        if (width) img.style.width = `${width}px`;
        if (height) img.style.height = `${height}px`;

        img.setAttribute("src", url);
        img.style.maxWidth = "100%";
        img.style.height = "auto";
      }
    }

    // Get the updated HTML
    const newHtml = (maybeEditor instanceof HTMLElement)
      ? root.innerHTML
      : (maybeEditor?.value !== undefined ? maybeEditor.value : root.innerHTML);

    setDescriptionHtml(newHtml);
    
    // If maybeEditor has a value property, update it too
    if (maybeEditor?.value !== undefined && !(maybeEditor instanceof HTMLElement)) {
      maybeEditor.value = newHtml;
    }
    
    return newHtml;
  } catch (e) {
    console.error("Error while processing pasted images", e);
    setError(parseError(e));
    return null;
  } finally {
    setProcessingImages(false);
  }
};
  async function handleUpload(files: FileList | null) {
    if (!files?.length) return;
    try {
      setUploading(true);
      setPickerError(null);
      let spImg = spfi('https://vaughnconstruction.sharepoint.com/news').using(SPFx(context));
      const file = files[0];
      const folderUrl = currentFolder || `${imagesWebServerRelativeUrl}/PublishingImages`;
      const folder = spImg.web.getFolderByServerRelativePath(folderUrl);

      const fileName = (file.name || `image-${Date.now()}.png`).trim();
      const addResult = await folder.files.addUsingPath(fileName, file, { Overwrite: true });

      const url =
        (addResult as any)?.data?.ServerRelativeUrl ||
        (addResult as any)?.file?.serverRelativeUrl ||
        (addResult as any)?.ServerRelativeUrl;

      if (!url) throw new Error("Upload succeeded, but no URL returned from SharePoint.");

      if (imagePickerTarget === "editor") {
        insertImageAtCursor(url, fileName);
      } else {
        setRollupImageUrl(url);
      }

      await refreshFolder();
    } catch (e) {
      setPickerError(parseError(e));
    } finally {
      setUploading(false);
    }
  }

  async function refreshFolder() {
    try {
      let spImg = spfi('https://vaughnconstruction.sharepoint.com/news').using(SPFx(context));
      const folder = spImg.web.getFolderByServerRelativePath(currentFolder);
      const [subs, files] = await Promise.all([folder.folders(), folder.files()]);
      setFolderSubfolders(subs);
      setFolderFiles(files);
    } catch (e) {
      setPickerError(parseError(e));
    }
  }

  function handleUrlInsert() {
    if (!urlToInsert) return;
    insertImageAtCursor(urlToInsert);
    setUrlToInsert("");
  }

  // CHANGE in saveItem(): run paste-image processing before save
  async function saveItem() {
    try {
      setLoading(true);
      setError(null);

      // Ensure pasted/dragged images are uploaded + URLs fixed
      const processedHtml = await processEditorImages();
      const contentToSave = processedHtml ?? descriptionHtml;

      const list = sp.web.lists.getById(listId);

      const body: any = {
        Title: title,
        VaughnContent: contentToSave,
        Abstract: abstractText,
        PublishingRollupImage: rollupImageUrl ? { Url: rollupImageUrl, Description: "Rollup" } : null,
        PublishingContactId: publishingContactId ?? null,
        Department: department || null,
        PublishingSource: publishingSource || null,
        ArticleDate: articleDate ? articleDate : null
        // DO NOT try to set OData__ModerationStatus directly
      };

      if (icymSelected.length > 0) {
        body.InCaseYouMissedId = icymSelected;
      }

      let itemId: number;

      if (editingItem) {
        await list.items.getById(editingItem.Id).update(body);
        itemId = editingItem.Id;

        if (icymSelected.length === 0) {
          await list.items.getById(itemId).update({ InCaseYouMissedId: [] });
        }
      } else {
        const addRes = await list.items.add(body);
        itemId = addRes.data.Id;
      }

      setIsPanelOpen(false);
      await loadItems();
    } catch (e) {
      setError(parseError(e));
      setLoading(false);
    }
  }

  // Delete flow
  const [deleteTarget, setDeleteTarget] = useState<IListItem | null>(null);
  function confirmDelete(item: IListItem) { setDeleteTarget(item); }
  async function doDelete() {
    if (!deleteTarget) return;
    try {
      setLoading(true);
      const list = sp.web.lists.getById(listId);
      await list.items.getById(deleteTarget.Id).recycle();
      setDeleteTarget(null);
      await loadItems();
    } catch (e) {
      setError(parseError(e));
      setLoading(false);
    }
  }

  // accessors for sorting per column
  const accessors: Record<string, (i: IListItem) => string> = {
    col_title: (i) => i.Title || "",
    col_author: (i) => i.Author?.Title || "",
    col_pub: (i) => (i.Published ? "Approved" : "Pending"),
    col_dept: (i) => i.Department || "None",
    col_src: (i) => i.PublishingSource || "",
    col_rollup: (i) => i.PublishingRollupImageUrl || "",
    col_artdate: (i) => i.ArticleDate || "",
  };

  function tsOf(i: IListItem): number {
    const s = accessors.col_artdate(i);
    const n = s ? new Date(s).getTime() : NaN;
    return Number.isFinite(n) ? n : 0;
  }

  function onColumnClick(ev?: React.MouseEvent<HTMLElement>, column?: IColumn) {
    if (!column) return;
    const newKey = column.key;
    if (newKey === sortKey) {
      setSortDesc(!sortDesc);
    } else {
      setSortKey(newKey);
      setSortDesc(false);
    }
  }

  function stripHtml(html?: string) {
    if (!html) return "";
    return html.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
  }

  const haystack = (i: IListItem) => {
    const parts = [
      i.Title,
      i.Author?.Title,
      i.Author?.EMail,
      i.Department,
      i.PublishingSource,
      i.Abstract,
      i.PublishingRollupImageUrl,
      i.Published ? "approved" : "pending",
      stripHtml(i.Description),
      i.ArticleDate,
      i.Created,
    ]
      .filter(Boolean)
      .join(" ")
      .toLowerCase();

    return parts;
  };

  const displayItems = useMemo(() => {
    let rows = items;

    if (search) {
      rows = rows.filter((i) => haystack(i).includes(search));
    }

    if (sortKey === "col_artdate") {
      const sorted = [...rows].sort((a, b) => tsOf(a) - tsOf(b));
      return sortDesc ? sorted.reverse() : sorted;
    }

    const acc = accessors[sortKey] || accessors["col_title"];
    const sorted = [...rows].sort((a, b) => {
      const av = acc(a).toLocaleLowerCase();
      const bv = acc(b).toLocaleLowerCase();
      if (av < bv) return -1;
      if (av > bv) return 1;
      return 0;
    });

    return sortDesc ? sorted.reverse() : sorted;
  }, [items, search, sortKey, sortDesc]);

  // Image Picker logic
  useEffect(() => {
    if (!isImagePickerOpen || !currentFolder) return;
    (async () => {
      try {
        setPickerError(null);
        let spImg = spfi('https://vaughnconstruction.sharepoint.com/news').using(SPFx(context));
        const folder = await spImg.web.getFolderByServerRelativePath(currentFolder);
        const [subs, files] = await Promise.all([folder.folders(), folder.files()]);
        setFolderSubfolders(subs);
        setFolderFiles(files);

        const parts = currentFolder.replace(/^\/+/, "").split("/");
        const baseIndex = parts.findIndex((p) => p.toLowerCase() === "publishingimages");
        const trailParts = parts.slice(0, baseIndex + 1);
        const crumbs: IBreadcrumbItem[] = trailParts.map((p, idx) => {
          const path = "/" + parts.slice(0, idx + 1).join("/");
          return { key: path, text: idx === 0 ? (imagesWebServerRelativeUrl.replace(/^\/+/, "")) : p, onClick: () => setCurrentFolder(path) } as IBreadcrumbItem;
        });
        const after = parts.slice(baseIndex + 1);
        after.forEach((seg, i) => {
          const path = "/" + parts.slice(0, baseIndex + 2 + i).join("/");
          crumbs.push({ key: path, text: seg, onClick: () => setCurrentFolder(path) });
        });
        setFolderTrail(crumbs);
      } catch (e) {
        setPickerError(parseError(e));
      }
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isImagePickerOpen, currentFolder]);

  const imageExtensions = new Set([".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg"]);

  return (
    <div className="p-4">
      <Stack tokens={{ childrenGap: 12 }}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <a href="https://vaughnconstruction.sharepoint.com/sites/Intranet/Lists/TestWeeklyWords/AllItems.aspx" target="_blank" rel="noopener noreferrer">
            <h2 style={{ margin: 0 }}>Weekly Words Management</h2>
          </a>
          <PrimaryButton text="Add new" onClick={openAddPanel} iconProps={{ iconName: "Add" }} />
        </Stack>

        {error && (
          <MessageBar messageBarType={MessageBarType.error} isMultiline={false} onDismiss={() => setError(null)}>
            {error}
          </MessageBar>
        )}
        <TextField
          placeholder="Search all columns…"
          value={searchRaw}
          onChange={(_, v) => setSearchRaw(v || "")}
          iconProps={{ iconName: "Search" }}
        />
        <DetailsList
          items={displayItems}
          columns={columns}
          compact={false}
          selectionMode={SelectionMode.none}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}
          isHeaderVisible
        />

        {loading && <span>Loading…</span>}
      </Stack>

      {/* Add/Edit Panel */}
      <Panel
        isOpen={isPanelOpen}
        onDismiss={() => setIsPanelOpen(false)}
        isBlocking={false}
        type={PanelType.largeFixed}
        headerText={editingItem ? "Edit item" : "Add new item"}
        closeButtonAriaLabel="Close"
      >
        <Stack tokens={{ childrenGap: 12 }}>
          <TextField label="Title" required value={title} onChange={(_, v) => setTitle(v || "")} />

          <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
            <PrimaryButton
              text="Insert image"
              iconProps={{ iconName: "Photo" }}
              onClick={() => { setImagePickerTarget("editor"); setIsImagePickerOpen(true); }}
            />
            <DefaultButton
              text="Insert URL"
              onClick={handleUrlInsert}
              iconProps={{ iconName: "Link" }}
            />
            <TextField
              placeholder="https://…"
              value={urlToInsert}
              onChange={(_, v) => setUrlToInsert(v || "")}
              styles={{ root: { width: 260 } }}
            />
            <DefaultButton
              text={processingImages ? "Processing images…" : "Process pasted images"}
              iconProps={{ iconName: "Photo2" }}
              disabled={processingImages}
              onClick={async () => { await processEditorImages(); }}
            />
          </Stack>

          <div>
            <label style={{ display: "block", fontWeight: 600, marginBottom: 6 }}>
              Page Content
            </label>

            <JoditEditor
              ref={(instance) => {
                editorRef.current = instance;
              }}
              value={descriptionHtml || ""}
              config={joditConfig}
              onBlur={(newContent) => setDescriptionHtml(newContent)}
              onChange={() => { }}
            />
          </div>

          {/* Publishing Contact, Department, and Publishing Source in one row */}
          <Stack horizontal tokens={{ childrenGap: 12 }}>
            <Stack.Item grow={1}>
              <div>
                <label style={{ display: "block", fontWeight: 600, marginBottom: 6 }}>Publishing Contact</label>
                <PeoplePicker
                  context={context}
                  personSelectionLimit={1}
                  principalTypes={[PrincipalType.User]}
                  ensureUser={true}
                  showHiddenInUI={false}
                  resolveDelay={300}
                  defaultSelectedUsers={publishingContactEmail ? [publishingContactEmail] : []}
                  onChange={(items: IPeoplePickerUserItem[]) => {
                    const first = items?.[0];
                    setPublishingContactId(first?.id ? Number(first.id) : null);
                    setPublishingContactEmail(first?.secondaryText);
                  }}
                  webAbsoluteUrl={siteUrl || context.pageContext.web.absoluteUrl}
                  key={`peoplePicker_${editingItem?.Id || 'new'}_${publishingContactEmail}`}
                />
              </div>
            </Stack.Item>

            {/* NEW: Publishing Date (ArticleDate) */}
            <Stack.Item grow={1}>
              <DatePicker
                label="Publishing Date"
                placeholder="Select a date"
                allowTextInput={false}
                value={articleDate ?? undefined}
                onSelectDate={(d) => setArticleDate(d ?? null)}
              />
            </Stack.Item>

            <Stack.Item grow={1}>
              <Dropdown
                label="Department"
                placeholder="Select department"
                options={departmentOptions}
                selectedKey={department ?? "__NONE__"}
                onChange={(_, opt) => {
                  const key = (opt?.key as string) || "__NONE__";
                  setDepartment(key === "__NONE__" ? undefined : (opt!.key as string));
                }}
              />
            </Stack.Item>

            <Stack.Item grow={1}>
              <Dropdown
                label="Publishing Source"
                placeholder="Select publishing source"
                options={publishingSourceOptions}
                selectedKey={publishingSource}
                onChange={(_, opt) => setPublishingSource((opt?.key as string) || undefined)}
                required
              />
            </Stack.Item>
          </Stack>

          <TextField
            label="Abstract"
            multiline
            rows={4}
            value={abstractText}
            onChange={(_, v) => setAbstractText(v || "")}
          />

          <Stack tokens={{ childrenGap: 8 }}>
            <label style={{ fontWeight: 600 }}>Preview Image</label>
            {rollupImageUrl ? (
              <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                <img src={rollupImageUrl} alt="Rollup" style={{ width: 120, height: 80, objectFit: "cover", borderRadius: 6 }} />
                <TextField
                  label="Image URL"
                  value={rollupImageUrl}
                  onChange={(_, v) => setRollupImageUrl(v || "")}
                  styles={{ root: { width: 420 } }}
                />
                <DefaultButton text="Clear" onClick={() => setRollupImageUrl("")} />
              </div>
            ) : (
              <TextField
                label="Image URL"
                placeholder="https://…"
                value={rollupImageUrl}
                onChange={(_, v) => setRollupImageUrl(v || "")}
                styles={{ root: { width: 420 } }}
              />
            )}
            <DefaultButton
              text="Select Preview Image"
              iconProps={{ iconName: "Photo2Fill" }}
              onClick={() => { setImagePickerTarget("rollup"); setIsImagePickerOpen(true); }}
            />
          </Stack>

          {/* In Case You Missed - Full Width */}
          <Stack tokens={{ childrenGap: 6 }}>
            <label style={{ fontWeight: 600 }}>In Case You Missed (pick up to 4)</label>
            <Dropdown
              placeholder="Select items"
              multiSelect
              selectedKeys={icymSelected}
              options={(editingItem
                ? (icymOptions as IDropdownOption[]).filter(o => Number(o.key) !== editingItem.Id)
                : (icymOptions as IDropdownOption[])
              )}
              onChange={(_, option) => {
                if (!option) return;
                const id = Number(option.key);
                const isSelected = icymSelected.includes(id);
                if (isSelected) {
                  setIcymSelected(icymSelected.filter(k => k !== id));
                } else {
                  if (icymSelected.length >= 4) {
                    alert("You can select at most 4 items in 'In Case You Missed'.");
                    return;
                  }
                  setIcymSelected([...icymSelected, id]);
                }
              }}
            />
          </Stack>

          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton text="Save" onClick={saveItem} />
            <DefaultButton text="Cancel" onClick={() => setIsPanelOpen(false)} />
          </Stack>
        </Stack>
      </Panel>

      {/* Image Picker Panel */}
      <Panel
        isOpen={isImagePickerOpen}
        onDismiss={() => setIsImagePickerOpen(false)}
        type={PanelType.large}
        isBlocking={false}
        headerText={`Select or upload image (${imagesLibraryTitle})`}
        closeButtonAriaLabel="Close"
      >
        <Stack tokens={{ childrenGap: 12 }}>
          {pickerError && (
            <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setPickerError(null)}>
              {pickerError}
            </MessageBar>
          )}

          <Breadcrumb items={folderTrail} maxDisplayedItems={8} ariaLabel="Breadcrumb" overflowAriaLabel="More" />

          <Stack horizontal tokens={{ childrenGap: 12 }}>
            <input
              type="file"
              accept="image/*"
              onChange={(e) => handleUpload(e.target.files)}
              disabled={uploading}
            />
            {uploading && <span>Uploading…</span>}
          </Stack>

          <div style={{ display: "grid", gridTemplateColumns: "240px 1fr", gap: 16 }}>
            <div>
              <h4 style={{ marginTop: 0 }}>Folders</h4>
              <ul>
                {folderSubfolders.map((f) => (
                  <li key={f.Name}>
                    <a
                      href="#"
                      onClick={(e) => {
                        e.preventDefault();
                        setCurrentFolder(f.ServerRelativeUrl);
                      }}
                    >
                      📁 {f.Name}
                    </a>
                  </li>
                ))}
                {folderSubfolders.length === 0 && <li>No subfolders</li>}
              </ul>
            </div>
            <div>
              <h4 style={{ marginTop: 0 }}>Images</h4>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(140px, 1fr))", gap: 12 }}>
                {folderFiles
                  .filter((file: any) =>
                    imageExtensions.has((file.Name || "").toLowerCase().replace(/^.*(\.[^.]+)$/, "$1"))
                  )
                  .map((file: any) => (
                    <div key={file.UniqueId} style={{ border: "1px solid #eee", borderRadius: 8, padding: 8 }}>
                      <div style={{ aspectRatio: "1/1", overflow: "hidden", borderRadius: 6, marginBottom: 6 }}>
                        <img src={file.ServerRelativeUrl} alt={file.Name} style={{ width: "100%", height: "100%", objectFit: "cover" }} />
                      </div>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                        <span title={file.Name} style={{ fontSize: 12, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: 100 }}>{file.Name}</span>
                        <DefaultButton
                          text="Use"
                          onClick={() => {
                            if (imagePickerTarget === "editor") {
                              insertImageAtCursor(file.ServerRelativeUrl, file.Name);
                            } else {
                              setRollupImageUrl(file.ServerRelativeUrl);
                            }
                            setIsImagePickerOpen(false);
                          }}
                        />
                      </div>
                    </div>
                  ))}
                {folderFiles.length === 0 && <div>No files</div>}
              </div>
            </div>
          </div>
        </Stack>
      </Panel>

      {/* Delete confirm */}
      <Dialog
        hidden={!deleteTarget}
        dialogContentProps={{ type: DialogType.normal, title: "Delete item", subText: deleteTarget ? `Are you sure you want to delete "${deleteTarget.Title}"?` : "" }}
        onDismiss={() => setDeleteTarget(null)}
      >
        <DialogFooter>
          <PrimaryButton text="Delete" onClick={doDelete} />
          <DefaultButton text="Cancel" onClick={() => setDeleteTarget(null)} />
        </DialogFooter>
      </Dialog>
    </div>
  );
};

export default NewsManager;

/** Utilities */
function parseError(e: any): string {
  if (!e) return "Unknown error";
  if (typeof e === "string") return e;
  if (e?.message) return e.message;
  try { return JSON.stringify(e); } catch { return String(e); }
}
