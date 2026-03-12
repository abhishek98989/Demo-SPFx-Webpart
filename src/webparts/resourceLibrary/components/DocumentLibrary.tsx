import * as React from 'react';
import { useState, useEffect, useCallback, useMemo } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
  SearchBox,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Stack,
  Text,
  CommandBar,
  ICommandBarItemProps,
  TextField,
  PrimaryButton,
  DefaultButton,
  IconButton,
  TooltipHost,
  Icon,
  DatePicker,
  Panel,
  PanelType,
} from '@fluentui/react';

import {
  PeoplePicker,
  PrincipalType,
  IPeoplePickerUserItem,
} from '@pnp/spfx-controls-react/lib/PeoplePicker';

// Interface for document item from SharePoint
interface IDocumentItem {
  Id: number;
  Title: string;
  FileLeafRef: string;
  FSObjType: number;
  FileRef: string;
  Training_x0020_Category?: string;

  // ✅ Add modified
  Modified?: string; // ISO string from SharePoint
}

// Props interface for the component
interface IDocumentLibraryProps {
  context: WebPartContext;
}

/** ✅ Power Automate HTTP trigger URL (When an HTTP request is received) */
const POWER_AUTOMATE_HTTP_URL =
  'https://defaulte0ccdb813c814947a03335cadd72e8.38.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/4ff49d38e22d48809f02abaf6fbd2ffb/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=GOQTTg-WUd0UStPZ7UccU6wd98k-zqLX3UYYP2ZdaYY';

/** Default report recipients */
const DEFAULT_REPORT_RECIPIENTS = [
  'msewell@vaughnconstruction.com',
  'JThomas@vaughnconstruction.com',
  'kjung@vaughnconstruction.com',
];

const DocumentLibrary: React.FC<IDocumentLibraryProps> = ({ context }) => {
  // State management
  const [items, setItems] = useState<IDocumentItem[]>([]);
  const [filteredItems, setFilteredItems] = useState<IDocumentItem[]>([]);
  const [currentPageItems, setCurrentPageItems] = useState<IDocumentItem[]>([]);
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<string>('');
  const [searchText, setSearchText] = useState<string>('');
  const [currentPage, setCurrentPage] = useState<number>(1);
  const [pageInput, setPageInput] = useState<string>('1');
  const [copiedId, setCopiedId] = useState<number | null>(null);
  const [selectedAlphabet, setSelectedAlphabet] = useState<string>('');
  const [columnsState, setColumnsState] = useState<IColumn[]>([]);
  const [sortState, setSortState] = useState<{ key: string | null; isSortedDescending: boolean }>({
    key: 'title',
    isSortedDescending: false,
  });

  /** ✅ Advanced filter: Modified date range */
  const [startDate, setStartDate] = useState<Date | null>(null);
  const [endDate, setEndDate] = useState<Date | null>(null);

  /** ✅ Share report panel + recipients */
  const [isSharePanelOpen, setIsSharePanelOpen] = useState<boolean>(false);
  const [reportRecipients, setReportRecipients] = useState<string[]>(DEFAULT_REPORT_RECIPIENTS);
  const [sendingReport, setSendingReport] = useState<boolean>(false);
  const [reportStatus, setReportStatus] = useState<string>('');

  /** ✅ Graph-powered search state */
  const [graphItems, setGraphItems] = useState<IDocumentItem[]>([]);
  const [graphSearching, setGraphSearching] = useState<boolean>(false);
  const [graphError, setGraphError] = useState<string>('');

  const itemsPerPage = 20;
  const alphabetLetters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');

  /** ✅ List ID (Resource Library) */
  const RESOURCE_LIBRARY_LIST_ID = '8177736d-9faa-49cc-82b3-4ff5a93fa02e';

  // Get file extension from filename
  const getFileExtension = (fileName: string): string => {
    if (!fileName) return 'Unknown';
    const parts = fileName.split('.');
    if (parts.length > 1) {
      return parts[parts.length - 1].toUpperCase();
    }
    return 'Unknown';
  };

  // Get file icon based on extension
  const getFileIcon = (fileName: string): string => {
    const ext = getFileExtension(fileName).toLowerCase();
    const iconMap: { [key: string]: string } = {
      pdf: 'PDF',
      doc: 'WordDocument',
      docx: 'WordDocument',
      xls: 'ExcelDocument',
      xlsx: 'ExcelDocument',
      ppt: 'PowerPointDocument',
      pptx: 'PowerPointDocument',
      txt: 'TextDocument',
      zip: 'ZipFolder',
      jpg: 'FileImage',
      jpeg: 'FileImage',
      png: 'FileImage',
      gif: 'FileImage',
      mp4: 'Video',
      avi: 'Video',
      mov: 'Video',
    };
    return iconMap[ext] || 'Page';
  };

  // Build open URL (adds ?web=1 for office docs)
  const buildOpenUrl = (fileRef: string, fileName: string) => {
    const ext = getFileExtension(fileName).toLowerCase();
    const downloadExtensions = ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'mp4', 'avi', 'mov', 'wmv', 'zip', 'rar', '7z'];

    let url = `${window.location.origin}${fileRef}`;
    if (!downloadExtensions.includes(ext)) {
      url += '?web=1';
    }
    return url;
  };

  // Copy URL to clipboard
  const copyToClipboard = async (fileRef: string, fileName: string, itemId: number) => {
    try {
      const fullUrl = buildOpenUrl(fileRef, fileName);
      await navigator.clipboard.writeText(fullUrl);
      setCopiedId(itemId);
      setTimeout(() => setCopiedId(null), 2000);
    } catch (err) {
      // eslint-disable-next-line no-console
      console.error('Failed to copy URL:', err);
    }
  };

  // Open document
  const openDocument = (fileRef: string, fileName: string) => {
    const url = buildOpenUrl(fileRef, fileName);
    window.open(url, '_blank');
  };

  /** Convert a full webUrl -> server relative file ref */
  const webUrlToFileRef = (webUrl?: string): string => {
    if (!webUrl) return '';
    try {
      const u = new URL(webUrl);
      return u.pathname; // "/sites/.../Shared%20Documents/.."
    } catch {
      return webUrl.startsWith('/') ? webUrl : '';
    }
  };

  // Fetch items from SharePoint document library (browse / no-search mode)
  const fetchItems = useCallback(async () => {
    setLoading(true);
    setError('');

    try {
      const sp = spfi('https://vaughnconstruction.sharepoint.com').using(SPFx(context));

      const result = await sp.web.lists
        .getById(RESOURCE_LIBRARY_LIST_ID)
        .items.select('Id', 'Title', 'FileLeafRef', 'FSObjType', 'FileRef', 'Training_x0020_Category', 'Modified')
        .top(5000)();

      // Filter out folders
      const documentsOnly = result.filter((item: IDocumentItem) => item.FSObjType === 0);

      // eslint-disable-next-line no-console
      console.log(`Fetched ${documentsOnly.length} documents from library`);
      setItems(documentsOnly);
    } catch (err) {
      setError('Failed to fetch documents from SharePoint');
      // eslint-disable-next-line no-console
      console.error('Error fetching documents:', err);
    } finally {
      setLoading(false);
    }
  }, [context]);

  // Fetch items on component mount
  useEffect(() => {
    fetchItems();
  }, [fetchItems]);

  /** ✅ Graph Search (accurate search) */
  const runGraphSearch = useCallback(
    async (queryText: string) => {
      const trimmed = (queryText || '').trim();
      if (!trimmed) {
        setGraphItems([]);
        setGraphError('');
        return;
      }

      setGraphSearching(true);
      setGraphError('');

      try {
        const client = await context.msGraphClientFactory.getClient('3');

        // Restrict to this library by ListId, and only documents (not folders)
        const kql = `${trimmed} AND (ListId:${RESOURCE_LIBRARY_LIST_ID} OR listid:${RESOURCE_LIBRARY_LIST_ID}) AND (IsDocument:1 OR isdocument:1)`;

        const body = {
          requests: [
            {
              entityTypes: ['driveItem'],
              query: { queryString: kql },
              from: 0,
              size: 200,
              fields: ['title', 'name', 'webUrl', 'lastModifiedDateTime', 'filetype', 'path'],
            },
          ],
        };

        const graphRes: any = await client.api('/search/query').version('v1.0').post(body);

        const hits: any[] = graphRes?.value?.[0]?.hitsContainers?.[0]?.hits || [];

        // Graph -> prelim
        const prelim: IDocumentItem[] = hits
          .map((h) => h?.resource)
          .filter(Boolean)
          .map((r: any) => {
            const name = r?.name || '';
            const webUrl = r?.webUrl || '';
            const fileRef = webUrlToFileRef(webUrl);

            const listItemId =
              Number(r?.sharepointIds?.listItemId) ||
              Number(r?.sharePointIds?.listItemId) ||
              0;

            return {
              Id: listItemId || 0,
              Title: r?.name || r?.title || '',
              FileLeafRef: name || r?.title || '',
              FSObjType: 0,
              FileRef: fileRef,
              Modified: r?.lastModifiedDateTime || undefined,
              Training_x0020_Category: undefined,
            };
          })
          // keep only items that look like files
          .filter((x) => !!x.FileLeafRef && x.FileLeafRef.includes('.'));

        // Enrich from SP list so we get Training Category + correct Title/FileRef etc.
        const ids = prelim.map((p) => p.Id).filter((id) => !!id);

        if (ids.length) {
          const sp = spfi('https://vaughnconstruction.sharepoint.com').using(SPFx(context));

          // "Id eq 1 or Id eq 2 or ..."
          const filter = ids.slice(0, 200).map((id) => `Id eq ${id}`).join(' or ');

          const details: any[] = await sp.web.lists
            .getById(RESOURCE_LIBRARY_LIST_ID)
            .items.select('Id', 'Title', 'FileLeafRef', 'FSObjType', 'FileRef', 'Training_x0020_Category', 'Modified')
            .filter(filter)();

          const byId = new Map<number, any>(details.map((d) => [d.Id, d]));

          const merged: IDocumentItem[] = prelim.map((p) => {
            const d = byId.get(p.Id);
            if (!d) return p;

            return {
              Id: d.Id,
              Title: d.Title || p.Title || d.FileLeafRef,
              FileLeafRef: d.FileLeafRef || p.FileLeafRef,
              FSObjType: d.FSObjType ?? 0,
              FileRef: d.FileRef || p.FileRef,
              Training_x0020_Category: d.Training_x0020_Category,
              Modified: d.Modified || p.Modified,
            };
          });

          setGraphItems(merged.filter((x) => x.FSObjType === 0));
        } else {
          // If Graph didn't provide listItemId, show basic results
          setGraphItems(prelim);
        }
      } catch (e: any) {
        // eslint-disable-next-line no-console
        console.error('Graph search failed', e);
        setGraphError(e?.message || String(e));
        setGraphItems([]);
      } finally {
        setGraphSearching(false);
      }
    },
    [context, RESOURCE_LIBRARY_LIST_ID]
  );

  /** ✅ debounce Graph search while typing */
  useEffect(() => {
    const t = setTimeout(() => {
      if (searchText.trim()) {
        runGraphSearch(searchText);
      } else {
        setGraphItems([]);
        setGraphError('');
      }
    }, 450);

    return () => clearTimeout(t);
  }, [searchText, runGraphSearch]);

  // Search function (still used for browse mode; Graph already did server-side search)
  const searchInData = (item: IDocumentItem, searchTerm: string): boolean => {
    if (!searchTerm.trim()) return true;

    const term = searchTerm.toLowerCase().trim();

    const searchFields = [
      item.Title,
      item.FileLeafRef,
      item.Training_x0020_Category,
      getFileExtension(item.FileLeafRef),
    ];

    return searchFields.some((field) => field?.toLowerCase().includes(term));
  };

  // Handle search
  const handleSearch = (newValue?: string) => {
    const value = newValue || '';
    setSearchText(value);
    setCurrentPage(1);
    setPageInput('1');
  };

  // Handle alphabet filter
  const handleAlphabetFilter = (letter: string) => {
    if (selectedAlphabet === letter) {
      setSelectedAlphabet('');
    } else {
      setSelectedAlphabet(letter);
    }
    setCurrentPage(1);
    setPageInput('1');
  };

  // Helper to get sortable value from item
  const getSortValue = (item: any, fieldName?: string) => {
    if (!fieldName) return '';

    if (fieldName === 'title') {
      return (item.Title || item.FileLeafRef || '').toLowerCase();
    }
    if (fieldName === 'Training_x0020_Category') {
      return (item.Training_x0020_Category || '').toLowerCase();
    }
    if (fieldName === 'documentType') {
      return getFileExtension(item.FileLeafRef).toLowerCase();
    }
    if (fieldName === 'Modified') {
      return item.Modified ? new Date(item.Modified).getTime() : 0;
    }

    const value = (item as any)[fieldName];
    if (value == null) return '';
    if (Array.isArray(value)) return value.length > 0 ? String(value[0]) : '';
    return String(value).toLowerCase();
  };

  /** ✅ Date range match (inclusive) */
  const isWithinModifiedRange = (item: IDocumentItem) => {
    if (!startDate || !endDate) return true;
    if (!item.Modified) return false;

    const mod = new Date(item.Modified);

    const start = new Date(startDate);
    start.setHours(0, 0, 0, 0);

    const end = new Date(endDate);
    end.setHours(23, 59, 59, 999);

    return mod >= start && mod <= end;
  };

  // Memoized filtered items
  const processedItems = useMemo(() => {
    // ✅ Use Graph results ONLY when user is searching
    const sourceItems = searchText.trim() ? graphItems : items;

    let result = [...sourceItems];

    // Determine active sort column
    const activeCol = sortState.key ? columnsState.find((c) => c.key === sortState.key) : null;
    const fieldToSort = activeCol?.fieldName || 'title';

    // Sort
    result = result.sort((a, b) => {
      const valA: any = getSortValue(a, fieldToSort);
      const valB: any = getSortValue(b, fieldToSort);

      if (valA === valB) return 0;
      if (sortState.isSortedDescending) {
        return valB > valA ? 1 : -1;
      } else {
        return valA > valB ? 1 : -1;
      }
    });

    // Apply alphabet filter
    if (selectedAlphabet) {
      result = result.filter((item) => {
        const name = item.Title || item.FileLeafRef || '';
        return name.toUpperCase().startsWith(selectedAlphabet);
      });
    }

    // Apply search filter only in browse mode (Graph already did the searching)
    if (searchText.trim() && !searchText.trim()) {
      result = result.filter((item) => searchInData(item, searchText));
    }

    // ✅ Apply Modified date range filter (only when both dates selected)
    if (startDate && endDate) {
      result = result.filter((item) => isWithinModifiedRange(item));
    }

    return result;
  }, [items, graphItems, searchText, selectedAlphabet, sortState, columnsState, startDate, endDate]);

  // Update filtered items
  useEffect(() => {
    setFilteredItems(processedItems);
    const totalPagesLocal = Math.ceil(processedItems.length / itemsPerPage);
    if (currentPage > totalPagesLocal && totalPagesLocal > 0) {
      setCurrentPage(1);
      setPageInput('1');
    }
  }, [processedItems, currentPage, itemsPerPage]);

  // Calculate pagination
  const totalPages = Math.ceil(filteredItems.length / itemsPerPage);

  // Update current page items
  useEffect(() => {
    const startIndex = (currentPage - 1) * itemsPerPage;
    const endIndex = startIndex + itemsPerPage;
    const pageItems = filteredItems.slice(startIndex, endIndex);
    setCurrentPageItems(pageItems);
  }, [filteredItems, currentPage, itemsPerPage]);

  // Handle page change
  const handlePageChange = (page: number) => {
    if (page >= 1 && page <= totalPages) {
      setCurrentPage(page);
      setPageInput(page.toString());
    }
  };

  // Handle page input change
  const handlePageInputChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
    setPageInput(newValue || '');
  };

  // Handle page input submit
  const handlePageInputSubmit = () => {
    const page = parseInt(pageInput, 10);
    if (!isNaN(page) && page >= 1 && page <= totalPages) {
      handlePageChange(page);
    } else {
      setPageInput(currentPage.toString());
    }
  };

  // Columns
  const columns: IColumn[] = [
    {
      key: 'icon',
      name: '',
      fieldName: 'icon',
      minWidth: 40,
      maxWidth: 40,
      onRender: (item: IDocumentItem) => (
        <Icon iconName={getFileIcon(item.FileLeafRef)} styles={{ root: { fontSize: '20px', color: '#0078d4' } }} />
      ),
    },
    {
      key: 'title',
      name: 'Title',
      fieldName: 'title',
      minWidth: 200,
      maxWidth: 300,
      isResizable: true,
      isSorted: true,
      isSortedDescending: false,
      onRender: (item: IDocumentItem) => (
        <Stack>
          <Text variant="medium" styles={{ root: { fontWeight: '600' } }}>
            {item.Title || item.FileLeafRef}
          </Text>
          {item.Title && item.Title !== item.FileLeafRef && (
            <Text variant="small" styles={{ root: { color: '#666' } }}>
              {item.FileLeafRef}
            </Text>
          )}
        </Stack>
      ),
    },
    {
      key: 'category',
      name: 'Training Category',
      fieldName: 'Training_x0020_Category',
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
      isSorted: false,
      isSortedDescending: false,
      onRender: (item: IDocumentItem) => <Text variant="medium">{item.Training_x0020_Category || '-'}</Text>,
    },
    {
      key: 'documentType',
      name: 'Document Type',
      fieldName: 'documentType',
      minWidth: 100,
      maxWidth: 120,
      isResizable: true,
      isSorted: false,
      isSortedDescending: false,
      onRender: (item: IDocumentItem) => (
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
          <div
            style={{
              padding: '4px 12px',
              background: '#f3f2f1',
              borderRadius: '4px',
              fontSize: '12px',
              fontWeight: '600',
              color: '#323130',
            }}
          >
            {getFileExtension(item.FileLeafRef)}
          </div>
        </Stack>
      ),
    },
    {
      key: 'modified',
      name: 'Modified',
      fieldName: 'Modified',
      minWidth: 140,
      maxWidth: 160,
      isResizable: true,
      isSorted: false,
      isSortedDescending: false,
      onRender: (item: IDocumentItem) => {
        const d = item.Modified ? new Date(item.Modified) : null;
        return <Text variant="medium">{d ? d.toLocaleDateString() : '-'}</Text>;
      },
    },
    {
      key: 'actions',
      name: 'Actions',
      fieldName: 'actions',
      minWidth: 80,
      maxWidth: 80,
      onRender: (item: IDocumentItem) => (
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <TooltipHost content={copiedId === item.Id ? 'Copied!' : 'Copy URL'}>
            <IconButton
              iconProps={{ iconName: copiedId === item.Id ? 'CheckMark' : 'Copy' }}
              onClick={() => copyToClipboard(item.FileRef, item.FileLeafRef, item.Id)}
              styles={{ root: { color: copiedId === item.Id ? '#107c10' : '#0078d4' } }}
            />
          </TooltipHost>
          <TooltipHost content="Open Document">
            <IconButton
              iconProps={{ iconName: 'OpenInNewWindow' }}
              onClick={() => openDocument(item.FileRef, item.FileLeafRef)}
              styles={{ root: { color: '#0078d4' } }}
            />
          </TooltipHost>
        </Stack>
      ),
    },
  ];

  // Column click handler for sorting
  const onColumnClick = (ev: React.MouseEvent<HTMLElement>, column?: IColumn) => {
    if (!column || column.key === 'icon' || column.key === 'actions') return;

    setColumnsState((prev) => {
      const newCols = prev.map((col) => {
        if (col.key === column.key) {
          const newDesc = col.isSorted ? !col.isSortedDescending : false;
          return { ...col, isSorted: true, isSortedDescending: newDesc };
        }
        return { ...col, isSorted: false, isSortedDescending: false };
      });

      const active = newCols.find((c) => c.isSorted);
      setSortState({
        key: active ? active.key : null,
        isSortedDescending: active ? active.isSortedDescending || false : false,
      });

      return newCols;
    });
  };

  // Initialize columns state with onColumnClick handler
  useEffect(() => {
    const initializedColumns = columns.map((c) => ({ ...c, onColumnClick }));
    setColumnsState(initializedColumns);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [copiedId, startDate, endDate]);

  /** ✅ Command bar items (Refresh + Share Report when date range is selected) */
  const canShareReport = !!startDate && !!endDate;

  const commandBarItems: ICommandBarItemProps[] = [
    {
      key: 'refresh',
      text: 'Refresh',
      iconProps: { iconName: 'Refresh' },
      onClick: () => fetchItems(),
      disabled: loading,
    },
    {
      key: 'shareReport',
      text: 'Share Report',
      iconProps: { iconName: 'Mail' },
      onClick: () => {
        setReportStatus('');
        setReportRecipients(DEFAULT_REPORT_RECIPIENTS);
        setIsSharePanelOpen(true);
      },
      disabled: loading || !canShareReport,
    },
  ];

  /** Build HTML table for the report (Title clickable link) */
  const buildReportHtmlTable = (reportItems: IDocumentItem[]) => {
    const formatDate = (d?: string | Date) => {
      if (!d) return '';
      const date = new Date(d);
      const mm = String(date.getMonth() + 1).padStart(2, '0');
      const dd = String(date.getDate()).padStart(2, '0');
      const yyyy = date.getFullYear();
      return `${mm}/${dd}/${yyyy}`;
    };

    const rows = reportItems
      .map((it) => {
        const title = (it.Title || it.FileLeafRef || '').replace(/</g, '&lt;').replace(/>/g, '&gt;');

        const url = buildOpenUrl(it.FileRef, it.FileLeafRef);
        const modified = it.Modified ? formatDate(it.Modified) : '';
        const category = (it.Training_x0020_Category || '').replace(/</g, '&lt;').replace(/>/g, '&gt;');
        const docType = getFileExtension(it.FileLeafRef);

        return `
        <tr>
          <td style="padding:8px;border:1px solid #ddd;">
            <a href="${url}" target="_blank" rel="noreferrer">${title}</a>
          </td>
          <td style="padding:8px;border:1px solid #ddd;">${docType}</td>
          <td style="padding:8px;border:1px solid #ddd;">${category || '-'}</td>
          <td style="padding:8px;border:1px solid #ddd;">${modified || '-'}</td>
        </tr>
      `;
      })
      .join('');

    return `
    <div style="font-family:Segoe UI, Arial, sans-serif;">
      <h3 style="margin:0 0 8px 0;">Resource Library – Modified Documents Report</h3>
      <div style="margin:0 0 12px 0;color:#444;">
        Date range:
        <b>${startDate ? formatDate(startDate) : ''}</b>
        to
        <b>${endDate ? formatDate(endDate) : ''}</b><br/>
        Total items: <b>${reportItems.length}</b>
      </div>
      <table style="border-collapse:collapse;width:100%;font-size:13px;">
        <thead>
          <tr>
            <th style="text-align:left;padding:8px;border:1px solid #ddd;background:#f5f5f5;">Title</th>
            <th style="text-align:left;padding:8px;border:1px solid #ddd;background:#f5f5f5;">Type</th>
            <th style="text-align:left;padding:8px;border:1px solid #ddd;background:#f5f5f5;">Category</th>
            <th style="text-align:left;padding:8px;border:1px solid #ddd;background:#f5f5f5;">Modified</th>
          </tr>
        </thead>
        <tbody>${rows || ''}</tbody>
      </table>
    </div>
  `;
  };

  const peoplePickerKey = `reportPicker_${startDate?.toDateString() || 'na'}_${endDate?.toDateString() || 'na'}_${isSharePanelOpen}`;

  /** Call Power Automate HTTP trigger */
  const sendReport = async () => {
    if (!startDate || !endDate) return;

    if (!POWER_AUTOMATE_HTTP_URL || POWER_AUTOMATE_HTTP_URL.includes('PASTE_YOUR_FLOW_HTTP_POST_URL_HERE')) {
      setReportStatus('❌ Please update POWER_AUTOMATE_HTTP_URL in DocumentLibrary.tsx.');
      return;
    }

    try {
      setSendingReport(true);
      setReportStatus('');

      const payload = {
        recipients: reportRecipients,
        subject: `Resource Library: Modified documents (${startDate.toLocaleDateString()} – ${endDate.toLocaleDateString()})`,
        reportHtml: buildReportHtmlTable([...filteredItems]),
      };

      const res = await fetch(POWER_AUTOMATE_HTTP_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });

      if (!res.ok) {
        const t = await res.text();
        throw new Error(`Flow returned ${res.status}. ${t}`);
      }

      setReportStatus(`✅ Report sent to: ${reportRecipients.join(', ')}`);
      setIsSharePanelOpen(false);
    } catch (e: any) {
      // eslint-disable-next-line no-console
      console.error(e);
      setReportStatus(`❌ Failed to send report. ${e?.message || e}`);
    } finally {
      setSendingReport(false);
    }
  };

  return (
    <div>
      <Stack tokens={{ childrenGap: 20 }}>
        {/* Header */}
        <Stack
          horizontal
          horizontalAlign="space-between"
          verticalAlign="center"
          styles={{ root: { cursor: 'pointer' } }}
          onClick={() => {
            window.location.href = 'https://vaughnconstruction.sharepoint.com/Forms/Forms/AllItems.aspx';
          }}
        >
          <Text variant="xxLarge" styles={{ root: { fontWeight: '600' } }}>
            Resource Library
          </Text>
          <CommandBar items={commandBarItems} />
        </Stack>

        {/* Search Box */}
        <Stack>
          <SearchBox
            placeholder="Search documents by title, category, or file type..."
            value={searchText}
            onChange={(_, newValue) => handleSearch(newValue)}
            styles={{ root: { maxWidth: '500px' } }}
          />
        </Stack>

        {/* Graph search feedback */}
        {graphSearching && (
          <Stack horizontalAlign="center" styles={{ root: { padding: '10px 0' } }}>
            <Spinner size={SpinnerSize.small} label="Searching (Graph)..." />
          </Stack>
        )}
        {!graphSearching && graphError && (
          <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
            Graph search failed: {graphError}
          </MessageBar>
        )}

        {/* ✅ Advanced Filter: Modified date range */}
        <Stack
          horizontal
          wrap
          tokens={{ childrenGap: 16 }}
          verticalAlign="end"
          styles={{ root: { padding: 12, border: '1px solid #edebe9', borderRadius: 6, background: '#faf9f8' } }}
        >
          <Stack styles={{ root: { minWidth: 220 } }}>
            <Text variant="medium" styles={{ root: { fontWeight: 600, marginBottom: 6 } }}>
              Modified Start Date
            </Text>
            <DatePicker
              placeholder="Select start date"
              value={startDate || undefined}
              onSelectDate={(d) => {
                setStartDate(d || null);
                setCurrentPage(1);
                setPageInput('1');
              }}
            />
          </Stack>

          <Stack styles={{ root: { minWidth: 220 } }}>
            <Text variant="medium" styles={{ root: { fontWeight: 600, marginBottom: 6 } }}>
              Modified End Date
            </Text>
            <DatePicker
              placeholder="Select end date"
              value={endDate || undefined}
              onSelectDate={(d) => {
                setEndDate(d || null);
                setCurrentPage(1);
                setPageInput('1');
              }}
            />
          </Stack>

          <Stack horizontal tokens={{ childrenGap: 10 }} styles={{ root: { paddingTop: 22 } }}>
            <DefaultButton
              text="Clear Dates"
              onClick={() => {
                setStartDate(null);
                setEndDate(null);
                setCurrentPage(1);
                setPageInput('1');
              }}
              disabled={!startDate && !endDate}
            />
            <PrimaryButton
              text="Share Report"
              iconProps={{ iconName: 'Mail' }}
              disabled={!canShareReport || loading}
              onClick={() => {
                setReportStatus('');
                setReportRecipients(DEFAULT_REPORT_RECIPIENTS);
                setIsSharePanelOpen(true);
              }}
            />
          </Stack>
        </Stack>

        {/* Alphabet Filter */}
        <Stack>
          <Text variant="medium" styles={{ root: { marginBottom: '8px', fontWeight: '600' } }}>
            Filter by Name:
          </Text>
          <Stack horizontal wrap tokens={{ childrenGap: 4 }}>
            <button
              style={{
                padding: '4px 8px',
                margin: '2px',
                background: selectedAlphabet === '' ? '#0078d4' : '#fff',
                color: selectedAlphabet === '' ? '#fff' : '#0078d4',
                border: '1px solid #0078d4',
                borderRadius: '4px',
                cursor: 'pointer',
                fontSize: '12px',
              }}
              onClick={() => {
                setSelectedAlphabet('');
                setCurrentPage(1);
                setPageInput('1');
              }}
            >
              All
            </button>
            {alphabetLetters.map((letter) => (
              <button
                key={letter}
                style={{
                  padding: '4px 8px',
                  margin: '2px',
                  background: selectedAlphabet === letter ? '#0078d4' : '#fff',
                  color: selectedAlphabet === letter ? '#fff' : '#0078d4',
                  border: '1px solid #0078d4',
                  borderRadius: '4px',
                  cursor: 'pointer',
                  fontSize: '12px',
                }}
                onClick={() => handleAlphabetFilter(letter)}
              >
                {letter}
              </button>
            ))}
          </Stack>
        </Stack>

        {/* Results Summary */}
        {!loading && (
          <Text variant="medium" styles={{ root: { color: '#666' } }}>
            Showing {currentPageItems.length} of {filteredItems.length} documents
            {(searchText || selectedAlphabet || (startDate && endDate)) && ` (filtered from ${items.length} total)`}
          </Text>
        )}

        {/* Error Message */}
        {error && (
          <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
            {error}
          </MessageBar>
        )}

        {/* Loading Spinner */}
        {loading && (
          <Stack horizontalAlign="center" styles={{ root: { padding: '40px' } }}>
            <Spinner size={SpinnerSize.large} label="Loading documents..." />
          </Stack>
        )}

        {/* Documents Table */}
        {!loading && !error && (
          <>
            <DetailsList
              items={currentPageItems}
              columns={columnsState}
              layoutMode={DetailsListLayoutMode.justified}
              selectionMode={SelectionMode.none}
              styles={{
                root: { border: '1px solid #edebe9', borderRadius: '4px' },
                headerWrapper: { backgroundColor: '#f8f9fa' },
              }}
              onShouldVirtualize={() => false}
            />

            {/* Pagination */}
            {totalPages > 1 && (
              <Stack
                horizontal
                horizontalAlign="center"
                verticalAlign="center"
                tokens={{ childrenGap: 10 }}
                styles={{ root: { marginTop: '20px' } }}
              >
                <DefaultButton
                  text="Previous"
                  disabled={currentPage === 1}
                  onClick={() => handlePageChange(currentPage - 1)}
                  iconProps={{ iconName: 'ChevronLeft' }}
                />

                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                  <Text variant="medium">Page</Text>
                  <TextField
                    value={pageInput}
                    onChange={handlePageInputChange}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') {
                        handlePageInputSubmit();
                      }
                    }}
                    styles={{ root: { width: '60px' }, fieldGroup: { height: '32px' } }}
                  />
                  <Text variant="medium">of {totalPages}</Text>
                  <PrimaryButton text="Go" onClick={handlePageInputSubmit} styles={{ root: { minWidth: '40px' } }} />
                </Stack>

                <DefaultButton
                  text="Next"
                  disabled={currentPage === totalPages}
                  onClick={() => handlePageChange(currentPage + 1)}
                  iconProps={{ iconName: 'ChevronRight' }}
                />
              </Stack>
            )}
          </>
        )}

        {/* No Results Message */}
        {!loading && !error && filteredItems.length === 0 && items.length > 0 && (
          <Stack horizontalAlign="center" styles={{ root: { padding: '40px' } }}>
            <Text variant="mediumPlus">No documents found matching your filter criteria.</Text>
          </Stack>
        )}

        {/* Empty Library Message */}
        {!loading && !error && items.length === 0 && (
          <Stack horizontalAlign="center" styles={{ root: { padding: '40px' } }}>
            <Text variant="mediumPlus">No documents available in this library.</Text>
          </Stack>
        )}

        {/* ✅ Share Panel */}
        <Panel
          isOpen={isSharePanelOpen}
          onDismiss={() => setIsSharePanelOpen(false)}
          type={PanelType.medium}
          headerText="Share Modified Documents Report"
          closeButtonAriaLabel="Close"
        >
          <Stack tokens={{ childrenGap: 14 }}>
            <Text>
              This will send a report for documents <b>Modified</b> between{' '}
              <b>{startDate ? startDate.toLocaleDateString() : '-'}</b> and{' '}
              <b>{endDate ? endDate.toLocaleDateString() : '-'}</b>.
            </Text>

            <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
              Recipients
            </Text>

            <PeoplePicker
              context={context as any}
              personSelectionLimit={10}
              principalTypes={[PrincipalType.User]}
              ensureUser={true}
              showHiddenInUI={false}
              resolveDelay={300}
              defaultSelectedUsers={reportRecipients?.length ? reportRecipients : DEFAULT_REPORT_RECIPIENTS}
              onChange={(items: IPeoplePickerUserItem[]) => {
                const mails = (items || [])
                  .map((p) => (p.secondaryText || '').trim())
                  .filter(Boolean);

                setReportRecipients(mails);
              }}
              webAbsoluteUrl={context.pageContext.web.absoluteUrl}
              key={peoplePickerKey}
            />

            {reportStatus && (
              <MessageBar messageBarType={reportStatus.startsWith('✅') ? MessageBarType.success : MessageBarType.error}>
                {reportStatus}
              </MessageBar>
            )}

            <PrimaryButton
              text={sendingReport ? 'Sending...' : 'Send Report'}
              iconProps={{ iconName: 'Send' }}
              disabled={sendingReport || !reportRecipients.length}
              onClick={sendReport}
            />
            <DefaultButton text="Cancel" onClick={() => setIsSharePanelOpen(false)} />
          </Stack>
        </Panel>
      </Stack>
    </div>
  );
};

export default DocumentLibrary;
