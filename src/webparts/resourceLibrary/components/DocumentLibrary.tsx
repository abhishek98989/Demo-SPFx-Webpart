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
} from '@fluentui/react';

// Interface for document item from SharePoint
interface IDocumentItem {
    Id: number;
    Title: string;
    FileLeafRef: string;
    FSObjType: number;
    FileRef: string;
    Training_x0020_Category?: string;
}

// Props interface for the component
interface IDocumentLibraryProps {
    context: WebPartContext;
}

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
        isSortedDescending: false
    });

    const itemsPerPage = 20;
    const alphabetLetters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');

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
            'pdf': 'PDF',
            'doc': 'WordDocument',
            'docx': 'WordDocument',
            'xls': 'ExcelDocument',
            'xlsx': 'ExcelDocument',
            'ppt': 'PowerPointDocument',
            'pptx': 'PowerPointDocument',
            'txt': 'TextDocument',
            'zip': 'ZipFolder',
            'jpg': 'FileImage',
            'jpeg': 'FileImage',
            'png': 'FileImage',
            'gif': 'FileImage',
            'mp4': 'Video',
            'avi': 'Video',
            'mov': 'Video',
        };
        return iconMap[ext] || 'Page';
    };

    // Copy URL to clipboard
    const copyToClipboard = async (fileRef: string, fileName: string, itemId: number) => {
        try {
            const ext = getFileExtension(fileName).toLowerCase();
            const downloadExtensions = ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'mp4', 'avi', 'mov', 'wmv', 'zip', 'rar', '7z'];

            let fullUrl = `${window.location.origin}${fileRef}`;
            // Add ?web=1 for Office documents to open in browser instead of download
            if (!downloadExtensions.includes(ext)) {
                fullUrl += '?web=1';
            }

            await navigator.clipboard.writeText(fullUrl);
            setCopiedId(itemId);
            setTimeout(() => setCopiedId(null), 2000);
        } catch (err) {
            console.error('Failed to copy URL:', err);
        }
    };

    // Open document with web=1 parameter for Office files
    const openDocument = (fileRef: string, fileName: string) => {
        const ext = getFileExtension(fileName).toLowerCase();
        const downloadExtensions = ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'mp4', 'avi', 'mov', 'wmv', 'zip', 'rar', '7z'];

        let url = `${window.location.origin}${fileRef}`;

        // Add ?web=1 for Office documents to open in browser instead of download
        if (!downloadExtensions.includes(ext)) {
            url += '?web=1';
        }

        window.open(url, '_blank');
    };

    // Fetch items from SharePoint document library
    const fetchItems = useCallback(async () => {
        setLoading(true);
        setError('');

        try {
            const sp = spfi('https://vaughnconstruction.sharepoint.com').using(SPFx(context));

            const result = await sp.web.lists
                .getById('8177736d-9faa-49cc-82b3-4ff5a93fa02e')
                .items
                .select('Id', 'Title', 'FileLeafRef', 'FSObjType', 'FileRef', 'Training_x0020_Category')
                .top(5000)();

            // Filter out folders (FSObjType = 1 means folder, 0 means file)
            const documentsOnly = result.filter((item: IDocumentItem) => item.FSObjType === 0);

            console.log(`Fetched ${documentsOnly.length} documents from library`);
            setItems(documentsOnly);
        } catch (err) {
            setError('Failed to fetch documents from SharePoint');
            console.error('Error fetching documents:', err);
        } finally {
            setLoading(false);
        }
    }, [context]);

    // Fetch items on component mount
    useEffect(() => {
        fetchItems();
    }, [fetchItems]);

    // Search function
    const searchInData = (item: IDocumentItem, searchTerm: string): boolean => {
        if (!searchTerm.trim()) return true;

        const term = searchTerm.toLowerCase().trim();

        const searchFields = [
            item.Title,
            item.FileLeafRef,
            item.Training_x0020_Category,
            getFileExtension(item.FileLeafRef)
        ];

        return searchFields.some(field =>
            field?.toLowerCase().includes(term)
        );
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

        const value = (item as any)[fieldName];
        if (value == null) return '';
        if (Array.isArray(value)) return value.length > 0 ? String(value[0]) : '';
        return String(value).toLowerCase();
    };

    // Memoized filtered items
    const processedItems = useMemo(() => {
        let result = [...items];

        // Determine active sort column
        const activeCol = sortState.key ? columnsState.find(c => c.key === sortState.key) : null;
        const fieldToSort = activeCol?.fieldName || 'title';

        // Sort
        result = result.sort((a, b) => {
            const valA = getSortValue(a, fieldToSort);
            const valB = getSortValue(b, fieldToSort);

            if (valA === valB) return 0;
            if (sortState.isSortedDescending) {
                return valB > valA ? 1 : -1;
            } else {
                return valA > valB ? 1 : -1;
            }
        });

        // Apply alphabet filter
        if (selectedAlphabet) {
            result = result.filter(item => {
                const name = item.Title || item.FileLeafRef || '';
                return name.toUpperCase().startsWith(selectedAlphabet);
            });
        }

        // Apply search filter
        if (searchText.trim()) {
            result = result.filter(item => searchInData(item, searchText));
        }

        return result;
    }, [items, searchText, selectedAlphabet, sortState, columnsState]);

    // Update filtered items
    useEffect(() => {
        setFilteredItems(processedItems);
        const totalPages = Math.ceil(processedItems.length / itemsPerPage);
        if (currentPage > totalPages && totalPages > 0) {
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

    // Table columns
    const columns: IColumn[] = [
        {
            key: 'icon',
            name: '',
            fieldName: 'icon',
            minWidth: 40,
            maxWidth: 40,
            onRender: (item: IDocumentItem) => (
                <Icon
                    iconName={getFileIcon(item.FileLeafRef)}
                    styles={{
                        root: {
                            fontSize: '20px',
                            color: '#0078d4'
                        }
                    }}
                />
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
            onRender: (item: IDocumentItem) => (
                <Text variant="medium">
                    {item.Training_x0020_Category || '-'}
                </Text>
            ),
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
                            color: '#323130'
                        }}
                    >
                        {getFileExtension(item.FileLeafRef)}
                    </div>
                </Stack>
            ),
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
                            iconProps={{
                                iconName: copiedId === item.Id ? 'CheckMark' : 'Copy'
                            }}
                            onClick={() => copyToClipboard(item.FileRef, item.FileLeafRef, item.Id)}
                            styles={{
                                root: {
                                    color: copiedId === item.Id ? '#107c10' : '#0078d4'
                                }
                            }}
                        />
                    </TooltipHost>
                    <TooltipHost content="Open Document">
                        <IconButton
                            iconProps={{ iconName: 'OpenInNewWindow' }}
                            onClick={() => openDocument(item.FileRef, item.FileLeafRef)}
                            styles={{
                                root: {
                                    color: '#0078d4'
                                }
                            }}
                        />
                    </TooltipHost>
                </Stack>
            ),
        },
    ];

    // Column click handler for sorting
    const onColumnClick = (ev: React.MouseEvent<HTMLElement>, column?: IColumn) => {
        if (!column || column.key === 'icon' || column.key === 'actions') return;

        setColumnsState(prev => {
            const newCols = prev.map(col => {
                if (col.key === column.key) {
                    const newDesc = col.isSorted ? !col.isSortedDescending : false;
                    return { ...col, isSorted: true, isSortedDescending: newDesc };
                }
                return { ...col, isSorted: false, isSortedDescending: false };
            });

            const active = newCols.find(c => c.isSorted);
            setSortState({
                key: active ? active.key : null,
                isSortedDescending: active ? active.isSortedDescending || false : false
            });

            return newCols;
        });
    };

    // Initialize columns state with onColumnClick handler
    useEffect(() => {
        const initializedColumns = columns.map(c => ({ ...c, onColumnClick }));
        setColumnsState(initializedColumns);
    }, [copiedId]); // Include copiedId to update when copy state changes

    // Command bar items
    const commandBarItems: ICommandBarItemProps[] = [
        {
            key: 'refresh',
            text: 'Refresh',
            iconProps: { iconName: 'Refresh' },
            onClick: () => {
                fetchItems();
            },
            disabled: loading,
        },
    ];

    return (
        <div>
            <Stack tokens={{ childrenGap: 20 }}>
                {/* Header */}
                <Stack
                    horizontal
                    horizontalAlign="space-between"
                    verticalAlign="center"
                    styles={{ root: { cursor: 'pointer' } }} // Show pointer on hover
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
                        {alphabetLetters.map(letter => (
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
                        {(searchText || selectedAlphabet) && ` (filtered from ${items.length} total)`}
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
                                root: {
                                    border: '1px solid #edebe9',
                                    borderRadius: '4px',
                                },
                                headerWrapper: {
                                    backgroundColor: '#f8f9fa',
                                },
                            }}
                            onShouldVirtualize={() => false}
                        />

                        {/* Pagination */}
                        {totalPages > 1 && (
                            <Stack horizontal horizontalAlign="center" verticalAlign="center" tokens={{ childrenGap: 10 }} styles={{ root: { marginTop: '20px' } }}>
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
                                        styles={{
                                            root: { width: '60px' },
                                            fieldGroup: { height: '32px' },
                                        }}
                                    />
                                    <Text variant="medium">of {totalPages}</Text>
                                    <PrimaryButton
                                        text="Go"
                                        onClick={handlePageInputSubmit}
                                        styles={{ root: { minWidth: '40px' } }}
                                    />
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
                        <Text variant="mediumPlus">No documents found matching your search criteria.</Text>
                    </Stack>
                )}

                {/* Empty Library Message */}
                {!loading && !error && items.length === 0 && (
                    <Stack horizontalAlign="center" styles={{ root: { padding: '40px' } }}>
                        <Text variant="mediumPlus">No documents available in this library.</Text>
                    </Stack>
                )}
            </Stack>
        </div>
    );
};

export default DocumentLibrary;