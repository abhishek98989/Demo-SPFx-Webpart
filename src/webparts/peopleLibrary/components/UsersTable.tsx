import * as React from 'react';
import { useState, useEffect, useCallback, useMemo, useRef } from 'react';
import { MSGraphClientV3, MSGraphClientFactory } from '@microsoft/sp-http';
import { ResponseType } from '@microsoft/microsoft-graph-client';
import { WebPartContext } from '@microsoft/sp-webpart-base';
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
    Persona,
    PersonaSize,
    PersonaPresence,
    Stack,
    Text,
    CommandBar,
    ICommandBarItemProps,
    TextField,
    PrimaryButton,
    DefaultButton,
    Toggle,
} from '@fluentui/react';
import { LivePersona } from '@pnp/spfx-controls-react/lib/LivePersona';
import { ServiceScope } from '@microsoft/sp-core-library';
import styles from './PeopleLibrary.module.scss';

// Interface for user data from Microsoft Graph
interface IUserData {
    id: string;
    displayName: string;
    mail: string;
    mobilePhone: string;
    businessPhones: string[];
    officeLocation: string;
    jobTitle: string;
    ipPhone: any;
    companyName: string;
    userPrincipalName: string;
    userType?: string;
    accountEnabled?: boolean;
    photo?: string;
    onPremisesDistinguishedName?: string;
}

// Props interface for the component
interface IUsersTableProps {
    context: WebPartContext;
}

// Enhanced photo cache with better state management
interface PhotoCacheEntry {
    url: string | null;
    loading: boolean;
    error: boolean;
    timestamp: number;
}

const photoCache = new Map<string, PhotoCacheEntry>();
const pendingPhotoRequests = new Map<string, Promise<PhotoCacheEntry>>();

// Photo loading queue to prevent race conditions
class PhotoLoadQueue {
    private queue: Array<{ email: string; resolve: (entry: PhotoCacheEntry) => void }> = [];
    private processing = false;
    private graphClient: MSGraphClientV3 | null = null;

    setGraphClient(client: MSGraphClientV3) {
        this.graphClient = client;
    }

    async loadPhoto(email: string): Promise<PhotoCacheEntry> {
        const normalizedEmail = email?.toLowerCase().trim();
        
        // Return cached result if available
        const cached = photoCache.get(normalizedEmail);
        if (cached && (cached.url || cached.error)) {
            return cached;
        }

        // Return pending request if exists
        const pending = pendingPhotoRequests.get(normalizedEmail);
        if (pending) {
            return pending;
        }

        // Create new request
        const promise = this.fetchPhoto(normalizedEmail);
        pendingPhotoRequests.set(normalizedEmail, promise);

        try {
            const result = await promise;
            return result;
        } finally {
            pendingPhotoRequests.delete(normalizedEmail);
        }
    }

    private async fetchPhoto(normalizedEmail: string): Promise<PhotoCacheEntry> {
        if (!this.graphClient || !normalizedEmail) {
            const errorEntry: PhotoCacheEntry = {
                url: null,
                loading: false,
                error: true,
                timestamp: Date.now()
            };
            photoCache.set(normalizedEmail, errorEntry);
            return errorEntry;
        }

        const loadingEntry: PhotoCacheEntry = {
            url: null,
            loading: true,
            error: false,
            timestamp: Date.now()
        };
        photoCache.set(normalizedEmail, loadingEntry);

        try {
            const photoResponse = await this.graphClient
                .api(`/users/${normalizedEmail}/photo/$value`)
                .responseType(ResponseType.BLOB)
                .get();

            let blob: Blob;
            if (photoResponse instanceof Blob) {
                blob = photoResponse;
            } else if (photoResponse.body instanceof Blob) {
                blob = photoResponse.body;
            } else {
                blob = new Blob([photoResponse.body || photoResponse], { type: 'image/jpeg' });
            }

            const url = URL.createObjectURL(blob);
            const successEntry: PhotoCacheEntry = {
                url,
                loading: false,
                error: false,
                timestamp: Date.now()
            };
            photoCache.set(normalizedEmail, successEntry);
            return successEntry;
        } catch (error) {
            console.log('Photo not available for user:', normalizedEmail);
            const errorEntry: PhotoCacheEntry = {
                url: null,
                loading: false,
                error: true,
                timestamp: Date.now()
            };
            photoCache.set(normalizedEmail, errorEntry);
            return errorEntry;
        }
    }
}

const photoLoadQueue = new PhotoLoadQueue();

// Memoized LivePersonaCard component to prevent unnecessary re-renders
interface ILivePersonaCardProps {
    serviceScope: any;
    userEmail: string;
    displayName: string;
    userId: string; // Add unique user ID for better tracking
}

const LivePersonaCard: React.FC<ILivePersonaCardProps> = React.memo(({
    serviceScope,
    userEmail,
    displayName,
    userId
}) => {
    const normalizedEmail = userEmail?.toLowerCase().trim();
    const [photoUrl, setPhotoUrl] = useState<string | null>(() => {
        const cached = photoCache.get(normalizedEmail);
        return cached?.url || null;
    });
    const [isLoading, setIsLoading] = useState<boolean>(false);
    const [hasError, setHasError] = useState<boolean>(false);
    const mountedRef = useRef(true);

    // Function to get initials from display name
    const getInitials = (name: string): string => {
        if (!name) return '??';
        return name.trim().split(/\s+/).map((word: string) => word?.charAt(0).toUpperCase()).join('').substring(0, 2);
    };

    // Load photo only once when component mounts
    useEffect(() => {
        mountedRef.current = true;

        const loadPhoto = async () => {
            if (!normalizedEmail) {
                setHasError(true);
                return;
            }

            // Check cache first
            const cached = photoCache.get(normalizedEmail);
            if (cached) {
                if (mountedRef.current) {
                    setPhotoUrl(cached.url);
                    setIsLoading(cached.loading);
                    setHasError(cached.error);
                }
                if (cached.url || cached.error) {
                    return; // Already loaded or failed
                }
            }

            // Load from queue
            setIsLoading(true);
            try {
                const result = await photoLoadQueue.loadPhoto(normalizedEmail);
                if (mountedRef.current) {
                    setPhotoUrl(result.url);
                    setIsLoading(false);
                    setHasError(result.error);
                }
            } catch (error) {
                if (mountedRef.current) {
                    setIsLoading(false);
                    setHasError(true);
                }
            }
        };

        loadPhoto();

        return () => {
            mountedRef.current = false;
        };
    }, [normalizedEmail]); // Only re-run if email changes

    // Memoize persona element to prevent re-creation
    const personaElement = useMemo(() => (
        <Persona
            text={displayName}
            size={PersonaSize.size40}
            hidePersonaDetails={true}
            imageUrl={photoUrl || undefined}
            imageInitials={photoUrl ? undefined : getInitials(displayName)}
            coinSize={40}
        />
    ), [displayName, photoUrl]);

    // If we have service scope and valid email, wrap with LivePersona for hover functionality
    if (serviceScope && normalizedEmail && !hasError) {
        return (
            <div style={{ position: 'relative' }}>
                {personaElement}
                <div style={{ 
                    position: 'absolute', 
                    top: 0, 
                    left: 0, 
                    width: '100%', 
                    height: '100%', 
                    opacity: 0,
                    pointerEvents: 'all'
                }}>
                    <LivePersona
                        serviceScope={serviceScope}
                        upn={normalizedEmail}
                        template={personaElement}
                    />
                </div>
            </div>
        );
    }

    return personaElement;
}, (prevProps, nextProps) => {
    // Custom comparison function - only re-render if these props change
    return (
        prevProps.userId === nextProps.userId &&
        prevProps.userEmail === nextProps.userEmail &&
        prevProps.displayName === nextProps.displayName
    );
});

LivePersonaCard.displayName = 'LivePersonaCard';

const UsersTable: React.FC<IUsersTableProps> = ({ context }) => {
    // State management
    const [users, setUsers] = useState<IUserData[]>([]);
    const [filteredUsers, setFilteredUsers] = useState<IUserData[]>([]);
    const [currentPageUsers, setCurrentPageUsers] = useState<IUserData[]>([]);
    const [loading, setLoading] = useState<boolean>(false);
    const [error, setError] = useState<string>('');
    const [searchText, setSearchText] = useState<string>('');
    const [currentPage, setCurrentPage] = useState<number>(1);
    const [pageInput, setPageInput] = useState<string>('1');
    const [selectedAlphabet, setSelectedAlphabet] = useState<string>('');
    
    const [sortedColumnKey, setSortedColumnKey] = useState<keyof IUserData | 'mobilePhone' | 'businessPhones' | 'ipPhone' | ''>('displayName');
    const [sortAscending, setSortAscending] = useState<boolean>(true);

    const [graphClient, setGraphClient] = useState<MSGraphClientV3 | null>(null);

    const itemsPerPage = 20;
    const alphabetLetters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');

    // Initialize Microsoft Graph client
    useEffect(() => {
        const initializeGraphClient = async () => {
            try {
                const client = await context.msGraphClientFactory.getClient('3');
                setGraphClient(client);
                photoLoadQueue.setGraphClient(client);
            } catch (err) {
                setError('Failed to initialize Graph client');
                console.error('Graph client initialization error:', err);
            }
        };

        initializeGraphClient();
    }, [context]);

    // Fetch all users from Microsoft Graph (recursive to get all users)
    const fetchAllUsers = async (
        graphClient: MSGraphClientV3,
        users: IUserData[] = [],
        nextLink?: string
    ): Promise<IUserData[]> => {
        try {
            let request;
           if (nextLink) {
    request = graphClient.api(nextLink.replace('https://graph.microsoft.com/v1.0', ''));
} else {
    request = graphClient
        .api('/users')
        .select([
            'id',
            'displayName',
            'givenName',
            'surname',
            'userPrincipalName',
            'mail',
            'otherMails',
            'proxyAddresses',
            'mobilePhone',
            'businessPhones',
            'jobTitle',
            'companyName',
            'department',
            'employeeId',
            'officeLocation',
            'preferredLanguage',
            'accountEnabled',
            'usageLocation',
            'createdDateTime',
            'lastPasswordChangeDateTime',
            'signInSessionsValidFromDateTime',
            'passwordPolicies',
            'onPremisesSyncEnabled',
            'onPremisesLastSyncDateTime',
            'onPremisesDistinguishedName',
            'onPremisesSamAccountName',
            'onPremisesUserPrincipalName',
            'onPremisesDomainName',
            'onPremisesImmutableId',
            'onPremisesSecurityIdentifier',
            'onPremisesExtensionAttributes'
        ])
        .top(999);
}

            const response = await request.get();
            const newUsers = response.value || [];
            const allUsers = [...users, ...newUsers];

            if (response['@odata.nextLink']) {
                return await fetchAllUsers(graphClient, allUsers, response['@odata.nextLink']);
            }

            return allUsers;
        } catch (error) {
            console.error('Error fetching users batch:', error);
            return users;
        }
    };

    // Normalize phone number for searching
    const normalizePhone = (phone: string): string => {
        if (!phone) return '';
        return phone.replace(/[\s\-\(\)\.]/g, '');
    };

    // Improved search function
    const searchInData = (user: IUserData, searchTerm: string): boolean => {
        if (!searchTerm.trim()) return true;
        
        const term = searchTerm.toLowerCase().trim();
        const normalizedSearchTerm = normalizePhone(term);
        
        const textFields = [
            user.displayName,
            user.mail,
            user.officeLocation,
            user.jobTitle,
            user.userPrincipalName
        ];
        
        const textMatch = textFields.some(field => 
            field?.toLowerCase().includes(term)
        );
        
        if (textMatch) return true;
        
        const phoneFields = [
            user.mobilePhone,
            user.ipPhone,
            ...(user.businessPhones || [])
        ];
        
        const phoneMatch = phoneFields.some(phone => {
            if (!phone) return false;
            const normalizedPhone = normalizePhone(phone.toString());
            return normalizedPhone.includes(normalizedSearchTerm) || 
                   phone.toString().toLowerCase().includes(term);
        });
        
        return phoneMatch;
    };

    // Sorting utility function
    const sortUsers = useCallback((
        data: IUserData[],
        columnKey: keyof IUserData | 'mobilePhone' | 'businessPhones' | 'ipPhone' | '',
        isAscending: boolean
    ): IUserData[] => {
        if (!columnKey) return data;

        const sortedData = [...data].sort((a, b) => {
            let aValue: string | any;
            let bValue: string | any;

            if (columnKey === 'businessPhones') {
                aValue = a.businessPhones && a.businessPhones.length > 0 ? a.businessPhones[0] : '';
                bValue = b.businessPhones && b.businessPhones.length > 0 ? b.businessPhones[0] : '';
            } else if (columnKey === 'mobilePhone') {
                aValue = a.mobilePhone || '';
                bValue = b.mobilePhone || '';
            } else if (columnKey === 'ipPhone') {
                aValue = a.ipPhone || '';
                bValue = b.ipPhone || '';
            } else {
                aValue = a[columnKey as keyof IUserData] as string || '';
                bValue = b[columnKey as keyof IUserData] as string || '';
            }

            if (typeof aValue === 'string' && typeof bValue === 'string') {
                const comparison = aValue.toLowerCase().localeCompare(bValue.toLowerCase());
                return isAscending ? comparison : -comparison;
            }

            return 0;
        });

        return sortedData;
    }, []);

    // Fetch users from Microsoft Graph
    const fetchUsers = useCallback(async () => {
        if (!graphClient) return;

        setLoading(true);
        setError('');

        try {
            // DON'T clear photo cache on refresh - keep loaded photos
            // photoCache.clear(); // REMOVED
            
            let usersData = await fetchAllUsers(graphClient);

            usersData = usersData.filter(user => {
                if (user.companyName && !user.companyName.toLowerCase().includes('vaughn')) {
                    return false;
                }
 if (user.mail?.toLowerCase() === 'kcotie@vaughnconstruction.com') {
        return false;
    }
if (
    user.onPremisesDistinguishedName &&
    user.onPremisesDistinguishedName.toLowerCase().includes('ou=~disabled')
) {
    return false;
}
                if (user.accountEnabled === false) {
                    return false;
                }

                if (user.userType === 'Guest') {
                    return false;
                }

                if (user.userPrincipalName && user.userPrincipalName.includes('#EXT#')) {
                    return false;
                }

                if ((user.mail && user.mail.includes('_')) || 
                    (user.userPrincipalName && user.userPrincipalName.includes('_')) || 
                    (user.userPrincipalName && user.userPrincipalName?.toLowerCase()?.includes('test'))) {
                    return false;
                }

                if (user?.jobTitle) {
                    const jobTitleLower = user.jobTitle.toLowerCase();
                    const excludedTitles = ['intern', 'construction work', 'layout', 'foreman', 'instrument', 'student'];
                    const excludedExactTitles = ['carpenter', 'driver', 'craft worker', 'rod man'];
                    
                    if (excludedTitles.some(title => jobTitleLower.includes(title)) ||
                        excludedExactTitles.includes(user.jobTitle.trim().toLowerCase())) {
                        return false;
                    }
                }

                if (user.userPrincipalName) {
                    const upnLower = user.userPrincipalName.toLowerCase();
                    if (upnLower.includes('service') || upnLower.includes('admin') || upnLower.includes('system')) {
                        return false;
                    }
                }

                if (!user.mail || user.mail.trim() === '') {
                    return false;
                }

                if (!user.jobTitle || user.jobTitle.trim() === '') {
                    return false;
                }

                return true;
            });

            console.log(`Fetched ${usersData.length} filtered users from Azure AD`);

            usersData = sortUsers(usersData, 'displayName', true);

            setUsers(usersData);
            setSortedColumnKey('displayName');
            setSortAscending(true);
        } catch (err) {
            setError('Failed to fetch users from Azure AD');
            console.error('Error fetching users:', err);
        } finally {
            setLoading(false);
        }
    }, [graphClient, sortUsers]);

    // Fetch users when graph client is ready
    useEffect(() => {
        if (graphClient) {
            fetchUsers();
        }
    }, [graphClient, fetchUsers]);

    // Memoized filtered and sorted users
    const processedUsers = useMemo(() => {
        let result = [...users];

        if (selectedAlphabet) {
            result = result.filter(user =>
                user.displayName?.toUpperCase().startsWith(selectedAlphabet)
            );
        }

        if (searchText.trim()) {
            result = result.filter(user => searchInData(user, searchText));
        }

        result = sortUsers(result, sortedColumnKey, sortAscending);

        return result;
    }, [users, searchText, selectedAlphabet, sortedColumnKey, sortAscending, sortUsers]);

    // Update filtered users and reset page only when filters change
    useEffect(() => {
        setFilteredUsers(processedUsers);
        const totalPages = Math.ceil(processedUsers.length / itemsPerPage);
        
        if (currentPage > totalPages && totalPages > 0 || currentPage === 0) {
            setCurrentPage(1);
            setPageInput('1');
        }
    }, [processedUsers, currentPage, itemsPerPage]);

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

    // Handle column header click for sorting
    const onColumnClick = useCallback((
        event: React.MouseEvent<HTMLElement>,
        column: IColumn
    ): void => {
        const columnKey = column.fieldName as keyof IUserData | 'mobilePhone' | 'businessPhones' | 'ipPhone' | '';
        
        let newSortAscending = sortAscending;
        
        if (sortedColumnKey === columnKey) {
            newSortAscending = !sortAscending;
        } else {
            newSortAscending = true;
        }

        setSortedColumnKey(columnKey);
        setSortAscending(newSortAscending);

        setCurrentPage(1);
        setPageInput('1');
    }, [sortedColumnKey, sortAscending]);

    // Calculate pagination
    const totalPages = Math.ceil(filteredUsers.length / itemsPerPage);

    // Update current page users when filtered users or current page changes
    useEffect(() => {
        const startIndex = (currentPage - 1) * itemsPerPage;
        const endIndex = startIndex + itemsPerPage;
        const pageUsers = filteredUsers.slice(startIndex, endIndex);

        setCurrentPageUsers(pageUsers);
    }, [filteredUsers, currentPage, itemsPerPage]);

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

    // Improved phone number display
    const formatPhoneNumber = (phone: string): string => {
        if (!phone) return '';
        const cleaned = phone.replace(/\D/g, '');
        if (cleaned.length === 10) {
            return `(${cleaned.slice(0, 3)}) ${cleaned.slice(3, 6)}-${cleaned.slice(6)}`;
        }
        return phone;
    };

    // Define table columns - memoized with stable keys
    const columns: IColumn[] = useMemo(() => [
        {
            key: 'photo',
            name: '',
            fieldName: 'photo',
            minWidth: 60,
            maxWidth: 60,
            isPadded: true,
            onRender: (item: IUserData) => {
                const userIdentifier = item.mail || item.userPrincipalName;
                return (
                    <LivePersonaCard
                        key={item.id}
                        serviceScope={context?.serviceScope}
                        userEmail={userIdentifier}
                        displayName={item.displayName}
                        userId={item.id}
                    />
                );
            },
        },
        {
            key: 'displayName',
            name: 'Name',
            fieldName: 'displayName',
            minWidth: 150,
            maxWidth: 200,
            isResizable: true,
            isSorted: sortedColumnKey === 'displayName',
            isSortedDescending: !sortAscending,
            onRender: (item: IUserData) => (
                <Text variant="medium" styles={{ root: { fontWeight: '600' } }}>
                    {item.displayName}
                </Text>
            ),
        },
        {
            key: 'mobilePhone',
            name: 'Mobile',
            fieldName: 'mobilePhone',
            minWidth: 120,
            maxWidth: 150,
            isResizable: true,
            isSorted: sortedColumnKey === 'mobilePhone',
            isSortedDescending: !sortAscending,
            onRender: (item: IUserData) => (
                <Text variant="medium">
                    {item.mobilePhone ? formatPhoneNumber(item.mobilePhone) : ''}
                </Text>
            ),
        },
        {
            key: 'businessPhones',
            name: 'Office Phone',
            fieldName: 'businessPhones',
            minWidth: 120,
            maxWidth: 150,
            isResizable: true,
            isSorted: sortedColumnKey === 'businessPhones',
            isSortedDescending: !sortAscending,
            onRender: (item: IUserData) => (
                <Text variant="medium">
                    {item.businessPhones && item.businessPhones.length > 0 && item.businessPhones[0]
                        ? formatPhoneNumber(item.businessPhones[0])
                        : ''}
                </Text>
            ),
        },
        {
            key: 'ipPhone',
            name: 'Extension',
            fieldName: 'ipPhone',
            minWidth: 100,
            maxWidth: 120,
            isResizable: true,
            isSorted: sortedColumnKey === 'ipPhone',
            isSortedDescending: !sortAscending,
            onRender: (item: IUserData) => {
                return <Text variant="medium">{item?.ipPhone || ''}</Text>;
            },
        },
        {
            key: 'mail',
            name: 'Email',
            fieldName: 'mail',
            minWidth: 200,
            maxWidth: 250,
            isResizable: true,
            isSorted: sortedColumnKey === 'mail',
            isSortedDescending: !sortAscending,
            onRender: (item: IUserData) => (
                <Text variant="medium">
                    {item.mail && (
                        <a href={`mailto:${item.mail}`} style={{ color: '#0078d4', textDecoration: 'none' }}>
                            {item.mail}
                        </a>
                    )}
                </Text>
            ),
        },
        {
            key: 'officeLocation',
            name: 'Location',
            fieldName: 'officeLocation',
            minWidth: 120,
            maxWidth: 150,
            isResizable: true,
            isSorted: sortedColumnKey === 'officeLocation',
            isSortedDescending: !sortAscending,
            onRender: (item: IUserData) => (
                <Text variant="medium">{item.officeLocation || ''}</Text>
            ),
        },
        {
            key: 'jobTitle',
            name: 'Title',
            fieldName: 'jobTitle',
            minWidth: 150,
            maxWidth: 200,
            isResizable: true,
            isSorted: sortedColumnKey === 'jobTitle',
            isSortedDescending: !sortAscending,
            onRender: (item: IUserData) => (
                <Text variant="medium">{item.jobTitle || ''}</Text>
            ),
        },
    ], [sortedColumnKey, sortAscending, context?.serviceScope]);

    // Command bar items
    const commandBarItems: ICommandBarItemProps[] = [
        {
            key: 'refresh',
            text: 'Refresh',
            iconProps: { iconName: 'Refresh' },
            onClick: () => {
                fetchUsers();
            },
            disabled: loading,
        },
    ];

    return (
        <div>
            <Stack tokens={{ childrenGap: 20 }}>
                <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                    <Text variant="xxLarge" styles={{ root: { fontWeight: '600' } }}>
                        People Library
                    </Text>
                    <CommandBar items={commandBarItems} />
                </Stack>

                <Stack horizontal tokens={{ childrenGap: 20 }} verticalAlign="end">
                    <SearchBox
                        placeholder="Search users by name, email, phone, location, or title..."
                        value={searchText}
                        onChange={(_, newValue) => handleSearch(newValue)}
                        styles={{ root: { maxWidth: '400px', flexGrow: 1 } }}
                    />
                </Stack>

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

                {!loading && (
                    <Text variant="medium" styles={{ root: { color: '#666' } }}>
                        Showing {currentPageUsers.length} of {filteredUsers.length} users
                        {(searchText || selectedAlphabet) && ` (filtered from ${users.length} total)`}
                        {sortedColumnKey && ` sorted by ${columns.find(c => c.key === sortedColumnKey)?.name} (${sortAscending ? 'A-Z' : 'Z-A'})`}
                    </Text>
                )}

                {error && (
                    <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
                        {error}
                    </MessageBar>
                )}

                {loading && (
                    <Stack horizontalAlign="center" styles={{ root: { padding: '40px' } }}>
                        <Spinner size={SpinnerSize.large} label="Loading users..." />
                    </Stack>
                )}

                {!loading && !error && (
                    <>
                        <DetailsList
                            items={currentPageUsers}
                            columns={columns}
                            layoutMode={DetailsListLayoutMode.justified}
                            selectionMode={SelectionMode.none}
                            onColumnHeaderClick={onColumnClick}
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

                {!loading && !error && filteredUsers.length === 0 && users.length > 0 && (
                    <Stack horizontalAlign="center" styles={{ root: { padding: '40px' } }}>
                        <Text variant="mediumPlus">No users found matching your search criteria.</Text>
                    </Stack>
                )}
            </Stack>
        </div>
    );
};

export default UsersTable;