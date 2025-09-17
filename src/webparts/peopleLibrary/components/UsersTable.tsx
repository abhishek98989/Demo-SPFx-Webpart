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
}

// Props interface for the component
interface IUsersTableProps {
    context: WebPartContext;
}

// Photo cache to prevent duplicate requests and wrong photos
const photoCache = new Map<string, { url: string | null; loading: boolean; error: boolean }>();

// LivePersonaCard component with improved photo loading and caching
interface ILivePersonaCardProps {
    serviceScope: any;
    userEmail: string;
    displayName: string;
    graphClient?: MSGraphClientV3;
}

const LivePersonaCard: React.FC<ILivePersonaCardProps> = ({
    serviceScope,
    userEmail,
    displayName,
    graphClient
}) => {
    const [photoState, setPhotoState] = useState<{ url: string | null; loading: boolean; error: boolean }>({
        url: null,
        loading: false,
        error: false
    });
    
    const mountedRef = useRef(true);
    const loadingRef = useRef(false);
    
    // Normalize email for consistent caching
    const normalizedEmail = userEmail?.toLowerCase().trim();

    // Function to get initials from display name
    const getInitials = (name: string): string => {
        if (!name) return '??';
        return name.trim().split(/\s+/).map((word: string) => word?.charAt(0).toUpperCase()).join('').substring(0, 2);
    };

    // Update state from cache
    useEffect(() => {
        if (!normalizedEmail) return;
        
        const cached = photoCache.get(normalizedEmail);
        if (cached) {
            setPhotoState(cached);
        }
    }, [normalizedEmail]);

    // Load user photo with improved caching and race condition handling
    useEffect(() => {
        const loadPhoto = async () => {
            if (!graphClient || !normalizedEmail || loadingRef.current) return;

            // Check if already in cache
            const cached = photoCache.get(normalizedEmail);
            if (cached && (cached.url || cached.error)) {
                setPhotoState(cached);
                return;
            }

            // Prevent multiple simultaneous requests for same user
            if (cached?.loading) return;

            loadingRef.current = true;
            const loadingState = { url: null, loading: true, error: false };
            photoCache.set(normalizedEmail, loadingState);
            setPhotoState(loadingState);

            try {
                // Try multiple identifiers to get the photo
                let photoResponse;
                try {
                    // First try with email
                    photoResponse = await graphClient
                        .api(`/users/${userEmail}/photo/$value`)
                        .responseType(ResponseType.BLOB)
                        .get();
                } catch (emailError) {
                    // If email fails, try with UPN if different
                    if (userEmail !== normalizedEmail) {
                        photoResponse = await graphClient
                            .api(`/users/${normalizedEmail}/photo/$value`)
                            .responseType(ResponseType.BLOB)
                            .get();
                    } else {
                        throw emailError;
                    }
                }

                if (photoResponse && mountedRef.current) {
                    let blob: Blob;
                    if (photoResponse instanceof Blob) {
                        blob = photoResponse;
                    } else if (photoResponse.body instanceof Blob) {
                        blob = photoResponse.body;
                    } else {
                        blob = new Blob([photoResponse.body || photoResponse], { type: 'image/jpeg' });
                    }
                    
                    const url = URL.createObjectURL(blob);
                    const successState = { url, loading: false, error: false };
                    photoCache.set(normalizedEmail, successState);
                    setPhotoState(successState);
                }
            } catch (error) {
                console.log('Photo not available for user:', normalizedEmail);
                const errorState = { url: null, loading: false, error: true };
                photoCache.set(normalizedEmail, errorState);
                if (mountedRef.current) {
                    setPhotoState(errorState);
                }
            } finally {
                loadingRef.current = false;
            }
        };

        loadPhoto();
    }, [graphClient, normalizedEmail, userEmail]);

    // Cleanup on unmount
    useEffect(() => {
        return () => {
            mountedRef.current = false;
        };
    }, []);

    // Create persona element with consistent photo state
    const personaElement = (
        <Persona
            text={displayName}
            size={PersonaSize.size40}
            hidePersonaDetails={true}
            imageUrl={photoState.url || undefined}
            imageInitials={photoState.url ? undefined : getInitials(displayName)}
            coinSize={40}
        />
    );

    // If we have service scope and valid email, wrap with LivePersona for hover functionality
    if (serviceScope && normalizedEmail && !photoState.error) {
        return (
            <div style={{ position: 'relative' }}>
                {/* Visible persona with photo */}
                {personaElement}
                {/* Invisible LivePersona positioned over the visible one for hover functionality */}
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
                        template={
                            <Persona
                                text={displayName}
                                size={PersonaSize.size40}
                                hidePersonaDetails={true}
                                imageUrl={photoState.url || undefined}
                                imageInitials={photoState.url ? undefined : getInitials(displayName)}
                                coinSize={40}
                            />
                        }
                    />
                </div>
            </div>
        );
    }

    // Fallback to regular Persona only
    return personaElement;
};

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
                    .select(
                        'id,displayName,mail,mobilePhone,ipPhone,businessPhones,officeLocation,jobTitle,userPrincipalName,userType,accountEnabled,companyName'
                    )
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

    // Normalize phone number for searching (remove spaces, dashes, parentheses)
    const normalizePhone = (phone: string): string => {
        if (!phone) return '';
        return phone.replace(/[\s\-\(\)\.]/g, '');
    };

    // Improved search function
    const searchInData = (user: IUserData, searchTerm: string): boolean => {
        if (!searchTerm.trim()) return true;
        
        const term = searchTerm.toLowerCase().trim();
        const normalizedSearchTerm = normalizePhone(term);
        
        // Search in text fields (case insensitive)
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
        
        // Search in phone numbers (normalized)
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

    // Fetch users from Microsoft Graph
    const fetchUsers = useCallback(async () => {
        if (!graphClient) return;

        setLoading(true);
        setError('');

        try {
            // Clear photo cache when refreshing data
            photoCache.clear();
            
            let usersData = await fetchAllUsers(graphClient);

            usersData = usersData.filter(user => {
                // 1. Only company contains "Vaughn"
                if (user.companyName && !user.companyName.toLowerCase().includes('vaughn')) {
                    return false;
                }

                // 2. Exclude Disabled accounts
                if (user.accountEnabled === false) {
                    return false;
                }

                // 3. Exclude Guest users
                if (user.userType === 'Guest') {
                    return false;
                }

                // 4. Exclude #EXT# (external)
                if (user.userPrincipalName && user.userPrincipalName.includes('#EXT#')) {
                    return false;
                }

                // 5. Exclude UPN or mail with underscores or test accounts
                if ((user.mail && user.mail.includes('_')) || 
                    (user.userPrincipalName && user.userPrincipalName.includes('_')) || 
                    (user.userPrincipalName && user.userPrincipalName?.toLowerCase()?.includes('test'))) {
                    return false;
                }

                // 6. Exclude job titles per AD filter
                if (user?.jobTitle) {
                    const jobTitleLower = user.jobTitle.toLowerCase();
                    const excludedTitles = ['intern', 'construction work', 'layout', 'foreman', 'instrument', 'student'];
                    const excludedExactTitles = ['carpenter', 'driver', 'craft worker', 'rod man'];
                    
                    if (excludedTitles.some(title => jobTitleLower.includes(title)) ||
                        excludedExactTitles.includes(user.jobTitle.trim().toLowerCase())) {
                        return false;
                    }
                }

                // 7. Exclude system/service/admin accounts
                if (user.userPrincipalName) {
                    const upnLower = user.userPrincipalName.toLowerCase();
                    if (upnLower.includes('service') || upnLower.includes('admin') || upnLower.includes('system')) {
                        return false;
                    }
                }

                // 8. Require valid email
                if (!user.mail || user.mail.trim() === '') {
                    return false;
                }

                // 9. Require jobTitle not empty
                if (!user.jobTitle || user.jobTitle.trim() === '') {
                    return false;
                }

                return true;
            });

            console.log(`Fetched ${usersData.length} filtered users from Azure AD`);

            // Sort alphabetically by displayName
            usersData = usersData.sort((a, b) => {
                const nameA = (a.displayName || '').toLowerCase();
                const nameB = (b.displayName || '').toLowerCase();
                return sortAscending ? nameA.localeCompare(nameB) : nameB.localeCompare(nameA);
            });

            setUsers(usersData);
        } catch (err) {
            setError('Failed to fetch users from Azure AD');
            console.error('Error fetching users:', err);
        } finally {
            setLoading(false);
        }
    }, [graphClient, sortAscending]);

    // Fetch users when graph client is ready
    useEffect(() => {
        if (graphClient) {
            fetchUsers();
        }
    }, [graphClient, fetchUsers]);

    // Memoized filtered and sorted users with improved search
    const processedUsers = useMemo(() => {
        let result = [...users];

        // Sort alphabetically
        result = result.sort((a, b) => {
            const nameA = (a.displayName || '').toLowerCase();
            const nameB = (b.displayName || '').toLowerCase();
            return sortAscending ? nameA.localeCompare(nameB) : nameB.localeCompare(nameA);
        });

        // Apply alphabet filter
        if (selectedAlphabet) {
            result = result.filter(user =>
                user.displayName?.toUpperCase().startsWith(selectedAlphabet)
            );
        }

        // Apply search filter with improved logic
        if (searchText.trim()) {
            result = result.filter(user => searchInData(user, searchText));
        }

        return result;
    }, [users, searchText, selectedAlphabet, sortAscending]);

    // Update filtered users and reset page only when filters change
    useEffect(() => {
        setFilteredUsers(processedUsers);
        // Only reset to page 1 if we're currently beyond the available pages
        const totalPages = Math.ceil(processedUsers.length / itemsPerPage);
        if (currentPage > totalPages && totalPages > 0) {
            setCurrentPage(1);
            setPageInput('1');
        }
    }, [processedUsers, currentPage, itemsPerPage]);

    // Handle search (reset page when search changes)
    const handleSearch = (newValue?: string) => {
        const value = newValue || '';
        setSearchText(value);
        setCurrentPage(1);
        setPageInput('1');
    };

    // Handle alphabet filter (reset page when alphabet filter changes)
    const handleAlphabetFilter = (letter: string) => {
        if (selectedAlphabet === letter) {
            setSelectedAlphabet(''); // Clear filter if same letter clicked
        } else {
            setSelectedAlphabet(letter);
        }
        setCurrentPage(1);
        setPageInput('1');
    };

    // Handle sort toggle
    const handleSortToggle = () => {
        setSortAscending(!sortAscending);
        // Don't reset page when sorting
    };

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
        // Basic formatting - you can enhance this based on your phone number format requirements
        const cleaned = phone.replace(/\D/g, '');
        if (cleaned.length === 10) {
            return `(${cleaned.slice(0, 3)}) ${cleaned.slice(3, 6)}-${cleaned.slice(6)}`;
        }
        return phone; // Return as-is if not standard format
    };

    // Define table columns with improved phone rendering
    const columns: IColumn[] = [
        {
            key: 'photo',
            name: '',
            fieldName: 'photo',
            minWidth: 60,
            maxWidth: 60,
            onRender: (item: IUserData) => {
                // Use mail as primary identifier, fallback to userPrincipalName
                const userIdentifier = item.mail || item.userPrincipalName;
                return (
                    <LivePersonaCard
                        serviceScope={context?.serviceScope}
                        userEmail={userIdentifier}
                        displayName={item.displayName}
                        graphClient={graphClient ?? undefined}
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
            onRender: (item: IUserData) => (
                <Text variant="medium">{item.jobTitle || ''}</Text>
            ),
        },
    ];

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
                {/* Header */}
                <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                    <Text variant="xxLarge" styles={{ root: { fontWeight: '600' } }}>
                        People Library
                    </Text>
                    <CommandBar items={commandBarItems} />
                </Stack>

                {/* Controls Row */}
                <Stack horizontal tokens={{ childrenGap: 20 }} verticalAlign="end">
                    <SearchBox
                        placeholder="Search users by name, email, phone, location, or title..."
                        value={searchText}
                        onChange={(_, newValue) => handleSearch(newValue)}
                        styles={{ root: { maxWidth: '400px', flexGrow: 1 } }}
                    />

                    <Toggle
                        label="Sort A-Z"
                        checked={sortAscending}
                        onChange={handleSortToggle}
                        onText="A-Z"
                        offText="Z-A"
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
                        Showing {currentPageUsers.length} of {filteredUsers.length} users
                        {(searchText || selectedAlphabet) && ` (filtered from ${users.length} total)`}
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
                        <Spinner size={SpinnerSize.large} label="Loading users..." />
                    </Stack>
                )}

                {/* Users Table */}
                {!loading && !error && (
                    <>
                        <DetailsList
                            items={currentPageUsers}
                            columns={columns}
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

                        {/* Enhanced Pagination */}
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