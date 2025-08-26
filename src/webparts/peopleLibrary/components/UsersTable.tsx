import * as React from 'react';
import { useState, useEffect, useCallback, useMemo } from 'react';
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
    companyName:string
    userPrincipalName: string;
    userType?: string; // Added to help with guest user filtering
    accountEnabled?: boolean; // Added to filter disabled accounts
    photo?: string;
}

// Props interface for the component
interface IUsersTableProps {
    context: WebPartContext;
}

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
    const [photoCache, setPhotoCache] = useState<Map<string, string | undefined>>(new Map());
    const [loadingPhotos, setLoadingPhotos] = useState<Set<string>>(new Set());

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

    // Fetch user photo with caching
    const fetchUserPhoto = async (userId: string): Promise<string | undefined> => {
        if (!graphClient) return undefined;

        // Check if photo is already cached
        if (photoCache.has(userId)) {
            return photoCache.get(userId);
        }

        // Check if photo is currently being loaded
        if (loadingPhotos.has(userId)) {
            return undefined;
        }

        try {
            setLoadingPhotos(prev => new Set(prev).add(userId));

            const photoResponse = await graphClient
                .api(`/users/${userId}/photo/$value`)
                .responseType(ResponseType.BLOB)
                .get();

            let photoUrl: string | undefined = undefined;

            // Handle different response formats
            if (photoResponse) {
                let blob: Blob;

                if (photoResponse instanceof Blob) {
                    blob = photoResponse;
                } else if (photoResponse.body) {
                    if (photoResponse.body instanceof Blob) {
                        blob = photoResponse.body;
                    } else {
                        // Convert ArrayBuffer or other format to Blob
                        blob = new Blob([photoResponse.body], { type: 'image/jpeg' });
                    }
                } else {
                    // Try to convert the entire response to blob
                    blob = new Blob([photoResponse], { type: 'image/jpeg' });
                }

                photoUrl = URL.createObjectURL(blob);
                console.log('Photo loaded successfully for user:', userId, 'URL:', photoUrl);
            }

            // Cache the photo (even if undefined) - this won't trigger re-renders
            setPhotoCache(prev => new Map(prev).set(userId, photoUrl));

            return photoUrl;
        } catch (err) {
            console.log('Photo not available for user:', userId, err);
            // Photo might not exist, cache undefined to avoid future requests
            setPhotoCache(prev => new Map(prev).set(userId, undefined));
            return undefined;
        } finally {
            setLoadingPhotos(prev => {
                const newSet = new Set(prev);
                newSet.delete(userId);
                return newSet;
            });
        }
    };

    // Load photos for visible users (optimized to not cause re-renders)
    // Batch load all images for current view and set at once to avoid flicker
    const loadPhotosForVisibleUsers = useCallback(async (visibleUsers: IUserData[]) => {
        const usersNeedingPhotos = visibleUsers.filter(user =>
            !photoCache.has(user.userPrincipalName) &&
            !loadingPhotos.has(user.userPrincipalName)
        );

        if (usersNeedingPhotos.length === 0) return;

        // Fetch all photos in parallel
        const photoPromises = usersNeedingPhotos.map(async user => {
            try {
                const photoUrl = await fetchUserPhoto(user.userPrincipalName);
                return [user.userPrincipalName, photoUrl] as [string, string | undefined];
            } catch (error) {
                console.error('Error loading photo for user:', user.userPrincipalName, error);
                return [user.userPrincipalName, undefined] as [string, undefined];
            }
        });

        const photoResults = await Promise.all(photoPromises);
        // Update photoCache in one go
        setPhotoCache(prev => {
            const newCache = new Map(prev);
            photoResults.forEach(([userId, photoUrl]) => {
                newCache.set(userId, photoUrl);
            });
            return newCache;
        });
    }, [photoCache, loadingPhotos, fetchUserPhoto]);

    // Fetch all users from Microsoft Graph (recursive to get all users)
  const fetchAllUsers = async (
  graphClient: MSGraphClientV3,
  users: IUserData[] = [],
  nextLink?: string
): Promise<IUserData[]> => {
  try {
    let request;
    if (nextLink) {
      // Use the nextLink for pagination
      request = graphClient.api(nextLink.replace('https://graph.microsoft.com/v1.0', ''));
    } else {
      // Initial request
      request = graphClient
        .api('/users')
        .select(
          'id,displayName,mail,mobilePhone,ipPhone,businessPhones,officeLocation,jobTitle,userPrincipalName,userType,accountEnabled,companyName'
        )
        .top(999); // Max per request
    }

    const response = await request.get();
    const newUsers = response.value || [];
    const allUsers = [...users, ...newUsers];

    // If there's a nextLink, recursively fetch more users
    if (response['@odata.nextLink']) {
      return await fetchAllUsers(graphClient, allUsers, response['@odata.nextLink']);
    }

    return allUsers;
  } catch (error) {
    console.error('Error fetching users batch:', error);
    return users; // Return what we have so far
  }
};

// Fetch users from Microsoft Graph
const fetchUsers = useCallback(async () => {
  if (!graphClient) return;

  setLoading(true);
  setError('');

  try {
    // Fetch all users using recursive pagination
    let usersData = await fetchAllUsers(graphClient);

    usersData = usersData.filter(user => {
      // 1. Only company contains "Vaughn"
      if (user.companyName && !user.companyName.toLowerCase().includes('vaughn')) {
        return false;
      }

      // 2. Exclude Disabled accounts → check accountEnabled
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

      // 5. Exclude UPN or mail with underscores
      if ((user.mail && user.mail.includes('_')) || (user.userPrincipalName && user.userPrincipalName.includes('_'))) {
        return false;
      }

      // 6. Exclude job titles per AD filter
      if (
        user?.jobTitle &&
        (
          user.jobTitle.toLowerCase().includes('intern') ||
          user.jobTitle.toLowerCase().includes('construction work') ||
          user.jobTitle.toLowerCase().includes('layout') ||
          user.jobTitle.toLowerCase().includes('foreman') ||
          user.jobTitle.toLowerCase().includes('instrument') ||
          user.jobTitle.toLowerCase().includes('student') ||
          ['carpenter', 'driver', 'craft worker', 'rod man'].includes(user.jobTitle.trim())
        )
      ) {
        return false;
      }

      // 7. Exclude system/service/admin accounts
      if (
        user.userPrincipalName &&
        (user.userPrincipalName.toLowerCase().includes('service') ||
          user.userPrincipalName.toLowerCase().includes('admin') ||
          user.userPrincipalName.toLowerCase().includes('system'))
      ) {
        return false;
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

    // Memoized filtered and sorted users (fixed to not reset pagination)
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

        // Apply search filter
        if (searchText.trim()) {
            result = result.filter(user =>
                user.displayName?.toLowerCase().includes(searchText.toLowerCase()) ||
                user.mail?.toLowerCase().includes(searchText.toLowerCase()) ||
                user.mobilePhone?.toLowerCase().includes(searchText.toLowerCase()) ||
                user.businessPhones?.some(phone => phone.toLowerCase().includes(searchText.toLowerCase())) ||
                user.officeLocation?.toLowerCase().includes(searchText.toLowerCase()) ||
                user.jobTitle?.toLowerCase().includes(searchText.toLowerCase()) ||
                user.userPrincipalName?.toLowerCase().includes(searchText.toLowerCase())
            );
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

    // Load photos for current page users (moved after currentPageUsers is set)
    useEffect(() => {
        if (currentPageUsers.length > 0 && !loading) {
            loadPhotosForVisibleUsers(currentPageUsers);
        }
    }, [currentPageUsers, loading, loadPhotosForVisibleUsers]);

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

    // Define table columns
    const columns: IColumn[] = [
        {
            key: 'photo',
            name: '',
            fieldName: 'photo',
            minWidth: 60,
            maxWidth: 60,
            onRender: (item: IUserData) => {
                const photoUrl = photoCache.get(item.userPrincipalName) || item.photo;
                return (
                    <Persona
                        size={PersonaSize.size40}
                        imageUrl={photoUrl}
                        text={item.displayName}
                        presence={PersonaPresence.none}
                        hidePersonaDetails={true}
                        imageInitials={item.displayName ? item.displayName.split(' ').map(n => n[0]).join('').substring(0, 2) : '??'}
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
                <Text variant="medium">{item.mobilePhone}</Text>
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
                    {item.businessPhones && item.businessPhones.length > 0 && item.businessPhones[0] }
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
                return <Text variant="medium">{item?.ipPhone}</Text>;
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
                <Text variant="medium">{item.officeLocation}</Text>
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
                <Text variant="medium">{item.jobTitle}</Text>
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
                // Clear photo cache on refresh
                setPhotoCache(new Map());
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