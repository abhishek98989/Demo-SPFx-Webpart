import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  SelectionMode,
  IColumn,
  SearchBox,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  Modal,
  IconButton,
  Slider,
  TextField,
  Panel,
  PanelType,
  Stack,
  Text
} from '@fluentui/react';
import '../../../globalCommon/fluent.css';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { useGlobalLoaderContext } from '../../../globalCommon/customLoader';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IPostsArchiveProps } from './IPostsArchiveProps';
import { IArchiveItem } from '../PostsArchiveWebPart';
import styles from './PostsArchive.module.scss';
import AttachmentViewer from './AttachmentViewer';
import AddEditForm from './AddEditForm';
import { SPFx, spfi } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { GlobalLoaderProvider } from '../../../globalCommon/customLoader';

// Extended interface to include image URL for slider view and preloading status
interface IExtendedArchiveItem extends IArchiveItem {
  imageUrl?: string;
  isImagePreloaded?: boolean; // New property to track if image has been preloaded
}
const PostsArchiveContent: React.FC<IPostsArchiveProps> = (props) => {


  
  const [items, setItems] = useState<IExtendedArchiveItem[]>([]);
  const sp = spfi().using(SPFx(props?.Context));
  const [filteredItems, setFilteredItems] = useState<IExtendedArchiveItem[]>([]);
  const [searchText, setSearchText] = useState<string>('');
  const [selectedRegions, setSelectedRegions] = useState<string[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [selectedItem, setSelectedItem] = useState<IArchiveItem | any>(null);
  const [isModalOpen, setIsModalOpen] = useState<boolean>(false);
  const [isAddPanelOpen, setIsAddPanelOpen] = useState<boolean>(false);
  const [sliderIndex, setSliderIndex] = useState<number>(0);
  const { showLoader, hideLoader } = useGlobalLoaderContext();
  // Auto-slider states
  const intervalRef = useRef<NodeJS.Timeout | null>(null);
  const [isSliderPaused, setIsSliderPaused] = useState<boolean>(false);
  
  // File viewer states
  const [isViewerOpen, setIsViewerOpen] = useState<boolean>(false);
  const [currentViewFile, setCurrentViewFile] = useState<{ url: string; name: string; type: string } | null>(null);

  // Mouse/Touch sliding states
  const [isDragging, setIsDragging] = useState<boolean>(false);
  const [startX, setStartX] = useState<number>(0);
  const [currentX, setCurrentX] = useState<number>(0);
  const [translateX, setTranslateX] = useState<number>(0);
  const sliderRef = useRef<HTMLDivElement>(null);
  const slideContentRef = useRef<HTMLDivElement>(null);

  const regionOptions: IDropdownOption[] = [
    { key: 'Houston', text: 'Houston' },
    { key: 'Austin', text: 'Austin' },
    { key: 'Lubbock', text: 'Lubbock' },
    { key: 'El Paso', text: 'El Paso' },
    { key: 'DFW', text: 'DFW' },
    { key: 'College Station', text: 'College Station' },
    { key: 'RGV', text: 'RGV' },
    { key: 'San Antonio', text: 'San Antonio' },
    { key: 'BCS/Stephenville', text: 'BCS/Stephenville' },
    { key: 'North Texas', text: 'North Texas' },
    { key: 'Galveston', text: 'Galveston' },
    { key: 'Louisiana', text: 'Louisiana' },
    { key: 'Rio Grande Valley', text: 'Rio Grande Valley' }
  ];

  useEffect(() => {
    loadItems();
  }, []);

  useEffect(() => {
    filterItems();
  }, [items, searchText, selectedRegions]);

  // Keyboard navigation effect
  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => {
      // Only handle keyboard navigation when slider view is active and no modals are open
      if (props.viewType === 'slider' && !isModalOpen && !isAddPanelOpen && !isViewerOpen) {
        if (event.key === 'ArrowLeft') {
          event.preventDefault();
          prevSlide();
        } else if (event.key === 'ArrowRight') {
          event.preventDefault();
          nextSlide();
        }
      }
    };

    // Add event listener to document
    document.addEventListener('keydown', handleKeyDown);

    // Cleanup
    return () => {
      document.removeEventListener('keydown', handleKeyDown);
    };
  }, [props.viewType, isModalOpen, isAddPanelOpen, isViewerOpen, filteredItems.length]);

  // Auto-slider effect
  useEffect(() => {
    if (props.viewType === 'slider' && props.slideAfter && !isSliderPaused && !isViewerOpen && !isDragging) {
      const slideTime = parseInt(props.slideAfter.toString());
      if (!isNaN(slideTime) && slideTime > 0) {
        intervalRef.current = setInterval(() => {
          nextSlide();
        }, slideTime * 1000);
      }
    }

    return () => {
      if (intervalRef.current) {
        clearInterval(intervalRef.current);
      }
    };
  }, [props.viewType, props.slideAfter, isSliderPaused, isViewerOpen, sliderIndex, isDragging]);

  // Pause slider when viewer opens or dragging
  useEffect(() => {
    if (isViewerOpen || isDragging) {
      setIsSliderPaused(true);
    } else {
      setIsSliderPaused(false);
    }
  }, [isViewerOpen, isDragging]);

  // Mouse/Touch event handlers
  const handleMouseDown = (e: React.MouseEvent) => {
    if (filteredItems.length <= 1) return;
    
    setIsDragging(true);
    setStartX(e.clientX);
    setCurrentX(e.clientX);
    
    // Prevent text selection during drag
    e.preventDefault();
  };

  const handleTouchStart = (e: React.TouchEvent) => {
    if (filteredItems.length <= 1) return;
    
    setIsDragging(true);
    setStartX(e.touches[0].clientX);
    setCurrentX(e.touches[0].clientX);
  };

  const handleMouseMove = (e: React.MouseEvent) => {
    if (!isDragging) return;
    
    setCurrentX(e.clientX);
    const diff = e.clientX - startX;
    setTranslateX(diff);
  };

  const handleTouchMove = (e: React.TouchEvent) => {
    if (!isDragging) return;
    
    setCurrentX(e.touches[0].clientX);
    const diff = e.touches[0].clientX - startX;
    setTranslateX(diff);
  };

  const handleMouseUp = () => {
    if (!isDragging) return;
    
    const diff = currentX - startX;
    const threshold = 50; // Minimum distance to trigger slide
    
    if (Math.abs(diff) > threshold) {
      if (diff > 0) {
        prevSlide();
      } else {
        nextSlide();
      }
    }
    
    setIsDragging(false);
    setTranslateX(0);
  };

  const handleTouchEnd = () => {
    handleMouseUp();
  };

  // Helper function to check if file is an image
  const isImageFile = (fileName: string): boolean => {
    const imageExtensions = ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'svg', 'webp'];
    const extension = fileName.split('.').pop()?.toLowerCase();
    return imageExtensions.includes(extension || '');
  };

  // Function to load and preload image URL for slider view items
  const preloadImage = async (item: IExtendedArchiveItem): Promise<IExtendedArchiveItem> => {
    if (!item.AttachmentFiles || item.AttachmentFiles.length === 0) {
      return { ...item, isImagePreloaded: true }; // No image to preload
    }

    const imageAttachment = item.AttachmentFiles.find(att => isImageFile(att.FileName));
    
    if (!imageAttachment) {
      return { ...item, isImagePreloaded: true }; // No image attachment to preload
    }

    const imageUrl = `${props.Context.pageContext.web.absoluteUrl}/_api/web/lists('${props.listId}')/items(${item.Id})/AttachmentFiles('${imageAttachment.FileName}')/$value`;
    
    return new Promise(resolve => {
      const img = new Image();
      img.src = imageUrl;
      img.onload = () => {
        resolve({
          ...item,
          imageUrl: imageUrl,
          isImagePreloaded: true // Mark as preloaded when image loads
        });
      };
      img.onerror = (e) => {
        console.error('Error preloading image:', imageUrl, e);
        resolve({
          ...item,
          imageUrl: undefined, // Clear imageUrl if loading fails
          isImagePreloaded: true // Still mark as attempted and finished
        });
      };
    });
  };

  const loadItems = async (): Promise<void> => {
    showLoader()
    try {
      let query = sp.web.lists
        .getById(props?.listId)
        .items
        .select(
          "Id", "Title", "PublishingDate", "Source", "Description", "Region", 
          "Images", "Attachments", "AttachmentFiles"
        )
        .expand("AttachmentFiles")
        .orderBy("PublishingDate", false);

      let allItems: IExtendedArchiveItem[] = [];

      if (props.viewType === 'slider') {
        const itemsToLoad = props.numberOfEvents || 10; 
        const fetchedItems = await query.top(itemsToLoad)();
        
        // Preload images for all fetched slider items that have single image attachments
        const itemsWithPreloadedImages = await Promise.all(
          fetchedItems.map(async (item) => {
            if (item.AttachmentFiles && 
                item.AttachmentFiles.length === 1 && 
                isImageFile(item.AttachmentFiles[0].FileName)) {
              return await preloadImage(item);
            }
            return {
              ...item,
              isImagePreloaded: true // Mark as preloaded if no image or multiple attachments
            };
          })
        );
        allItems = itemsWithPreloadedImages;
      } else {
        // For table view, load in batches to handle large datasets
        let hasMore = true;
        let skip = 0;
        const batchSize = 2000;

        while (hasMore) {
          const batchQuery = query.skip(skip).top(batchSize);
          const batchItems = await batchQuery();
          
          if (batchItems.length === 0) {
            hasMore = false;
          } else {
            // For table view, don't preload images, just mark as loaded
            allItems.push(...(batchItems.map(item => ({ ...item, isImagePreloaded: true }))));
            skip += batchSize;
            
            // If we got less than batch size, we've reached the end
            if (batchItems.length < batchSize) {
              hasMore = false;
            }
          }
        }
      }
hideLoader()
      setItems(allItems);

    } catch (error) {
      console.error("Error loading items:", error);
    } finally {
      setLoading(false);
    }
  };

  const filterItems = (): void => {
    let filtered = [...items];

    // Filter by search text - check if fields exist before filtering
    if (searchText && searchText.trim() !== '') {
      const searchLower = searchText.toLowerCase().trim();
      filtered = filtered.filter(item => {
        const title = item.Title ? item.Title.toLowerCase() : '';
        const source = item.Source ? item.Source.toLowerCase() : '';
        const description = item.Description ? item.Description.toLowerCase() : '';
        
        return title.includes(searchLower) || 
               source.includes(searchLower) || 
               description.includes(searchLower);
      });
    }

    // Filter by regions - check if Region field exists and has value
    if (selectedRegions.length > 0) {
      filtered = filtered.filter(item => 
        item.Region && selectedRegions.includes(item.Region)
      );
    }

    setFilteredItems(filtered);
    // Reset slider index when filters change
    setSliderIndex(0);
  };

  const columns: IColumn[] = [
    {
      key: 'title',
      name: 'Title',
      fieldName: 'Title',
      minWidth: 200,
      maxWidth: 300,
      isResizable: true,
      onRender: (item: IArchiveItem) => (
        <button
          className={styles.linkButton}
          onClick={() => openItemModal(item)}
        >
          {item.Title}
        </button>
      )
    },
    {
      key: 'publishingDate',
      name: 'Publishing Date',
      fieldName: 'PublishingDate',
      minWidth: 120,
      maxWidth: 150,
      isResizable: true,
      onRender: (item: IArchiveItem) =>
        new Date(item.PublishingDate).toLocaleDateString()
    },
    {
      key: 'source',
      name: 'Source',
      fieldName: 'Source',
      minWidth: 150,
      maxWidth: 200,
      isResizable: true
    },
    {
      key: 'region',
      name: 'Region',
      fieldName: 'Region',
      minWidth: 120,
      maxWidth: 150,
      isResizable: true
    }
  ];

  const openItemModal = (item: IArchiveItem): void => {
    setSelectedItem(item);
    setIsModalOpen(true);
  };

  const closeModal = (): void => {
    setIsModalOpen(false);
    setIsAddPanelOpen(false);
    setSelectedItem(null);
  };

  const nextSlide = (): void => {
    const maxIndex = filteredItems.length - 1;
    setSliderIndex(prev => (prev >= maxIndex ? 0 : prev + 1));
  };

  const prevSlide = (): void => {
    const maxIndex = filteredItems.length - 1;
    setSliderIndex(prev => (prev <= 0 ? maxIndex : prev - 1));
  };

  const getFileIcon = (fileName: string): string => {
    const extension = fileName.split('.').pop()?.toLowerCase();
    if (extension === 'pdf') {
      return 'PDF';
    } else if (['jpg', 'jpeg', 'png', 'gif', 'bmp', 'svg'].includes(extension || '')) {
      return 'FileImage';
    }
    return 'Attach';
  };

  // Enhanced viewAttachment function with SharePoint authentication headers
  const viewAttachment = async (attachment: any, itemId: number): Promise<void> => {
    try {
      const fileUrl = `${props.Context.pageContext.web.absoluteUrl}/_api/web/lists('${props.listId}')/items(${itemId})/AttachmentFiles('${attachment.FileName}')/$value`;
      const fileExtension = attachment.FileName.split('.').pop()?.toLowerCase();
      
      if (fileExtension === 'pdf') {
        // For PDFs, create a blob URL to avoid direct download
        const response = await props.Context.httpClient.get(fileUrl, SPHttpClient.configurations.v1);
        
        if (response.ok) {
          const blob = await response.blob();
          const blobUrl = URL.createObjectURL(blob);
          
          setCurrentViewFile({
            url: blobUrl,
            name: attachment.FileName,
            type: 'application/pdf'
          });
          setIsViewerOpen(true);
        } else {
          console.error('Failed to fetch PDF file:', response.statusText);
        }
      } else {
        // For images, use direct URL
        const fileType = 'image';
        
        setCurrentViewFile({
          url: fileUrl,
          name: attachment.FileName,
          type: fileType
        });
        setIsViewerOpen(true);
      }
    } catch (error) {
      console.error('Error viewing file:', error);
    }
  };

  const closeViewer = (): void => {
    // Clean up blob URL if it exists
    if (currentViewFile?.url.startsWith('blob:')) {
      URL.revokeObjectURL(currentViewFile.url);
    }    
    setIsViewerOpen(false);
    setCurrentViewFile(null);
  };


  const renderTableView = (): JSX.Element => (
    <div className={styles.tableView}>
      <div className={styles.headerControls}>
        <PrimaryButton
          text="Add New Post"
          onClick={() => setIsAddPanelOpen(true)}
          className={styles.addButton}
        />
      </div>

      <div className={styles.filters}>
        <SearchBox
          placeholder="Search posts..."
          value={searchText}
          onChange={(_, newValue) => setSearchText(newValue || '')}
          className={styles.searchBox}
        />

        <Dropdown
          placeholder="Filter by regions"
          options={regionOptions}
          multiSelect
          selectedKeys={selectedRegions}
          onChange={(_, option) => {
            if (option) {
              const newRegions = option.selected
                ? [...selectedRegions, option.key as string]
                : selectedRegions.filter(region => region !== option.key);
              setSelectedRegions(newRegions);
            }
          }}
          className={styles.regionDropdown}
        />
      </div>

      <DetailsList
        items={filteredItems}
        columns={columns}
        setKey="set"
        layoutMode={DetailsListLayoutMode.justified}
        selectionMode={SelectionMode.none}
        isHeaderVisible={true}
        className={styles.detailsList}
      />
    </div>
  );

  const renderSliderView = (): JSX.Element => {
    const currentItems = filteredItems;

    if (currentItems.length === 0) {
      return <div>No items to display</div>;
    }

    const currentItem = currentItems[sliderIndex];

    // Show loading indicator if the current image isn't yet preloaded
    if (props.viewType === 'slider' && !currentItem.isImagePreloaded) {
        return <div>Loading image...</div>;
    }

    return (
      <div className={styles.sliderView}>
        <div className={styles.sliderContainer}>
          <IconButton
            iconProps={{ iconName: 'ChevronLeft' }}
            onClick={prevSlide}
            className={styles.sliderArrow}
            disabled={currentItems.length <= 1}
          />

          <div 
            className={styles.slideContent}
            ref={slideContentRef}
            onMouseDown={handleMouseDown}
            onMouseMove={isDragging ? handleMouseMove : undefined}
            onMouseUp={handleMouseUp}
            onMouseLeave={handleMouseUp}
            onTouchStart={handleTouchStart}
            onTouchMove={isDragging ? handleTouchMove : undefined}
            onTouchEnd={handleTouchEnd}
            style={{
              transform: `translateX(${translateX}px)`,
              transition: isDragging ? 'none' : 'transform 0.3s ease',
              cursor: isDragging ? 'grabbing' : 'grab',
              userSelect: 'none'
            }}
          >
            <div className={styles.slideCard}>
              {/* Title */}
              <h3 className={styles?.slideTitle}>{currentItem?.Title}</h3>
              
              {/* Region */}
              <p className={styles?.slideRegion}>{currentItem?.Region}</p>
              
              {/* Description - Truncated */}
              <div className={styles?.slideDescription}>
                <div dangerouslySetInnerHTML={{ __html: currentItem?.Description }} />
              </div>

              {/* Single Image Display - Make clickable to open in modal */}
              {(currentItem?.imageUrl && currentItem?.AttachmentFiles &&
               currentItem?.AttachmentFiles?.length === 1 && isImageFile(currentItem.AttachmentFiles[0].FileName)) && (
                <div 
                  className={styles.inlineAttachmentImage}
                  onClick={(e) => {
                    e.stopPropagation();
                    // Find the image attachment and open it in viewer
                    const imageAttachment = currentItem.AttachmentFiles?.find(att => isImageFile(att.FileName));
                    if (imageAttachment) {
                      viewAttachment(imageAttachment, currentItem.Id);
                    }
                  }}
                  style={{ cursor: 'pointer' }}
                >
                  <img
                    src={currentItem.imageUrl}
                    alt={currentItem.Title}
                    style={{ 
                      maxWidth: '100%', 
                      maxHeight: '300px', 
                      objectFit: 'contain',
                      borderRadius: '4px',
                      marginTop: '12px',
                      marginBottom: '12px',
                      pointerEvents: 'none', // Prevent dragging of images
                      transition: 'opacity 0.2s ease'
                    }}
                    onError={(e) => {
                      // Hide image if it fails to load
                      (e.target as HTMLImageElement).style.display = 'none';
                    }}
                  />
                  {/* Overlay indicator */}
                  <div style={{
                    position: 'absolute',
                    top: '16px',
                    right: '8px',
                    background: 'rgba(0,0,0,0.6)',
                    color: 'white',
                    padding: '4px 8px',
                    borderRadius: '4px',
                    fontSize: '12px',
                    pointerEvents: 'none'
                  }}>
                    Click to view
                  </div>
                </div>
              )}


              {/* Attachments - Show all attachments including single images */}
              {currentItem?.AttachmentFiles && 
               currentItem?.AttachmentFiles?.length > 1 && ( // Changed from > 1 to > 0 to show all attachments
                <div className={styles.attachments}>
                  <Text variant="small" style={{ fontWeight: 600, marginBottom: '8px' }}>
                    Attachments:
                  </Text>
                  <div className={styles?.attachmentList}>
                    {currentItem?.AttachmentFiles?.map((attachment, index) => (
                      <div key={index} className={styles?.attachmentItem}>
                        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                          <IconButton
                            iconProps={{ iconName: getFileIcon(attachment?.FileName) }}
                            onClick={(e) => {
                              e.stopPropagation(); // Prevent slider drag
                              viewAttachment(attachment, currentItem?.Id);
                            }}
                            title={`View ${attachment?.FileName}`}
                            className={styles?.viewButton}
                          />
                          <span 
                            style={{ cursor: 'pointer', flex: 1, fontSize: '12px' }}
                            onClick={(e) => {
                              e.stopPropagation(); // Prevent slider drag
                              viewAttachment(attachment, currentItem?.Id);
                            }}
                            title="Click to view"
                          >
                            {attachment?.FileName}
                          </span>
                        </Stack>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {/* Bottom row with Date, Source, and View Details button */}
              <div className={styles.slideFooter}>
                <div className={styles.slideMetadata}>
                  <span className={styles.slideDate}>
                    Date: {new Date(currentItem?.PublishingDate).toLocaleDateString()}
                  </span>
                  <span className={styles.slideSource}>
                    Source: {currentItem?.Source}
                  </span>
                </div>
                <PrimaryButton
                  text="View Details"
                  onClick={(e) => {
                    e.stopPropagation(); // Prevent slider drag
                    openItemModal(currentItem);
                  }}
                  className={styles.viewDetailsButton}
                />
              </div>
            </div>
          </div>

          <IconButton
            iconProps={{ iconName: 'ChevronRight' }}
            onClick={nextSlide}
            className={styles.sliderArrow}
            disabled={currentItems.length <= 1}
          />
        </div>

        <div className={styles.sliderIndicators}>
          {currentItems.map((_, index) => (
            <button
              key={index}
              className={`${styles.indicator} ${index === sliderIndex ? styles.active : ''}`}
              onClick={() => setSliderIndex(index)}
            />
          ))}
        </div>

        {/* Auto-slide controls */}
        {props.slideAfter && parseInt(props.slideAfter.toString()) > 0 && (
          <div className={styles.autoSlideControls}>
            <IconButton
              iconProps={{ iconName: isSliderPaused ? 'Play' : 'Pause' }}
              onClick={() => setIsSliderPaused(!isSliderPaused)}
              title={isSliderPaused ? 'Resume auto-slide' : 'Pause auto-slide'}
              className={styles.pauseButton}
            />
          </div>
        )}
      </div>
    );
  };

  if (loading) {
    return <div>Loading...</div>;
  }

  return (
    <GlobalLoaderProvider>
    <FluentProvider theme={webLightTheme}>
    <div className={styles.postsArchive}>
      {props.viewType === 'table' ? renderTableView() : renderSliderView()}
      <AddEditForm 
        isOpen={isModalOpen || isAddPanelOpen} 
        Context={props?.Context} 
        listId={props?.listId} 
        onDismiss={closeModal} 
        onSave={() => {
          loadItems(); // Refresh items after save
          closeModal();
        }} 
        item={selectedItem} 
      />
      
      {/* File Viewer Modal */}
      <AttachmentViewer isOpen={isViewerOpen} currentFile={currentViewFile} onClose={closeViewer}/>
    </div>
    </FluentProvider>

 </GlobalLoaderProvider>
  );
};
const PostsArchive: React.FC<IPostsArchiveProps> = (props) => {
  return (
    <GlobalLoaderProvider>
      <PostsArchiveContent {...props} />
    </GlobalLoaderProvider>
  );
};

export default PostsArchive;