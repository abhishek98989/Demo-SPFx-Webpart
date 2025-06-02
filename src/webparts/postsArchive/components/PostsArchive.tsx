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
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IPostsArchiveProps } from './IPostsArchiveProps';
import { IArchiveItem } from '../PostsArchiveWebPart';
import styles from './PostsArchive.module.scss';
import AddEditForm from './AddEditForm';
import { SPFx, spfi } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

const PostsArchive: React.FC<IPostsArchiveProps> = (props) => {
  const [items, setItems] = useState<IArchiveItem[]>([]);
  const sp = spfi().using(SPFx(props?.Context));
  const [filteredItems, setFilteredItems] = useState<IArchiveItem[]>([]);
  const [searchText, setSearchText] = useState<string>('');
  const [selectedRegions, setSelectedRegions] = useState<string[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [selectedItem, setSelectedItem] = useState<IArchiveItem | any>(null);
  const [isModalOpen, setIsModalOpen] = useState<boolean>(false);
  const [isAddPanelOpen, setIsAddPanelOpen] = useState<boolean>(false);
  const [sliderIndex, setSliderIndex] = useState<number>(0);
  
  // Auto-slider states
  const intervalRef = useRef<NodeJS.Timeout | null>(null);
  const [isSliderPaused, setIsSliderPaused] = useState<boolean>(false);
  
  // File viewer states
  const [isViewerOpen, setIsViewerOpen] = useState<boolean>(false);
  const [currentViewFile, setCurrentViewFile] = useState<{ url: string; name: string; type: string } | null>(null);

  const regionOptions: IDropdownOption[] = [
    { key: 'North America', text: 'North America' },
    { key: 'Europe', text: 'Europe' },
    { key: 'Asia Pacific', text: 'Asia Pacific' },
    { key: 'Latin America', text: 'Latin America' },
    { key: 'Middle East & Africa', text: 'Middle East & Africa' },
    { key: 'Global', text: 'Global' }
  ];

  useEffect(() => {
    loadItems();
  }, []);

  useEffect(() => {
    filterItems();
  }, [items, searchText, selectedRegions]);

  // Auto-slider effect
  useEffect(() => {
    if (props.viewType === 'slider' && props.slideAfter && !isSliderPaused && !isViewerOpen) {
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
  }, [props.viewType, props.slideAfter, isSliderPaused, isViewerOpen, sliderIndex]);

  // Pause slider when viewer opens
  useEffect(() => {
    if (isViewerOpen) {
      setIsSliderPaused(true);
    } else {
      setIsSliderPaused(false);
    }
  }, [isViewerOpen]);

  const loadItems = async (): Promise<void> => {
    try {
      const items = await sp.web.lists
        .getById(props?.listId)
        .items
        .select(
          "Id", "Title", "PublishingDate", "Source", "Description", "Region", 
          "Images", "Attachments", "AttachmentFiles"
        )
        .expand("AttachmentFiles")
        .orderBy("PublishingDate", false)();
  
      setItems(items as IArchiveItem[]);
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
    const maxIndex = Math.min(props.numberOfEvents, filteredItems.length) - 1;
    setSliderIndex(prev => (prev >= maxIndex ? 0 : prev + 1));
  };

  const prevSlide = (): void => {
    const maxIndex = Math.min(props.numberOfEvents, filteredItems.length) - 1;
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

  const viewAttachment = async (attachment: any, itemId: number): Promise<void> => {
    try {
      const fileUrl = `${props.Context.pageContext.web.absoluteUrl}/_api/web/lists('${props.listId}')/items(${itemId})/AttachmentFiles('${attachment.FileName}')/$value`;
      const fileExtension = attachment.FileName.split('.').pop()?.toLowerCase();
      const fileType = fileExtension === 'pdf' ? 'application/pdf' : 'image';
      
      setCurrentViewFile({
        url: fileUrl,
        name: attachment.FileName,
        type: fileType
      });
      setIsViewerOpen(true);
    } catch (error) {
      console.error('Error viewing file:', error);
    }
  };

  const closeViewer = (): void => {
    setIsViewerOpen(false);
    setCurrentViewFile(null);
  };

  const truncateText = (text: string, maxLength: number = 200): string => {
    if (!text) return '';
    const plainText = text.replace(/<[^>]*>/g, ''); // Remove HTML tags
    return plainText.length > maxLength ? plainText.substring(0, maxLength) + '...' : plainText;
  };

  const renderFileViewer = (): JSX.Element => {
    if (!currentViewFile) return <></>;

    return (
      <Modal
        isOpen={isViewerOpen}
        onDismiss={closeViewer}
        isBlocking={true}
        containerClassName={styles.fileViewerModal}
      >
        <div className={styles.fileViewerContainer}>
          <div className={styles.fileViewerHeader}>
            <Text variant="large">{currentViewFile.name}</Text>
            <IconButton
              iconProps={{ iconName: 'Cancel' }}
              onClick={closeViewer}
              title="Close"
            />
          </div>
          <div className={styles.fileViewerContent}>
            {currentViewFile.type === 'image' ? (
              <img
                src={currentViewFile.url}
                alt={currentViewFile.name}
                style={{ maxWidth: '100%', maxHeight: '80vh', objectFit: 'contain' }}
              />
            ) : (
              <iframe
                src={currentViewFile.url}
                title={currentViewFile.name}
                style={{ width: '100%', height: '80vh', border: 'none' }}
              />
            )}
          </div>
        </div>
      </Modal>
    );
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
    const slidesToShow = Math.min(props.numberOfEvents, filteredItems.length);
    const currentItems = filteredItems.slice(0, slidesToShow);

    if (currentItems.length === 0) {
      return <div>No items to display</div>;
    }

    const currentItem = currentItems[sliderIndex];

    return (
      <div className={styles.sliderView}>
        <div className={styles.sliderContainer}>
          <IconButton
            iconProps={{ iconName: 'ChevronLeft' }}
            onClick={prevSlide}
            className={styles.sliderArrow}
            disabled={currentItems.length <= 1}
          />

          <div className={styles.slideContent}>
            <div className={styles.slideCard} style={{ maxHeight: '400px', overflow: 'hidden' }}>
              {/* Title */}
              <h3 className={styles.slideTitle}>{currentItem.Title}</h3>
              
              {/* Region */}
              <p className={styles.slideRegion}>{currentItem.Region}</p>
              
              {/* Description - Truncated */}
              {/* <div className={styles.slideDescription}>
                <div dangerouslySetInnerHTML={{ __html: currentItem?.Description }} />
              </div> */}

              {/* Images */}
              {currentItem.Images && (
                <div className={styles.slideImages}>
                  <div dangerouslySetInnerHTML={{ __html: currentItem.Images }} />
                </div>
              )}

              {/* Attachments */}
              {currentItem.AttachmentFiles && currentItem.AttachmentFiles.length > 0 && (
                <div className={styles.attachments}>
                  <Text variant="small" style={{ fontWeight: 600, marginBottom: '8px' }}>
                    Attachments:
                  </Text>
                  <div className={styles.attachmentList}>
                    {currentItem.AttachmentFiles.map((attachment, index) => (
                      <div key={index} className={styles.attachmentItem}>
                        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                          <IconButton
                            iconProps={{ iconName: getFileIcon(attachment.FileName) }}
                            onClick={() => viewAttachment(attachment, currentItem.Id)}
                            title={`View ${attachment.FileName}`}
                            className={styles.viewButton}
                          />
                          <span 
                            style={{ cursor: 'pointer', flex: 1, fontSize: '12px' }}
                            onClick={() => viewAttachment(attachment, currentItem.Id)}
                            title="Click to view"
                          >
                            {attachment.FileName}
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
                    Date: {new Date(currentItem.PublishingDate).toLocaleDateString()}
                  </span>
                  <span className={styles.slideSource}>
                    Source: {currentItem.Source}
                  </span>
                </div>
                <PrimaryButton
                  text="View Details"
                  onClick={() => openItemModal(currentItem)}
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
      {renderFileViewer()}
    </div>
  );
};

export default PostsArchive;