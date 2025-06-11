import * as React from 'react';
import { useState, useRef } from 'react';
import {
  Modal,
  IconButton,
  Text
} from '@fluentui/react';
import styles from './PostsArchive.module.scss';

interface IAttachmentViewerProps {
  isOpen: boolean;
  currentFile: {
    url: string;
    name: string;
    type: string;
  } | null;
  onClose: () => void;
}

const AttachmentViewer: React.FC<IAttachmentViewerProps> = ({
  isOpen,
  currentFile,
  onClose
}) => {
  // Image viewer states
  const [zoomLevel, setZoomLevel] = useState<number>(1);
  const [panPosition, setPanPosition] = useState<{ x: number; y: number }>({ x: 0, y: 0 });
  const [isPanning, setIsPanning] = useState<boolean>(false);
  const [panStart, setPanStart] = useState<{ x: number; y: number }>({ x: 0, y: 0 });
  const imageViewerRef = useRef<HTMLDivElement>(null);

  // Reset states when modal closes or file changes
  React.useEffect(() => {
    if (!isOpen || !currentFile) {
      setZoomLevel(1);
      setPanPosition({ x: 0, y: 0 });
      setIsPanning(false);
    }
  }, [isOpen, currentFile]);

  const handleClose = (): void => {
    // Clean up blob URL if it exists
    if (currentFile?.url.startsWith('blob:')) {
      URL.revokeObjectURL(currentFile.url);
    }
    
    // Reset zoom and pan
    setZoomLevel(1);
    setPanPosition({ x: 0, y: 0 });
    setIsPanning(false);
    
    onClose();
  };

  // Image viewer functions
  const handleZoomIn = (): void => {
    setZoomLevel(prev => Math.min(prev + 0.10, 3)); // Max zoom 300%
  };

  const handleZoomOut = (): void => {
    setZoomLevel(prev => {
      const newZoom = Math.max(prev - 0.10, 0.5); // Min zoom 50%
      // Reset pan if zooming out to fit
      if (newZoom <= 1) {
        setPanPosition({ x: 0, y: 0 });
      }
      return newZoom;
    });
  };

  const handleZoomReset = (): void => {
    setZoomLevel(1);
    setPanPosition({ x: 0, y: 0 });
  };

  const handleImagePanStart = (e: React.MouseEvent | React.TouchEvent): void => {
    if (zoomLevel <= 1) return; // Only allow panning when zoomed in
    
    setIsPanning(true);
    const clientX = 'touches' in e ? e.touches[0].clientX : e.clientX;
    const clientY = 'touches' in e ? e.touches[0].clientY : e.clientY;
    
    setPanStart({
      x: clientX - panPosition.x,
      y: clientY - panPosition.y
    });
  };

  const handleImagePanMove = (e: React.MouseEvent | React.TouchEvent): void => {
    if (!isPanning || zoomLevel <= 1) return;
    
    const clientX = 'touches' in e ? e.touches[0].clientX : e.clientX;
    const clientY = 'touches' in e ? e.touches[0].clientY : e.clientY;
    
    setPanPosition({
      x: clientX - panStart.x,
      y: clientY - panStart.y
    });
  };

  const handleImagePanEnd = (): void => {
    setIsPanning(false);
  };

  const handleImageWheel = (e: React.WheelEvent): void => {
    e.preventDefault();
    
    if (e.deltaY < 0) {
      handleZoomIn();
    } else {
      handleZoomOut();
    }
  };

  if (!currentFile) return null;

  const isImage = currentFile.type === 'image';

  return (
    <Modal
      isOpen={isOpen}
      onDismiss={handleClose}
      isBlocking={true}
      containerClassName={styles.fileViewerModal}
    >
      <div className={styles.fileViewerContainer}>
        <div className={styles.fileViewerHeader}>
          <Text variant="large">{currentFile.name}</Text>
          <div className={styles.fileViewerControls}>
            {isImage && (
              <>
                <IconButton
                  iconProps={{ iconName: 'ZoomIn' }}
                  onClick={handleZoomIn}
                  title="Zoom In"
                  disabled={zoomLevel >= 3}
                />
                <IconButton
                  iconProps={{ iconName: 'ZoomOut' }}
                  onClick={handleZoomOut}
                  title="Zoom Out"
                  disabled={zoomLevel <= 0.5}
                />
                <IconButton
                  iconProps={{ iconName: 'FitPage' }}
                  onClick={handleZoomReset}
                  title="Reset Zoom"
                />
                <Text variant="small" style={{ margin: '0 8px' }}>
                  {Math.round(zoomLevel * 100)}%
                </Text>
              </>
            )}
            <IconButton
              iconProps={{ iconName: 'Cancel' }}
              onClick={handleClose}
              title="Close"
            />
          </div>
        </div>
        <div 
          className={styles.fileViewerContent}
          ref={imageViewerRef}
          style={{
            overflow: zoomLevel > 1 ? 'hidden' : 'auto',
            cursor: zoomLevel > 1 ? (isPanning ? 'grabbing' : 'grab') : 'default'
          }}
        >
          {isImage ? (
            <img
              src={currentFile.url}
              alt={currentFile.name}
              style={{ 
                maxWidth: zoomLevel <= 1 ? '100%' : 'none',
                maxHeight: zoomLevel <= 1 ? '80vh' : 'none',
                objectFit: zoomLevel <= 1 ? 'contain' : 'none',
                transform: `scale(${zoomLevel}) translate(${panPosition.x}px, ${panPosition.y}px)`,
                transformOrigin: 'center center',
                transition: isPanning ? 'none' : 'transform 0.2s ease',
                userSelect: 'none'
              }}
              onMouseDown={handleImagePanStart}
              onMouseMove={handleImagePanMove}
              onMouseUp={handleImagePanEnd}
              onMouseLeave={handleImagePanEnd}
              onTouchStart={handleImagePanStart}
              onTouchMove={handleImagePanMove}
              onTouchEnd={handleImagePanEnd}
              onWheel={handleImageWheel}
              draggable={false}
            />
          ) : (
            <iframe
              src={currentFile.url}
              title={currentFile.name}
              style={{ width: '100%', height: '80vh', border: 'none' }}
            />
          )}
        </div>
      </div>
    </Modal>
  );
};

export default AttachmentViewer;