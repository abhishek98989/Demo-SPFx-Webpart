import * as React from 'react';
import { Spinner, SpinnerSize, Overlay, Text } from '@fluentui/react';
import { mergeStyles } from '@fluentui/react/lib/Styling';

export interface IGlobalLoaderProps {
  isLoading: boolean;
  loadingText?: string;
  spinnerSize?: SpinnerSize;
  isGlobal?: boolean; // New prop to control global vs container scoping
}

// Global overlay styles (covers entire viewport)
const globalOverlayStyles = mergeStyles({
  position: 'fixed',
  top: 0,
  left: 0,
  width: '100%',
  height: '100%',
  backgroundColor: 'rgba(255, 255, 255, 0.8)',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  zIndex: 9999,
  flexDirection: 'column'
});

// Container-scoped overlay styles (covers only the parent container)
const containerOverlayStyles = mergeStyles({
  position: 'absolute',
  top: 0,
  left: 0,
  width: '100%',
  height: '100%',
  backgroundColor: 'rgba(255, 255, 255, 0.8)',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  zIndex: 999,
  flexDirection: 'column'
});

const spinnerContainerStyles = mergeStyles({
  display: 'flex',
  flexDirection: 'column',
  alignItems: 'center',
  gap: '16px',
  padding: '24px',
  borderRadius: '8px',
  minWidth: '200px'
});

const loadingTextStyles = mergeStyles({
  fontSize: '16px',
  fontWeight: '400',
  color: '#323130'
});

export const GlobalLoader: React.FC<IGlobalLoaderProps> = ({
  isLoading,
  loadingText = 'Loading...',
  spinnerSize = SpinnerSize.large,
  isGlobal = false // Default to container-scoped
}) => {
  if (!isLoading) {
    return null;
  }

  const overlayClass = isGlobal ? globalOverlayStyles : containerOverlayStyles;

  return (
    <div className={overlayClass}>
      <div className={spinnerContainerStyles}>
        <Spinner 
          size={spinnerSize} 
          ariaLabel={loadingText}
        />
        {loadingText && <Text className={loadingTextStyles}>
          {loadingText}
        </Text>}
      </div>
    </div>
  );
};

// Hook for managing global loader state
export const useGlobalLoader = () => {
  const [isLoading, setIsLoading] = React.useState<boolean>(false);
  const [loadingText, setLoadingText] = React.useState<string>('Loading...');

  const showLoader = React.useCallback((text?: string) => {
    if (text) {
      setLoadingText(text);
    }
    setIsLoading(true);
  }, []);

  const hideLoader = React.useCallback(() => {
    setIsLoading(false);
  }, []);

  return {
    isLoading,
    loadingText,
    showLoader,
    hideLoader
  };
};

// Context for global loader state management
export interface IGlobalLoaderContext {
  isLoading: boolean;
  loadingText: string;
  showLoader: (text?: string) => void;
  hideLoader: () => void;
}

export const GlobalLoaderContext = React.createContext<IGlobalLoaderContext | undefined>(undefined);

export const GlobalLoaderProvider: React.FC<{ 
  children: React.ReactNode;
  isGlobal?: boolean; // Allow configuration at provider level
}> = ({ children, isGlobal = false }) => {
  const loaderState = useGlobalLoader();

  return (
    <GlobalLoaderContext.Provider value={loaderState}>
      {children}
      <GlobalLoader 
        isLoading={loaderState.isLoading} 
        loadingText={loaderState.loadingText}
        isGlobal={isGlobal}
      />
    </GlobalLoaderContext.Provider>
  );
};

// Hook to use the global loader context
export const useGlobalLoaderContext = (): IGlobalLoaderContext => {
  const context = React.useContext(GlobalLoaderContext);
  if (!context) {
    throw new Error('useGlobalLoaderContext must be used within a GlobalLoaderProvider');
  }
  return context;
};

// Wrapper component to easily add container-scoped loading to any element
export const LoadingContainer: React.FC<{
  children: React.ReactNode;
  isLoading: boolean;
  loadingText?: string;
  spinnerSize?: SpinnerSize;
  style?: React.CSSProperties;
  className?: string;
}> = ({ 
  children, 
  isLoading, 
  loadingText, 
  spinnerSize, 
  style, 
  className 
}) => {
  return (
    <div 
      style={{ position: 'relative', ...style }} 
      className={className}
    >
      {children}
      <GlobalLoader 
        isLoading={isLoading}
        loadingText={loadingText}
        spinnerSize={spinnerSize}
        isGlobal={false}
      />
    </div>
  );
};