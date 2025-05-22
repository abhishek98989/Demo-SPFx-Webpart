import * as React from 'react';
import { Spinner, SpinnerSize, Overlay, Text } from '@fluentui/react';
import { mergeStyles } from '@fluentui/react/lib/Styling';

export interface IGlobalLoaderProps {
  isLoading: boolean;
  loadingText?: string;
  spinnerSize?: SpinnerSize;
}

const overlayStyles = mergeStyles({
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

const spinnerContainerStyles = mergeStyles({
  display: 'flex',
  flexDirection: 'column',
  alignItems: 'center',
  gap: '16px',
  padding: '24px',
  backgroundColor: 'white',
  borderRadius: '8px',
  boxShadow: '0 4px 16px rgba(0, 0, 0, 0.12)',
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
  spinnerSize = SpinnerSize.large
}) => {
  if (!isLoading) {
    return null;
  }

  return (
    <div className={overlayStyles}>
      <div className={spinnerContainerStyles}>
        <Spinner 
          size={spinnerSize} 
          ariaLabel={loadingText}
        />
        {loadingText&&<Text className={loadingTextStyles}>
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

export const GlobalLoaderProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const loaderState = useGlobalLoader();

  return (
    <GlobalLoaderContext.Provider value={loaderState}>
      {children}
      <GlobalLoader 
        isLoading={loaderState.isLoading} 
        loadingText={loaderState.loadingText}
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