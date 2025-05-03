import React, { createContext, useState, useContext } from 'react';

export const LoadingContext = createContext<{
  isLoading: boolean;
  setIsLoading: (val: boolean) => void;
}>({
  isLoading: false,
  setIsLoading: () => {}
});

export const LoadingProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [isLoading, _setIsLoading] = useState(false);

  // Wrapped version with logging
  const setIsLoading = (val: boolean): void => {
    console.log('[LoadingContext] setIsLoading called with:', val);
    _setIsLoading(val);
  };

  return (
    <LoadingContext.Provider value={{ isLoading, setIsLoading }}>
      {children}
    </LoadingContext.Provider>
  );
};

export const useLoading = (): { isLoading: boolean; setIsLoading: (val: boolean) => void } => {
  const context = useContext(LoadingContext);
  if (!context) {
    throw new Error('useLoading must be used within a LoadingProvider');
  }
  return context;
};
