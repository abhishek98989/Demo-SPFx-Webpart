// ICustomDocumentSearchProps.ts
export interface ICustomDocumentSearchProps {
  description: string;
  context: any;

  /** New: controls UI mode from property pane */
  searchMode: 'SearchBar' | 'SearchWithResult';

  // default sample props (keep them if already there)
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
