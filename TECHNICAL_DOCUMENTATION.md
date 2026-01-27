# Technical Documentation 

## Overview
This repository is a SharePoint Framework (SPFx) solution named `vaughnconstruction` that ships multiple client-side web parts built with React. The solution bundles 15 web parts, uses PnPjs for SharePoint data access, and includes several rich-text/editor and UI libraries.

## Technology Stack
- SPFx build toolchain: 1.18.2 (`@microsoft/sp-build-web`).
- SPFx runtime packages: 1.21.1 (`@microsoft/sp-*` dependencies).
- React: 17.0.1
- TypeScript: 4.7.4
- UI: Fluent UI v8, Office UI Fabric v7, Bootstrap 5
- Data access: PnPjs (`@pnp/sp`)
- Editors: CKEditor 5, TinyMCE, Toast UI, Jodit, RoosterJS
- Utilities: `xlsx`, `xml2js`, `dompurify`

## Repository Layout
- `src/`
  - `index.ts`: SPFx entry point.
  - `globalCommon/`: shared components, styles, and helpers.
  - `webparts/`: individual web parts (each has a manifest, webpart class, components, and localization).
- `config/`: SPFx build and packaging configuration.
- `lib/`: compiled JS output (generated).
- `dist/`, `temp/`, `release/`, `sharepoint/solution/`: build artifacts (generated).

## Web Parts
The solution bundles the following web parts (from `config/config.json` and manifest titles):

| Web part | Folder | Manifest title |
| --- | --- | --- |
| LearnerBlogPost | `src/webparts/learnerBlogPost` | LearnerBlogPost |
| MarketingCalendar | `src/webparts/marketingCalendar` | MarketingCalendar |
| PostsArchive | `src/webparts/postsArchive` | PostsArchive |
| ConsolidatedCalendar | `src/webparts/consolidatedCalendar` | ConsolidatedCalendar |
| WeeklyWords | `src/webparts/weeklyWords` | weeklyWords |
| SharepointFeedbackForm | `src/webparts/sharepointFeedbackForm` | sharepointFeedbackForm |
| PeopleLibrary | `src/webparts/peopleLibrary` | PeopleLibrary |
| CorporateResponsibilities | `src/webparts/corporateResponsibilities` | CorporateResponsibilities |
| Locations | `src/webparts/locations` | Locations |
| ResourceLibrary | `src/webparts/resourceLibrary` | ResourceLibrary |
| WeeklyWordsManagement | `src/webparts/weeklyWordsManagement` | WeeklyWordsManagement |
| WeeklyWordsViewer | `src/webparts/weeklyWordsViewer` | WeeklyWordsViewer |
| LessonLearnedTable | `src/webparts/lessonLearnedTable` | LessonLearnedTable |
| WeeklyWordsPreview | `src/webparts/weeklyWordsPreview` | WeeklyWordsPreview |
| CustomDocumentSearch | `src/webparts/customDocumentSearch` | CustomDocumentSearch |

Each web part directory follows a standard SPFx pattern:
- `<Name>WebPart.ts`: web part class (initialization, property pane, render).
- `components/`: React components for the UI.
- `loc/`: localization strings.
- `<Name>WebPart.manifest.json`: metadata and manifest definition.

### LearnerBlogPost
Technical:
- Data source: Lessons Learned posts list via PnPjs (`PostsListId`), loads a single item by query param `LessonsLearnedId`.
- Fields rendered: Title, Author, PublishedDate, CSI divisions, Contacts, and HTML body sanitized with DOMPurify.
- External CSS: loads `ModernSPFxStyle.css` from the tenant (`SPComponentLoader.loadCss`).
Usability:
- Use on a details page that is navigated to from a list/table view.
- Ensure the URL includes `?LessonsLearnedId=<id>` or a "No blog post found" message appears.
- Configure the posts list ID in web part properties or in the web part's property bag.

### LessonLearnedTable
Technical:
- Data source: Lessons Learned posts list (`postsListId`), CSI divisions list hard-coded to `90df1643-9bbd-4dfa-bc72-327e6263048f`.
- Views: toggle between card and table view; supports search, sorting, pagination, and CSI division filtering.
- Permissions: checks `EditListItems` to enable editing links.
Usability:
- Configure the posts list ID in the property pane.
- Users can filter by CSI divisions, search full content, and switch between views.
- Edit links appear only for users with list edit permissions.

### Locations
Technical:
- Data source: `Locations` list on root site (`450dda32-a7e7-4439-8779-e1cc1523e7fd`).
- Displays `Office/Project`, `Address`, `Phone`, `Fax` in a sortable table.
- Permissions: checks `EditListItems` and links header to the list when allowed.
Usability:
- Intended as a read-only directory for most users.
- Edit users can jump to the list from the header link.

### CorporateResponsibilities
Technical:
- Data source: `WhoDoesWhat` list (`8c97c9e1-ecaf-4fcb-93a9-4016509d185c`).
- Shows `Title` (What) and `Who` fields in `GlobalTable` with optional link to list items.
- Permissions: checks `EditListItems` to enable clickable rows and header link.
Usability:
- Use as a read-only responsibility matrix for general users.
- Edit users can open list items directly for updates.

### PeopleLibrary
Technical:
- Data source: Microsoft Graph `/users` with pagination; requires `User.Read.All` and `User.ReadBasic.All` permissions.
- Filters: excludes guests and disabled users; search by name, email, phone, location, title.
- UI: paginated table with sorting, alphabet filter, refresh, and LivePersona hover cards with cached photos.
Usability:
- Use for a company-wide directory; supports quick search and browsing.
- If Graph permissions are not approved, the web part will show an error state.

### ResourceLibrary
Technical:
- Data source: hard-coded site `https://vaughnconstruction.sharepoint.com` and list ID `8177736d-9faa-49cc-82b3-4ff5a93fa02e`.
- Features: search across title/name/category/extension, A-Z filter, sorting, pagination, copy link, open in new tab.
- Advanced filters: modified date range filtering.
- Reporting: "Share report" panel with PeoplePicker and a Power Automate HTTP trigger URL in `DocumentLibrary.tsx`.
Usability:
- Use as a searchable document index for training or resource content.
- Update the Power Automate URL and default recipients to enable reporting.
- Users can copy links or open documents directly from the grid.

### SharepointFeedbackForm
Technical:
- Data source: feedback list ID `b1d451db-4d9e-40e7-af56-9cf204d81664`.
- Form writes `Title` and `Feedback` fields; uses current user name/email from context.
- UI: collapsible contact form with inline validation and success/error messages.
Usability:
- Use on intranet pages to capture SharePoint support requests.
- Minimal configuration; ensure list exists and users have add item rights.

### PostsArchive
Technical:
- Data source: list ID from property pane (`listId`), items include attachments and images.
- Views: `table` (search + region filters + file viewer) and `slider` (auto-rotate, keyboard, touch/drag).
- Editing: add/edit panel for items and attachments; uses PnPjs and SPHttpClient.
- Property pane: list ID, view type, number of events, slide interval.
Usability:
- Use table view for full archive search/browse; slider view for homepage highlights.
- Authors can add or update items with images and attachments from the UI.

### MarketingCalendar
Technical:
- Data source: list ID and site URL from property pane.
- Calendar types: `Marketing Calendar`, `Marketing Calendar-Internal`, `Company Calendar` (affects categories/colors).
- UI: react-big-calendar with month/week/day views, category filter, date panel, add/edit/delete events.
- Permissions: checks list permissions for add/edit/delete.
Usability:
- Set list ID and site URL for the target calendar.
- Users can view by category and, if permitted, create/update events.

### ConsolidatedCalendar
Technical:
- Data source: multiple calendar lists across sites via `siteCalendarCombos` (site URL + calendar ID).
- Merges events into a single calendar, color-coded by site.
- Supports viewing and editing via `ConsolidatedEventForm`.
Usability:
- Configure multiple site/calendar pairs in the property pane custom field.
- Use as a unified view of department or regional calendars.

### WeeklyWords
Technical:
- Data source: list ID and site URL from property pane.
- Fetches the latest approved article where `ArticleDate <= today`.
- Renders content and optional banners; error and empty states included.
Usability:
- Place on a home page to surface the most recent approved article.
- Ensure moderation status is set to approved for items to appear.

### WeeklyWordsManagement
Technical:
- Data source: list ID and site URL from property pane.
- CRUD UI with Jodit editor for HTML content; supports image upload and picker from `Publishing Images` library.
- Fields: Title, VaughnContent, Department, PublishingSource, Publishing Rollup Image, Abstract, PublishingContact, InCaseYouMissed, ArticleDate, Published.
- Includes PeoplePicker and list item recycling for delete.
Usability:
- Use by content editors to manage Weekly Words and department-specific posts.
- Requires edit permissions on the list and image library.

### WeeklyWordsPreview
Technical:
- Data source: list ID + site URL from property pane.
- Dropdown-based preview of available articles, sorted by date.
- Displays HTML content with formatted date/time and error handling.
Usability:
- Use for reviewers to validate content before publishing.
- Select a post from the dropdown to preview full content.

### WeeklyWordsViewer
Technical:
- Data source: list ID + site URL from property pane.
- View modes: `Tiles`, `Post`, `Redirected` (uses `WeeklyWordsId` from URL).
- Filters by `PublishingSource` and optional `Department`.
- Uses Tailwind CSS CDN and CKEditor CSS for styling.
Usability:
- Tiles view for homepage rollups; Post view for latest article; Redirected view for deep links.
- Link format: `/SitePages/Weekly-Words-Viewer.aspx?WeeklyWordsId=<id>`.

### CustomDocumentSearch
Technical:
- Data sources: SharePoint search + Graph search across:
  - Resource Library (`8177736d-9faa-49cc-82b3-4ff5a93fa02e`)
  - Weekly Words (`b2b9bc28-8e97-4716-9b88-3c456f8d2b6a`) at `/news`
  - Lessons Learned (`42F972A2-3D9D-4DBE-86C8-32EBAF4ACFDB`) at `/ll`
  - IT Blogs (`7dd1ac79-1372-41a6-972f-0e6eca0d9f0e`) at `/sites/ITSite`
  - Company Calendar (`3c14f133-7cb6-4582-ac0a-4e5c401a5952`)
  - Corporate Responsibilities (`8c97c9e1-ecaf-4fcb-93a9-4016509d185c`)
  - Marketing Calendar (`9be28777-0981-4627-a366-779f280711ef`)
  - Locations (`450dda32-a7e7-4439-8779-e1cc1523e7fd`)
  - People Search (SharePoint People result source)
- Supports exact/all/any matching, highlighting, and LivePersona with photo lookup.
- Property pane: `searchMode` toggles `SearchBar` (redirect to results page) or `SearchWithResult` (inline results).
Usability:
- Use SearchBar mode for global search entry points.
- Use SearchWithResult mode for intranet pages needing inline results and filters.
- Requires Graph permissions to show people and certain list sources.

## Shared Components and Utilities
`src/globalCommon/` contains reusable UI and helpers used across web parts:
- `GlobalTable.tsx`: shared table component.
- `customLoader.tsx`: loading indicator UI.
- `Banners.tsx`: reusable banner/notice UI.
- `EventRecurrenceControls/` and `reccurenceStringToText.ts`: recurrence controls and helpers.
- `fluent.css`, `style.css`: shared styles.

## Configuration and Packaging
Key configuration files:
- `config/config.json`: maps bundle names to each web part entrypoint and manifest.
- `config/package-solution.json`: solution metadata and output package name.
- `config/serve.json`: local workbench URL and dev server port (4321, https).

Solution packaging details from `config/package-solution.json`:
- Package: `sharepoint/solution/vaughnconstruction.sppkg`
- `includeClientSideAssets`: `true`
- `skipFeatureDeployment`: `true`
- Graph permissions requested:
  - `User.Read.All`
  - `User.ReadBasic.All`

## Build, Serve, and Package
Prerequisites:
- Node.js `>=16.13.0 <17.0.0` or `>=18.17.1 <19.0.0`
- npm (via Node)

Common commands:
```bash
npm install
gulp serve
```

Production bundle and package:
```bash
gulp clean
gulp bundle --ship
gulp package-solution --ship
```

The packaged `.sppkg` will be in `sharepoint/solution/`.

## Deployment (SharePoint App Catalog)
1. Run the production packaging commands above.
2. Upload `sharepoint/solution/vaughnconstruction.sppkg` to the tenant app catalog.
3. Approve API permission requests in the SharePoint admin center (if required).
4. Add web parts to pages from the SharePoint toolbox.

## Localization
Each web part contains `loc/` resources referenced in `config/config.json` under `localizedResources`.

## Linting and Tests
The repository uses ESLint via SPFx tooling:
- `npm run test` runs `gulp test`.
- `@microsoft/eslint-config-spfx` is configured for React + SPFx.

## Generated vs. Source Files
- Source code lives in `src/`.
- `lib/`, `dist/`, `temp/`, `release/`, and `sharepoint/solution/` are generated outputs and should not be edited directly.

## Notes on SPFx Versions
Build tooling references SPFx 1.18.2 packages, while runtime dependencies in `package.json` reference SPFx 1.21.1 packages. If you upgrade tooling or dependencies, align versions across the toolchain to avoid incompatibilities.
