// import * as React from 'react';
// import { spfi, SPFx } from "@pnp/sp";
// import "@pnp/sp/lists";
// import "@pnp/sp/items";
// import "@pnp/sp/batching";
// import "@pnp/sp/webs";
// import "@pnp/sp/security";
// import { PrimaryButton, DetailsList, DetailsListLayoutMode, IColumn, MessageBar, MessageBarType } from "@fluentui/react";

// export const Scripting = (props: any) => {

//   const [items, setItems] = React.useState<any[]>([]);
//   const [status, setStatus] = React.useState<string>('');
//   const [loading, setLoading] = React.useState<boolean>(false);


//   // 🧭 Define table columns
//   const columns: IColumn[] = [
//     { key: 'col1', name: 'Title', fieldName: 'Title', minWidth: 150, maxWidth: 200, isResizable: true },
//     { key: 'col2', name: 'Name', fieldName: 'FileLeafRef', minWidth: 150, maxWidth: 200, isResizable: true },
//     { key: 'col3', name: 'Type', minWidth: 80, maxWidth: 100, isResizable: true,
//       onRender: (item) => (item.FSObjType === 1 ? '📁 Folder' : '📄 File')
//     },
//   ];

//   // 📥 Load items initially
//   const fetchItems = async () => {
//       const sp = spfi('https://vaughnconstruction.sharepoint.com').using(SPFx(props?.context));

//     try {
//       setLoading(true);
//       setStatus('Fetching items...');
//       const result = await sp.web.lists
//         .getById('8177736d-9faa-49cc-82b3-4ff5a93fa02e')
//         .items
//         .select('Id', 'Title', 'FileLeafRef', 'FSObjType')
//         .top(5000)();
//       setItems(result);
//       setStatus(`Fetched ${result.length} items`);
//     } catch (error: any) {
//       console.error('Error fetching items:', error);
//       setStatus('Error fetching items.');
//     } finally {
//       setLoading(false);
//     }
//   };

//   React.useEffect(() => {
//     fetchItems();
//   }, []);

//   // 🚀 Batch update logic
//   const updateTitlesInBatches = async () => {
//        const sp = spfi('https://vaughnconstruction.sharepoint.com').using(SPFx(props?.context));

//     try {
//       setLoading(true);
//       setStatus('Preparing to update missing titles...');

//       const itemsToUpdate = items.filter(
//         (item) => (!item.Title || item.Title.trim() === '') && item.FileLeafRef
//       );

//       if (itemsToUpdate.length === 0) {
//         setStatus('No items with missing titles found.');
//         setLoading(false);
//         return;
//       }

//       const chunkSize = 50;
//       for (let i = 0; i < itemsToUpdate.length; i += chunkSize) {
//         const chunk = itemsToUpdate.slice(i, i + chunkSize);
//         const batchIndex = Math.floor(i / chunkSize) + 1;
//         const totalBatches = Math.ceil(itemsToUpdate.length / chunkSize);

//         setStatus(`Processing batch ${batchIndex} of ${totalBatches}...`);

//         const [batchedSP, execute] = sp.batched();

//         for (const item of chunk) {
//           batchedSP.web.lists
//             .getById('8177736d-9faa-49cc-82b3-4ff5a93fa02e')
//             .items.getById(item.Id)
//             .update({
//               Title: item.FileLeafRef,
//             });
//         }

//         await execute();
//         console.log(`✅ Batch ${batchIndex} complete`);
//       }

//       setStatus('✅ All batches completed. Refreshing items...');
//       await fetchItems();
//     } catch (error: any) {
//       console.error('Error updating titles:', error);
//       setStatus(`❌ Error: ${error.message}`);
//     } finally {
//       setLoading(false);
//     }
//   };

//   return (
//     <div style={{ padding: 20 }}>
//       <h2>📂 Document Library Title Sync</h2>

//       {status && (
//         <MessageBar
//           messageBarType={status.startsWith('❌') ? MessageBarType.error : MessageBarType.info}
//           isMultiline={false}
//           style={{ marginBottom: 10 }}
//         >
//           {status}
//         </MessageBar>
//       )}

//       <div style={{ marginBottom: 10 }}>
//         <PrimaryButton
//           text="Start Update Missing Titles"
//           onClick={updateTitlesInBatches}
//           disabled={loading}
//         />
//       </div>

//       <DetailsList
//         items={items}
//         columns={columns}
//         layoutMode={DetailsListLayoutMode.fixedColumns}
//         compact={true}
//       />
//     </div>
//   );
// };
// import * as React from 'react';
// import { spfi, SPFx } from "@pnp/sp";
// import "@pnp/sp/webs";
// import "@pnp/sp/lists";
// import "@pnp/sp/items";
// import "@pnp/sp/files";
// import "@pnp/sp/folders";
// import { PrimaryButton, DetailsList, DetailsListLayoutMode, IColumn, MessageBar, MessageBarType } from "@fluentui/react";

// export const Scripting = (props: any) => {

//   const [items, setItems] = React.useState<any[]>([]);
//   const [status, setStatus] = React.useState<string>('');
//   const [loading, setLoading] = React.useState<boolean>(false);

//   // ⚙️ CONFIG — CHANGE THESE
//   const SOURCE_SITE_URL = "https://vaughnconstruction.sharepoint.com/sites/modernhr"; // 🔁 change
//   const DEST_SITE_URL = "https://vaughnconstruction.sharepoint.com"; // 🔁 change
//   const SOURCE_LIST_ID = "d8fce1ab-a7c3-4bfb-b6eb-5119bb8b83f4"; // 🔁 change
//   const DEST_LIST_ID = "8177736d-9faa-49cc-82b3-4ff5a93fa02e"; // 🔁 change

//   // 🧭 Define table columns
//   const columns: IColumn[] = [
//     { key: 'col1', name: 'Title', fieldName: 'Title', minWidth: 150, maxWidth: 200, isResizable: true },
//     { key: 'col2', name: 'File Name', fieldName: 'FileLeafRef', minWidth: 150, maxWidth: 250, isResizable: true },
//     { key: 'col3', name: 'Type', minWidth: 80, maxWidth: 100, isResizable: true,
//       onRender: (item) => (item.FSObjType === 1 ? '📁 Folder' : '📄 File')
//     },
//   ];

//   // 📥 Load all items from the source document library
//  const fetchSourceItems = async () => {
//     const spSource = spfi(SOURCE_SITE_URL).using(SPFx(props.context));

//     try {
//       setLoading(true);
//       setStatus("Fetching documents from source library...");

//       const result = await spSource.web.lists
//         .getById(SOURCE_LIST_ID)
//         .items
//         .select("Id", "Title", "FileLeafRef", "FileRef", "FSObjType", "File/ServerRelativeUrl")
//         .expand("File")
//         .top(5000)();

//       setItems(result);
//       setStatus(`✅ Fetched ${result.length} items from source library`);
//     } catch (error: any) {
//       console.error("❌ Error fetching items:", error);
//       setStatus(`❌ Error fetching items: ${error.message}`);
//     } finally {
//       setLoading(false);
//     }
//   };


//   React.useEffect(() => {
//     fetchSourceItems();
//   }, []);

//   // 🚀 Copy logic
//    const copyDocuments = async () => {
//     const spSource = spfi(SOURCE_SITE_URL).using(SPFx(props.context));
//     const spDest = spfi(DEST_SITE_URL).using(SPFx(props.context));

//     try {
//       setLoading(true);
//       setStatus("Preparing to copy documents...");

//       const filesToCopy = items.filter(item => item.FSObjType === 0); // only files
//       if (filesToCopy.length === 0) {
//         setStatus("No documents found to copy.");
//         setLoading(false);
//         return;
//       }

//       for (let i = 0; i < filesToCopy.length; i++) {
//         const file = filesToCopy[i];
//         const fileName = file.FileLeafRef;

//         setStatus(`📄 Copying (${i + 1}/${filesToCopy.length}): ${fileName} ...`);

//         try {
//           // Get file binary from source
//           const arrayBuffer = await spSource.web
//             .getFileByServerRelativePath(file.File.ServerRelativeUrl)
//             .getBuffer();

//           // Upload to destination
//           await spDest.web.lists
//             .getById(DEST_LIST_ID)
//             .rootFolder
//             .files
//             .addUsingPath(fileName, arrayBuffer, { Overwrite: true });

//           console.log(`✅ Copied: ${fileName}`);
//         } catch (fileError: any) {
//           console.error(`❌ Failed to copy ${fileName}:`, fileError);
//         }
//       }

//       setStatus("✅ All documents copied successfully!");
//     } catch (error: any) {
//       console.error("❌ Error copying documents:", error);
//       setStatus(`❌ Error: ${error.message}`);
//     } finally {
//       setLoading(false);
//     }
//   };

//   return (
//     <div style={{ padding: 20 }}>
//       <h2>📂 Copy Documents Between Libraries</h2>

//       {status && (
//         <MessageBar
//           messageBarType={status.startsWith('❌') ? MessageBarType.error : MessageBarType.info}
//           isMultiline={false}
//           style={{ marginBottom: 10 }}
//         >
//           {status}
//         </MessageBar>
//       )}

//       <div style={{ marginBottom: 10 }}>
//         <PrimaryButton
//           text="Start Copy Documents"
//           onClick={copyDocuments}
//           disabled={loading}
//         />
//       </div>

//       <DetailsList
//         items={items}
//         columns={columns}
//         layoutMode={DetailsListLayoutMode.fixedColumns}
//         compact={true}
//       />
//     </div>
//   );
// };
// // import * as React from 'react';
// // import { spfi, SPFx } from "@pnp/sp";
// // import "@pnp/sp/webs";
// // import "@pnp/sp/lists";
// // import "@pnp/sp/items";
// // import "@pnp/sp/files";
// // import "@pnp/sp/folders";
// // import { PrimaryButton, DetailsList, DetailsListLayoutMode, IColumn, MessageBar, MessageBarType } from "@fluentui/react";

// // export const Scripting = (props: any) => {

// //   const [items, setItems] = React.useState<any[]>([]);
// //   const [status, setStatus] = React.useState<string>('');
// //   const [loading, setLoading] = React.useState<boolean>(false);

// //   // ⚙️ CONFIG — CHANGE THESE
// //   const SOURCE_SITE_URL = "https://vaughnconstruction.sharepoint.com/news"; // 🔁 change
// //   const DEST_SITE_URL = "https://vaughnconstruction.sharepoint.com/sites/intranet"; // 🔁 change
// //   const SOURCE_LIST_ID = "216B6A79-4A9F-4E15-962E-23133411D96B"; // 🔁 change
// //   const DEST_LIST_ID = "4270cc97-6d4b-4406-8d01-63e15959946b"; // 🔁 change

// //   // 🧭 Define table columns
// //   const columns: IColumn[] = [
// //     { key: 'col1', name: 'Title', fieldName: 'Title', minWidth: 150, maxWidth: 200, isResizable: true },
// //     // { key: 'col2', name: 'File Name', fieldName: 'FileLeafRef', minWidth: 150, maxWidth: 250, isResizable: true },
// //     // { key: 'col3', name: 'Type', minWidth: 80, maxWidth: 100, isResizable: true,
// //     //   onRender: (item) => (item.FSObjType === 1 ? '📁 Folder' : '📄 File')
// //     // },
// //   ];

// //   // 📥 Load all items from the source document library
// //   const fetchSourceItems = async () => {
// //     const spSource = spfi(SOURCE_SITE_URL).using(SPFx(props.context));

// //     try {
// //       setLoading(true);
// //       setStatus("Fetching documents from source library...");

// //       const result = await spSource.web.lists
// //         .getById(SOURCE_LIST_ID)
// //         .items
// //         .select("Id", "Title", "ArticleDate", "VaughnContent", "PublishingContact/Title", "OData__ModerationStatus")
// //         .expand("PublishingContact")
// //         .top(5000)();

// //       setItems(result);
// //       setStatus(`✅ Fetched ${result.length} items from source library`);
// //     } catch (error: any) {
// //       console.error("❌ Error fetching items:", error);
// //       setStatus(`❌ Error fetching items: ${error.message}`);
// //     } finally {
// //       setLoading(false);
// //     }
// //   };


// //   React.useEffect(() => {
// //     fetchSourceItems();
// //   }, []);

// //   // 🚀 Copy logic
// //   const copyDocuments = async () => {
// //     const spSource = spfi(SOURCE_SITE_URL).using(SPFx(props.context));
// //     const spDest = spfi(DEST_SITE_URL).using(SPFx(props.context));

// //     try {
// //       setLoading(true);
// //       setStatus("Preparing to copy documents...");

// //       const filesToCopy = items
// //       if (filesToCopy.length === 0) {
// //         setStatus("No documents found to copy.");
// //         setLoading(false);
// //         return;
// //       }

// //       for (let i = 0; i < filesToCopy.length; i++) {
// //         const file = filesToCopy[i];
// //         const fileName = file.FileLeafRef;

// //         setStatus(`📄 Copying (${i + 1}/${filesToCopy.length}): ${fileName} ...`);

// //         try {
// //           // Get file binary from source
     
// //           // Upload to destination
// //           await spDest.web.lists
// //             .getById(DEST_LIST_ID)
// //             .items
// //             .add({
// //               Title: file.Title,
// //               ArticleDate: file.ArticleDate,
// //               VaughnContent: file.VaughnContent,
// //               PublishingContactId: file.PublishingContact ? file.PublishingContact.Id : null,
// //               OData__ModerationStatus: file.OData__ModerationStatus
// //             })

// //           console.log(`✅ Copied: ${fileName}`);
// //         } catch (fileError: any) {
// //           console.error(`❌ Failed to copy ${fileName}:`, fileError);
// //         }
// //       }

// //       setStatus("✅ All documents copied successfully!");
// //     } catch (error: any) {
// //       console.error("❌ Error copying documents:", error);
// //       setStatus(`❌ Error: ${error.message}`);
// //     } finally {
// //       setLoading(false);
// //     }
// //   };

// //   return (
// //     <div style={{ padding: 20 }}>
// //       <h2>📂 Copy Documents Between Libraries</h2>

// //       {status && (
// //         <MessageBar
// //           messageBarType={status.startsWith('❌') ? MessageBarType.error : MessageBarType.info}
// //           isMultiline={false}
// //           style={{ marginBottom: 10 }}
// //         >
// //           {status}
// //         </MessageBar>
// //       )}

// //       <div style={{ marginBottom: 10 }}>
// //         <PrimaryButton
// //           text="Start Copy Documents"
// //           onClick={copyDocuments}
// //           disabled={loading}
// //         />
// //       </div>

// //       <DetailsList
// //         items={items}
// //         columns={columns}
// //         layoutMode={DetailsListLayoutMode.fixedColumns}
// //         compact={true}
// //       />
// //     </div>
// //   );
// // };



import * as React from 'react';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";

import {
  PrimaryButton,
  DefaultButton,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  MessageBar,
  MessageBarType,
  Checkbox,
  Stack,
  Link,
  Selection,
  SelectionMode,
  Dialog,
  DialogType,
  DialogFooter
} from "@fluentui/react";

export const Scripting = (props: any) => {

  const [items, setItems] = React.useState<any[]>([]);
  const [duplicateFiles, setDuplicateFiles] = React.useState<any[]>([]);
  const [status, setStatus] = React.useState<string>('');
  const [loading, setLoading] = React.useState<boolean>(false);
  const [selectedItems, setSelectedItems] = React.useState<any[]>([]);
  const [showConfirmDialog, setShowConfirmDialog] = React.useState<boolean>(false);

  // ============================================
  // CONFIG
  // ============================================
  const SITE_URL = "https://vaughnconstruction.sharepoint.com";
  const LIBRARY_ID = "8177736d-9faa-49cc-82b3-4ff5a93fa02e";

  // ============================================
  // SELECTION
  // ============================================
  const selection = React.useMemo(() => {
    return new Selection({
      onSelectionChanged: () => {
        const selected = selection.getSelection();
        setSelectedItems(selected as any[]);
      }
    });
  }, []);

  // ============================================
  // COLUMNS
  // ============================================
  const columns: IColumn[] = [
    {
      key: 'col1',
      name: 'File Name',
      fieldName: 'FileLeafRef',
      minWidth: 220,
      maxWidth: 320,
      isResizable: true
    },
    {
      key: 'col2',
      name: 'Title',
      fieldName: 'Title',
      minWidth: 180,
      maxWidth: 250,
      isResizable: true
    },
    {
      key: 'col3',
      name: 'Folder Location',
      minWidth: 350,
      isResizable: true,
      onRender: (item) => {

        const folderPath = item.FileDirRef;

        return (
          <Link
            href={`${SITE_URL}${item.FileRef}`}
            target="_blank"
          >
            {folderPath}
          </Link>
        );
      }
    },
    {
      key: 'col4',
      name: 'Full File Path',
      minWidth: 400,
      isResizable: true,
      onRender: (item) => {

        return (
          <Link
            href={`${SITE_URL}${item.FileRef}`}
            target="_blank"
          >
            {item.FileRef}
          </Link>
        );
      }
    },
    {
      key: 'col5',
      name: 'Modified',
      fieldName: 'Modified',
      minWidth: 150,
      onRender: (item) => {
        return new Date(item.Modified).toLocaleString();
      }
    }
  ];

  // ============================================
  // FETCH LIBRARY ITEMS
  // ============================================
  const fetchItems = async () => {

    const sp = spfi(SITE_URL).using(SPFx(props.context));

    try {

      setLoading(true);
      setStatus('Fetching library items...');

      const result = await sp.web.lists
        .getById(LIBRARY_ID)
        .items
        .select(
          'Id',
          'Title',
          'FileLeafRef',
          'FileRef',
          'FileDirRef',
          'Modified',
          'FSObjType',
        )
        .top(5000)();

      // Only files
      const filesOnly = result.filter((item: any) => item.FSObjType === 0);

      setItems(filesOnly);

      setStatus(`Loaded ${filesOnly.length} files.`);

      findDuplicates(filesOnly);

    } catch (error: any) {

      console.error(error);
      setStatus(`❌ Error fetching items: ${error.message}`);

    } finally {

      setLoading(false);
    }
  };

  // ============================================
  // FIND DUPLICATES
  // ============================================
  const findDuplicates = (libraryItems: any[]) => {

    const groupedFiles: any = {};

    libraryItems.forEach((item) => {

      const fileName = item.FileLeafRef?.toLowerCase()?.trim();

      if (!groupedFiles[fileName]) {
        groupedFiles[fileName] = [];
      }

      groupedFiles[fileName].push(item);
    });

    const duplicates: any[] = [];

    Object.keys(groupedFiles).forEach((key) => {

      if (groupedFiles[key].length > 1) {

        groupedFiles[key].forEach((duplicateItem: any) => {
          duplicates.push(duplicateItem);
        });
      }
    });

    setDuplicateFiles(duplicates);

    setStatus(
      `✅ Found ${duplicates.length} duplicate files across the library.`
    );
  };

  React.useEffect(() => {
    fetchItems();
  }, []);

  // ============================================
  // DELETE SELECTED FILES
  // ============================================
  const deleteSelectedFiles = async () => {

    const sp = spfi(SITE_URL).using(SPFx(props.context));

    try {

      setLoading(true);
      setShowConfirmDialog(false);

      if (selectedItems.length === 0) {

        setStatus('Please select files to delete.');
        setLoading(false);
        return;
      }

      setStatus(`Deleting ${selectedItems.length} selected files...`);

      for (let i = 0; i < selectedItems.length; i++) {

        const file = selectedItems[i];

        setStatus(
          `Deleting ${i + 1} of ${selectedItems.length}: ${file.FileLeafRef}`
        );

        try {

          await sp.web
            .getFileByServerRelativePath(file.FileRef)
            .recycle();

          console.log(`Deleted: ${file.FileRef}`);

        } catch (fileError: any) {

          console.error(`Failed deleting ${file.FileRef}`, fileError);
        }
      }

      setStatus('✅ Selected duplicate files moved to recycle bin successfully.');

      await fetchItems();

    } catch (error: any) {

      console.error(error);
      setStatus(`❌ Error deleting files: ${error.message}`);

    } finally {

      setLoading(false);
    }
  };

  // ============================================
  // UI
  // ============================================
  return (

    <div style={{ padding: 20 }}>

      <h2>📂 Duplicate Files Manager</h2>

      {status && (
        <MessageBar
          messageBarType={
            status.startsWith('❌')
              ? MessageBarType.error
              : MessageBarType.info
          }
          isMultiline={false}
          style={{ marginBottom: 12 }}
        >
          {status}
        </MessageBar>
      )}

      <Stack
        horizontal
        tokens={{ childrenGap: 10 }}
        style={{ marginBottom: 15 }}
      >

        <PrimaryButton
          text="Refresh Duplicate Files"
          onClick={fetchItems}
          disabled={loading}
        />

        <DefaultButton
          text={`Delete Selected (${selectedItems.length})`}
          onClick={() => setShowConfirmDialog(true)}
          disabled={loading || selectedItems.length === 0}
        />

      </Stack>

      <div style={{ marginBottom: 10 }}>
        <strong>Total Duplicate Files:</strong> {duplicateFiles.length}
      </div>

      <DetailsList
        items={duplicateFiles}
        columns={columns}
        layoutMode={DetailsListLayoutMode.fixedColumns}
        compact={true}
        selection={selection}
        selectionMode={SelectionMode.multiple}
        checkboxVisibility={2}
      />

      {/* CONFIRM DELETE DIALOG */}
      <Dialog
        hidden={!showConfirmDialog}
        onDismiss={() => setShowConfirmDialog(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Delete Duplicate Files',
          subText: `Are you sure you want to delete ${selectedItems.length} selected duplicate files?`
        }}
      >

        <DialogFooter>

          <PrimaryButton
            text="Delete"
            onClick={deleteSelectedFiles}
          />

          <DefaultButton
            text="Cancel"
            onClick={() => setShowConfirmDialog(false)}
          />

        </DialogFooter>

      </Dialog>

    </div>
  );
};