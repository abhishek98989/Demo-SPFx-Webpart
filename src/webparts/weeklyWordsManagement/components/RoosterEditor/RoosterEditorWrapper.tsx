// import * as React from "react";
// import {
//     createContentModelEditor,
// } from "roosterjs-content-model-core";
// import { toggleBold, toggleItalic, toggleUnderline, toggleBullets, toggleNumbering, clearFormat, setFontFamily, setFontSize, setAlignment } from "roosterjs-content-model-api";
// import { pluginWordContentModelPaste, WordPaste } from "roosterjs-content-model-plugins";
// import { IEditor } from "roosterjs-content-model-types";
// import { pluginImagePasteHandler, pluginExtractPastedImages } from "./plugins/imagePasteHandler";
// import { pluginAutoList } from "./plugins/autoList";
// import { IconButton } from "@fluentui/react/lib/Button";
// import "./toolbar.css";

// /** Props */
// export interface IRoosterEditorWrapperProps {
//     context: any;
//     imagesWebServerRelativeUrl: string;
//     descriptionHtml: string;
//     onChange: (html: string) => void;
// }

// const RoosterEditorWrapper: React.FC<IRoosterEditorWrapperProps> = ({
//     context,
//     imagesWebServerRelativeUrl,
//     descriptionHtml,
//     onChange,
// }) => {

//     const editorDivRef = React.useRef<HTMLDivElement | null>(null);
//     const editorRef = React.useRef<IEditor | null>(null);

//     /** boot editor */
//     React.useEffect(() => {

//         if (!editorDivRef.current) return;

//         // expose SP upload config globally so plugins can use
//         (window as any).__roosterUploadConfig = {
//             context,
//             targetFolder: `${imagesWebServerRelativeUrl}/PublishingImages`,
//         };

//         const editor = createContentModelEditor(editorDivRef.current, {
//             plugins: [
//                 pluginWordContentModelPaste(new WordPaste()),   // Word/outlook cleanup
//                 pluginAutoList(),                              // auto bullets & numbering
//                 pluginExtractPastedImages(),                   // blob/data detection
//                 pluginImagePasteHandler({
//                     context,
//                     targetFolder: `${imagesWebServerRelativeUrl}/PublishingImages`,
//                 }),
//             ],
//         });

//         editorRef.current = editor;

//         // load existing html if editing
//         if (descriptionHtml) {
//             editor.setContent(descriptionHtml);
//         }

//         return () => {
//             editor.dispose();
//         };
//     }, []);

//     /** when content changes, extract html */
//     const handleContentChanged = React.useCallback(() => {
//         if (!editorRef.current) return;
//         const html = editorRef.current.getContent(true);
//         if (html) onChange(html);
//     }, [onChange]);

//     React.useEffect(() => {
//         if (!editorRef.current) return;
//         editorRef.current.addContentChangedListener(handleContentChanged);

//         return () => {
//             editorRef.current?.removeContentChangedListener(handleContentChanged);
//         };
//     }, [handleContentChanged]);

//     /** formatting helpers */
//     const format = (action: any) => {
//         if (!editorRef.current) return;
//         editorRef.current.focus();
//         editorRef.current.formatContentModel((model) => {
//             action(editorRef.current, model);
//             return model;
//         });
//     };

//     const execModel = (action: any) => {
//         if (!editorRef.current) return;
//         editorRef.current.focus();
//         editorRef.current.formatContentModel((model) => {
//             action(model);
//             return model;
//         });
//     };

//     /** API to insert image from picker */
//     const insertImage = async (url: string, w?: number, h?: number) => {
//         if (!editorRef.current) return;
//         editorRef.current.focus();
//         editorRef.current.formatContentModel((model) => {

//             model.blocks.push({
//                 blockType: "Paragraph",
//                 segments: [
//                     {
//                         segmentType: "Image",
//                         src: url,
//                         format: {
//                             width: w ? `${w}px` : undefined,
//                             height: h ? `${h}px` : undefined,
//                             maxWidth: "100%",
//                             objectFit: "contain",
//                         },
//                     }
//                 ],
//             });

//             return model;
//         });
//     };

//     // allow global to use it temporarily
//     (window as any).insertImageFromRooster = insertImage;

//     return (
//         <div className="rooster-wrapper">
//             {/* MAIN TOOLBAR */}
//             <div className="rooster-toolbar">
//                 <IconButton iconProps={{ iconName: "Bold" }} title="Bold" onClick={() => format(toggleBold)} />
//                 <IconButton iconProps={{ iconName: "Italic" }} title="Italic" onClick={() => format(toggleItalic)} />
//                 <IconButton iconProps={{ iconName: "Underline" }} title="Underline" onClick={() => format(toggleUnderline)} />

//                 <IconButton iconProps={{ iconName: "BulletedList" }} title="Bullets" onClick={() => format(toggleBullets)} />
//                 <IconButton iconProps={{ iconName: "NumberedList" }} title="Numbering" onClick={() => format(toggleNumbering)} />

//                 {/* Alignment */}
//                 <IconButton iconProps={{ iconName: "AlignLeft" }} title="Left" onClick={() => format(() => setAlignment(editorRef.current!, "left"))} />
//                 <IconButton iconProps={{ iconName: "AlignCenter" }} title="Center" onClick={() => format(() => setAlignment(editorRef.current!, "center"))} />
//                 <IconButton iconProps={{ iconName: "AlignRight" }} title="Right" onClick={() => format(() => setAlignment(editorRef.current!, "right"))} />

//                 {/* Font Family */}
//                 <select title="Font Family" onChange={(e) => format(() => setFontFamily(editorRef.current!, e.target.value))}>
//                     <option value="Calibri">Calibri</option>
//                     <option value="Arial">Arial</option>
//                     <option value="Georgia">Georgia</option>
//                     <option value="Times New Roman">Times New Roman</option>
//                 </select>

//                 {/* Font Size */}
//                 <select title="Font size" onChange={(e) => format(() => setFontSize(editorRef.current!, e.target.value))}>
//                     <option value="12px">12</option>
//                     <option value="14px">14</option>
//                     <option value="16px">16</option>
//                     <option value="18px">18</option>
//                     <option value="20px">20</option>
//                     <option value="24px">24</option>
//                 </select>

//                 <IconButton iconProps={{ iconName: "ClearFormatting" }} title="Clear formatting" onClick={() => format(clearFormat)} />
//             </div>

//             {/* EDITOR */}
//             <div
//                 ref={editorDivRef}
//                 className="rooster-editor-surface"
//                 style={{
//                     border: "1px solid #ddd",
//                     minHeight: 300,
//                     padding: 12,
//                     overflow: "auto",
//                     background: "white",
//                     borderRadius: 4,
//                 }}
//             ></div>
//         </div>
//     );
// };

// export default RoosterEditorWrapper;
