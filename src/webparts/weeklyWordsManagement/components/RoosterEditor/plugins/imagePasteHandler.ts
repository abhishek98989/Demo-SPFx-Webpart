import { IEditor } from "roosterjs-content-model-types";

interface IImagePasteHandlerOptions {
  context: any;
  targetFolder: string;
}

export function pluginImagePasteHandler(opts: IImagePasteHandlerOptions) {
  return {
    onPasted: async (editor: IEditor, event: ClipboardEvent) => {
      if (!event.clipboardData) return;

      const items:any = event.clipboardData.items;
      for (const item of items) {
        if (item.type.startsWith("image/")) {
          const file = item.getAsFile();
          if (!file) continue;

          // Upload to SP
          const url = await uploadToSharePoint(file, opts);
          if (!url) continue;

          // Insert into editor at cursor
          insertImageSegment(editor, url);
          event.preventDefault();
        }
      }
    },
  };
}

export function pluginExtractPastedImages() {
  return {
    onContentChanged: (editor: IEditor) => {
      const doc = editor.getDocument();
      const imgs = doc.querySelectorAll("img[src^='data:image'], img[src^='blob:']");
      imgs.forEach(async (img: HTMLImageElement) => {
        try {
          const src = img.src;
          const blob = await fetch(src).then((r) => r.blob());
          const width = img.getAttribute("width");
          const height = img.getAttribute("height");
          const url = await uploadToSharePoint(blob, (window as any).__roosterUploadConfig);

          img.src = url;
          if (width) img.style.width = width + "px";
          if (height) img.style.height = height + "px";
          img.style.maxWidth = "100%";
          img.style.height = "auto";
        } catch { }
      });
    },
  };
}

async function uploadToSharePoint(file: Blob | File, opts: IImagePasteHandlerOptions) {
  const sp = await import("@pnp/sp");
  const spfi = sp.spfi;
  const SPFx = (await import("@pnp/sp")).SPFx;

  const spImg = spfi("https://vaughnconstruction.sharepoint.com/news").using(
    SPFx(opts.context)
  );

  const folder = spImg.web.getFolderByServerRelativePath(opts.targetFolder);
  const fileName = `paste-${Date.now()}-${Math.floor(Math.random() * 1000)}.png`;

  const result :any= await folder.files.addUsingPath(fileName, file, { Overwrite: true });

  return (
    result?.data?.ServerRelativeUrl ||
    result?.file?.serverRelativeUrl ||
    result?.ServerRelativeUrl ||
    null
  );
}

function insertImageSegment(editor: IEditor, url: string) {
  editor.formatContentModel((model:any) => {
    model.blocks.push({
      blockType: "Paragraph",
      segments: [
        {
          segmentType: "Image",
          src: url,
          format: {
            maxWidth: "100%",
            height: "auto",
          },
        },
      ],
    });
    return model;
  });
}
