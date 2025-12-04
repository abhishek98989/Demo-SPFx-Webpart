export function pluginAutoList() {
  return {
    onKeyDown: (editor:any, ev:any) => {
      if (ev.key !== " ") return;

      const selection = editor.getContentModelCopy();
      if (!selection) return;

      const text = selection.dom?.textContent || "";
      const trimmed = text.trim();

      if (trimmed === "*") {
        ev.preventDefault();
        editor.setContent(`<ul><li></li></ul>`);
        return;
      }

      if (/^\d+\.$/gm.test(trimmed)) {
        ev.preventDefault();
        editor.setContent(`<ol><li></li></ol>`);
      }
    },
  };
}
