import { WorkbookWorkbench } from "@doffi/editor/react";
import { createElement } from "react";
import { createRoot } from "react-dom/client";
import { startWorkbookHost } from "./browser.mjs";

await startWorkbookHost((container, session, options) => {
  const root = createRoot(container);
  return new Promise(resolve => {
    root.render(createElement(WorkbookWorkbench, {
      ...options, session,
      onViewChange(view) {
        if (view) resolve({ view, unmount: () => root.unmount() });
      },
    }));
  });
});
