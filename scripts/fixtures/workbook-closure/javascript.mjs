import { mountWorkbookWorkbench } from "@doffi/editor";
import { startWorkbookHost } from "./browser.mjs";

await startWorkbookHost((container, session, options) => {
  const view = mountWorkbookWorkbench(container, session, options);
  return { view, unmount: () => view.dispose() };
});
