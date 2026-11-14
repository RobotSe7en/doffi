import type { WordDocumentMetadata, AssetDescriptor } from "../contracts/document.js";
import type { Annotation } from "../contracts/annotations.js";
import type { ParsedHtml } from "./html.js";
import { invalid } from "./validation.js";
import { wordDrawingTransforms } from '../word-drawing.js';
import { validateWordMath } from '../word-math.js';

/** Validate native semantic references against the same HTML source, never an XML copy. */
export function validateWord(
  metadata: WordDocumentMetadata,
  annotations: readonly Annotation[],
  html: ParsedHtml,
  assets: readonly AssetDescriptor[] = [],
): void {
  if (annotations.some((annotation) => annotation.replies !== undefined)) invalid("Word comment replies are not implemented.");
  if (annotations.some((annotation) => annotation.status === "closed")) invalid("Closed native comment state is supported only for presentation threads.");
  const word = metadata.word;
  const node = (id: string) => {
    const result = html.peekNode(id);
    if (!result) invalid("A Word semantic reference targets a missing node.");
    return result;
  };
  const unique = <T>(
    values: readonly T[],
    key: (item: T) => string,
  ): Map<string, T> => {
    const map = new Map<string, T>();
    for (const item of values) {
      const id = key(item);
      if (map.has(id)) invalid("Word semantic identities must be unique.");
      map.set(id, item);
    }
    return map;
  };
  if (word) {
    if (word.nativeContent) {
      const native = word.nativeContent, paths = new Set(assets.map(asset => asset.path));
      const ids = new Set<string>(), parts = new Set<string>();
      for (const fragment of native.fragments) {
        if (ids.has(fragment.nativeId) || !paths.has(fragment.assetPath)) invalid("Native fragments require unique identities and declared resources.");
        ids.add(fragment.nativeId);
        if (fragment.nodeId) {
          const owner = node(fragment.nodeId);
          if (fragment.kind !== "property" && (owner.attributes["data-word-native-object"] !== fragment.nativeId || html.getNode(fragment.nodeId)!.text !== "\ufffc"))
            invalid("Native objects require an intact atomic source placeholder.");
        } else if (fragment.kind !== "property") invalid("Native objects must have a source node.");
      }
      for (const part of native.parts) {
        if (parts.has(part.sourcePart) || !paths.has(part.assetPath)) invalid("Native parts require unique paths and declared resources.");
        parts.add(part.sourcePart);
      }
      const relations = new Set<string>();
      for (const relation of native.relationships) {
        const key = relation.sourcePart + "\n" + relation.id;
        if (relations.has(key)) invalid("Native relationships require unique identities within each part.");
        relations.add(key);
      }
    }
    for (const id of html.nodeIds()) {
      const nativeId = node(id).attributes["data-word-native-object"];
      if (nativeId && !word.nativeContent?.fragments.some(fragment => fragment.nativeId === nativeId && fragment.nodeId === id && fragment.kind !== "property"))
        invalid("A native object marker requires its matching typed descriptor.");
    }
    const stories = unique(word.stories, (item) => item.storyId);
    const storyRoots = unique(word.stories, (item) => item.rootNodeId);
    const order = new Map(html.nodeIds().map((id, index) => [id, index]));
    const storyOf = (id: string): string | undefined => {
      let current: string | null = id;
      while (current !== null) {
        if (storyRoots.has(current)) return current;
        current = node(current).parentNodeId;
      }
      return undefined;
    };
    if (word.stories.filter((story) => story.kind === "main").length !== 1)
      invalid("Word requires exactly one main story.");
    for (const story of word.stories) {
      const root = node(story.rootNodeId);
      if (
        !html.isContainer(root.nodeId) ||
        root.attributes["data-office-story"] !== story.kind
      )
        invalid("A Word story requires a matching content container.");
      for (
        let parent = root.parentNodeId;
        parent !== null;
        parent = node(parent).parentNodeId
      ) {
        if (storyRoots.has(parent)) invalid("Word stories cannot be nested.");
      }
    }
    const drawings=unique(word.drawings??[],item=>item.nodeId),textOwners=new Set<string>();
    for(const drawing of drawings.values())if(drawing.kind==='chart'){
      const chart=drawing.chart;if(new Set(chart.series.map(series=>series.seriesId)).size!==chart.series.length)invalid('Chart series require distinct identities.');
      const range=(min:number|undefined,max:number|undefined):void=>{if(min!==undefined&&max!==undefined&&min>=max)invalid('A chart axis maximum must exceed its minimum.');};range(chart.minimum,chart.maximum);range(chart.xMinimum,chart.xMaximum);range(chart.secondaryMinimum,chart.secondaryMaximum);range(chart.secondaryXMinimum,chart.secondaryXMaximum);
      if(chart.kind==='combo'){if(chart.series.some(series=>series.kind==='bar')&&!chart.series.every(series=>series.kind==='bar'))invalid('Horizontal bar series require a horizontal chart group.');for(const axis of ['primary','secondary']){const members=chart.series.filter(series=>(series.axis??'primary')===axis);if(members.some(series=>series.kind==='scatter')&&members.some(series=>series.kind!=='scatter'))invalid('Scatter and category series require separate axis pairs.');}}
      for(const series of chart.series){if(series.values.length!==chart.categories.length||series.xValues&&series.xValues.length!==series.values.length||series.pointColors&&series.pointColors.length>series.values.length)invalid('Chart vectors must match their category and point positions.');if((chart.kind==='scatter'||chart.kind==='combo'&&series.kind==='scatter')&&!series.xValues)invalid('Scatter series require explicit x values.');if(chart.kind==='combo'&&!series.kind)invalid('Combination charts require each series kind.');}
      for(const[side,options]of [['primary',chart.axes?.value],['secondary',chart.axes?.secondaryValue]]as const)if(options?.crossesAt!==undefined&&!Number.isInteger(options.crossesAt)&&chart.kind!=='scatter'&&!chart.series.some(series=>(series.axis??'primary')===side&&series.kind==='scatter'))invalid('A category crossing requires an integer category position.');
      for(const style of [chart.textStyle,chart.titleStyle,chart.legendStyle,chart.labelStyle,...Object.values(chart.axes??{}).map(axis=>axis?.textStyle)])if(style?.fonts?.ascii&&style.fonts.hAnsi&&style.fonts.ascii!==style.fonts.hAnsi)invalid('DrawingML chart text uses one Latin face for ASCII and high-ANSI characters.');
      if(chart.native){const paths=new Set(assets.map(asset=>asset.path)),parts=new Set(chart.native.parts.map(part=>part.sourcePart));for(const part of chart.native.parts)if(!paths.has(part.assetPath))invalid('Native chart resources must be declared.');if(parts.size!==chart.native.parts.length||new Set(chart.native.relationships.map(relation=>relation.sourcePart+'\n'+relation.id)).size!==chart.native.relationships.length)invalid('Native chart parts and relationship identities must be unique.');if(!parts.has(chart.native.chartPart)||chart.native.workbookPart&&!parts.has(chart.native.workbookPart))invalid('Native chart template parts are missing.');for(const relation of chart.native.relationships)if(!parts.has(relation.sourcePart)||!relation.external&&!parts.has(relation.target))invalid('Native chart relationships must remain within their declared graph.');}
    }
    const controls=unique(word.controls??[],control=>control.nodeId);for(const control of controls.values()){const source=node(control.nodeId);if(source.attributes['data-word-content-control']!==control.kind||source.tagName!==(control.level==='block'?'div':'span'))invalid('A content control requires its matching source wrapper.');for(const value of [control.checkedSymbol,control.uncheckedSymbol])if(value!==undefined&&(value===0||value>=0xd800&&value<=0xdfff))invalid('A checkbox symbol must be a valid BMP character.');if(control.fullDate!==undefined&&!Number.isFinite(Date.parse(control.fullDate)))invalid('The content control date is invalid.');}
    for(const id of html.nodeIds()){const kind=node(id).attributes['data-word-content-control'];if(kind&&kind!=='toc'&&!controls.has(id))invalid('A content control marker requires its typed definition.');}
    const equations=unique(word.equations??[],equation=>equation.nodeId);for(const equation of equations.values()){if(node(equation.nodeId).attributes['data-word-equation']!=='true'||html.getNode(equation.nodeId)!.text!=='\ufffc')invalid('An equation requires its typed atomic source.');validateWordMath(equation);}for(const id of html.nodeIds())if(node(id).attributes['data-word-equation']&&!equations.has(id))invalid('An equation marker requires its math definition.');
    const notes=unique(word.notes??[],item=>item.noteId),references=unique(word.noteReferences??[],item=>item.nodeId),noteOwners=new Set<string>();
    if(word.footnoteOptions?.position&&!['pageBottom','beneathText'].includes(word.footnoteOptions.position)||word.endnoteOptions?.position&&!['sectEnd','docEnd'].includes(word.endnoteOptions.position))invalid('The note position must match its note kind.');
    unique(word.notes??[],note=>note.storyId);unique((word.notes??[]).filter(note=>(note.type??'normal')!=='normal'),note=>note.kind+' '+note.type);
    for(const note of notes.values())if(stories.get(note.storyId)?.kind!==note.kind)invalid('A note requires its own matching Word story.');
    for(const reference of references.values()){const note=notes.get(reference.noteId),source=node(reference.nodeId);if(!note||(note.type??'normal')!=='normal'||noteOwners.has(note.noteId)||source.attributes['data-word-note-reference']!==note.kind||html.getNode(reference.nodeId)!.text!=='\ufffc')invalid('A note reference requires one matching atomic marker and independently owned note.');noteOwners.add(note.noteId);const main=word.stories.find(story=>story.kind==='main')!;if(storyOf(reference.nodeId)!==main.rootNodeId)invalid('A note reference must belong to the main story.');}
    for(const id of html.nodeIds())if(node(id).attributes['data-word-note-reference']&&!references.has(id))invalid('A note marker requires typed reference metadata.');
    for(const id of html.nodeIds()){const marker=node(id).attributes['data-word-note-separator'];if(marker&&!['separator','continuationSeparator'].includes(marker))invalid('The note separator marker is invalid.');}
    const transforms=wordDrawingTransforms(word,{getNode:id=>html.getNode(id)});
    const parentDrawing=(id:string):string|undefined=>{for(let parent=node(id).parentNodeId;parent;parent=node(parent).parentNodeId)if(drawings.has(parent))return parent;return;};
    for(const drawing of drawings.values()){
      const source=node(drawing.nodeId);if(source.attributes['data-word-drawing']!==drawing.kind)invalid('A Word drawing requires its matching source marker.');
      if(drawing.kind==='image'?source.tagName!=='img':source.tagName!=='span')invalid('A Word drawing has the wrong source element.');
      if(drawing.kind!=='image'&&drawing.kind!=='group'&&html.getNode(source.nodeId)!.text!=='\ufffc')invalid('A Word drawing must retain its atomic source character.');
      if(drawing.kind!=='connector'&&(drawing.geometry.width<=0||drawing.geometry.height<=0))invalid('Drawing dimensions must be positive.');
      if(drawing.kind==='image'&&drawing.crop&&(drawing.crop.left+drawing.crop.right>=1||drawing.crop.top+drawing.crop.bottom>=1))invalid('Cropping must retain a positive image area.');
      const parent=parentDrawing(drawing.nodeId);if(parent){if(drawings.get(parent)!.kind!=='group'||drawing.layout.mode!=='inline')invalid('Only a group can contain Word drawing objects.');}
      if(drawing.kind==='group'&&(!source.childNodeIds.length||source.childNodeIds.some(id=>!drawings.has(id))))invalid('A group must contain only drawing sources.');
      if('textStoryId'in drawing&&drawing.textStoryId){if(stories.get(drawing.textStoryId)?.kind!=='textBox'||textOwners.has(drawing.textStoryId))invalid('Drawing text requires an independently owned text box story.');textOwners.add(drawing.textStoryId);}
      if(drawing.kind==='connector')for(const id of [drawing.connector.startNodeId,drawing.connector.endNodeId])if(id&&(!drawings.has(id)||!transforms.sameCanvas(drawing.nodeId,id)||storyOf(id)!==storyOf(drawing.nodeId)))invalid('Connector endpoints require other floating drawings in the same source canvas.');
    }
    for(const id of html.nodeIds())if(node(id).attributes['data-word-drawing']&&!drawings.has(id))invalid('A drawing marker requires typed drawing metadata.');
    const activeStories=new Set<string>(),completeStories=new Set<string>();const inspectStory=(id:string):void=>{if(activeStories.has(id))invalid('Drawing text story references cannot be cyclic.');if(completeStories.has(id))return;activeStories.add(id);for(const drawing of drawings.values())if(storyOf(drawing.nodeId)===stories.get(id)?.rootNodeId&&'textStoryId'in drawing&&drawing.textStoryId)inspectStory(drawing.textStoryId);activeStories.delete(id);completeStories.add(id);};for(const story of word.stories)inspectStory(story.storyId);
    const main = word.stories.find((story) => story.kind === "main")!;
    for (const section of metadata.sections) {
      const setup=section.pageSetup??metadata.pageSetup;
      if(setup.width<=setup.margins.left+setup.margins.right||setup.height<=setup.margins.top+setup.margins.bottom)invalid("Page margins must leave a positive printable area.");
      if(section.gutter && ((section.gutterPosition === 'top' ? setup.height-setup.margins.top-setup.margins.bottom : setup.width-setup.margins.left-setup.margins.right) <= section.gutter)) invalid("The binding gutter must leave a positive printable area.");
      if(section.columns){const columns=section.columns;if(columns.gap*(columns.count-1)>=setup.width-setup.margins.left-setup.margins.right-(section.gutterPosition==='top'?0:section.gutter??0))invalid("Column spacing must leave a positive text width.");}
      let parent: string | null = section.rootNodeId;
      while (parent !== null && parent !== main.rootNodeId)
        parent = node(parent).parentNodeId;
      if (parent === null) invalid("A section must belong to the main story.");
      for (const [kind, references] of [
        ["header", section.headers],
        ["footer", section.footers],
      ] as const) {
        for (const id of Object.values(references ?? {})) {
          if (id !== undefined && stories.get(id)?.kind !== kind)
            invalid(
              "A section refers to a missing or wrong-kind auxiliary story.",
            );
        }
      }
    }
    const styles = unique(word.styles, (item) => item.styleId);
    const visited = new Set<string>();
    for (const style of word.styles) {
      for (const reference of [
        style.basedOn,
        style.linkedStyleId,
        style.nextStyleId,
      ]) {
        if (reference !== undefined && !styles.has(reference))
          invalid("A native style refers to a missing style.");
      }
      let current: typeof style | undefined = style;
      const chain = new Set<string>();
      while (current && !visited.has(current.styleId)) {
        if (chain.has(current.styleId))
          invalid("Native style inheritance is cyclic.");
        chain.add(current.styleId);
        current = current.basedOn ? styles.get(current.basedOn) : undefined;
      }
      for (const id of chain) visited.add(id);
    }
    const numbering = unique(word.numbering, (item) => item.numberingId);
    for (const definition of word.numbering)
      unique(definition.levels, (level) => String(level.level));
    const links = unique(word.hyperlinks, (item) => item.nodeId);
    const linkEnds=new Map(word.hyperlinks.filter(link=>link.endNodeId).map(link=>[link.endNodeId!,link]));
    unique(word.hyperlinks.filter(link=>link.endNodeId),link=>link.endNodeId!);
    const fields = unique(word.fields, (item) => item.nodeId);
    const revisions = unique(word.revisions, (item) => item.nodeId);
    const tabs = unique(word.tabs, (item) => item.paragraphNodeId);
    const bookmarks = unique(word.bookmarks, (item) => item.bookmarkId);
    unique(word.bookmarks, (item) => item.name);
    const names = new Set(
      [...bookmarks.values()].map((bookmark) => bookmark.name),
    );
    for (const bookmark of bookmarks.values()) {
      node(bookmark.startNodeId);
      node(bookmark.endNodeId);
      if (
        storyOf(bookmark.startNodeId) !== storyOf(bookmark.endNodeId) ||
        order.get(bookmark.startNodeId)! > order.get(bookmark.endNodeId)!
      )
        invalid("A bookmark range crosses stories or is reversed.");
    }
    for (const link of links.values()) {
      node(link.nodeId);
      if(link.endNodeId){node(link.endNodeId);if(storyOf(link.nodeId)!==storyOf(link.endNodeId)||order.get(link.nodeId)!>=order.get(link.endNodeId)!||node(link.nodeId).attributes['data-word-link']!=='start'||node(link.endNodeId).attributes['data-word-link']!=='end')invalid('A hyperlink range requires ordered markers in one story.');}
      if (!link.target && !link.bookmark)
        invalid("A hyperlink requires a target.");
      if (
        link.target &&
        (!/^(https?:|mailto:)/i.test(link.target) ||
          /[\u0000-\u001f\u007f]/.test(link.target))
      )
        invalid("A hyperlink has an unsupported or executable target.");
      if (!link.target && link.bookmark && !names.has(link.bookmark))
        invalid("A hyperlink refers to a missing bookmark.");
    }
    const lastChild=(id:string):number=>{const children=node(id).childNodeIds;return children.length?lastChild(children.at(-1)!):order.get(id)!;};const linkRanges=word.hyperlinks.map(link=>({from:order.get(link.nodeId)!,to:link.endNodeId?order.get(link.endNodeId)!:lastChild(link.nodeId)})).sort((a,b)=>a.from-b.from);for(let i=1;i<linkRanges.length;i++)if(linkRanges[i]!.from<=linkRanges[i-1]!.to)invalid('Hyperlink ranges cannot overlap or nest.');
    for (const field of fields.values()) {
      if (node(field.nodeId).attributes["data-word-field"] === undefined)
        invalid("A field requires its source marker or container.");
      if (
        node(field.nodeId).attributes["data-word-field"] === "begin" &&
        field.endNodeId === undefined
      )
        invalid("A complex field requires its closing marker.");
      if (field.endNodeId !== undefined) {
        if (
          node(field.endNodeId).attributes["data-word-field-end"] !==
            field.nodeId ||
          storyOf(field.endNodeId) !== storyOf(field.nodeId) ||
          order.get(field.endNodeId)! < order.get(field.nodeId)!
        )
          invalid("A field end does not match its start.");
      }
      for (const id of field.resultNodeIds ?? []) {
        node(id);
        if (storyOf(id) !== storyOf(field.nodeId))
          invalid("A field result crosses a story boundary.");
      }
      for (const nested of field.instructionFields ?? []) {
        if (
          !fields.has(nested.fieldNodeId) ||
          nested.offset > field.instruction.length ||
          nested.fieldNodeId === field.nodeId
        )
          invalid("A nested field instruction reference is invalid.");
      }
    }
    const fieldDone = new Set<string>();
    const activeFields = new Set<string>();
    for (const root of fields.keys()) {
      const stack: { id: string; exit: boolean }[] = [
        { id: root, exit: false },
      ];
      while (stack.length) {
        const entry = stack.pop()!;
        if (entry.exit) {
          activeFields.delete(entry.id);
          fieldDone.add(entry.id);
          continue;
        }
        if (fieldDone.has(entry.id)) continue;
        if (activeFields.has(entry.id))
          invalid("Field instruction references are cyclic.");
        activeFields.add(entry.id);
        stack.push({ id: entry.id, exit: true });
        for (const nested of fields.get(entry.id)!.instructionFields ?? [])
          stack.push({ id: nested.fieldNodeId, exit: false });
      }
    }
    for (const revision of revisions.values()) {
      if (
        node(revision.nodeId).tagName !==
        (revision.kind === "insert" ? "ins" : "del")
      )
        invalid("A revision must refer to the matching revision container.");
    }
    for (const tab of tabs.values()) {
      if (
        !["p", "h1", "h2", "h3", "h4", "h5", "h6"].includes(
          node(tab.paragraphNodeId).tagName,
        )
      )
        invalid("Tab stops require a paragraph.");
      unique(tab.stops, (stop) => String(stop.position));
    }
    for (const table of unique(word.tables, (item) => item.nodeId).values()) {
      if (node(table.nodeId).tagName !== "table")
        invalid("A table grid requires a table node.");
    }
    for (const id of html.nodeIds()) {
      const current = node(id);
      const attrs = current.attributes;
      if (attrs["data-word-style"] && !styles.has(attrs["data-word-style"]))
        invalid("A node refers to a missing native style.");
      if (attrs["data-word-numbering"]) {
        const level = Number(attrs["data-word-level"]);
        if (
          !numbering
            .get(attrs["data-word-numbering"])
            ?.levels.some((item) => item.level === level)
        )
          invalid("A paragraph refers to a missing numbering level.");
      }
      if (attrs["data-word-field"] !== undefined && !fields.has(id))
        invalid("A field has no semantic definition.");
      if (attrs["data-word-link"] !== undefined && !links.has(id)&&!linkEnds.has(id))
        invalid("A link has no semantic definition.");
      if (["ins", "del"].includes(current.tagName) && !revisions.has(id))
        invalid("A revision has no semantic definition.");
    }
  } else {
    for (const id of html.nodeIds()) {
      const current = node(id);
      if (["ins", "del"].includes(current.tagName))
        invalid("Native revision containers require Word metadata.");
    }
  }
  for (const annotation of annotations) {
    if (annotation.bodyNodeId !== undefined) {
      if (
        !word?.stories.some(
          (story) =>
            story.kind === "comment" &&
            story.rootNodeId === annotation.bodyNodeId,
        )
      )
        invalid("A rich annotation requires its comment story.");
    }
  }
}
