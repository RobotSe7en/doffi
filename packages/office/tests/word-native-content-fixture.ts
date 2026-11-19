import { strFromU8, strToU8 } from "fflate";
import {
  createAssetDescriptor,
  openDocument,
  WordNativeFragmentCollectionSchema,
} from "@doffi/core";
import type { WordContentSemantics, WordNativeContent, DocumentPackage } from "@doffi/core";
import { importDocx } from "@doffi/office";
import { body, mutateDocx, styles } from "./regression-fixtures.js";
export const STORE_ID = "{A0000000-0000-4000-8000-000000000001}";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export function wordNativeContentFixture(bound = false): Uint8Array {
  const selected = bound
    ? '<w:sdt><w:sdtPr><w:id w:val="42"/><w:dataBinding w:storeItemID="' +
      STORE_ID +
      '" w:xpath="/data/value" w:prefixMappings=""/><w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>BOUND</w:t></w:r></w:p></w:sdtContent></w:sdt>'
    : '<w:customXml w:element="selected" w:uri="urn:doffi:native-copy"><w:p><w:r><w:t>SELECTED-NATIVE</w:t></w:r></w:p><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>SELECTED-CELL</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:customXml>';
  return mutateDocx("01-text-numbering.docx", {
    "word/styles.xml": () =>
      styles('<w:style w:type="paragraph" w:styleId="Normal" w:default="1"/>'),
    "word/document.xml": (s) =>
      body(
        s,
        "<w:p><w:r><w:t>BEFORE</w:t></w:r></w:p>" +
          selected +
          '<w:customXml w:element="private" w:uri="urn:doffi:native-copy"><w:p><w:r><w:t>PRIVATE-NATIVE-OUTSIDE</w:t></w:r></w:p></w:customXml>' +
          "<w:p><w:r><w:t>AFTER</w:t></w:r></w:p>",
      ),
    "customXml/item1.xml": () => "<data><value>SELECTED-BOUND-DATA</value></data>",
    "customXml/itemProps1.xml": () =>
      '<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" ds:itemID="' +
      STORE_ID +
      '"><ds:schemaRefs/></ds:datastoreItem>',
    "customXml/item2.xml": () => "<data>PRIVATE-UNBOUND-DATA</data>",
    "customXml/itemProps2.xml": () =>
      '<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" ds:itemID="{B0000000-0000-4000-8000-000000000002}"><ds:schemaRefs/></ds:datastoreItem>',
    "customXml/_rels/item1.xml.rels": () =>
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="props" Type="' +
      R +
      '/customXmlProps" Target="itemProps1.xml"/></Relationships>',
    "customXml/_rels/item2.xml.rels": () =>
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="props" Type="' +
      R +
      '/customXmlProps" Target="itemProps2.xml"/></Relationships>',
    "word/_rels/document.xml.rels": (s) =>
      s.replace(
        "</Relationships>",
        [1, 2]
          .map(
            (i) =>
              '<Relationship Id="custom' +
              i +
              '" Type="' +
              R +
              '/customXml" Target="../customXml/item' +
              i +
              '.xml"/>',
          )
          .join("") + "</Relationships>",
      ),
    "[Content_Types].xml": (s) =>
      s.replace(
        "</Types>",
        [1, 2]
          .map(
            (i) =>
              '<Override PartName="/customXml/item' +
              i +
              '.xml" ContentType="application/xml"/><Override PartName="/customXml/itemProps' +
              i +
              '.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>',
          )
          .join("") + "</Types>",
      ),
  });
}
export function legacyNativeContent(files: DocumentPackage): DocumentPackage {
  const model = openDocument(files, { kind: "document" }),
    snapshot = model.snapshot(),
    metadata = JSON.parse(JSON.stringify(snapshot.metadata)),
    manifest = JSON.parse(JSON.stringify(snapshot.manifest)),
    result = new Map(files),
    native: WordNativeContent = metadata.word.nativeContent;
  for (const path of new Set(native.fragments.map((f) => f.assetPath))) {
    const collection = WordNativeFragmentCollectionSchema.parse(
        JSON.parse(strFromU8(result.get(path)!)),
      ),
      next = path.replace(/\.json$/, ".xml"),
      bytes = strToU8(
        '<native:fragments xmlns:native="urn:doffi:word:native-fragments">' +
          Object.entries(collection.fragments)
            .map(([id, xml]) => '<native:item id="' + id + '">' + xml + "</native:item>")
            .join("") +
          "</native:fragments>",
      );
    for (const f of native.fragments)
      if (f.assetPath === path) {
        f.assetPath = next;
        delete f.fragmentFormat;
        delete f.relationshipIds;
        delete f.dependencyParts;
      }
    manifest.assets = manifest.assets.filter((a: { path: string }) => a.path !== path);
    manifest.assets.push(createAssetDescriptor(next, "application/xml", bytes));
    result.delete(path);
    result.set(next, bytes);
  }
  for (const p of native.parts) delete p.customXml;
  result.set(snapshot.manifest.metadata, strToU8(JSON.stringify(metadata)));
  result.set("manifest.json", strToU8(JSON.stringify(manifest)));
  return result;
}
/** Direct public-operation fixture: independent of the browser copy collector. */
export function nativeContentGraph() {
  const source = openDocument(
      importDocx(wordNativeContentFixture(), { documentId: "native-source" }),
      { kind: "document" },
    ),
    native = source.snapshot().metadata.word!.nativeContent!,
    fragment = native.fragments.find((f) => f.kind === "block")!,
    collection = WordNativeFragmentCollectionSchema.parse(
      JSON.parse(strFromU8(source.serialize().get(fragment.assetPath)!)),
    ),
    bytes = strToU8(
      JSON.stringify({
        ...collection,
        fragments: { [fragment.fragmentKey!]: collection.fragments[fragment.fragmentKey!] },
      }),
    ),
    descriptor = createAssetDescriptor("assets/selected-native.json", "application/json", bytes),
    semantics: WordContentSemantics = {
      nativeContent: {
        mainPart: native.mainPart,
        fragments: [{ ...JSON.parse(JSON.stringify(fragment)), assetPath: descriptor.path }],
        parts: [],
        relationships: [],
      },
    };
  return {
    html: source.getNode(fragment.nodeId!)!.html,
    semantics,
    assets: [{ descriptor, bytes }],
  };
}

export function nativePropertyContentFixture(): Uint8Array {
  return mutateDocx("01-text-numbering.docx", {
    "word/styles.xml": () =>
      styles(
        '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"/><w:style w:type="paragraph" w:styleId="NativeStyle"><w:name w:val="Native style"/><w:rPr><w:shadow/></w:rPr></w:style>',
      ),
    "word/document.xml": (s) =>
      body(
        s,
        '<w:p xmlns:x="urn:doffi:native-properties" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x" x:paragraph="selected-property"><w:pPr><w:pStyle w:val="NativeStyle"/></w:pPr><w:r x:run="selected-run"><w:rPr><w:eastAsianLayout w:id="3" w:combine="1"/></w:rPr><w:t>PROPERTY-TEXT</w:t></w:r></w:p><w:p><w:r><w:t>PRIVATE-PROPERTY-OUTSIDE</w:t></w:r></w:p>',
      ),
  });
}

export function nativePreviewContentFixture(): Uint8Array {
  return mutateDocx("03-images.docx", {
    "word/document.xml": (s) => {
      const drawing = s.match(/<w:drawing[\s>][\s\S]*?<\/w:drawing>/)![0];
      return s.replace(
        drawing,
        '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:doffi:native-choice"><mc:Choice Requires="x"><x:widget r:id="payload"/></mc:Choice><mc:Fallback>' +
          drawing +
          "</mc:Fallback></mc:AlternateContent>",
      );
    },
    "word/embeddings/selected.bin": () => "SELECTED-NATIVE-PAYLOAD",
    "word/embeddings/private.bin": () => "PRIVATE-NATIVE-PAYLOAD",
    "word/data/shared.xml": () => "<data>SELECTED-SHARED-DEPENDENCY</data>",
    "word/embeddings/_rels/selected.bin.rels": () =>
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="shared" Type="urn:doffi:shared" Target="../data/shared.xml"/><Relationship Id="external" Type="' +
      R +
      '/hyperlink" Target="https://example.invalid/native" TargetMode="External"/></Relationships>',
    "word/data/_rels/shared.xml.rels": () =>
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="cycle" Type="urn:doffi:cycle" Target="../embeddings/selected.bin"/></Relationships>',
    "word/_rels/document.xml.rels": (s) =>
      s.replace(
        "</Relationships>",
        '<Relationship Id="payload" Type="urn:doffi:payload" Target="embeddings/selected.bin"/><Relationship Id="private" Type="urn:doffi:payload" Target="embeddings/private.bin"/></Relationships>',
      ),
    "[Content_Types].xml": (s) =>
      s.replace(
        "</Types>",
        '<Override PartName="/word/embeddings/selected.bin" ContentType="application/octet-stream"/><Override PartName="/word/embeddings/private.bin" ContentType="application/octet-stream"/><Override PartName="/word/data/shared.xml" ContentType="application/xml"/></Types>',
      ),
  });
}

export function nativeNumberingContentFixture(): Uint8Array {
  return mutateDocx("01-text-numbering.docx", {
    "word/styles.xml": () =>
      styles(
        '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"/><w:style w:type="paragraph" w:styleId="NativeStyle"><w:rPr><w:shadow/></w:rPr></w:style>',
      ),
    "word/document.xml": (s) =>
      body(
        s,
        '<w:customXml w:element="numbered" w:uri="urn:doffi:native-copy"><w:p><w:pPr><w:pStyle w:val="NativeStyle"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>OPAQUE-LIST-ONE</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>OPAQUE-LIST-FOUR</w:t></w:r></w:p></w:customXml><w:p><w:r><w:t>AFTER</w:t></w:r></w:p>',
      ),
    "word/numbering.xml": (s) =>
      s.replace(
        "</w:numbering>",
        '<w:abstractNum w:abstractNumId="99"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="PRIVATE-UNSELECTED-NUMBERING"/></w:lvl></w:abstractNum><w:num w:numId="99"><w:abstractNumId w:val="99"/></w:num></w:numbering>',
      ),
  });
}
