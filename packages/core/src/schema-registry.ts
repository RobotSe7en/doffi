import { z } from "zod";
import { AnchorSchema, AnnotationSchema, AnnotationsSchema } from "./contracts/annotations.js";
import { AssetDescriptorSchema, DocumentManifestSchema, DocumentMetadataSchema } from "./contracts/document.js";
import { DocumentErrorDataSchema, DocumentEventSchema } from "./contracts/events.js";
import type { DocumentErrorCode } from "./contracts/events.js";
import { DocumentInteractionRequestSchema, EditProposalSchema } from "./contracts/interaction.js";
import { PresentationInteractionRequestSchema,PresentationEditProposalSchema } from "./contracts/presentation-interaction.js";
import { DocumentContentOperationSchema, DocumentOperationSchema, DocumentTransactionSchema } from "./contracts/operations.js";
import { FORMAT_VERSION, PROTOCOL_VERSION } from "./contracts/primitives.js";
import { DocumentError, summarizeIssues } from "./errors.js";

export const contractSchemas = Object.freeze({
  "asset-descriptor": AssetDescriptorSchema,
  "document-manifest": DocumentManifestSchema,
  "document-metadata": DocumentMetadataSchema,
  anchor: AnchorSchema,
  annotation: AnnotationSchema,
  annotations: AnnotationsSchema,
  "content-operation": DocumentContentOperationSchema,
  "document-operation": DocumentOperationSchema,
  "document-transaction": DocumentTransactionSchema,
  "interaction-request": DocumentInteractionRequestSchema,
  "edit-proposal": EditProposalSchema,
  "presentation-interaction-request":PresentationInteractionRequestSchema,
  "presentation-edit-proposal":PresentationEditProposalSchema,
  "document-error": DocumentErrorDataSchema,
  "document-event": DocumentEventSchema,
});
export type ContractName = keyof typeof contractSchemas;
export type ContractValue<Name extends ContractName> = z.infer<(typeof contractSchemas)[Name]>;
export const CONTRACT_NAMES: readonly ContractName[] = Object.freeze(
  Object.keys(contractSchemas) as ContractName[],
);
const formatContracts: ReadonlySet<ContractName> = new Set([
  "asset-descriptor", "document-manifest", "document-metadata", "anchor", "annotation", "annotations",
]);

export interface ContractJsonSchema {
  readonly $schema: string;
  readonly $id: string;
  readonly [key: string]: unknown;
}

function schemaFor<Name extends ContractName>(name: Name): (typeof contractSchemas)[Name] {
  if (!Object.hasOwn(contractSchemas, name)) {
    throw new DocumentError("UnsupportedFeature", "Unknown structural contract.");
  }
  return contractSchemas[name];
}

function object(value: unknown): Record<string, unknown> | undefined {
  return value !== null && typeof value === "object" && !Array.isArray(value)
    ? value as Record<string, unknown> : undefined;
}

function checkVersion(value: unknown, supported: number): void {
  if (typeof value === "number" && Number.isSafeInteger(value) && value !== supported) {
    throw new DocumentError("UnsupportedVersion", "The declared version is not supported.");
  }
}

function checkEnvelope(name: ContractName, input: unknown): void {
  const data = object(input);
  if (!data) return;
  if (name === "document-manifest") {
    if (typeof data.format === "string" && data.format !== "doffi") {
      throw new DocumentError("UnsupportedFormat", "The document format is not supported.");
    }
    checkVersion(data.formatVersion, FORMAT_VERSION);
  }
  if (["document-transaction", "interaction-request", "edit-proposal", "presentation-interaction-request","presentation-edit-proposal","document-event"].includes(name)) {
    checkVersion(data.protocolVersion, PROTOCOL_VERSION);
  }
  if (name === "interaction-request") {
    const context = object(data.context);
    if (context) {
      checkVersion(context.formatVersion, FORMAT_VERSION);
      if (context.kind === "presentation" || context.kind === "workbook") {
        throw new DocumentError("UnsupportedFeature", "This interaction context is not implemented yet.");
      }
    }
  }
  if(name==="presentation-interaction-request"){const context=object(data.context);if(context)checkVersion(context.formatVersion,FORMAT_VERSION);}
}

// Zod collects every array item's issues before returning. Preflight arrays one
// item at a time, including arrays nested in objects and contract variants, so
// invalid input cannot build an issue list proportional to the array's length.
// Cache schema structure only. A subtree without any preflight array has no
// work here; its ordinary Zod validation below remains complete. In particular,
// scalar cell unions need no repeated discriminator-error construction.
const arrayPaths = new WeakMap<z.core.$ZodType, boolean>();
function hasArrayPath(schema: z.core.$ZodType): boolean {
  const known = arrayPaths.get(schema); if (known !== undefined) return known;
  const found = schema instanceof z.ZodArray || schema instanceof z.ZodOptional && hasArrayPath(schema.unwrap())
    || (schema instanceof z.ZodDiscriminatedUnion || schema instanceof z.ZodUnion) && schema.options.some(option => hasArrayPath(option))
    || schema instanceof z.ZodObject && Object.values(schema.shape).some(child => hasArrayPath(child as z.core.$ZodType));
  arrayPaths.set(schema, found); return found;
}
function checkArrayItems(schema: z.core.$ZodType, input: unknown, path: PropertyKey[], fail: (issues: z.ZodIssue[]) => never): void {
  if (!hasArrayPath(schema)) return;
  if (schema instanceof z.ZodOptional) {
    if (input !== undefined) checkArrayItems(schema.unwrap(), input, path, fail);
  } else if (schema instanceof z.ZodDiscriminatedUnion) {
    const data = object(input);
    if (!data) return;
    const discriminator = schema.def.discriminator;
    for (const option of schema.options) {
      if (option instanceof z.ZodObject && option.shape[discriminator].safeParse(data[discriminator]).success) {
        checkArrayItems(option, input, path, fail);
        break;
      }
    }
  } else if (schema instanceof z.ZodUnion) {
    // Metadata variants have different nested arrays. Preflight each bounded
    // alternative without parsing a whole variant (or collecting every issue).
    for (const option of schema.options) checkArrayItems(option, input, path, fail);
  } else if (schema instanceof z.ZodObject) {
    const data = object(input);
    if (!data) return;
    for (const [key, child] of Object.entries(schema.shape)) {
      checkArrayItems(child as z.core.$ZodType, data[key], [...path, key], fail);
    }
  } else if (schema instanceof z.ZodArray && Array.isArray(input)) {
    for (let index = 0; index < input.length; index++) {
      const itemPath = [...path, index];
      checkArrayItems(schema.element, input[index], itemPath, fail);
      const result = z.safeParse(schema.element, input[index]);
      if (!result.success) {
        fail(result.error.issues.map((issue) => ({ ...issue, path: [...itemPath, ...issue.path] })));
      }
    }
  }
}

/** Structural validation only; it does not open a package or execute an operation. */
export function parseContract<Name extends ContractName>(name: Name, input: unknown): ContractValue<Name> {
  const schema = schemaFor(name);
  checkEnvelope(name, input);
  const code: DocumentErrorCode = ["edit-proposal","presentation-edit-proposal"].includes(name) ? "InvalidProposal"
    : ["content-operation", "document-operation", "document-transaction"].includes(name)
      ? "InvalidOperation" : "InvalidDocument";
  const fail = (issues: z.ZodIssue[]): never => {
    throw new DocumentError(code, "Invalid " + name + " data.", summarizeIssues(issues));
  };
  checkArrayItems(schema, input, [], fail);
  const result = schema.safeParse(input);
  if (result.success) {
    // The selected registry entry determines the corresponding generic output.
    return result.data as ContractValue<Name>;
  }
  return fail(result.error.issues);
}

/** Each call produces a self-contained schema with no network references or relaxed checks. */
export function getContractJsonSchema(name: ContractName): ContractJsonSchema {
  const schema = z.toJSONSchema(schemaFor(name), {
    target: "draft-2020-12",
    io: "input",
    unrepresentable: "throw",
    cycles: "throw",
    reused: "inline",
  });
  return {
    ...schema,
    $schema: "https://json-schema.org/draft/2020-12/schema",
    $id: "urn:doffi:contract:" + name + ":"
      + (formatContracts.has(name) ? FORMAT_VERSION : PROTOCOL_VERSION),
  };
}
