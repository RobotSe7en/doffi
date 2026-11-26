/** Non-visual properties and explicit resets still affect a native consumer's
 * style cascade. Preserve their scalar source values without a proofing or
 * text-effects engine. Active visual effects continue through native recovery. */
export const passiveWordRunBooleans = [
  "noProof",
  "outline",
  "shadow",
  "emboss",
  "imprint",
  "vanish",
  "webHidden",
  "rtl",
  "cs",
  "specVanish",
] as const;
export const passiveWordRunValues = ["position", "em", "effect"] as const;
export const passiveWordRunCss = (name: string): string => "--word-run-native-" + name;

export const wordUnderlineValues = [
  "none",
  "single",
  "words",
  "double",
  "thick",
  "dotted",
  "dottedHeavy",
  "dash",
  "dashedHeavy",
  "dashLong",
  "dashLongHeavy",
  "dotDash",
  "dashDotHeavy",
  "dotDotDash",
  "dashDotDotHeavy",
  "wave",
  "wavyHeavy",
  "wavyDouble",
] as const;

export const wordEmphasisValues = ["none", "dot", "comma", "circle", "underDot"] as const;

export const wordOpenTypeLigatures = [
  "none",
  "standard",
  "contextual",
  "historical",
  "discretional",
  "standardContextual",
  "standardHistorical",
  "contextualHistorical",
  "standardDiscretional",
  "contextualDiscretional",
  "historicalDiscretional",
  "standardContextualHistorical",
  "standardContextualDiscretional",
  "standardHistoricalDiscretional",
  "contextualHistoricalDiscretional",
  "all",
] as const;

export const wordOpenTypeNumSpacings = ["default", "proportional", "tabular"] as const;
export const wordOpenTypeNumForms = ["default", "lining", "oldStyle"] as const;
