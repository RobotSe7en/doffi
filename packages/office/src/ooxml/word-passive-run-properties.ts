/** Non-visual properties and explicit resets still affect a native consumer's
 * style cascade. Preserve their scalar source values without a proofing or
 * text-effects engine. Active visual effects continue through native recovery. */
export const passiveWordRunBooleans = ['noProof', 'outline', 'shadow', 'emboss', 'imprint', 'vanish', 'webHidden', 'rtl', 'cs', 'specVanish'] as const;
export const passiveWordRunValues = ['w', 'position', 'em', 'effect'] as const;
export const passiveWordRunCss = (name: string): string => '--word-run-native-' + name;
