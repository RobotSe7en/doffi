/** Native source flags without additional paragraph paint. Active bidi or
 * alternate text directions continue through native compatibility recovery. */
export const passiveWordParagraphBooleans = ["suppressLineNumbers", "bidi"] as const;
export const passiveWordParagraphValues = ["textDirection", "divId"] as const;
export const passiveWordParagraphCss = (name: string): string => "--word-paragraph-native-" + name;
