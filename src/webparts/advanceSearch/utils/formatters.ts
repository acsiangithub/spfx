import { IChipStyle } from "../types/advanceSearchTypes";

export const choiceToString = (value: string | string[] | undefined | null): string => {
  if (!value) return "";
  return Array.isArray(value) ? value.join(", ") : value;
};

export const sanitizeKqlValue = (value: string): string =>
  value.replace(/"/g, '\\"').trim();

// Deterministic pastel color generator for any string choice
export const getDynamicChipStyle = (str: string): IChipStyle => {
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    hash = str.charCodeAt(i) + ((hash << 5) - hash);
  }
  const h = Math.abs(hash % 360);
  return {
    bg: `hsl(${h}, 65%, 94%)`,
    border: `hsl(${h}, 50%, 80%)`,
    text: `hsl(${h}, 75%, 25%)`,
  };
};

// SharePoint column formatting CSS class mapping
export const SP_FORMAT_CLASS_MAP: Record<string, IChipStyle> = {
  "sp-css-backgroundcolor-warningbackground1": { bg: "#fff4ce", border: "#fde37f", text: "#795b00" },
  "sp-css-backgroundcolor-warningbackground2": { bg: "#fff4ce", border: "#fde37f", text: "#795b00" },
  "sp-css-backgroundcolor-warningbackground3": { bg: "#fff4ce", border: "#fde37f", text: "#795b00" },
  "sp-css-backgroundcolor-severewarningbackground1": { bg: "#fed9cc", border: "#fca385", text: "#a4262c" },
  "sp-css-backgroundcolor-severewarningbackground3": { bg: "#fed9cc", border: "#fca385", text: "#a4262c" },
  "sp-css-backgroundcolor-errorbackground1": { bg: "#fde7e9", border: "#f19999", text: "#a80000" },
  "sp-css-backgroundcolor-errorbackground3": { bg: "#fde7e9", border: "#f19999", text: "#a80000" },
  "sp-css-backgroundcolor-successbackground1": { bg: "#dff6dd", border: "#92c353", text: "#107c10" },
  "sp-css-backgroundcolor-successbackground3": { bg: "#dff6dd", border: "#92c353", text: "#107c10" },
  "sp-css-backgroundcolor-neutralbackground1": { bg: "#f3f2f1", border: "#edebe9", text: "#323130" },
};

// Parse color choices from SharePoint field CustomFormatter JSON if present
export const parseSpCustomFormatter = (
  customFormatterJson?: string
): Record<string, IChipStyle> => {
  const result: Record<string, IChipStyle> = {};
  if (!customFormatterJson) return result;

  try {
    const raw = customFormatterJson;
    const ifRegex = /@currentField\s*==\s*['"]([^'"]+)['"]\s*,\s*['"]([^'"]+)['"]/gi;
    let match: RegExpExecArray | null;
    while ((match = ifRegex.exec(raw)) !== null) {
      const choiceVal = match[1].trim();
      const styleOrClass = match[2].trim().toLowerCase();
      if (SP_FORMAT_CLASS_MAP[styleOrClass]) {
        result[choiceVal.toLowerCase()] = SP_FORMAT_CLASS_MAP[styleOrClass];
      } else if (styleOrClass.startsWith("#") || styleOrClass.startsWith("rgb")) {
        result[choiceVal.toLowerCase()] = {
          bg: styleOrClass,
          border: styleOrClass,
          text: "#1a1918",
        };
      }
    }
  } catch (err) {
    console.warn("Could not parse field CustomFormatter:", err);
  }

  return result;
};

// Fallback semantic styles for confidentiality values if no SharePoint JSON formatter exists
export const getSemanticConfidentialityStyle = (val: string): IChipStyle | null => {
  const clean = val.trim().toLowerCase();
  if (clean.includes("strictly") || clean.includes("high") || clean.includes("secret")) {
    return { bg: "#fde7e9", border: "#f19999", text: "#a80000" };
  }
  if (clean.includes("confidential") || clean.includes("restricted")) {
    return { bg: "#fed9cc", border: "#fca385", text: "#a4262c" };
  }
  if (clean.includes("internal")) {
    return { bg: "#fff4ce", border: "#fde37f", text: "#795b00" };
  }
  if (clean.includes("public") || clean.includes("general")) {
    return { bg: "#dff6dd", border: "#92c353", text: "#107c10" };
  }
  return null;
};
