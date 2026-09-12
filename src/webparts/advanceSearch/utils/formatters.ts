import { IChipStyle, IAlertRule } from "../types/advanceSearchTypes";

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

// Semantic style for Alerts (e.g. "Update in progress" matching SharePoint column formatting JSON: #FFA500)
export const getSemanticAlertStyle = (val: string): IChipStyle | null => {
  const clean = val.trim().toLowerCase();
  if (clean.includes("update in progress")) {
    return { bg: "#fff4e5", border: "#FFA500", text: "#d97706" };
  }
  return null;
};

// Dynamically parse SharePoint column formatting JSON for Alerts field
export const parseAlertsCustomFormatter = (
  customFormatterJson?: string
): IAlertRule | null => {
  if (!customFormatterJson) return null;

  try {
    const parsed = JSON.parse(customFormatterJson);

    let foundTxtContent = "";
    let foundColor = "";

    const walk = (node: any): void => {
      if (!node || typeof node !== "object") return;
      if (typeof node.txtContent === "string") {
        foundTxtContent = node.txtContent;
      }
      if (node.style && typeof node.style.color === "string") {
        foundColor = node.style.color;
      }
      if (Array.isArray(node.children)) {
        node.children.forEach(walk);
      }
    };

    walk(parsed);

    if (!foundTxtContent && !foundColor) return null;

    let text = "";
    let durationMinutes = 1;
    let requiresEditorMe = false;

    // Handles expressions like =if((Number(@now) - Number([$Modified]))/60000 < 1 && [$Editor.email] == @me, 'Update in progress', '')
    const ifMatch = /=if\s*\((.+?),\s*['"]([^'"]+)['"]\s*,\s*['"]([^'"]*)['"]\s*\)/i.exec(foundTxtContent);
    if (ifMatch) {
      const condition = ifMatch[1];
      text = ifMatch[2].trim();

      const minMatch = /\/\s*60000\s*<\s*(\d+(?:\.\d+)?)/i.exec(condition);
      if (minMatch) {
        durationMinutes = parseFloat(minMatch[1]);
      } else {
        const secMatch = /\/\s*1000\s*<\s*(\d+(?:\.\d+)?)/i.exec(condition);
        if (secMatch) {
          durationMinutes = parseFloat(secMatch[1]) / 60;
        }
      }

      requiresEditorMe = condition.includes("@me") && /editor/i.test(condition);
    } else if (foundTxtContent && !foundTxtContent.startsWith("@currentField")) {
      text = foundTxtContent.trim();
    }

    if (!text) return null;

    let color = "#FFA500";
    if (foundColor) {
      const colorIfMatch = /=if\s*\(.*?,\s*['"]([^'"]+)['"]/i.exec(foundColor);
      if (colorIfMatch) {
        color = colorIfMatch[1].trim();
      } else if (foundColor.startsWith("#") || foundColor.startsWith("rgb")) {
        color = foundColor.trim();
      }
    }

    const style: IChipStyle = {
      bg: color.toLowerCase() === "#ffa500" ? "#fff4e5" : "rgba(255, 165, 0, 0.12)",
      border: color,
      text: color.toLowerCase() === "#ffa500" ? "#d97706" : color,
    };

    return {
      durationMinutes,
      requiresEditorMe,
      text,
      color,
      style,
    };
  } catch (err) {
    console.warn("Could not parse Alerts CustomFormatter:", err);
    return null;
  }
};
