import dayjs, { Dayjs } from "dayjs";
import { IChipStyle, IAlertRule } from "../types/advanceSearchTypes";

export const choiceToString = (value: string | string[] | undefined | null): string => {
  if (!value) return "";
  return Array.isArray(value) ? value.join(", ") : value;
};

export const sanitizeKqlValue = (value: string): string =>
  value.replace(/"/g, '\\"').trim();

export const formatKeywordTerm = (raw: string): string => {
  let val = raw.trim();
  if (!val) return "";

  // If already wrapped in quotes
  if (val.startsWith('"') && val.endsWith('"') && val.length >= 2) {
    val = val.slice(1, -1).trim();
  }

  // Check if it contains wildcard *
  if (val.includes("*")) {
    // In SharePoint KQL, wildcards MUST NOT be wrapped in quotes for prefix expansion to work
    return sanitizeKqlValue(val);
  }

  return `"${sanitizeKqlValue(val)}"`;
};

export const compileKeywordsToKql = (chips: string[], pendingInput?: string): string => {
  const rawList: string[] = [...chips];
  if (pendingInput && pendingInput.trim()) {
    pendingInput
      .split(",")
      .map((s) => s.trim())
      .filter(Boolean)
      .forEach((s) => {
        if (!rawList.includes(s)) {
          rawList.push(s);
        }
      });
  }

  if (rawList.length === 0) return "";

  type Token = { type: "op"; value: "AND" | "OR" | "NOT" } | { type: "term"; value: string };
  const tokens: Token[] = [];

  const tokenRegex = /"([^"]+)"|(\S+)/g;

  rawList.forEach((entry) => {
    let match: RegExpExecArray | null;
    tokenRegex.lastIndex = 0;
    while ((match = tokenRegex.exec(entry)) !== null) {
      if (match[1] !== undefined) {
        const quotedContent = match[1].trim();
        if (quotedContent) {
          tokens.push({
            type: "term",
            value: formatKeywordTerm(`"${quotedContent}"`),
          });
        }
      } else if (match[2] !== undefined) {
        const word = match[2].trim();
        const upper = word.toUpperCase();
        if (upper === "AND" || upper === "&&") {
          tokens.push({ type: "op", value: "AND" });
        } else if (upper === "OR" || upper === "||") {
          tokens.push({ type: "op", value: "OR" });
        } else if (upper === "NOT" || upper === "!") {
          tokens.push({ type: "op", value: "NOT" });
        } else {
          tokens.push({
            type: "term",
            value: formatKeywordTerm(word),
          });
        }
      }
    }
  });

  if (tokens.length === 0) return "";

  // 1. Remove leading binary operators (AND, OR)
  while (tokens.length > 0 && tokens[0].type === "op" && (tokens[0].value === "AND" || tokens[0].value === "OR")) {
    tokens.shift();
  }

  // 2. Remove trailing operators (AND, OR, NOT)
  while (tokens.length > 0 && tokens[tokens.length - 1].type === "op") {
    tokens.pop();
  }

  if (tokens.length === 0) return "";

  // 3. Build string with default AND when two terms are adjacent without an operator
  const resultParts: string[] = [];
  for (let i = 0; i < tokens.length; i++) {
    const curr = tokens[i];
    const prev = i > 0 ? tokens[i - 1] : null;

    if (curr.type === "term") {
      if (prev && prev.type === "term") {
        resultParts.push("AND");
      }
      resultParts.push(curr.value);
    } else if (curr.type === "op") {
      if (curr.value === "NOT") {
        resultParts.push("NOT");
      } else {
        if (prev && prev.type === "op") {
          resultParts.pop();
        }
        resultParts.push(curr.value);
      }
    }
  }

  if (resultParts.length === 0) return "";
  if (resultParts.length === 1) return resultParts[0];

  return `(${resultParts.join(" ")})`;
};

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

export const getSemanticDocumentStatusStyle = (val: string): IChipStyle | null => {
  const clean = val.trim().toLowerCase();
  if (
    clean === "approved" ||
    clean === "published" ||
    clean === "active" ||
    clean === "valid" ||
    clean === "released" ||
    clean.includes("success")
  ) {
    return { bg: "#dff6dd", border: "#92c353", text: "#107c10" };
  }
  if (
    clean.includes("review") ||
    clean.includes("pending") ||
    clean.includes("progress") ||
    clean.includes("waiting")
  ) {
    return { bg: "#fff4ce", border: "#fde37f", text: "#795b00" };
  }
  if (clean === "draft" || clean === "new" || clean === "created") {
    return { bg: "#e0f2fe", border: "#bae6fd", text: "#0369a1" };
  }
  if (
    clean.includes("reject") ||
    clean.includes("obsolete") ||
    clean.includes("expired") ||
    clean.includes("cancel") ||
    clean.includes("inactive")
  ) {
    return { bg: "#fde7e9", border: "#f19999", text: "#a80000" };
  }
  return null;
};

export const evaluateDateCustomFormatter = (
  date: Dayjs,
  customFormatterJson?: string
): IChipStyle => {
  if (customFormatterJson) {
    try {
      const raw = customFormatterJson;
      const now = dayjs();
      const diffMs = date.diff(now, "millisecond");

      const pastMatch = /@currentField\s*<=\s*@now\s*,\s*['"]([^'"]+)['"]/i.exec(raw);
      if (date.isBefore(now, "day") && pastMatch) {
        const cls = pastMatch[1].trim().toLowerCase();
        if (SP_FORMAT_CLASS_MAP[cls]) return SP_FORMAT_CLASS_MAP[cls];
        if (cls.startsWith("#") || cls.startsWith("rgb")) {
          return { bg: cls, border: cls, text: "#1a1918" };
        }
      }

      const windowRegex = /@currentField\s*<=\s*@now\s*\+\s*(\d+)\s*,\s*['"]([^'"]+)['"]/gi;
      let winMatch: RegExpExecArray | null;
      while ((winMatch = windowRegex.exec(raw)) !== null) {
        const offsetMs = parseInt(winMatch[1], 10);
        if (diffMs > 0 && diffMs <= offsetMs) {
          const cls = winMatch[2].trim().toLowerCase();
          if (SP_FORMAT_CLASS_MAP[cls]) return SP_FORMAT_CLASS_MAP[cls];
          if (cls.startsWith("#") || cls.startsWith("rgb")) {
            return { bg: cls, border: cls, text: "#1a1918" };
          }
        }
      }

      const lastBranch = /,\s*['"](sp-css-[^'"]+|#[0-9a-fA-F]{3,8})['"]\s*\)+/i.exec(raw);
      if (date.isAfter(now, "day") && lastBranch) {
        const cls = lastBranch[1].trim().toLowerCase();
        if (SP_FORMAT_CLASS_MAP[cls]) return SP_FORMAT_CLASS_MAP[cls];
        if (cls.startsWith("#") || cls.startsWith("rgb")) {
          return { bg: cls, border: cls, text: "#1a1918" };
        }
      }
    } catch (err) {
      console.warn("Could not evaluate date custom formatter:", err);
    }
  }

  if (date.isBefore(dayjs(), "day")) {
    return { bg: "#fde7e9", border: "#f19999", text: "#a80000" };
  }
  if (date.diff(dayjs(), "day") <= 30) {
    return { bg: "#fff4ce", border: "#fde37f", text: "#795b00" };
  }
  return { bg: "#dff6dd", border: "#92c353", text: "#107c10" };
};
