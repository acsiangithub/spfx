import dayjs from "dayjs";
import { doclib_AllProducts } from "../types/advanceSearchTypes";

export const isValueEmpty = (value: unknown): boolean => {
  if (value === null || value === undefined) return true;
  if (Array.isArray(value)) {
    if (value.length === 0) return true;
    return value.every((v) => isValueEmpty(v));
  }
  const str = String(value).trim();
  return str === "";
};

export const multiSelectFilterFn = (
  row: { getValue: (columnId: string) => unknown },
  columnId: string,
  filterValue: unknown
): boolean => {
  if (filterValue === undefined || filterValue === null || filterValue === "") return true;
  if (Array.isArray(filterValue) && filterValue.length === 0) return true;

  const selectedValues = (Array.isArray(filterValue) ? filterValue : [filterValue])
    .map((v) => String(v).trim())
    .filter((v) => v !== "");

  if (selectedValues.length === 0) return true;

  const raw = row.getValue(columnId);
  const empty = isValueEmpty(raw);

  const matchEmpty = selectedValues.some((v) => v.toLowerCase() === "(empty)");
  const regularSelections = selectedValues.filter((v) => v.toLowerCase() !== "(empty)");

  if (empty) {
    return matchEmpty;
  }

  if (regularSelections.length === 0) {
    return false;
  }

  const rowStr = String(raw).trim();
  const tokens = rowStr
    .split(/[\r\n;,]+/)
    .map((t) => t.trim().toLowerCase())
    .filter(Boolean);

  return regularSelections.some((sel) => {
    const selLower = sel.toLowerCase();
    if (tokens.includes(selLower)) return true;
    if (columnId === "PIMProductSearchText") {
      return rowStr.toLowerCase().includes(selLower);
    }
    return false;
  });
};

export const documentDateFilter = (
  row: { getValue: (columnId: string) => unknown },
  columnId: string,
  filterValue: unknown
): boolean => {
  const rowValue = row.getValue(columnId);
  if (!rowValue || !filterValue) return true;

  const rowDate = dayjs(rowValue as string | Date);
  const filterDate = dayjs(filterValue as string);

  if (!rowDate.isValid() || !filterDate.isValid()) return true;

  return rowDate.isSame(filterDate, "day") || rowDate.isAfter(filterDate, "day");
};

export const exactDateFilter = (
  row: { getValue: (columnId: string) => unknown },
  columnId: string,
  filterValue: unknown
): boolean => {
  const rowValue = row.getValue(columnId);
  if (!rowValue || !filterValue) return true;

  const rowDate = dayjs(rowValue as string | Date);
  const filterDate = dayjs(filterValue as string);

  if (!rowDate.isValid() || !filterDate.isValid()) return true;

  return rowDate.isSame(filterDate, "day");
};

export const isPersonMe = (
  email: string | undefined,
  title: string | undefined,
  curEmail: string | undefined,
  curName: string | undefined
): boolean => {
  const e = (email || "").toLowerCase().trim();
  const t = (title || "").toLowerCase().trim();
  const cE = (curEmail || "").toLowerCase().trim();
  const cN = (curName || "").toLowerCase().trim();
  return (
    (Boolean(cE) && Boolean(e) && (e === cE || e.includes(cE) || cE.includes(e))) ||
    (Boolean(cN) && Boolean(t) && (t === cN || t.includes(cN) || cN.includes(t)))
  );
};

export type FileTypeCategory = "word" | "excel" | "powerpoint" | "pdf";

export const FILE_TYPE_EXTENSIONS: Record<FileTypeCategory, string[]> = {
  word: [".doc", ".docx", ".docm", ".dot", ".dotx"],
  excel: [".xls", ".xlsx", ".xlsm", ".xlsb", ".csv"],
  powerpoint: [".ppt", ".pptx", ".pptm", ".pot", ".potx", ".pps", ".ppsx"],
  pdf: [".pdf"],
};

export const matchesFileType = (filename: string, fileTypes: string[]): boolean => {
  if (!filename || !fileTypes || fileTypes.length === 0) return true;
  const lower = filename.toLowerCase().trim();
  return fileTypes.some((type) => {
    const exts = FILE_TYPE_EXTENSIONS[type as FileTypeCategory];
    if (exts) {
      return exts.some((ext) => lower.endsWith(ext));
    }
    return false;
  });
};

export const filenameFilterFn = (
  row: { original: doclib_AllProducts; getValue: (columnId: string) => unknown },
  _columnId: string,
  filterValue: unknown
): boolean => {
  if (filterValue === undefined || filterValue === null || filterValue === "") return true;
  const filename = (row.original.filename || "").toLowerCase();

  if (typeof filterValue === "string") {
    return filename.includes(filterValue.toLowerCase().trim());
  }

  if (typeof filterValue === "object" && filterValue !== null) {
    const fVal = filterValue as { text?: string; fileTypes?: string[] };
    if (fVal.text && fVal.text.trim()) {
      if (!filename.includes(fVal.text.trim().toLowerCase())) return false;
    }
    if (fVal.fileTypes && fVal.fileTypes.length > 0) {
      return matchesFileType(filename, fVal.fileTypes);
    }
  }

  return true;
};

export const itemMatchesFilter = (
  item: doclib_AllProducts,
  colId: string,
  filterValue: unknown,
  currentUserEmail?: string,
  currentUserName?: string
): boolean => {
  if (filterValue === undefined || filterValue === null || filterValue === "") return true;
  if (Array.isArray(filterValue) && filterValue.length === 0) return true;

  if (colId === "id") {
    if (item.id === undefined || item.id === null) return false;
    return String(item.id).trim() === String(filterValue).trim();
  }

  if (colId === "filename") {
    const fname = (item.filename || "").toLowerCase();
    if (typeof filterValue === "string") {
      return fname.includes(filterValue.toLowerCase().trim());
    }
    if (typeof filterValue === "object" && filterValue !== null) {
      const fVal = filterValue as { text?: string; fileTypes?: string[] };
      if (fVal.text && fVal.text.trim()) {
        if (!fname.includes(fVal.text.trim().toLowerCase())) return false;
      }
      if (fVal.fileTypes && fVal.fileTypes.length > 0) {
        return matchesFileType(fname, fVal.fileTypes);
      }
      return true;
    }
    return true;
  }
  if (colId === "Alerts") {
    return (item.Alerts || "").toLowerCase().includes(String(filterValue).toLowerCase());
  }
  if (colId === "DocumentDate" || colId === "ExpiryDate" || colId === "NextReviewDate") {
    const rawDate = (item as any)[colId];
    if (!rawDate) return false;
    const rowDate = dayjs(rawDate);
    const filterDate = dayjs(filterValue as string);
    if (!rowDate.isValid() || !filterDate.isValid()) return true;
    return rowDate.isSame(filterDate, "day") || rowDate.isAfter(filterDate, "day");
  }
  if (colId === "Created" || colId === "Modified") {
    const rawDate = (item as any)[colId];
    if (!rawDate) return false;
    const rowDate = dayjs(rawDate);
    const filterDate = dayjs(filterValue as string);
    if (!rowDate.isValid() || !filterDate.isValid()) return true;
    return rowDate.isSame(filterDate, "day");
  }

  const selectedValues = (Array.isArray(filterValue) ? filterValue : [filterValue])
    .map((v) => String(v).trim())
    .filter((v) => v !== "");

  if (selectedValues.length === 0) return true;

  if (colId === "AuthorTitle" || colId === "EditorTitle" || colId === "ReviewedByTitle") {
    let personEmail = "";
    let personTitle = "";
    if (colId === "AuthorTitle") {
      personEmail = item.AuthorEmail || "";
      personTitle = item.AuthorTitle || "";
    } else if (colId === "EditorTitle") {
      personEmail = item.EditorEmail || "";
      personTitle = item.EditorTitle || "";
    } else if (colId === "ReviewedByTitle") {
      personEmail = item.ReviewedByEmail || "";
      personTitle = item.ReviewedByTitle || "";
    }

    const empty = !personTitle && !personEmail;
    const matchEmpty = selectedValues.some((v) => v.toLowerCase() === "(empty)");
    if (empty) return matchEmpty;

    const matchMe = selectedValues.some((v) => v.toLowerCase() === "me");
    if (matchMe && isPersonMe(personEmail, personTitle, currentUserEmail, currentUserName)) {
      return true;
    }

    const regularSelections = selectedValues
      .filter((v) => v.toLowerCase() !== "(empty)" && v.toLowerCase() !== "me")
      .map((v) => v.toLowerCase());

    if (regularSelections.length === 0) return false;

    const tokens = personTitle
      .split(/[\r\n;,]+/)
      .map((t) => t.trim().toLowerCase())
      .filter(Boolean);

    return regularSelections.some((sel) => tokens.includes(sel));
  }

  const raw = (item as Record<string, unknown>)[colId];
  const empty = isValueEmpty(raw);

  const matchEmpty = selectedValues.some((v) => v.toLowerCase() === "(empty)");
  const regularSelections = selectedValues.filter((v) => v.toLowerCase() !== "(empty)");

  if (empty) {
    return matchEmpty;
  }

  if (regularSelections.length === 0) {
    return false;
  }

  const rowStr = String(raw).trim();
  const tokens = rowStr
    .split(/[\r\n;,]+/)
    .map((t) => t.trim().toLowerCase())
    .filter(Boolean);

  return regularSelections.some((sel) => {
    const selLower = sel.toLowerCase();
    if (tokens.includes(selLower)) return true;
    if (colId === "PIMProductSearchText") {
      return rowStr.toLowerCase().includes(selLower);
    }
    return false;
  });
};
