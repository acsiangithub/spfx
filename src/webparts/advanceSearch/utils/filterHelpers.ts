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

export const itemMatchesFilter = (
  item: doclib_AllProducts,
  colId: string,
  filterValue: unknown
): boolean => {
  if (filterValue === undefined || filterValue === null || filterValue === "") return true;
  if (Array.isArray(filterValue) && filterValue.length === 0) return true;

  if (colId === "filename") {
    return (item.filename || "").toLowerCase().includes(String(filterValue).toLowerCase());
  }
  if (colId === "Alerts") {
    return (item.Alerts || "").toLowerCase().includes(String(filterValue).toLowerCase());
  }
  if (colId === "DocumentDate") {
    if (!item.DocumentDate) return false;
    const rowDate = dayjs(item.DocumentDate);
    const filterDate = dayjs(filterValue as string);
    if (!rowDate.isValid() || !filterDate.isValid()) return true;
    return rowDate.isSame(filterDate, "day") || rowDate.isAfter(filterDate, "day");
  }

  const selectedValues = (Array.isArray(filterValue) ? filterValue : [filterValue])
    .map((v) => String(v).trim())
    .filter((v) => v !== "");

  if (selectedValues.length === 0) return true;

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
