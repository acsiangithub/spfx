import dayjs from "dayjs";
import { doclib_AllProducts } from "../types/advanceSearchTypes";

export const multiSelectFilterFn = (
  row: { getValue: (columnId: string) => unknown },
  columnId: string,
  filterValue: unknown
): boolean => {
  if (!filterValue || (Array.isArray(filterValue) && filterValue.length === 0)) return true;
  const rowValue = String(row.getValue(columnId) || "").toLowerCase();
  const selectedValues = Array.isArray(filterValue) ? filterValue : [filterValue];
  return selectedValues.some((val) => rowValue.includes(String(val).toLowerCase()));
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

  const rawValue = (item as Record<string, unknown>)[colId];
  const itemVal = String(rawValue || "").toLowerCase();
  const selectedValues = Array.isArray(filterValue) ? filterValue : [filterValue];
  return selectedValues.some((val) => itemVal.includes(String(val).toLowerCase()));
};
