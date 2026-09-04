import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/search";
import "@pnp/sp/sharing";
import "@pnp/sp/fields/list";
import { SharingRole } from "@pnp/sp/sharing";
import {
  doclib_AllProducts,
  IProductLookupItem,
  IClientLookupItem,
  IDocumentTypeItem,
  ISubDocumentTypeItem,
  IFieldFormatters,
  ISharingConfig,
  IChipStyle,
} from "../webparts/advanceSearch/types/advanceSearchTypes";
import {
  choiceToString,
  parseSpCustomFormatter,
} from "../webparts/advanceSearch/utils/formatters";

export const mapSharePointItemsToProducts = (items: any[]): doclib_AllProducts[] => {
  return items.map((item: any) => ({
    id: item.Id,
    filename: item.FileLeafRef ?? "",
    fileUrl: item.FileRef,
    PIMProduct: item.PIMProductCode ?? [],
    PIMProductSearchText: (item.PIMProductCode ?? [])
      .map((p: any) => `${p.Title} ${p.PIMProductName}`)
      .join(" "),
    ManufacturerSearchText: item.Manufacturer ?? "",
    BusinessLine: choiceToString(item.Business_x0020_Line),
    CountrySoldTo: choiceToString(item.Country),
    DocumentTypeSearchText: item.Document_x0020_Type ?? "",
    SubDocumentTypeSearchText: item.Sub_x0020_Document_x0020_Type ?? "",
    DocumentDate: item.Document_x0020_Date
      ? new Date(item.Document_x0020_Date)
      : null,
    Confidentiality: choiceToString(item.Confidentiality),
    Alerts: item.Alerts ?? "",
  }));
};

export const searchProducts = async (
  sp: SPFI,
  searchText: string
): Promise<IProductLookupItem[]> => {
  const escapedText = searchText.replace(/'/g, "''");

  const results = await sp.web.lists
    .getByTitle("PIM Product")
    .items.select(
      "ID",
      "Title",
      "PIMProductName",
      "Manufacturer",
      "BusinessLine",
      "ManufacturerLookupId"
    )
    .filter(`substringof('${escapedText}', PIMProductName)`)
    .orderBy("PIMProductName")
    .top(5000)();

  return (results as IProductLookupItem[]).sort((a, b) => {
    const nameA = `${a.Title || ""} ${a.PIMProductName || ""}`.trim();
    const nameB = `${b.Title || ""} ${b.PIMProductName || ""}`.trim();
    return nameA.localeCompare(nameB, undefined, { sensitivity: "base" });
  });
};

export const searchClients = async (
  sp: SPFI,
  searchText: string
): Promise<IClientLookupItem[]> => {
  const escapedText = searchText.replace(/'/g, "''");

  const results = await sp.web.lists
    .getByTitle("PIM Global Client")
    .items.select("ID", "Title")
    .filter(`substringof('${escapedText}', Title)`)
    .orderBy("Title")
    .top(5000)();

  return (results as IClientLookupItem[]).sort((a, b) =>
    (a.Title || "").localeCompare(b.Title || "", undefined, { sensitivity: "base" })
  );
};

export const loadTaxonomy = async (
  sp: SPFI
): Promise<{
  docTypes: IDocumentTypeItem[];
  subDocTypes: ISubDocumentTypeItem[];
}> => {
  if (!sp) {
    console.warn("loadTaxonomy: SPFI instance is undefined or null");
    return { docTypes: [], subDocTypes: [] };
  }

  let docTypes: IDocumentTypeItem[] = [];
  let subDocTypes: ISubDocumentTypeItem[] = [];

  try {
    const rawDocTypes = await sp.web.lists
      .getByTitle("Document Type")
      .items.select("ID", "Title", "ShortTitle")
      .top(5000)
      .orderBy("Title")();

    docTypes = (rawDocTypes as IDocumentTypeItem[]).sort((a, b) =>
      (a.Title || "").localeCompare(b.Title || "", undefined, { sensitivity: "base" })
    );
  } catch (error) {
    console.warn("loadTaxonomy 'Document Type' error:", error);
  }

  try {
    const rawSubDocTypes = await sp.web.lists
      .getByTitle("Sub Document Type")
      .items.select("ID", "Title", "DocumentType/Title")
      .expand("DocumentType")
      .top(2000)
      .orderBy("Title")();

    subDocTypes = (rawSubDocTypes as ISubDocumentTypeItem[]).sort((a, b) =>
      (a.Title || "").localeCompare(b.Title || "", undefined, { sensitivity: "base" })
    );
  } catch (error) {
    console.warn("loadTaxonomy 'Sub Document Type' error:", error);
  }

  return { docTypes, subDocTypes };
};

export const loadSharingConfiguration = async (sp: SPFI): Promise<ISharingConfig> => {
  if (!sp) return { subject: "", message: "" };
  try {
    const items = await sp.web.lists
      .getByTitle("Configuration")
      .items.select("Title", "Subject", "Message")
      .filter("Title eq 'External Sharing'")
      .top(1)();

    if (items && items.length > 0) {
      const config = items[0];
      return {
        subject: config.Subject ?? "",
        message: config.Message ?? "",
      };
    }
  } catch (err) {
    console.warn("Could not load 'Configuration' list item for 'External Sharing':", err);
  }
  return { subject: "", message: "" };
};

export const loadListFieldFormatting = async (sp: SPFI): Promise<IFieldFormatters> => {
  if (!sp) return { businessLine: {}, confidentiality: {} };
  try {
    const fields = await sp.web.lists
      .getByTitle("Clients & Products")
      .fields.select("InternalName", "CustomFormatter")
      .filter("InternalName eq 'Confidentiality' or InternalName eq 'Business_x0020_Line'")();

    let blFormat: Record<string, IChipStyle> = {};
    let confFormat: Record<string, IChipStyle> = {};

    fields.forEach((f: any) => {
      if (f.InternalName === "Business_x0020_Line") {
        blFormat = parseSpCustomFormatter(f.CustomFormatter);
      } else if (f.InternalName === "Confidentiality") {
        confFormat = parseSpCustomFormatter(f.CustomFormatter);
      }
    });

    return { businessLine: blFormat, confidentiality: confFormat };
  } catch (err) {
    console.warn("Could not load field CustomFormatter:", err);
    return { businessLine: {}, confidentiality: {} };
  }
};

export interface IBatchLoadResult {
  items: doclib_AllProducts[];
  nextSkipId?: number;
  hasMore: boolean;
}

export interface ISearchBatchResult {
  items: doclib_AllProducts[];
  nextStartRow?: number;
  hasMore: boolean;
  totalRows?: number;
}

export const loadRecordsBatch = async (
  sp: SPFI,
  lastId?: number,
  pageSize: number = 1000
): Promise<IBatchLoadResult> => {
  let itemsQuery = sp.web.lists
    .getByTitle("Clients & Products")
    .items.select(
      "Id",
      "Title",
      "FileLeafRef",
      "FileRef",
      "Country",
      "Business_x0020_Line",
      "PIMProductCode/Title",
      "PIMProductCode/PIMProductName",
      "Manufacturer",
      "Document_x0020_Type",
      "Sub_x0020_Document_x0020_Type",
      "Document_x0020_Date",
      "Alerts",
      "Confidentiality"
    )
    .expand("PIMProductCode")
    .orderBy("Id", false)
    .top(pageSize);

  if (lastId !== undefined && lastId > 0) {
    itemsQuery = itemsQuery.filter(`Id lt ${lastId}`);
  }

  const batch = await itemsQuery();
  const mapped = mapSharePointItemsToProducts(batch);
  const lowestId = batch.length > 0 ? batch[batch.length - 1].Id : undefined;

  return {
    items: mapped,
    nextSkipId: lowestId,
    hasMore: batch.length >= pageSize,
  };
};

export const loadAllRecords = async (
  sp: SPFI,
  pageSize: number = 5000
): Promise<doclib_AllProducts[]> => {
  const allProducts: any[] = [];
  let lastId = 0;

  while (true) {
    const batch = await sp.web.lists
      .getByTitle("Clients & Products")
      .items.select(
        "Id",
        "Title",
        "FileLeafRef",
        "FileRef",
        "Country",
        "Business_x0020_Line",
        "PIMProductCode/Title",
        "PIMProductCode/PIMProductName",
        "Manufacturer",
        "Document_x0020_Type",
        "Sub_x0020_Document_x0020_Type",
        "Document_x0020_Date",
        "Alerts",
        "Confidentiality"
      )
      .expand("PIMProductCode")
      .filter(`Id gt ${lastId}`)
      .orderBy("Id")
      .top(pageSize)();

    if (batch.length === 0) break;

    allProducts.push(...batch);
    lastId = batch[batch.length - 1].Id;

    if (batch.length < pageSize) break;
  }

  return mapSharePointItemsToProducts(allProducts);
};

export const searchRecords = async (
  sp: SPFI,
  queryText: string,
  startRow: number = 0,
  pageSize: number = 1000
): Promise<ISearchBatchResult> => {
  let currentRow = startRow;
  let totalRows: number | undefined = undefined;
  const allOrderedItems: doclib_AllProducts[] = [];
  const seenIds = new Set<number>();
  let hitEnd = false;

  while (allOrderedItems.length < pageSize && !hitEnd) {
    const needed = pageSize - allOrderedItems.length;
    // Fetch up to 500 at a time (SharePoint REST Search limit per request) or what is needed
    const currentLimit = Math.min(500, needed);

    const results = await sp.search({
      Querytext: queryText,
      RowLimit: currentLimit,
      StartRow: currentRow,
      TrimDuplicates: false,
      SelectProperties: [
        "Title",
        "ListItemID",
        "Path",
        "DocumentDateOWSTDATE",
        "BusinessLineOWSCHCM",
        "CountryOWSCHCM",
        "ManufacturerOWSTEXT",
        "LongProductNameOWSMTXT",
        "DocumentTypeOWSTEXT",
        "SubDocumentTypeOWSMTXT",
        "owstaxIdPIMProductTermSet",
        "PIMProductCodeOWSTEXT",
        "ConfidentialityOWSCHCS",
        "AlertsOWSMTXT",
      ],
    });

    if (results && typeof results.TotalRows === "number") {
      totalRows = results.TotalRows;
    }

    const currentResults = results?.PrimarySearchResults ?? [];
    currentRow += currentResults.length;

    if (currentResults.length < currentLimit) {
      hitEnd = true;
    }

    const chunkIds: number[] = [];
    currentResults.forEach((r: any) => {
      const id = Number(r.ListItemID);
      if (!isNaN(id) && id > 0 && !seenIds.has(id)) {
        seenIds.add(id);
        chunkIds.push(id);
      }
    });

    if (chunkIds.length > 0) {
      const chunkItems: any[] = [];
      for (let i = 0; i < chunkIds.length; i += 100) {
        const slice = chunkIds.slice(i, i + 100);
        const filter = slice.map((id) => `Id eq ${id}`).join(" or ");

        const items = await sp.web.lists
          .getByTitle("Clients & Products")
          .items.select(
            "Id",
            "Title",
            "FileLeafRef",
            "FileRef",
            "Country",
            "Business_x0020_Line",
            "PIMProductCode/Title",
            "PIMProductCode/PIMProductName",
            "Manufacturer",
            "Document_x0020_Type",
            "Sub_x0020_Document_x0020_Type",
            "Document_x0020_Date",
            "Alerts",
            "Confidentiality"
          )
          .expand("PIMProductCode")
          .filter(filter)();

        chunkItems.push(...items);
      }

      const itemMap = new Map<number, any>();
      chunkItems.forEach((item) => itemMap.set(item.Id, item));

      const orderedChunk: any[] = [];
      chunkIds.forEach((id) => {
        const item = itemMap.get(id);
        if (item) orderedChunk.push(item);
      });

      const mapped = mapSharePointItemsToProducts(orderedChunk);
      allOrderedItems.push(...mapped);
    }
  }

  const hasMore = !hitEnd && allOrderedItems.length >= pageSize;

  return {
    items: allOrderedItems,
    nextStartRow: hasMore ? currentRow : undefined,
    hasMore,
    totalRows,
  };
};

export const shareFilesByEmail = async (
  sp: SPFI,
  selectedItems: any[],
  toEmails: string[],
  ccEmails: string[],
  bccEmails: string[],
  subject: string,
  message: string
): Promise<void> => {
  const origin = window.location.origin;
  const allRecipients = Array.from(new Set([...toEmails, ...ccEmails, ...bccEmails]));

  let cleanMessage = (message || "").replace(/{}/g, "").trim();
  if (cleanMessage.length > 490) {
    cleanMessage = cleanMessage.substring(0, 487) + "...";
  }

  for (let index = 0; index < selectedItems.length; index++) {
    const item = selectedItems[index];
    const itemUrl: string = item.fileUrl || "";
    const fullUrl = itemUrl.startsWith("http") ? itemUrl : `${origin}${itemUrl}`;

    const result = await sp.web.shareObject(
      fullUrl,
      allRecipients,
      SharingRole.View,
      {
        subject: (subject || "Shared document").substring(0, 200),
        body: cleanMessage || "Please find the shared document.",
      }
    );

    if (result && result.ErrorMessage) {
      throw new Error(result.ErrorMessage);
    }
  }
};
