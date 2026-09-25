import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/views/list";
import "@pnp/sp/items";
import "@pnp/sp/search";
import "@pnp/sp/sharing";
import "@pnp/sp/fields/list";
import "@pnp/sp/security";
import { SharingRole } from "@pnp/sp/sharing";
import { PermissionKind } from "@pnp/sp/security";
import {
  doclib_AllProducts,
  IProductLookupItem,
  IClientLookupItem,
  IDocumentTypeItem,
  ISubDocumentTypeItem,
  IFieldFormatters,
  IAlertRule,
  ISharingConfig,
  IChipStyle,
} from "../webparts/advanceSearch/types/advanceSearchTypes";
import {
  choiceToString,
  parseSpCustomFormatter,
  parseAlertsCustomFormatter,
} from "../webparts/advanceSearch/utils/formatters";

export const mapSharePointItemsToProducts = (
  items: any[]
): doclib_AllProducts[] => {
  return items.map((item: any) => {
    const editorEmail = item.Editor?.EMail || item.Editor?.Email || "";
    const editorTitle = item.Editor?.Title || "";
    const modifiedDate = item.Modified ? new Date(item.Modified) : null;
    const authorEmail = item.Author?.EMail || item.Author?.Email || "";
    const authorTitle = item.Author?.Title || "";
    const createdDate = item.Created ? new Date(item.Created) : null;
    const reviewedByEmail = Array.isArray(item.ReviewActionTakenBy)
      ? item.ReviewActionTakenBy.map((u: any) => u.EMail || u.Email).filter(Boolean).join(", ")
      : (item.ReviewActionTakenBy?.EMail || item.ReviewActionTakenBy?.Email || "");
    const reviewedByTitle = Array.isArray(item.ReviewActionTakenBy)
      ? item.ReviewActionTakenBy.map((u: any) => u.Title).filter(Boolean).join(", ")
      : (item.ReviewActionTakenBy?.Title || "");
    const expiryDate = item.Expiry_x0020_Date ? new Date(item.Expiry_x0020_Date) : null;
    const nextReviewDate = item.Next_x0020_Review_x0020_Date ? new Date(item.Next_x0020_Review_x0020_Date) : null;

    // Resolve Products from lookup PIMProductCode
    const pimProducts = (item.PIMProductCode ?? []).map((p: any) => ({
      ID: p.ID ?? p.Id ?? 0,
      Title: p.Title ?? "",
      PIMProductName: p.PIMProductName ?? "",
      Manufacturer: p.Manufacturer ?? "",
      BusinessLine: p.BusinessLine ?? "",
      ManufacturerLookupId: p.ManufacturerLookupId,
    }));

    const pimProductSearchText = (item.PIMProductCode ?? [])
      .map((p: any) => `${p.Title || ""} ${p.PIMProductName || ""}`.trim())
      .filter(Boolean)
      .join(" ");

    const longProductNames = pimProducts
      .map((p: IProductLookupItem) => p.PIMProductName || "")
      .filter(Boolean)
      .join("; ");

    return {
      id: item.Id,
      filename: item.FileLeafRef ?? "",
      fileUrl: item.FileRef,
      PIMProduct: pimProducts,
      PIMProductSearchText: pimProductSearchText,
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
      Modified: modifiedDate,
      EditorEmail: editorEmail,
      EditorTitle: editorTitle,
      Created: createdDate,
      AuthorEmail: authorEmail,
      AuthorTitle: authorTitle,
      IssuedBy: choiceToString(item.Issued_x0020_By),
      Supplier: choiceToString(item.Supplier) || item.Supplier || "",
      SupplierEmail: item.Supplier_x0020_Email ?? "",
      CustomerName: item.Customer_x0020_Name ?? "",
      BatchNumber: item.Batch_x0020_Number ?? "",
      OriginalFilename: item.Original_x0020_Filename ?? item.OriginalFilename ?? "",
      OData__dlc_DocId:
        (typeof item.OData__dlc_DocId === "string" ? item.OData__dlc_DocId : "") ||
        (item.OData__dlc_DocIdUrl?.Description || item.OData__dlc_DocIdUrl?.Url || "") ||
        (typeof item._dlc_DocId === "string" ? item._dlc_DocId : "") ||
        (item._dlc_DocIdUrl?.Description || item._dlc_DocIdUrl?.Url || "") ||
        "",
      DocVersion: item.DocVersion !== undefined && item.DocVersion !== null ? item.DocVersion : (item.OData__UIVersionString ?? null),
      LongProductName: longProductNames || item.LongProductName || "",
      DocumentStatus: choiceToString(item.Document_x0020_Status),
      ReviewedByTitle: reviewedByTitle,
      ReviewedByEmail: reviewedByEmail,
      DocumentLanguage: choiceToString(item.Document_x0020_Language),
      ExpiryDate: expiryDate,
      NextReviewDate: nextReviewDate,
    };
  });
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
      "ManufacturerLookupId",
      "PIMProductTermSet"
    )
    .filter(`substringof('${escapedText}', PIMProductName)`)
    .orderBy("PIMProductName")
    .top(5000)();

  return (results as any[]).map((item) => {
    const rawTerm = item.PIMProductTermSet;
    const termGuid = Array.isArray(rawTerm) && rawTerm.length > 0
      ? rawTerm[0].TermGuid
      : rawTerm?.TermGuid || undefined;
    const wssId = Array.isArray(rawTerm) && rawTerm.length > 0
      ? rawTerm[0].WssId
      : rawTerm?.WssId || undefined;

    return {
      ID: item.ID,
      Title: item.Title || "",
      PIMProductName: item.PIMProductName || "",
      Manufacturer: item.Manufacturer || "",
      BusinessLine: item.BusinessLine || "",
      ManufacturerLookupId: item.ManufacturerLookupId,
      TermGuid: termGuid,
      WssId: wssId,
    };
  }).sort((a, b) => {
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
    .items.select("ID", "Title", "GlobalClientTermSet")
    .filter(`substringof('${escapedText}', Title)`)
    .orderBy("Title")
    .top(5000)();

  return (results as any[]).map((item) => {
    const rawTerm = item.GlobalClientTermSet;
    const termGuid = Array.isArray(rawTerm) && rawTerm.length > 0
      ? rawTerm[0].TermGuid
      : rawTerm?.TermGuid || undefined;
    const wssId = Array.isArray(rawTerm) && rawTerm.length > 0
      ? rawTerm[0].WssId
      : rawTerm?.WssId || undefined;

    return {
      ID: item.ID,
      Title: item.Title || "",
      TermGuid: termGuid,
      WssId: wssId,
    };
  }).sort((a, b) =>
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

export interface ILibraryColumnChoices {
  businessLine: string[];
  country: string[];
  confidentiality: string[];
  issuedBy: string[];
  documentStatus: string[];
  documentLanguage: string[];
}

export interface IListFieldMetadata {
  formatters: IFieldFormatters;
  choices: ILibraryColumnChoices;
  defaultViewUrl?: string;
}

export const loadListFieldMetadata = async (sp: SPFI): Promise<IListFieldMetadata> => {
  const defaultResult: IListFieldMetadata = {
    formatters: { businessLine: {}, confidentiality: {}, alerts: null, documentStatus: {} },
    choices: {
      businessLine: [],
      country: [],
      confidentiality: [],
      issuedBy: [],
      documentStatus: [],
      documentLanguage: [],
    },
  };

  if (!sp) return defaultResult;

  try {
    const list = sp.web.lists.getByTitle("Clients & Products");

    const [fields, defaultView] = await Promise.all([
      list.fields
        .select("InternalName", "CustomFormatter", "Choices")
        .filter(
          "InternalName eq 'Business_x0020_Line' or InternalName eq 'Country' or InternalName eq 'Confidentiality' or InternalName eq 'Alerts' or InternalName eq 'Document_x0020_Status' or InternalName eq 'Expiry_x0020_Date' or InternalName eq 'Next_x0020_Review_x0020_Date' or InternalName eq 'Issued_x0020_By' or InternalName eq 'Document_x0020_Language'"
        )(),
      list.views
        .filter("DefaultView eq true")
        .select("ServerRelativeUrl")()
        .then((views: any[]) => (views && views.length > 0 ? views[0] : null))
        .catch((err: any) => {
          console.warn("Could not retrieve defaultView for Clients & Products:", err);
          return null;
        }),
    ]);

    let blFormat: Record<string, IChipStyle> = {};
    let confFormat: Record<string, IChipStyle> = {};
    let docStatusFormat: Record<string, IChipStyle> = {};
    let expiryDateFormatter: string | undefined = undefined;
    let nextReviewDateFormatter: string | undefined = undefined;
    let alertsRule: IAlertRule | null = null;
    let businessLineChoices: string[] = [];
    let countryChoices: string[] = [];
    let confidentialityChoices: string[] = [];
    let issuedByChoices: string[] = [];
    let docStatusChoices: string[] = [];
    let docLangChoices: string[] = [];

    fields.forEach((f: any) => {
      const choices: string[] = Array.isArray(f.Choices) ? f.Choices : [];
      if (f.InternalName === "Business_x0020_Line") {
        if (f.CustomFormatter) {
          blFormat = parseSpCustomFormatter(f.CustomFormatter);
        }
        businessLineChoices = choices;
      } else if (f.InternalName === "Country") {
        countryChoices = choices;
      } else if (f.InternalName === "Confidentiality") {
        if (f.CustomFormatter) {
          confFormat = parseSpCustomFormatter(f.CustomFormatter);
        }
        confidentialityChoices = choices;
      } else if (f.InternalName === "Alerts") {
        if (f.CustomFormatter) {
          alertsRule = parseAlertsCustomFormatter(f.CustomFormatter);
        }
      } else if (f.InternalName === "Document_x0020_Status") {
        if (f.CustomFormatter) {
          docStatusFormat = parseSpCustomFormatter(f.CustomFormatter);
        }
        docStatusChoices = choices;
      } else if (f.InternalName === "Issued_x0020_By") {
        issuedByChoices = choices;
      } else if (f.InternalName === "Document_x0020_Language") {
        docLangChoices = choices;
      } else if (f.InternalName === "Expiry_x0020_Date") {
        if (f.CustomFormatter) {
          expiryDateFormatter = f.CustomFormatter;
        }
      } else if (f.InternalName === "Next_x0020_Review_x0020_Date") {
        if (f.CustomFormatter) {
          nextReviewDateFormatter = f.CustomFormatter;
        }
      }
    });

    return {
      formatters: {
        businessLine: blFormat,
        confidentiality: confFormat,
        alerts: alertsRule,
        documentStatus: docStatusFormat,
        expiryDateCustomFormatter: expiryDateFormatter,
        nextReviewDateCustomFormatter: nextReviewDateFormatter,
      },
      choices: {
        businessLine: businessLineChoices.sort((a, b) =>
          a.localeCompare(b, undefined, { sensitivity: "base" })
        ),
        country: countryChoices.sort((a, b) =>
          a.localeCompare(b, undefined, { sensitivity: "base" })
        ),
        confidentiality: confidentialityChoices.sort((a, b) =>
          a.localeCompare(b, undefined, { sensitivity: "base" })
        ),
        issuedBy: issuedByChoices.sort((a, b) =>
          a.localeCompare(b, undefined, { sensitivity: "base" })
        ),
        documentStatus: docStatusChoices.sort((a, b) =>
          a.localeCompare(b, undefined, { sensitivity: "base" })
        ),
        documentLanguage: docLangChoices.sort((a, b) =>
          a.localeCompare(b, undefined, { sensitivity: "base" })
        ),
      },
      defaultViewUrl: defaultView?.ServerRelativeUrl || undefined,
    };
  } catch (err) {
    console.warn("Could not load field metadata from Clients & Products:", err);
    return defaultResult;
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
      "PIMProductCode/Id",
      "PIMProductCode/Title",
      "PIMProductCode/PIMProductName",
      "Manufacturer",
      "Document_x0020_Type",
      "Sub_x0020_Document_x0020_Type",
      "Document_x0020_Date",
      "Alerts",
      "Confidentiality",
      "Modified",
      "Editor/Id",
      "Editor/Title",
      "Editor/EMail",
      "Created",
      "Author/Id",
      "Author/Title",
      "Author/EMail",
      "Issued_x0020_By",
      "Supplier",
      "Supplier_x0020_Email",
      "Customer_x0020_Name",
      "Batch_x0020_Number",
      "Original_x0020_Filename",
      "OData__dlc_DocId",
      "DocVersion",
      "Document_x0020_Status",
      "ReviewActionTakenBy/Id",
      "ReviewActionTakenBy/Title",
      "ReviewActionTakenBy/EMail",
      "Document_x0020_Language",
      "Expiry_x0020_Date",
      "Next_x0020_Review_x0020_Date"
    )
    .expand("PIMProductCode", "Editor", "Author", "ReviewActionTakenBy")
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

export const fetchSingleProductItem = async (
  sp: SPFI,
  itemId: number
): Promise<doclib_AllProducts | null> => {
  if (!sp || !itemId || itemId <= 0) return null;
  try {
    const raw = await sp.web.lists
      .getByTitle("Clients & Products")
      .items.getById(itemId)
      .select(
        "Id",
        "Title",
        "FileLeafRef",
        "FileRef",
        "Country",
        "Business_x0020_Line",
        "PIMProductCode/Id",
        "PIMProductCode/Title",
        "PIMProductCode/PIMProductName",
        "Manufacturer",
        "Document_x0020_Type",
        "Sub_x0020_Document_x0020_Type",
        "Document_x0020_Date",
        "Alerts",
        "Confidentiality",
        "Modified",
        "Editor/Id",
        "Editor/Title",
        "Editor/EMail",
        "Created",
        "Author/Id",
        "Author/Title",
        "Author/EMail",
        "Issued_x0020_By",
        "Supplier",
        "Supplier_x0020_Email",
        "Customer_x0020_Name",
        "Batch_x0020_Number",
        "Original_x0020_Filename",
        "OData__dlc_DocId",
        "DocVersion",
        "Document_x0020_Status",
        "ReviewActionTakenBy/Id",
        "ReviewActionTakenBy/Title",
        "ReviewActionTakenBy/EMail",
        "Document_x0020_Language",
        "Expiry_x0020_Date",
        "Next_x0020_Review_x0020_Date"
      )
      .expand("PIMProductCode", "Editor", "Author", "ReviewActionTakenBy")();

    if (!raw) return null;
    const mapped = mapSharePointItemsToProducts([raw]);
    return mapped && mapped.length > 0 ? mapped[0] : null;
  } catch (err) {
    console.warn(`fetchSingleProductItem error for item ${itemId}:`, err);
    return null;
  }
};

export const searchRecords = async (
  sp: SPFI,
  queryText: string,
  startRow: number = 0,
  pageSize: number = 1000,
  existingIds?: Set<number>
): Promise<ISearchBatchResult> => {
  let currentRow = startRow;
  let totalRows: number | undefined = undefined;
  const allOrderedItems: doclib_AllProducts[] = [];
  const seenIds = new Set<number>(existingIds);
  let hitEnd = false;

  const selectProperties = [
    "Title",
    "ListItemID",
    "Path",
    "DocumentDateOWSTDATE",
    "ExpiryDateOWSTDATE",
    "CreatedOWSDATE",
    "Created",
    "LastModifiedTime",
    "BusinessLineOWSCHCM",
    "CountryOWSCHCM",
    "ManufacturerOWSTEXT", // Client
    "LongProductNameOWSMTXT",
    "DocumentTypeOWSTEXT",
    "SubDocumentTypeOWSMTXT",
    "PIMProductCodeOWSTEXT",
    "ConfidentialityOWSCHCS",
    "AlertsOWSMTXT",
  ];

  while (allOrderedItems.length < pageSize && !hitEnd) {
    const needed = pageSize - allOrderedItems.length;
    let currentResults: any[] = [];

    if (needed > 500) {
      // Parallelize into 2 concurrent search requests (500 each)
      const [res1, res2] = await Promise.all([
        sp.search({
          Querytext: queryText,
          RowLimit: 500,
          StartRow: currentRow,
          TrimDuplicates: false,
          SelectProperties: selectProperties,
        }),
        sp.search({
          Querytext: queryText,
          RowLimit: 500,
          StartRow: currentRow + 500,
          TrimDuplicates: false,
          SelectProperties: selectProperties,
        }),
      ]);

      if (res1 && typeof res1.TotalRows === "number") {
        totalRows = res1.TotalRows;
      } else if (res2 && typeof res2.TotalRows === "number") {
        totalRows = res2.TotalRows;
      }

      const list1 = res1?.PrimarySearchResults ?? [];
      const list2 = res2?.PrimarySearchResults ?? [];

      currentResults = [...list1, ...list2];
      currentRow += list1.length + list2.length;

      if (list1.length < 500 || list2.length < 500) {
        hitEnd = true;
      }
    } else {
      const currentLimit = Math.min(500, Math.max(needed, 50));
      const results = await sp.search({
        Querytext: queryText,
        RowLimit: currentLimit,
        StartRow: currentRow,
        TrimDuplicates: false,
        SelectProperties: selectProperties,
      });

      if (results && typeof results.TotalRows === "number") {
        totalRows = results.TotalRows;
      }

      currentResults = results?.PrimarySearchResults ?? [];
      currentRow += currentResults.length;

      if (currentResults.length < currentLimit) {
        hitEnd = true;
      }
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
      const CHUNK_SIZE = 60;
      const CONCURRENCY_LIMIT = 10;
      const chunkSlices: number[][] = [];
      for (let i = 0; i < chunkIds.length; i += CHUNK_SIZE) {
        chunkSlices.push(chunkIds.slice(i, i + CHUNK_SIZE));
      }

      const chunkItems: any[] = [];
      for (let i = 0; i < chunkSlices.length; i += CONCURRENCY_LIMIT) {
        const batchSlices = chunkSlices.slice(i, i + CONCURRENCY_LIMIT);
        const batchPromises = batchSlices.map((slice) => {
          const filter = slice.map((id) => `Id eq ${id}`).join(" or ");
          return sp.web.lists
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
              "Confidentiality",
              "Modified",
              "Editor/Id",
              "Editor/Title",
              "Editor/EMail",
              "Created",
              "Author/Id",
              "Author/Title",
              "Author/EMail",
              "Issued_x0020_By",
              "Supplier",
              "Supplier_x0020_Email",
              "Customer_x0020_Name",
              "Batch_x0020_Number",
              "Original_x0020_Filename",
              "OData__dlc_DocId",
              "DocVersion",
              "Document_x0020_Status",
              "ReviewActionTakenBy/Id",
              "ReviewActionTakenBy/Title",
              "ReviewActionTakenBy/EMail",
              "Document_x0020_Language",
              "Expiry_x0020_Date",
              "Next_x0020_Review_x0020_Date"
            )
            .expand("PIMProductCode", "Editor", "Author", "ReviewActionTakenBy")
            .filter(filter)
            .top(CHUNK_SIZE)();
        });

        const batchResults = await Promise.all(batchPromises);
        batchResults.forEach((items) => {
          chunkItems.push(...items);
        });
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

export const checkUserWritePermissions = async (
  sp: SPFI,
  listTitle: string = "Clients & Products",
  itemIds?: number[]
): Promise<boolean> => {
  if (!sp) return false;
  try {
    const list = sp.web.lists.getByTitle(listTitle);

    // 1. Check list/library-level EditListItems permission
    const hasListPermission = await list.currentUserHasPermissions(PermissionKind.EditListItems);
    if (!hasListPermission) {
      return false;
    }

    // 2. If specific item IDs are provided, verify item-level EditListItems permissions
    if (itemIds && itemIds.length > 0) {
      const targetIds = itemIds.slice(0, 10);
      const itemChecks = await Promise.all(
        targetIds.map(async (id) => {
          try {
            return await list.items.getById(id).currentUserHasPermissions(PermissionKind.EditListItems);
          } catch (itemErr) {
            console.warn(`Could not check permissions for item ${id}:`, itemErr);
            return false;
          }
        })
      );

      if (itemChecks.some((canEdit) => !canEdit)) {
        return false;
      }
    }

    return true;
  } catch (error) {
    console.warn(`checkUserWritePermissions failed for "${listTitle}":`, error);
    return false;
  }
};

export interface ILoadedItemForEdit {
  id: number;
  products: IProductLookupItem[];
  clients: IClientLookupItem[];
  documentType: IDocumentTypeItem | null;
  subDocumentTypes: ISubDocumentTypeItem[];
  issuedBy?: string;
  supplier?: string;
  supplierEmail?: string;
  confidentiality?: string;
  documentDate?: Date | null;
  expiryDate?: Date | null;
  nextReviewDate?: Date | null;
  documentLanguage?: string[];
  documentStatus?: string;
  customerName?: string;
  batchNumber?: string;
}

export const loadItemDetailsForEdit = async (
  sp: SPFI,
  itemId: number
): Promise<ILoadedItemForEdit> => {
  const raw = await sp.web.lists
    .getByTitle("Clients & Products")
    .items.getById(itemId)
    .select(
      "Id",
      "PIMProductCode/Id",
      "PIMProductCode/Title",
      "PIMProductCode/PIMProductName",
      "Manufacturer",
      "GlobalClientTermSet",
      "PIMProductTermSet",
      "Document_x0020_Type",
      "Sub_x0020_Document_x0020_Type",
      "Issued_x0020_By",
      "Supplier",
      "Supplier_x0020_Email",
      "Confidentiality",
      "Document_x0020_Date",
      "Expiry_x0020_Date",
      "Next_x0020_Review_x0020_Date",
      "Document_x0020_Language",
      "Document_x0020_Status",
      "Customer_x0020_Name",
      "Batch_x0020_Number"
    )
    .expand("PIMProductCode")();

  // 1. Products: Load directly from PIMProductTermSet (Isahang batch query!)
  let products: IProductLookupItem[] = [];
  const rawProductTerms = raw.PIMProductTermSet;
  if (Array.isArray(rawProductTerms) && rawProductTerms.length > 0) {
    const termCodeMap = new Map<string, { label: string; name: string; termGuid?: string; wssId?: number }>();

    for (const t of rawProductTerms) {
      const label = (t.Label || "").trim();
      let code = "";
      let name = "";

      if (label.includes(" : ")) {
        const parts = label.split(" : ");
        code = parts[0].trim();
        name = parts.slice(1).join(" : ").trim();
      } else if (label.includes(":")) {
        const parts = label.split(":");
        code = parts[0].trim();
        name = parts.slice(1).join(":").trim();
      } else if (label.includes(" ")) {
        const spaceIdx = label.indexOf(" ");
        code = label.substring(0, spaceIdx).trim();
        name = label.substring(spaceIdx + 1).trim();
      } else {
        code = label;
      }

      if (code) {
        termCodeMap.set(code.toLowerCase(), {
          label,
          name,
          termGuid: t.TermGuid,
          wssId: t.WssId,
        });
      }
    }

    // Isahang batch filter query para sa lahat ng product codes sa halip na paisa-isa (chunked into slices of 25)
    const codes = Array.from(termCodeMap.keys());
    if (codes.length > 0) {
      const CHUNK = 25;
      for (let i = 0; i < codes.length; i += CHUNK) {
        const slice = codes.slice(i, i + CHUNK);
        try {
          const filterClause = slice
            .map((c) => `Title eq '${c.replace(/'/g, "''")}'`)
            .join(" or ");

          const masterItems = await sp.web.lists
            .getByTitle("PIM Product")
            .items.filter(filterClause)
            .select("ID", "Title", "PIMProductName", "Manufacturer", "BusinessLine")
            .top(slice.length + 10)();

          masterItems.forEach((mItem: any) => {
            const mCode = (mItem.Title || "").trim().toLowerCase();
            const info = termCodeMap.get(mCode);
            products.push({
              ID: mItem.ID,
              Title: mItem.Title,
              PIMProductName: mItem.PIMProductName || info?.name || "",
              Manufacturer: mItem.Manufacturer || "",
              BusinessLine: mItem.BusinessLine || "",
              TermGuid: info?.termGuid,
              WssId: info?.wssId,
            });
            termCodeMap.delete(mCode);
          });
        } catch (err) {
          console.warn("Batch resolving products failed:", err);
        }
      }
    }

    // Fallback for any product code not found in masterlist
    termCodeMap.forEach((info, codeKey) => {
      products.push({
        ID: -(products.length + 1),
        Title: codeKey.toUpperCase(),
        PIMProductName: info.name,
        TermGuid: info.termGuid,
        WssId: info.wssId,
      });
    });
  }

  // Fallback if PIMProductTermSet was empty: use PIMProductCode lookup
  if (products.length === 0 && Array.isArray(raw.PIMProductCode) && raw.PIMProductCode.length > 0) {
    products = raw.PIMProductCode.map((p: any) => ({
      ID: p.Id ?? p.ID ?? 0,
      Title: p.Title ?? "",
      PIMProductName: p.PIMProductName ?? "",
    }));
  }

  // 2. Clients: Load directly from GlobalClientTermSet (Isahang batch query!)
  let clients: IClientLookupItem[] = [];
  const rawClientTerms = raw.GlobalClientTermSet;
  if (Array.isArray(rawClientTerms) && rawClientTerms.length > 0) {
    const clientTitleMap = new Map<string, { termGuid?: string; wssId?: number }>();
    for (const t of rawClientTerms) {
      const clientTitle = (t.Label || "").trim();
      if (clientTitle) {
        clientTitleMap.set(clientTitle.toLowerCase(), {
          termGuid: t.TermGuid,
          wssId: t.WssId,
        });
      }
    }

    const clientTitles = Array.from(clientTitleMap.keys());
    if (clientTitles.length > 0) {
      const CHUNK = 25;
      for (let i = 0; i < clientTitles.length; i += CHUNK) {
        const slice = clientTitles.slice(i, i + CHUNK);
        try {
          const filterClause = slice
            .map((c) => `Title eq '${c.replace(/'/g, "''")}'`)
            .join(" or ");

          const masterClients = await sp.web.lists
            .getByTitle("PIM Global Client")
            .items.filter(filterClause)
            .select("ID", "Title")
            .top(slice.length + 10)();

          masterClients.forEach((mClient: any) => {
            const cTitle = (mClient.Title || "").trim().toLowerCase();
            const info = clientTitleMap.get(cTitle);
            clients.push({
              ID: mClient.ID,
              Title: mClient.Title,
              TermGuid: info?.termGuid,
              WssId: info?.wssId,
            });
            clientTitleMap.delete(cTitle);
          });
        } catch (err) {
          console.warn("Batch resolving clients failed:", err);
        }
      }
    }

    clientTitleMap.forEach((info, cTitle) => {
      clients.push({
        ID: -(clients.length + 1),
        Title: cTitle,
        TermGuid: info.termGuid,
        WssId: info.wssId,
      });
    });
  }

  // Fallback: If termset was empty on document, read from Manufacturer text field
  if (clients.length === 0 && raw.Manufacturer) {
    const names = String(raw.Manufacturer)
      .split(/[\r\n;,]+/)
      .map((c) => c.trim())
      .filter(Boolean);

    for (const name of names) {
      clients.push({ ID: -(clients.length + 1), Title: name });
    }
  }

  // 3. Document Type from text field
  const docType: IDocumentTypeItem | null = raw.Document_x0020_Type
    ? {
        ID: 0,
        Title: String(raw.Document_x0020_Type).trim(),
      }
    : null;

  // 4. Sub Document Type from text field
  const subDocTypes: ISubDocumentTypeItem[] = raw.Sub_x0020_Document_x0020_Type
    ? String(raw.Sub_x0020_Document_x0020_Type)
        .split(/[\r\n;,]+/)
        .map((s) => s.trim())
        .filter(Boolean)
        .map((title, idx) => ({ ID: -(idx + 1), Title: title }))
    : [];

  const parseDateOrNull = (dateVal: any): Date | null => {
    if (!dateVal) return null;
    const d = new Date(dateVal);
    return isNaN(d.getTime()) ? null : d;
  };

  return {
    id: raw.Id,
    products,
    clients,
    documentType: docType,
    subDocumentTypes: subDocTypes,
    issuedBy: choiceToString(raw.Issued_x0020_By),
    supplier: raw.Supplier || "",
    supplierEmail: raw.Supplier_x0020_Email || "",
    confidentiality: choiceToString(raw.Confidentiality),
    documentDate: parseDateOrNull(raw.Document_x0020_Date),
    expiryDate: parseDateOrNull(raw.Expiry_x0020_Date),
    nextReviewDate: parseDateOrNull(raw.Next_x0020_Review_x0020_Date),
    documentLanguage: Array.isArray(raw.Document_x0020_Language)
      ? raw.Document_x0020_Language
      : raw.Document_x0020_Language
      ? String(raw.Document_x0020_Language).split(/[\r\n;,]+/).map((s) => s.trim()).filter(Boolean)
      : [],
    documentStatus: choiceToString(raw.Document_x0020_Status),
    customerName: raw.Customer_x0020_Name || "",
    batchNumber: raw.Batch_x0020_Number || "",
  };
};

export interface IEditPropertiesPayload {
  productsModified?: boolean;
  selectedProducts?: IProductLookupItem[];

  clientsModified?: boolean;
  selectedClients?: IClientLookupItem[];

  documentTypeModified?: boolean;
  selectedDocumentType?: IDocumentTypeItem | null;

  subDocumentTypesModified?: boolean;
  selectedSubDocumentTypes?: ISubDocumentTypeItem[];

  issuedByModified?: boolean;
  issuedBy?: string;

  supplierModified?: boolean;
  supplier?: string;

  supplierEmailModified?: boolean;
  supplierEmail?: string;

  confidentialityModified?: boolean;
  confidentiality?: string;

  documentDateModified?: boolean;
  documentDate?: Date | null;

  expiryDateModified?: boolean;
  expiryDate?: Date | null;

  nextReviewDateModified?: boolean;
  nextReviewDate?: Date | null;

  documentLanguageModified?: boolean;
  documentLanguage?: string[];

  documentStatusModified?: boolean;
  documentStatus?: string;

  customerNameModified?: boolean;
  customerName?: string;

  batchNumberModified?: boolean;
  batchNumber?: string;
}

const safeText = (val: string | undefined | null, maxLen: number = 255): string | null => {
  if (val === undefined || val === null) return null;
  const trimmed = String(val).trim();
  if (!trimmed) return null;
  return trimmed.length > maxLen ? trimmed.substring(0, maxLen) : trimmed;
};

export const updateItemProperties = async (
  sp: SPFI,
  itemIds: number[],
  payload: IEditPropertiesPayload
): Promise<void> => {
  if (!sp || !itemIds || itemIds.length === 0) return;

  const list = sp.web.lists.getByTitle("Clients & Products");

  const combinedUpdatePayload: Record<string, any> = {};
  const taxonomyValues: { FieldName: string; FieldValue: string }[] = [];

  // 1. Product (multi value)
  if (payload.productsModified) {
    const products = payload.selectedProducts || [];
    const longNames = products.map((p) => p.PIMProductName || "").filter(Boolean).join("; ");
    const productCodes = products.map((p) => p.Title || "").filter(Boolean).join("; ");
    const pIds = products.map((p) => p.ID).filter((id) => id > 0);

    // Limit text fields to max 255 characters to avoid SPException -2130575336 on Single Line Text fields
    combinedUpdatePayload["LongProductName"] = safeText(longNames);
    combinedUpdatePayload["PIM_x0020_Product_x0020_Code"] = safeText(productCodes);
    combinedUpdatePayload["PIMProductCodeId"] = pIds;

    // Resolve any missing TermGuids from PIM Product masterlist in chunked batches
    const missingTermProducts = products.filter((p) => !p.TermGuid && p.Title);
    if (missingTermProducts.length > 0) {
      const CHUNK = 25;
      for (let i = 0; i < missingTermProducts.length; i += CHUNK) {
        const slice = missingTermProducts.slice(i, i + CHUNK);
        const filterClause = slice.map((p) => `Title eq '${p.Title.replace(/'/g, "''")}'`).join(" or ");
        try {
          const found = await sp.web.lists
            .getByTitle("PIM Product")
            .items.filter(filterClause)
            .select("ID", "Title", "PIMProductTermSet")
            .top(slice.length + 5)();
          found.forEach((mItem: any) => {
            const match = slice.find((p) => (p.Title || "").toLowerCase() === (mItem.Title || "").toLowerCase());
            if (match && mItem.PIMProductTermSet) {
              const rawTerm = mItem.PIMProductTermSet;
              match.TermGuid = Array.isArray(rawTerm) && rawTerm.length > 0
                ? rawTerm[0].TermGuid
                : rawTerm?.TermGuid || undefined;
            }
          });
        } catch (err) {
          console.warn("Could not batch resolve TermGuids for products:", err);
        }
      }
    }

    const termSetParts = products.map((p) => {
      const code = (p.Title || "").trim();
      const name = (p.PIMProductName || "").trim();
      const label = code && name ? `${code} : ${name}` : code || name;
      if (p.TermGuid) {
        return `${label}|${p.TermGuid}`;
      }
      return "";
    }).filter(Boolean);

    taxonomyValues.push({ FieldName: "PIMProductTermSet", FieldValue: termSetParts.join(";") });
  }

  // 2. Client (multi value)
  if (payload.clientsModified) {
    const clients = payload.selectedClients || [];
    const clientTitles = clients.map((c) => c.Title || "").filter(Boolean).join("; ");
    combinedUpdatePayload["Manufacturer"] = safeText(clientTitles);

    const resolvedClientIds: number[] = [];
    const missingClients = clients.filter((c) => (!c.ID || c.ID <= 0 || !c.TermGuid) && c.Title);
    if (missingClients.length > 0) {
      const CHUNK = 25;
      for (let i = 0; i < missingClients.length; i += CHUNK) {
        const slice = missingClients.slice(i, i + CHUNK);
        const filterClause = slice.map((c) => `Title eq '${c.Title.replace(/'/g, "''")}'`).join(" or ");
        try {
          const found = await sp.web.lists
            .getByTitle("PIM Global Client")
            .items.filter(filterClause)
            .select("ID", "Title", "GlobalClientTermSet")
            .top(slice.length + 5)();
          found.forEach((mClient: any) => {
            const match = slice.find((c) => (c.Title || "").toLowerCase() === (mClient.Title || "").toLowerCase());
            if (match) {
              if (!match.ID || match.ID <= 0) match.ID = mClient.ID;
              if (!match.TermGuid && mClient.GlobalClientTermSet) {
                const rawTerm = mClient.GlobalClientTermSet;
                match.TermGuid = Array.isArray(rawTerm) && rawTerm.length > 0
                  ? rawTerm[0].TermGuid
                  : rawTerm?.TermGuid || undefined;
              }
            }
          });
        } catch (err) {
          console.warn("Could not batch resolve clients:", err);
        }
      }
    }

    clients.forEach((c) => {
      if (c.ID && c.ID > 0) {
        resolvedClientIds.push(c.ID);
      }
    });

    combinedUpdatePayload["ManufacturerLookupId"] = resolvedClientIds;

    const termSetParts = clients.map((c) => {
      const label = (c.Title || "").trim();
      if (c.TermGuid) {
        return `${label}|${c.TermGuid}`;
      }
      return "";
    }).filter(Boolean);

    taxonomyValues.push({ FieldName: "GlobalClientTermSet", FieldValue: termSetParts.join(";") });
  }

  // 3. Document Type (single value)
  if (payload.documentTypeModified) {
    const docType = payload.selectedDocumentType;
    combinedUpdatePayload["Document_x0020_Type"] = safeText(docType?.Title);
    combinedUpdatePayload["DocumentTypeId"] = docType?.ID && docType.ID > 0 ? docType.ID : null;
  }

  // 4. Sub Document Type (multi value)
  if (payload.subDocumentTypesModified) {
    const subTypes = payload.selectedSubDocumentTypes || [];
    const subTitles = subTypes.map((s) => s.Title || "").filter(Boolean).join("; ");
    const sIds = subTypes.map((s) => s.ID).filter((id) => id > 0);

    combinedUpdatePayload["Sub_x0020_Document_x0020_Type"] = safeText(subTitles);
    combinedUpdatePayload["SubDocumentTypeId"] = sIds;
  }

  // 5. Issued by (choice)
  if (payload.issuedByModified) {
    combinedUpdatePayload["Issued_x0020_By"] = safeText(payload.issuedBy);
  }

  // 6. Issuer Name (Supplier - text)
  if (payload.supplierModified) {
    combinedUpdatePayload["Supplier"] = safeText(payload.supplier);
  }

  // 7. Document Provider Email (Supplier_x0020_Email - text)
  if (payload.supplierEmailModified) {
    combinedUpdatePayload["Supplier_x0020_Email"] = safeText(payload.supplierEmail);
  }

  // 8. Confidentiality (choice)
  if (payload.confidentialityModified) {
    combinedUpdatePayload["Confidentiality"] = safeText(payload.confidentiality);
  }

  // 9. Document Date (date)
  if (payload.documentDateModified) {
    combinedUpdatePayload["Document_x0020_Date"] = payload.documentDate
      ? payload.documentDate.toISOString()
      : null;
  }

  // 10. Expiry Date (date)
  if (payload.expiryDateModified) {
    combinedUpdatePayload["Expiry_x0020_Date"] = payload.expiryDate
      ? payload.expiryDate.toISOString()
      : null;
  }

  // 11. Next Review Date (date)
  if (payload.nextReviewDateModified) {
    combinedUpdatePayload["Next_x0020_Review_x0020_Date"] = payload.nextReviewDate
      ? payload.nextReviewDate.toISOString()
      : null;
  }

  // 12. Document Language (multi choice)
  if (payload.documentLanguageModified) {
    const langs = payload.documentLanguage && payload.documentLanguage.length > 0
      ? payload.documentLanguage
      : [];
    combinedUpdatePayload["Document_x0020_Language"] = langs;
  }

  // 13. Document Status (choice)
  if (payload.documentStatusModified) {
    combinedUpdatePayload["Document_x0020_Status"] = safeText(payload.documentStatus);
  }

  // 14. Customer Name (text)
  if (payload.customerNameModified) {
    combinedUpdatePayload["Customer_x0020_Name"] = safeText(payload.customerName);
  }

  // 15. Batch Number (text)
  if (payload.batchNumberModified) {
    combinedUpdatePayload["Batch_x0020_Number"] = safeText(payload.batchNumber);
  }

  console.log("Saving item properties to SharePoint (unified payload):", {
    itemIds,
    combinedUpdatePayload,
    taxonomyValues,
  });

  const CHUNK_SIZE = 10;
  for (let i = 0; i < itemIds.length; i += CHUNK_SIZE) {
    const chunk = itemIds.slice(i, i + CHUNK_SIZE);
    await Promise.all(
      chunk.map(async (id) => {
        const item = list.items.getById(id);

        // Step 1: Update Text, Date, Choice and Lookup columns via item.update
        if (Object.keys(combinedUpdatePayload).length > 0) {
          try {
            await item.update(combinedUpdatePayload);
            console.log(`item.update succeeded with unified payload for item ${id}`);
          } catch (updateErr: any) {
            console.error(`item.update error for item ${id}:`, updateErr);
            throw updateErr;
          }
        }

        // Step 2: Update Managed Metadata TermSets via validateUpdateListItem
        if (taxonomyValues.length > 0) {
          try {
            const results = await item.validateUpdateListItem(taxonomyValues);
            console.log(`validateUpdateListItem result for item ${id}:`, results);

            // Check if validateUpdateListItem reported any field validation errors
            if (Array.isArray(results)) {
              results.forEach((r: any) => {
                if (r.HasException) {
                  console.error(`Taxonomy validation exception on field ${r.FieldName}: ${r.ErrorMessage}`);
                }
              });
            }
          } catch (taxErr) {
            console.warn(`Taxonomy update notice for item ${id}:`, taxErr);
          }
        }
      })
    );
  }
};
