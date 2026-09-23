import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/views/list";
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

    return {
      id: item.Id,
      filename: item.FileLeafRef ?? "",
      fileUrl: item.FileRef,
      PIMProduct: (item.PIMProductCode ?? []).map((p: any) => ({
        ID: p.ID ?? p.Id ?? 0,
        Title: p.Title ?? "",
        PIMProductName: p.PIMProductName ?? "",
        Manufacturer: p.Manufacturer ?? "",
        BusinessLine: p.BusinessLine ?? "",
        ManufacturerLookupId: p.ManufacturerLookupId,
      })),
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
      Modified: modifiedDate,
      EditorEmail: editorEmail,
      EditorTitle: editorTitle,
      Created: createdDate,
      AuthorEmail: authorEmail,
      AuthorTitle: authorTitle,
      IssuedBy: choiceToString(item.Issued_x0020_By),
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
        "ExpiryDateOWSTDATE",
        "CreatedOWSDATE",
        "Created",
        //"ModifiedOWSDate",
        "LastModifiedTime",
        "BusinessLineOWSCHCM",
        "CountryOWSCHCM",
        //"CountryOWSCHM",
        "ManufacturerOWSTEXT",  //Client
        "LongProductNameOWSMTXT",
        "DocumentTypeOWSTEXT",
        "SubDocumentTypeOWSMTXT",        
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
      const CHUNK_SIZE = 40;
      const CONCURRENCY_LIMIT = 8;
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

    // Isahang batch filter query para sa lahat ng product codes sa halip na paisa-isa
    const codes = Array.from(termCodeMap.keys());
    if (codes.length > 0) {
      try {
        const filterClause = codes
          .map((c) => `Title eq '${c.replace(/'/g, "''")}'`)
          .join(" or ");

        const masterItems = await sp.web.lists
          .getByTitle("PIM Product")
          .items.filter(filterClause)
          .select("ID", "Title", "PIMProductName", "Manufacturer", "BusinessLine")
          .top(codes.length + 10)();

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
      try {
        const filterClause = clientTitles
          .map((c) => `Title eq '${c.replace(/'/g, "''")}'`)
          .join(" or ");

        const masterClients = await sp.web.lists
          .getByTitle("PIM Global Client")
          .items.filter(filterClause)
          .select("ID", "Title")
          .top(clientTitles.length + 10)();

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

    combinedUpdatePayload["LongProductName"] = longNames;
    combinedUpdatePayload["PIM_x0020_Product_x0020_Code"] = productCodes;
    combinedUpdatePayload["PIMProductCodeId"] = pIds;

    // Resolve any missing TermGuids from PIM Product masterlist
    for (const p of products) {
      if (!p.TermGuid && p.Title) {
        try {
          const found = await sp.web.lists
            .getByTitle("PIM Product")
            .items.filter(`Title eq '${p.Title.replace(/'/g, "''")}'`)
            .select("ID", "Title", "PIMProductTermSet")
            .top(1)();
          if (found && found.length > 0 && found[0].PIMProductTermSet) {
            const rawTerm = found[0].PIMProductTermSet;
            p.TermGuid = Array.isArray(rawTerm) && rawTerm.length > 0
              ? rawTerm[0].TermGuid
              : rawTerm?.TermGuid || undefined;
          }
        } catch (err) {
          console.warn(`Could not resolve TermGuid for product ${p.Title}:`, err);
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

    const termSetVal = termSetParts.join(";");
    if (termSetVal) {
      taxonomyValues.push({ FieldName: "PIMProductTermSet", FieldValue: termSetVal });
    }
  }

  // 2. Client (multi value)
  if (payload.clientsModified) {
    const clients = payload.selectedClients || [];
    const clientTitles = clients.map((c) => c.Title || "").filter(Boolean).join("; ");
    combinedUpdatePayload["Manufacturer"] = clientTitles;

    const resolvedClientIds: number[] = [];
    for (const c of clients) {
      if (c.ID && c.ID > 0) {
        resolvedClientIds.push(c.ID);
      } else if (c.Title) {
        try {
          const found = await sp.web.lists
            .getByTitle("PIM Global Client")
            .items.filter(`Title eq '${c.Title.replace(/'/g, "''")}'`)
            .select("ID", "Title", "GlobalClientTermSet")
            .top(1)();
          if (found && found.length > 0) {
            resolvedClientIds.push(found[0].ID);
            if (!c.TermGuid && found[0].GlobalClientTermSet) {
              const rawTerm = found[0].GlobalClientTermSet;
              c.TermGuid = Array.isArray(rawTerm) && rawTerm.length > 0
                ? rawTerm[0].TermGuid
                : rawTerm?.TermGuid || undefined;
            }
          }
        } catch {
          // ignore
        }
      }

      // Also ensure TermGuid is populated if still missing
      if (!c.TermGuid && c.Title) {
        try {
          const found = await sp.web.lists
            .getByTitle("PIM Global Client")
            .items.filter(`Title eq '${c.Title.replace(/'/g, "''")}'`)
            .select("ID", "Title", "GlobalClientTermSet")
            .top(1)();
          if (found && found.length > 0 && found[0].GlobalClientTermSet) {
            const rawTerm = found[0].GlobalClientTermSet;
            c.TermGuid = Array.isArray(rawTerm) && rawTerm.length > 0
              ? rawTerm[0].TermGuid
              : rawTerm?.TermGuid || undefined;
          }
        } catch {
          // ignore
        }
      }
    }
    if (resolvedClientIds.length > 0) {
      combinedUpdatePayload["ManufacturerLookupId"] = resolvedClientIds;
    }

    const termSetParts = clients.map((c) => {
      const label = (c.Title || "").trim();
      if (c.TermGuid) {
        return `${label}|${c.TermGuid}`;
      }
      return "";
    }).filter(Boolean);

    const termSetVal = termSetParts.join(";");
    if (termSetVal) {
      taxonomyValues.push({ FieldName: "GlobalClientTermSet", FieldValue: termSetVal });
    }
  }

  // 3. Document Type (single value)
  if (payload.documentTypeModified) {
    const docType = payload.selectedDocumentType;
    combinedUpdatePayload["Document_x0020_Type"] = docType?.Title || "";
    if (docType?.ID && docType.ID > 0) {
      combinedUpdatePayload["DocumentTypeId"] = docType.ID;
    }
  }

  // 4. Sub Document Type (multi value)
  if (payload.subDocumentTypesModified) {
    const subTypes = payload.selectedSubDocumentTypes || [];
    const subTitles = subTypes.map((s) => s.Title || "").filter(Boolean).join("; ");
    const sIds = subTypes.map((s) => s.ID).filter((id) => id > 0);

    combinedUpdatePayload["Sub_x0020_Document_x0020_Type"] = subTitles;
    if (sIds.length > 0) {
      combinedUpdatePayload["SubDocumentTypeId"] = sIds;
    }
  }

  // 5. Issued by (choice)
  if (payload.issuedByModified) {
    combinedUpdatePayload["Issued_x0020_By"] = payload.issuedBy || null;
  }

  // 6. Issuer Name (Supplier - text)
  if (payload.supplierModified) {
    combinedUpdatePayload["Supplier"] = payload.supplier || "";
  }

  // 7. Document Provider Email (Supplier_x0020_Email - text)
  if (payload.supplierEmailModified) {
    combinedUpdatePayload["Supplier_x0020_Email"] = payload.supplierEmail || "";
  }

  // 8. Confidentiality (choice)
  if (payload.confidentialityModified) {
    combinedUpdatePayload["Confidentiality"] = payload.confidentiality || null;
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
    combinedUpdatePayload["Document_x0020_Language"] =
      payload.documentLanguage && payload.documentLanguage.length > 0
        ? payload.documentLanguage
        : null;
  }

  // 13. Document Status (choice)
  if (payload.documentStatusModified) {
    combinedUpdatePayload["Document_x0020_Status"] = payload.documentStatus || null;
  }

  // 14. Customer Name (text)
  if (payload.customerNameModified) {
    combinedUpdatePayload["Customer_x0020_Name"] = payload.customerName || "";
  }

  // 15. Batch Number (text)
  if (payload.batchNumberModified) {
    combinedUpdatePayload["Batch_x0020_Number"] = payload.batchNumber || "";
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

        // Unified Step 1: Isahang POST update para sa lahat ng Text at Lookup columns
        if (Object.keys(combinedUpdatePayload).length > 0) {
          try {
            await item.update(combinedUpdatePayload);
            console.log(`item.update succeeded with unified payload for item ${id}`);
          } catch (updateErr: any) {
            console.warn(`item.update error for item ${id}, trying results-object format fallback:`, updateErr);
            const resultsPayload: Record<string, any> = { ...combinedUpdatePayload };
            if (Array.isArray(resultsPayload.PIMProductCodeId)) {
              resultsPayload.PIMProductCodeId = { results: resultsPayload.PIMProductCodeId };
            }
            if (Array.isArray(resultsPayload.ManufacturerLookupId)) {
              resultsPayload.ManufacturerLookupId = { results: resultsPayload.ManufacturerLookupId };
            }
            if (Array.isArray(resultsPayload.SubDocumentTypeId)) {
              resultsPayload.SubDocumentTypeId = { results: resultsPayload.SubDocumentTypeId };
            }
            try {
              await item.update(resultsPayload);
              console.log(`item.update resultsPayload succeeded for item ${id}`);
            } catch (updateErr2) {
              console.warn(`item.update fallback failed, saving text fields only:`, updateErr2);
              const textOnlyPayload = { ...combinedUpdatePayload };
              delete textOnlyPayload.PIMProductCodeId;
              delete textOnlyPayload.ManufacturerLookupId;
              delete textOnlyPayload.DocumentTypeId;
              delete textOnlyPayload.SubDocumentTypeId;
              await item.update(textOnlyPayload);
              console.log(`item.update textOnlyPayload succeeded for item ${id}`);
            }
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
