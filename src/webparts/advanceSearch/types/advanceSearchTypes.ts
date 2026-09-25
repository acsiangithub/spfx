export type doclib_AllProducts = {
  id?: number;
  filename: string;
  PIMProduct: IProductLookupItem[];
  PIMProductSearchText: string;
  BusinessLine: string;
  CountrySoldTo: string;
  ManufacturerSearchText: string;
  DocumentTypeSearchText: string;
  SubDocumentTypeSearchText: string;
  DocumentDate: Date | null;
  fileUrl: string;
  Confidentiality: string;
  Alerts: string;
  Modified?: Date | null;
  EditorEmail?: string;
  EditorTitle?: string;
  Created?: Date | null;
  AuthorEmail?: string;
  AuthorTitle?: string;
  IssuedBy?: string;
  Supplier?: string;
  SupplierEmail?: string;
  CustomerName?: string;
  BatchNumber?: string;
  OriginalFilename?: string;
  OData__dlc_DocId?: string;
  DocVersion?: number | string | null;
  LongProductName?: string;
  DocumentStatus?: string;
  ReviewedByTitle?: string;
  ReviewedByEmail?: string;
  DocumentLanguage?: string;
  ExpiryDate?: Date | null;
  NextReviewDate?: Date | null;
};

export type IProductLookupItem = {
  ID: number;
  Title: string;
  PIMProductName: string;
  Manufacturer?: string;
  BusinessLine?: string;
  ManufacturerLookupId?: number;
  TermGuid?: string;
  WssId?: number;
};

export type IClientLookupItem = {
  ID: number;
  Title: string;
  TermGuid?: string;
  WssId?: number;
};

export type IDocumentTypeItem = {
  ID: number;
  Title: string;
  ShortTitle?: string;
};

export type ISubDocumentTypeItem = {
  ID: number;
  Title: string;
  DocumentType?: {
    Title?: string;
  } | null;
};

export interface IChipStyle {
  bg: string;
  border: string;
  text: string;
}

export interface IAlertRule {
  durationMinutes: number;
  requiresEditorMe: boolean;
  text: string;
  color: string;
  style: IChipStyle;
}

export interface IFieldFormatters {
  businessLine: Record<string, IChipStyle>;
  confidentiality: Record<string, IChipStyle>;
  alerts?: IAlertRule | null;
  documentStatus: Record<string, IChipStyle>;
  issuedBy: Record<string, IChipStyle>;
  expiryDateCustomFormatter?: string;
  nextReviewDateCustomFormatter?: string;
}

export interface ISharingConfig {
  subject: string;
  message: string;
}
