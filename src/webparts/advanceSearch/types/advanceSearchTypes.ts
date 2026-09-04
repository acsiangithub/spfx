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
};

export type IProductLookupItem = {
  ID: number;
  Title: string;
  PIMProductName: string;
  Manufacturer?: string;
  BusinessLine?: string;
  ManufacturerLookupId?: number;
};

export type IClientLookupItem = {
  ID: number;
  Title: string;
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

export interface IFieldFormatters {
  businessLine: Record<string, IChipStyle>;
  confidentiality: Record<string, IChipStyle>;
}

export interface ISharingConfig {
  subject: string;
  message: string;
}
