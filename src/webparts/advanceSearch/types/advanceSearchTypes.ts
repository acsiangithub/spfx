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
}

export interface ISharingConfig {
  subject: string;
  message: string;
}
