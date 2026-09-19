import * as React from "react";
import { useMemo } from "react";
import { IAdvanceSearchProps } from "./IAdvanceSearchProps";
import { sp } from "../AdvanceSearchWebPart";
import { spfi, SPFx, SPFI } from "@pnp/sp";
import {
  MaterialReactTable,
  useMaterialReactTable,
  MRT_ShowHideColumnsButton,
  MRT_ToggleDensePaddingButton,
  MRT_ToggleGlobalFilterButton,
  type MRT_ColumnDef,
  type MRT_ColumnFiltersState,
  type MRT_GroupingState,
  type MRT_VisibilityState,
  type MRT_SortingState,
  type MRT_ColumnPinningState,
} from "material-react-table";

import Autocomplete from "@mui/material/Autocomplete";
import TextField from "@mui/material/TextField";
import Chip from "@mui/material/Chip";
import CircularProgress from "@mui/material/CircularProgress";
import Button from "@mui/material/Button";
import IconButton from "@mui/material/IconButton";
import ShareIcon from "@mui/icons-material/Share";
import FilterListOffIcon from "@mui/icons-material/FilterListOff";
import BookmarkBorderIcon from "@mui/icons-material/BookmarkBorder";
import DeleteOutlineIcon from "@mui/icons-material/DeleteOutline";
import KeyboardArrowDownIcon from "@mui/icons-material/KeyboardArrowDown";
import CheckIcon from "@mui/icons-material/Check";
import Dialog from "@mui/material/Dialog";
import DialogTitle from "@mui/material/DialogTitle";
import DialogContent from "@mui/material/DialogContent";
import DialogActions from "@mui/material/DialogActions";
import FormControl from "@mui/material/FormControl";
import FormControlLabel from "@mui/material/FormControlLabel";
import Switch from "@mui/material/Switch";
import InputLabel from "@mui/material/InputLabel";
import MenuItem from "@mui/material/MenuItem";
import Menu from "@mui/material/Menu";
import Select from "@mui/material/Select";
import Checkbox from "@mui/material/Checkbox";
import ListItemText from "@mui/material/ListItemText";
import OutlinedInput from "@mui/material/OutlinedInput";
import Divider from "@mui/material/Divider";
import Box from "@mui/material/Box";
import Typography from "@mui/material/Typography";
import Accordion from "@mui/material/Accordion";
import AccordionSummary from "@mui/material/AccordionSummary";
import AccordionDetails from "@mui/material/AccordionDetails";
import ExpandMoreIcon from "@mui/icons-material/ExpandMore";
import Link from "@mui/material/Link";
import CloseIcon from "@mui/icons-material/Close";
import RefreshIcon from "@mui/icons-material/Refresh";
import Tooltip from "@mui/material/Tooltip";
import { ThemeProvider } from "@mui/material/styles";
import { DatePicker } from "@mui/x-date-pickers/DatePicker";
import { LocalizationProvider } from "@mui/x-date-pickers/LocalizationProvider";
import { AdapterDayjs } from "@mui/x-date-pickers/AdapterDayjs";
import dayjs, { Dayjs } from "dayjs";

import {
  doclib_AllProducts,
  IProductLookupItem,
  IClientLookupItem,
  IDocumentTypeItem,
  ISubDocumentTypeItem,
  IFieldFormatters,
  ISharingConfig,
} from "../types/advanceSearchTypes";
import {
  sanitizeKqlValue,
  getDynamicChipStyle,
  getSemanticConfidentialityStyle,
  getSemanticAlertStyle,
  getSemanticDocumentStatusStyle,
  evaluateDateCustomFormatter,
  compileKeywordsToKql,
} from "../utils/formatters";
import {
  multiSelectFilterFn,
  documentDateFilter,
  exactDateFilter,
  isPersonMe,
  itemMatchesFilter,
  filenameFilterFn,
  isFilterActive,
} from "../utils/filterHelpers";
import { compactTheme } from "../theme/compactTheme";
import {
  searchProducts as searchProductsService,
  searchClients as searchClientsService,
  loadTaxonomy as loadTaxonomyService,
  loadSharingConfiguration as loadSharingConfigService,
  loadListFieldMetadata as loadListFieldMetadataService,
  loadRecordsBatch as loadRecordsBatchService,
  searchRecords as searchRecordsService,
  ILibraryColumnChoices,
} from "../../../services/sharePointService";
import EmailShareDialog from "./EmailShareDialog";

const DateFilterControl: React.FC<{
  column: { getFilterValue: () => unknown; setFilterValue: (value: unknown) => void };
  label?: string;
}> = ({ column, label = "Min Date" }) => {
  const filterValue = column.getFilterValue() as string | null;
  const pickerValue = filterValue ? dayjs(filterValue) : null;

  return (
    <DatePicker
      format="DD/MM/YYYY"
      label={label}
      value={pickerValue && pickerValue.isValid() ? pickerValue : null}
      onChange={(newValue) => {
        column.setFilterValue(newValue?.isValid() ? newValue.toISOString() : undefined);
      }}
      slotProps={{
        field: { clearable: true },
        textField: {
          size: "small",
          sx: {
            width: "100%",
            minWidth: "130px",
            "& .MuiInputBase-root": {
              paddingRight: "8px",
              paddingLeft: 0,
            },
            "& .MuiInputBase-input": {
              fontSize: "12px",
              padding: "6px 2px 6px 8px",
              minWidth: 0,
            },
            "& .MuiInputAdornment-root": {
              marginLeft: "2px",
              marginRight: 0,
              gap: "2px",
            },
            "& .MuiIconButton-root": {
              padding: "3px",
            },
            "& .MuiSvgIcon-root": {
              fontSize: "17px",
            },
          },
        },
      }}
    />
  );
};

const WordFileTypeIcon: React.FC<{ size?: number }> = ({ size = 16 }) => (
  <svg width={size} height={size} viewBox="0 0 16 16" fill="none" style={{ flexShrink: 0, display: "inline-block", verticalAlign: "middle" }}>
    <path d="M2.5 1.5A1.5 1.5 0 0 1 4 0h6a1.5 1.5 0 0 1 1.06.44l3.5 3.5A1.5 1.5 0 0 1 15 5v9.5A1.5 1.5 0 0 1 13.5 16H4A1.5 1.5 0 0 1 2.5 14.5v-13z" fill="#185ABD" />
    <path d="M10 0v3.5a1.5 1.5 0 0 0 1.5 1.5H15" fill="#0D47A1" opacity="0.6" />
    <text x="8.5" y="12.5" textAnchor="middle" fill="#ffffff" fontSize="8" fontWeight="bold" fontFamily="Segoe UI, Roboto, Helvetica, Arial, sans-serif">W</text>
  </svg>
);

const ExcelFileTypeIcon: React.FC<{ size?: number }> = ({ size = 16 }) => (
  <svg width={size} height={size} viewBox="0 0 16 16" fill="none" style={{ flexShrink: 0, display: "inline-block", verticalAlign: "middle" }}>
    <path d="M2.5 1.5A1.5 1.5 0 0 1 4 0h6a1.5 1.5 0 0 1 1.06.44l3.5 3.5A1.5 1.5 0 0 1 15 5v9.5A1.5 1.5 0 0 1 13.5 16H4A1.5 1.5 0 0 1 2.5 14.5v-13z" fill="#107C41" />
    <path d="M10 0v3.5a1.5 1.5 0 0 0 1.5 1.5H15" fill="#0A5C2F" opacity="0.6" />
    <text x="8.5" y="12.5" textAnchor="middle" fill="#ffffff" fontSize="8" fontWeight="bold" fontFamily="Segoe UI, Roboto, Helvetica, Arial, sans-serif">X</text>
  </svg>
);

const PptFileTypeIcon: React.FC<{ size?: number }> = ({ size = 16 }) => (
  <svg width={size} height={size} viewBox="0 0 16 16" fill="none" style={{ flexShrink: 0, display: "inline-block", verticalAlign: "middle" }}>
    <path d="M2.5 1.5A1.5 1.5 0 0 1 4 0h6a1.5 1.5 0 0 1 1.06.44l3.5 3.5A1.5 1.5 0 0 1 15 5v9.5A1.5 1.5 0 0 1 13.5 16H4A1.5 1.5 0 0 1 2.5 14.5v-13z" fill="#C43E1C" />
    <path d="M10 0v3.5a1.5 1.5 0 0 0 1.5 1.5H15" fill="#9E2F13" opacity="0.6" />
    <text x="8.5" y="12.5" textAnchor="middle" fill="#ffffff" fontSize="8" fontWeight="bold" fontFamily="Segoe UI, Roboto, Helvetica, Arial, sans-serif">P</text>
  </svg>
);

const PdfFileTypeIcon: React.FC<{ size?: number }> = ({ size = 16 }) => (
  <svg width={size} height={size} viewBox="0 0 16 16" fill="none" style={{ flexShrink: 0, display: "inline-block", verticalAlign: "middle" }}>
    <path d="M2.5 1.5A1.5 1.5 0 0 1 4 0h6a1.5 1.5 0 0 1 1.06.44l3.5 3.5A1.5 1.5 0 0 1 15 5v9.5A1.5 1.5 0 0 1 13.5 16H4A1.5 1.5 0 0 1 2.5 14.5v-13z" fill="#B30B00" />
    <path d="M10 0v3.5a1.5 1.5 0 0 0 1.5 1.5H15" fill="#7F0800" opacity="0.6" />
    <text x="8.5" y="12" textAnchor="middle" fill="#ffffff" fontSize="5.5" fontWeight="bold" fontFamily="Segoe UI, Roboto, Helvetica, Arial, sans-serif">PDF</text>
  </svg>
);

const GenericFileTypeIcon: React.FC<{ size?: number }> = ({ size = 16 }) => (
  <svg width={size} height={size} viewBox="0 0 16 16" fill="none" style={{ flexShrink: 0, display: "inline-block", verticalAlign: "middle" }}>
    <path d="M2.5 1.5A1.5 1.5 0 0 1 4 0h6a1.5 1.5 0 0 1 1.06.44l3.5 3.5A1.5 1.5 0 0 1 15 5v9.5A1.5 1.5 0 0 1 13.5 16H4A1.5 1.5 0 0 1 2.5 14.5v-13z" fill="#797775" />
    <path d="M10 0v3.5a1.5 1.5 0 0 0 1.5 1.5H15" fill="#484644" opacity="0.6" />
    <rect x="5" y="7" width="6" height="1" rx="0.5" fill="#fff" />
    <rect x="5" y="9.5" width="6" height="1" rx="0.5" fill="#fff" />
    <rect x="5" y="12" width="4" height="1" rx="0.5" fill="#fff" />
  </svg>
);

const getFilenameIcon = (filename: string, size: number = 16): React.ReactNode => {
  const lower = (filename || "").toLowerCase();
  if (lower.endsWith(".doc") || lower.endsWith(".docx") || lower.endsWith(".docm") || lower.endsWith(".dotx")) {
    return <WordFileTypeIcon size={size} />;
  }
  if (lower.endsWith(".xls") || lower.endsWith(".xlsx") || lower.endsWith(".xlsm") || lower.endsWith(".csv")) {
    return <ExcelFileTypeIcon size={size} />;
  }
  if (lower.endsWith(".ppt") || lower.endsWith(".pptx") || lower.endsWith(".pptm")) {
    return <PptFileTypeIcon size={size} />;
  }
  if (lower.endsWith(".pdf")) {
    return <PdfFileTypeIcon size={size} />;
  }
  // Generic file icon if the file type is not among the 4
  return <GenericFileTypeIcon size={size} />;
};

const exportToExcelCsv = (items: doclib_AllProducts[], customFileName: string = "Clients_Products_Export"): void => {
  if (!items || items.length === 0) return;

  const headers = [
    "File Name",
    "Clients",
    "Products",
    "Document Type",
    "Sub Document Type",
    "Business Line",
    "Country Sold To",
    "Issued By",
    "Document Status",
    "Document Language",
    "Confidentiality",
    "Document Date",
    "Expiry Date",
    "Next Review Date",
    "Created",
    "Created by",
    "Modified",
    "Modified by",
    "Reviewed By",
    "Alerts",
    "File URL",
    "ID",
  ];

  const escapeCsv = (val: unknown): string => {
    if (val === null || val === undefined) return '""';
    let str = "";
    if (val instanceof Date) {
      const d = dayjs(val);
      str = d.isValid() ? d.format("DD/MM/YYYY") : "";
    } else {
      str = String(val);
    }
    return `"${str.replace(/"/g, '""')}"`;
  };

  const csvRows: string[] = [];
  csvRows.push(headers.map((h) => `"${h}"`).join(","));

  items.forEach((item) => {
    const productsStr = (item.PIMProduct || [])
      .map((p) => `${p.Title || ""} ${p.PIMProductName || ""}`.trim())
      .filter(Boolean)
      .join("; ");

    const row = [
      escapeCsv(item.filename),
      escapeCsv(item.ManufacturerSearchText),
      escapeCsv(productsStr),
      escapeCsv(item.DocumentTypeSearchText),
      escapeCsv(item.SubDocumentTypeSearchText),
      escapeCsv(item.BusinessLine),
      escapeCsv(item.CountrySoldTo),
      escapeCsv(item.IssuedBy),
      escapeCsv(item.DocumentStatus),
      escapeCsv(item.DocumentLanguage),
      escapeCsv(item.Confidentiality),
      escapeCsv(item.DocumentDate),
      escapeCsv(item.ExpiryDate),
      escapeCsv(item.NextReviewDate),
      escapeCsv(item.Created),
      escapeCsv(item.AuthorTitle),
      escapeCsv(item.Modified),
      escapeCsv(item.EditorTitle),
      escapeCsv(item.ReviewedByTitle),
      escapeCsv(item.Alerts),
      escapeCsv(item.fileUrl),
      escapeCsv(item.id),
    ];
    csvRows.push(row.join(","));
  });

  // UTF-8 BOM (\uFEFF) forces Excel to interpret as UTF-8 so characters and column alignment are preserved
  const blob = new Blob(["\uFEFF" + csvRows.join("\r\n")], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  const timestamp = dayjs().format("YYYYMMDD_HHmmss");
  link.href = url;
  link.setAttribute("download", `${customFileName}_${timestamp}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
};

const FilenameFilterInput: React.FC<{
  column: { getFilterValue: () => unknown; setFilterValue: (value: unknown) => void };
}> = ({ column }) => {
  const raw = column.getFilterValue();
  const filterVal: { text?: string; fileTypes?: string[] } =
    typeof raw === "string"
      ? { text: raw, fileTypes: [] }
      : (raw as { text?: string; fileTypes?: string[] }) || {};

  const currentText = filterVal.text || "";
  const currentFileTypes = filterVal.fileTypes || [];

  return (
    <Box sx={{ width: "100%", display: "flex", flexDirection: "column", gap: 0.5 }}>
      <TextField
        size="small"
        placeholder="Filter file name..."
        variant="standard"
        value={currentText}
        onChange={(e) => {
          const newText = e.target.value;
          if (!newText && currentFileTypes.length === 0) {
            column.setFilterValue(undefined);
          } else {
            column.setFilterValue({ ...filterVal, text: newText });
          }
        }}
        InputProps={{
          endAdornment: currentText ? (
            <IconButton
              size="small"
              onClick={() => {
                if (currentFileTypes.length === 0) {
                  column.setFilterValue(undefined);
                } else {
                  column.setFilterValue({ ...filterVal, text: "" });
                }
              }}
              sx={{ p: "2px" }}
            >
              <CloseIcon sx={{ fontSize: 13 }} />
            </IconButton>
          ) : null,
        }}
        sx={{
          minWidth: "120px",
          "& .MuiInputBase-root": {
            fontSize: "12px",
            padding: "2px 4px",
          },
          "& .MuiInputBase-input": {
            fontSize: "12px",
            padding: "2px 4px",
          },
        }}
      />
      {currentFileTypes.length > 0 && (
        <Box sx={{ display: "flex", flexWrap: "wrap", gap: 0.5, mt: 0.25 }}>
          {currentFileTypes.map((ft) => (
            <Chip
              key={ft}
              size="small"
              label={ft.toUpperCase()}
              onDelete={() => {
                const next = currentFileTypes.filter((t) => t !== ft);
                if (next.length === 0 && !currentText) {
                  column.setFilterValue(undefined);
                } else {
                  column.setFilterValue({ ...filterVal, fileTypes: next });
                }
              }}
              sx={{
                height: "18px",
                fontSize: "10px",
                fontWeight: 600,
                bgcolor:
                  ft === "word"
                    ? "#e8f0fe"
                    : ft === "excel"
                    ? "#e6f4ea"
                    : ft === "powerpoint"
                    ? "#fce8e6"
                    : "#fde8e8",
                color:
                  ft === "word"
                    ? "#185abd"
                    : ft === "excel"
                    ? "#107c41"
                    : ft === "powerpoint"
                    ? "#c43e1c"
                    : "#b30b00",
                "& .MuiChip-deleteIcon": {
                  fontSize: "12px",
                },
              }}
            />
          ))}
        </Box>
      )}
    </Box>
  );
};

const MultiSelectAutocompleteFilter: React.FC<{
  column: { getFilterValue: () => unknown; setFilterValue: (value: unknown) => void };
  options: string[];
  placeholder?: string;
}> = ({ column, options, placeholder = "Filter..." }) => {
  const filterValue = column.getFilterValue() as string[] | string | undefined;
  const selectedValues = Array.isArray(filterValue)
    ? filterValue
    : filterValue
    ? [filterValue]
    : [];

  return (
    <Autocomplete
      multiple
      freeSolo
      size="small"
      options={options}
      value={selectedValues}
      onChange={(_e, newValue) => {
        const values = Array.isArray(newValue)
          ? newValue.map((v) => (typeof v === "string" ? v : "")).filter(Boolean)
          : [];
        column.setFilterValue(values.length > 0 ? values : undefined);
      }}
      renderOption={(props, option) => (
        <li
          {...props}
          style={
            option === "(Empty)"
              ? { fontStyle: "italic", color: "#666" }
              : undefined
          }
        >
          {option}
        </li>
      )}
      renderInput={(params) => (
        <TextField
          {...params}
          size="small"
          placeholder={selectedValues.length === 0 ? placeholder : ""}
          variant="standard"
          sx={{
            minWidth: "120px",
            "& .MuiInputBase-root": {
              fontSize: "12px",
              padding: "2px 4px",
            },
            "& .MuiInputBase-input": {
              fontSize: "12px",
              padding: "2px 4px",
            },
          }}
        />
      )}
      sx={{
        width: "100%",
        minWidth: "130px",
        "& .MuiChip-root": {
          height: "20px",
          fontSize: "11px",
          margin: "1px",
        },
      }}
    />
  );
};

interface ICollapsibleItemListProps {
  items: string[];
  filterValues?: unknown;
  renderItem?: (item: string, idx: number) => React.ReactNode;
}

const CollapsibleItemList: React.FC<ICollapsibleItemListProps> = ({
  items,
  filterValues,
  renderItem,
}) => {
  const [expanded, setExpanded] = React.useState(false);

  if (!items || items.length === 0) return <span>-</span>;
  if (items.length === 1) {
    return <>{renderItem ? renderItem(items[0], 0) : <span>{items[0]}</span>}</>;
  }

  const rawFilters = Array.isArray(filterValues)
    ? filterValues
    : filterValues !== undefined && filterValues !== null && filterValues !== ""
    ? [filterValues]
    : [];

  const activeFilters = rawFilters
    .map((f) => String(f).trim().toLowerCase())
    .filter((f) => f && f !== "(empty)");

  let visibleItems: string[];
  let hiddenItems: string[];

  if (activeFilters.length > 0) {
    const matching = items.filter((item) => {
      const lower = item.toLowerCase();
      return activeFilters.some((f) => lower.includes(f) || f.includes(lower));
    });
    const nonMatching = items.filter((item) => !matching.includes(item));

    visibleItems = matching.length > 0 ? matching : [items[0]];
    hiddenItems = nonMatching;
  } else {
    visibleItems = [items[0]];
    hiddenItems = items.slice(1);
  }

  const itemsToDisplay = expanded ? items : visibleItems;
  const hasHidden = hiddenItems.length > 0;

  return (
    <div style={{ display: "inline-flex", flexWrap: "wrap", alignItems: "center", gap: "4px" }}>
      {itemsToDisplay.map((item, idx) =>
        renderItem ? (
          <React.Fragment key={idx}>{renderItem(item, idx)}</React.Fragment>
        ) : (
          <span key={idx}>
            {item}
            {idx < itemsToDisplay.length - 1 ? "," : ""}
          </span>
        )
      )}

      {hasHidden && (
        <button
          type="button"
          onClick={(e) => {
            e.stopPropagation();
            setExpanded(!expanded);
          }}
          title={
            expanded
              ? "Show less"
              : `Show ${hiddenItems.length} more:\n${hiddenItems.join(", ")}`
          }
          style={{
            cursor: "pointer",
            border: "1px solid #c0c0c0",
            backgroundColor: expanded ? "#e3f2fd" : "#f5f5f5",
            color: "#1976d2",
            borderRadius: "10px",
            padding: "0 6px",
            fontSize: "10px",
            fontWeight: 700,
            lineHeight: "16px",
            height: "16px",
            display: "inline-flex",
            alignItems: "center",
          }}
        >
          {expanded ? "▴ less" : `... +${hiddenItems.length}`}
        </button>
      )}
    </div>
  );
};

const ProductListbox = React.forwardRef<
  HTMLUListElement,
  React.HTMLAttributes<HTMLElement>
>(function ProductListbox(props, ref) {
  const { children, ...other } = props;

  return (
    <ul
      {...other}
      ref={ref}
      style={{
        padding: 0,
        margin: 0,
        listStyle: "none",
        overflowX: "auto",
      }}
    >
      <Box
        sx={{
          display: "grid",
          gridTemplateColumns: "2.2fr 1.3fr 1.2fr",
          gap: 0.5,
          alignItems: "center",
          minWidth: "500px",
          width: "100%",
          px: 1,
          py: 0.5,
          borderBottom: "1px solid",
          borderColor: "divider",
          backgroundColor: "rgba(0, 0, 0, 0.02)",
          color: "text.primary",
          fontSize: "10.5px",
          fontWeight: 600,
        }}
      >
        <span>Product</span>
        <span>Client</span>
        <span>Business Line</span>
      </Box>
      {children}
    </ul>
  );
});

interface ITableViewPreset {
  id: string;
  name: string;
  isBuiltIn?: boolean;
  grouping?: MRT_GroupingState;
  sorting?: MRT_SortingState;
  columnVisibility?: MRT_VisibilityState;
  columnOrder?: string[];
  columnFilters?: MRT_ColumnFiltersState;
  columnPinning?: MRT_ColumnPinningState;
}

const STORAGE_KEY_CUSTOM_VIEWS = "advanceSearch_user_views";

const BUILT_IN_VIEWS: ITableViewPreset[] = [
  {
    id: "default",
    name: "All Documents",
    isBuiltIn: true,
    grouping: [],
    sorting: [{ id: "DocumentDate", desc: true }],
    columnVisibility: {},
    columnFilters: [],
  },
  {
    id: "byClient",
    name: "Grouped by Client",
    isBuiltIn: true,
    grouping: ["ManufacturerSearchText"],
    sorting: [{ id: "ManufacturerSearchText", desc: false }],
    columnVisibility: {},
    columnFilters: [],
  },
  {
    id: "byDocType",
    name: "Grouped by Document Type",
    isBuiltIn: true,
    grouping: ["DocumentTypeSearchText"],
    sorting: [{ id: "DocumentTypeSearchText", desc: false }],
    columnVisibility: {},
    columnFilters: [],
  },
  {
    id: "summary",
    name: "Summary (Compact)",
    isBuiltIn: true,
    grouping: [],
    sorting: [{ id: "DocumentDate", desc: true }],
    columnVisibility: {
      Alerts: false,
      CountrySoldTo: false,
      Confidentiality: false,
    },
    columnFilters: [],
  },
];

export interface IDateRange {
  from: Dayjs | null;
  to: Dayjs | null;
}

export interface IDateFiltersState {
  DocumentDate: IDateRange;
  ExpiryDate: IDateRange;
  NextReviewDate: IDateRange;
  Created: IDateRange;
  Modified: IDateRange;
}

const INITIAL_DATE_FILTERS: IDateFiltersState = {
  DocumentDate: { from: null, to: null },
  ExpiryDate: { from: null, to: null },
  NextReviewDate: { from: null, to: null },
  Created: { from: null, to: null },
  Modified: { from: null, to: null },
};

const AdvanceSearch: React.FC<IAdvanceSearchProps> = (props) => {
  const [items_AllProducts, setItems_AllProducts] =
    React.useState<doclib_AllProducts[]>([]);

  const [products, setProducts] = React.useState<IProductLookupItem[]>([]);
  const [clients, setClients] = React.useState<IClientLookupItem[]>([]);
  const [selectedProducts, setSelectedProducts] = React.useState<
    IProductLookupItem[]
  >([]);
  const [selectedClients, setSelectedClients] = React.useState<
    IClientLookupItem[]
  >([]);
  const [documentTypes, setDocumentTypes] = React.useState<IDocumentTypeItem[]>([]);
  const [allSubDocumentTypes, setAllSubDocumentTypes] = React.useState<
    ISubDocumentTypeItem[]
  >([]);
  const [selectedDocumentTypes, setSelectedDocumentTypes] = React.useState<string[]>([]);
  const [selectedSubDocumentTypes, setSelectedSubDocumentTypes] = React.useState<string[]>([]);

  const [productSearchText, setProductSearchText] = React.useState("");
  const [clientSearchText, setClientSearchText] = React.useState("");
  const [additionalKeywords, setAdditionalKeywords] = React.useState<string[]>([]);
  const [keywordInput, setKeywordInput] = React.useState<string>("");
  const [dateFilters, setDateFilters] = React.useState<IDateFiltersState>(INITIAL_DATE_FILTERS);

  const handleDateChange = React.useCallback(
    (field: keyof IDateFiltersState, bound: "from" | "to", value: Dayjs | null) => {
      setDateFilters((prev) => ({
        ...prev,
        [field]: {
          ...prev[field],
          [bound]: value,
        },
      }));
    },
    []
  );

  const [showMoreFilters, setShowMoreFilters] = React.useState<boolean>(false);
  const [selectedBusinessLines, setSelectedBusinessLines] = React.useState<string[]>([]);
  const [selectedCountries, setSelectedCountries] = React.useState<string[]>([]);
  const [selectedConfidentialities, setSelectedConfidentialities] = React.useState<string[]>([]);
  const [libraryChoices, setLibraryChoices] = React.useState<ILibraryColumnChoices>({
    businessLine: [],
    country: [],
    confidentiality: [],
  });

  const [lookupLoading, setLookupLoading] = React.useState(false);
  const [resultsLoading, setResultsLoading] = React.useState(true);
  const [isSearchDialogOpen, setIsSearchDialogOpen] = React.useState(false);
  const [taxonomyLoading, setTaxonomyLoading] = React.useState(false);
  const [isShareDialogOpen, setIsShareDialogOpen] = React.useState(false);
  const [selectedRowsData, setSelectedRowsData] = React.useState<any[]>([]);
  const [sharingConfig, setSharingConfig] = React.useState<ISharingConfig>({
    subject: "",
    message: "",
  });

  const currentUserEmail =
    props.context?.pageContext?.user?.email ||
    props.context?.pageContext?.user?.loginName ||
    "";
  const currentUserName = props.context?.pageContext?.user?.displayName || "";

  const [fieldFormatters, setFieldFormatters] = React.useState<IFieldFormatters>({
    businessLine: {},
    confidentiality: {},
    alerts: null,
    documentStatus: {},
  });

  const getRowAlerts = React.useCallback(
    (row: doclib_AllProducts): string[] => {
      const results: string[] = [];
      const rule = fieldFormatters.alerts;

      if (rule) {
        // Strict SharePoint Column Formatter evaluation:
        // txtContent: =if((Number(@now) - Number([$Modified]))/60000 < 1 && [$Editor.email] == @me, 'Update in progress', '')
        if (row.Modified) {
          const modTime = new Date(row.Modified).getTime();
          const diffMinutes = (Date.now() - modTime) / 60000;

          let isMe = true;
          if (rule.requiresEditorMe) {
            const editorEmail = (row.EditorEmail || "").toLowerCase();
            const editorTitle = (row.EditorTitle || "").toLowerCase();
            const curEmail = currentUserEmail.toLowerCase();
            const curName = currentUserName.toLowerCase();
            isMe =
              (Boolean(curEmail) && Boolean(editorEmail) && editorEmail === curEmail) ||
              (Boolean(curName) && Boolean(editorTitle) && editorTitle === curName);
          }

          if (diffMinutes >= -0.5 && diffMinutes < rule.durationMinutes && isMe) {
            results.push(rule.text);
          }
        }
        // When column format rule exists, it outputs '' when condition is false, hiding underlying text
        return results;
      }

      if (row.Alerts && row.Alerts.trim()) {
        row.Alerts.split(/[\r\n;,]+/).forEach((a) => {
          const trimmed = a.trim();
          if (trimmed && !results.includes(trimmed)) {
            results.push(trimmed);
          }
        });
      }

      return results;
    },
    [fieldFormatters.alerts, currentUserEmail, currentUserName]
  );

  const [columnFilters, setColumnFilters] = React.useState<MRT_ColumnFiltersState>([]);
  const activeColumnFilters = React.useMemo(() => {
    return columnFilters.filter((f) => isFilterActive(f.value));
  }, [columnFilters]);
  const [applyColumnFilters, setApplyColumnFilters] = React.useState<boolean>(false);
  const savedDialogValuesRef = React.useRef<{
    selectedProducts: IProductLookupItem[];
    productSearchText: string;
    selectedClients: IClientLookupItem[];
    clientSearchText: string;
    selectedDocumentTypes: string[];
    selectedSubDocumentTypes: string[];
    dateFilters: IDateFiltersState;
    additionalKeywords: string[];
    keywordInput: string;
    selectedBusinessLines: string[];
    selectedCountries: string[];
    selectedConfidentialities: string[];
  } | null>(null);
  const [grouping, setGrouping] = React.useState<MRT_GroupingState>([]);

  // --- View Management State ---
  const [selectedViewId, setSelectedViewId] = React.useState<string>("default");
  const [viewMenuAnchorEl, setViewMenuAnchorEl] = React.useState<null | HTMLElement>(null);
  const [customViews, setCustomViews] = React.useState<ITableViewPreset[]>(() => {
    try {
      const stored = localStorage.getItem(STORAGE_KEY_CUSTOM_VIEWS);
      return stored ? JSON.parse(stored) : [];
    } catch {
      return [];
    }
  });

  const [sorting, setSorting] = React.useState<MRT_SortingState>([
    { id: "DocumentDate", desc: true },
  ]);
  const [columnVisibility, setColumnVisibility] = React.useState<MRT_VisibilityState>({});
  const [columnOrder, setColumnOrder] = React.useState<string[]>([]);
  const [columnPinning, setColumnPinning] = React.useState<MRT_ColumnPinningState>({
    left: [],
    right: [],
  });

  // Dialog state for "Save Current View"
  const [isSaveViewDialogOpen, setIsSaveViewDialogOpen] = React.useState(false);
  const [newViewName, setNewViewName] = React.useState("");

  const allAvailableViews = React.useMemo(() => {
    return [...BUILT_IN_VIEWS, ...customViews];
  }, [customViews]);

  const currentView = React.useMemo(() => {
    return allAvailableViews.find((v) => v.id === selectedViewId) || BUILT_IN_VIEWS[0];
  }, [allAvailableViews, selectedViewId]);

  const handleApplyView = (viewId: string): void => {
    setSelectedViewId(viewId);
    const view = allAvailableViews.find((v) => v.id === viewId);
    if (!view) return;

    setGrouping(view.grouping ?? []);
    setSorting(view.sorting ?? []);
    setColumnVisibility(view.columnVisibility ?? {});
    if (view.columnOrder && view.columnOrder.length > 0) {
      setColumnOrder(view.columnOrder);
    }
    if (view.columnFilters !== undefined) {
      setColumnFilters(view.columnFilters.filter((f) => isFilterActive(f.value)));
    }
    if (view.columnPinning !== undefined) {
      setColumnPinning(view.columnPinning);
    } else {
      setColumnPinning({ left: [], right: [] });
    }
  };

  const handleSaveCurrentView = (): void => {
    const trimmed = newViewName.trim();
    if (!trimmed) return;

    const newView: ITableViewPreset = {
      id: `custom_${Date.now()}`,
      name: trimmed,
      isBuiltIn: false,
      grouping,
      sorting,
      columnVisibility,
      columnOrder,
      columnFilters: activeColumnFilters,
      columnPinning,
    };

    const updated = [...customViews.filter((v) => v.name.toLowerCase() !== trimmed.toLowerCase()), newView];
    setCustomViews(updated);
    try {
      localStorage.setItem(STORAGE_KEY_CUSTOM_VIEWS, JSON.stringify(updated));
    } catch (e) {
      console.warn("Could not save view to localStorage", e);
    }

    setSelectedViewId(newView.id);
    setIsSaveViewDialogOpen(false);
    setNewViewName("");
  };

  const handleDeleteView = (viewId: string, e: React.MouseEvent): void => {
    e.stopPropagation();
    const updated = customViews.filter((v) => v.id !== viewId);
    setCustomViews(updated);
    try {
      localStorage.setItem(STORAGE_KEY_CUSTOM_VIEWS, JSON.stringify(updated));
    } catch (e) {
      console.warn("Could not update localStorage", e);
    }
    if (selectedViewId === viewId) {
      handleApplyView("default");
    }
  };

  // Batch / pagination state for loadAllRecords and searchRecords
  const [nextSkipId, setNextSkipId] = React.useState<number | undefined>(undefined);
  const [nextSearchStartRow, setNextSearchStartRow] = React.useState<number | undefined>(undefined);
  const [currentSearchQuery, setCurrentSearchQuery] = React.useState<string>("");
  const [hasMoreRecords, setHasMoreRecords] = React.useState<boolean>(false);
  const [isLoadingMore, setIsLoadingMore] = React.useState<boolean>(false);
  const [isBrowseMode, setIsBrowseMode] = React.useState<boolean>(true);

  // State for File Context Menu & Edit Modal Dialog
  const [fileMenuAnchorEl, setFileMenuAnchorEl] = React.useState<null | HTMLElement>(null);
  const [selectedFileForAction, setSelectedFileForAction] = React.useState<doclib_AllProducts | null>(null);
  const [editModalUrl, setEditModalUrl] = React.useState<string | null>(null);

  // Dynamic table container height to fill available vertical space cleanly
  const paperRef = React.useRef<HTMLDivElement | null>(null);
  const tableContainerRef = React.useRef<HTMLDivElement | null>(null);
  const [tableHeight, setTableHeight] = React.useState<string>("calc(100vh - 160px)");

  React.useEffect(() => {
    const updateHeight = (): void => {
      const targetEl = paperRef.current || tableContainerRef.current;
      if (targetEl) {
        const rect = targetEl.getBoundingClientRect();
        // Reserve 8px for bottom padding/margin so table reaches bottom without window scrollbar
        const availableHeight = window.innerHeight - rect.top - 8;
        setTableHeight(`${Math.max(300, Math.floor(availableHeight))}px`);
      }
    };

    updateHeight();

    const scheduleUpdate = (): void => {
      updateHeight();
      setTimeout(updateHeight, 60);
      setTimeout(updateHeight, 200);
      setTimeout(updateHeight, 400);
    };

    window.addEventListener("resize", updateHeight);
    window.addEventListener("scroll", updateHeight);
    document.addEventListener("click", scheduleUpdate);
    document.addEventListener("transitionend", updateHeight);

    let ro: ResizeObserver | null = null;
    if (typeof ResizeObserver !== "undefined") {
      ro = new ResizeObserver(() => {
        updateHeight();
      });
      if (document.body) ro.observe(document.body);
      if (document.documentElement) ro.observe(document.documentElement);
      const spCanvas = document.getElementById("spPageCanvasContent");
      if (spCanvas) ro.observe(spCanvas);
    }

    let mo: MutationObserver | null = null;
    if (typeof MutationObserver !== "undefined") {
      mo = new MutationObserver(() => {
        updateHeight();
      });
      mo.observe(document.body, { attributes: true, subtree: false });
    }

    const timer = setTimeout(updateHeight, 300);

    return () => {
      window.removeEventListener("resize", updateHeight);
      window.removeEventListener("scroll", updateHeight);
      document.removeEventListener("click", scheduleUpdate);
      document.removeEventListener("transitionend", updateHeight);
      if (ro) ro.disconnect();
      if (mo) mo.disconnect();
      clearTimeout(timer);
    };
  }, []);

  const activeSp: SPFI = React.useMemo(() => {
    if (sp) return sp;
    const initialWebUrl = props.urlSite?.trim() || props.context?.pageContext?.web?.absoluteUrl;
    if (initialWebUrl && props.context) {
      return spfi(initialWebUrl).using(SPFx(props.context));
    }
    return sp;
  }, [props.urlSite, props.context]);

  React.useEffect(() => {
    const timer = setTimeout(async () => {
      if (productSearchText.trim().length >= 3) {
        setLookupLoading(true);
        try {
          const results = await searchProductsService(activeSp, productSearchText.trim());
          setProducts(results);
        } catch (error) {
          console.error("searchProducts error:", error);
        } finally {
          setLookupLoading(false);
        }
      } else {
        setProducts([]);
      }
    }, 300);

    return () => clearTimeout(timer);
  }, [productSearchText, activeSp]);

  React.useEffect(() => {
    const timer = setTimeout(async () => {
      if (clientSearchText.trim().length >= 3) {
        setLookupLoading(true);
        try {
          const results = await searchClientsService(activeSp, clientSearchText.trim());
          setClients(results);
        } catch (error) {
          console.error("searchClients error:", error);
        } finally {
          setLookupLoading(false);
        }
      } else {
        setClients([]);
      }
    }, 300);

    return () => clearTimeout(timer);
  }, [clientSearchText, activeSp]);

  React.useEffect(() => {
    const loadTaxonomy = async (): Promise<void> => {
      setTaxonomyLoading(true);
      try {
        const taxonomy = await loadTaxonomyService(activeSp);
        setDocumentTypes(taxonomy.docTypes);
        setAllSubDocumentTypes(taxonomy.subDocTypes);
      } catch (error) {
        console.error("loadTaxonomy error:", error);
        setDocumentTypes([]);
        setAllSubDocumentTypes([]);
      } finally {
        setTaxonomyLoading(false);
      }
    };

    const loadSharingConfig = async (): Promise<void> => {
      try {
        const config = await loadSharingConfigService(activeSp);
        setSharingConfig(config);
      } catch (err) {
        console.warn("loadSharingConfiguration error:", err);
      }
    };

    const loadFieldMetadata = async (): Promise<void> => {
      try {
        const { formatters, choices } = await loadListFieldMetadataService(activeSp);
        setFieldFormatters(formatters);
        setLibraryChoices(choices);
      } catch (err) {
        console.warn("loadListFieldMetadata error:", err);
      }
    };

    const loadInitialRecords = async (): Promise<void> => {
      try {
        setResultsLoading(true);
        const result = await loadRecordsBatchService(activeSp, undefined, 1000);
        setItems_AllProducts(result.items);
        setNextSkipId(result.nextSkipId);
        setHasMoreRecords(result.hasMore);
      } catch (error) {
        console.error("loadInitialRecords error:", error);
      } finally {
        setResultsLoading(false);
      }
    };

    void loadTaxonomy();
    void loadSharingConfig();
    void loadFieldMetadata();
    void loadInitialRecords();
  }, [activeSp]);

  // In-memory filtered sub-document types based on selected Document Types
  const availableSubDocumentTypes = useMemo(() => {
    let list: ISubDocumentTypeItem[];
    if (selectedDocumentTypes.length === 0) {
      list = allSubDocumentTypes;
    } else {
      const selectedSet = new Set(selectedDocumentTypes);
      list = allSubDocumentTypes.filter(
        (sub) => sub.DocumentType?.Title && selectedSet.has(sub.DocumentType.Title)
      );
    }
    return list.slice().sort((a, b) =>
      (a.Title || "").localeCompare(b.Title || "", undefined, { sensitivity: "base" })
    );
  }, [allSubDocumentTypes, selectedDocumentTypes]);

  // Prune any selected sub-document types if their parent document type was unselected
  React.useEffect(() => {
    if (selectedDocumentTypes.length > 0) {
      const validTitles = new Set(availableSubDocumentTypes.map((i) => i.Title));
      setSelectedSubDocumentTypes((prev) =>
        prev.filter((title) => validTitles.has(title))
      );
    }
  }, [availableSubDocumentTypes, selectedDocumentTypes]);

  const buildSearchQuery = (): string => {
    const clauses: string[] = [];
    const searchPath = `${props.urlSite.replace(/\/$/, "")}/Products/*`;
    clauses.push(`Path:"${searchPath}"`);
    clauses.push("IsDocument:1");

    const productClauses: string[] = [];
    selectedProducts.forEach((item) => {
      const title = (item.Title || "").trim();
      const productName = (item.PIMProductName || "").trim();
      if (title && productName) {
        productClauses.push(
          `("${sanitizeKqlValue(title)}" OR "${sanitizeKqlValue(productName)}")`
        );
      } else if (title) {
        productClauses.push(`"${sanitizeKqlValue(title)}"`);
      } else if (productName) {
        productClauses.push(`"${sanitizeKqlValue(productName)}"`);
      }
    });

    const clientValues = selectedClients
      .map((item) => (item.Title || "").trim())
      .filter(Boolean);

    const documentTypeValues = selectedDocumentTypes
      .map((item) => item.trim())
      .filter(Boolean);

    const subDocumentTypeValues = selectedSubDocumentTypes
      .map((item) => item.trim())
      .filter(Boolean);

    if (productClauses.length > 0) {
      clauses.push(`(${productClauses.join(" OR ")})`);
    }

    if (clientValues.length > 0) {
      clauses.push(
        `(${clientValues
          .map((value) => `ManufacturerOWSTEXT:"${sanitizeKqlValue(value)}"`)
          .join(" OR ")})`
      );
    }

    if (documentTypeValues.length > 0) {
      clauses.push(
        `(${documentTypeValues
          .map((value) => `"${sanitizeKqlValue(value)}"`)
          .join(" OR ")})`
      );
    }

    if (subDocumentTypeValues.length > 0) {
      clauses.push(
        `(${subDocumentTypeValues
          .map((value) => `"${sanitizeKqlValue(value)}"`)
          .join(" OR ")})`
      );
    }

    const businessLineValues = selectedBusinessLines
      .map((item) => item.trim())
      .filter(Boolean);

    const countryValues = selectedCountries
      .map((item) => item.trim())
      .filter(Boolean);

    const confidentialityValues = selectedConfidentialities
      .map((item) => item.trim())
      .filter(Boolean);

    if (businessLineValues.length > 0) {
      clauses.push(
        `(${businessLineValues
          .map((value) => `BusinessLineOWSCHCM:"${sanitizeKqlValue(value)}"`)
          .join(" OR ")})`
      );
    }

    if (countryValues.length > 0) {
      clauses.push(
        `(${countryValues
          .map((value) => `CountryOWSCHCM:"${sanitizeKqlValue(value)}"`)
          .join(" OR ")})`
      );
    }

    if (confidentialityValues.length > 0) {
      clauses.push(
        `(${confidentialityValues
          .map((value) => `ConfidentialityOWSCHCS:"${sanitizeKqlValue(value)}"`)
          .join(" OR ")})`
      );
    }

    const keywordClause = compileKeywordsToKql(additionalKeywords, keywordInput);
    if (keywordClause) {
      clauses.push(keywordClause);
    }

    const datePropertyMap: Record<keyof IDateFiltersState, string> = {
      DocumentDate: "DocumentDateOWSTDATE",
      ExpiryDate: "ExpiryDateOWSTDATE",
      NextReviewDate: "NextReviewDateOWSTDATE",
      Created: "Created",
      Modified: "LastModifiedTime",
    };

    (Object.keys(dateFilters) as (keyof IDateFiltersState)[]).forEach((fieldKey) => {
      const { from, to } = dateFilters[fieldKey];
      const targetProp = datePropertyMap[fieldKey];

      if (from && to) {
        clauses.push(
          `${targetProp}:${from.startOf("day").toISOString()}..${to.endOf("day").toISOString()}`
        );
      } else if (from) {
        clauses.push(
          `${targetProp}>=${from.startOf("day").toISOString()}`
        );
      } else if (to) {
        clauses.push(
          `${targetProp}<=${to.endOf("day").toISOString()}`
        );
      }
    });

    return clauses.join(" AND ");
  };

  const handleLoadAllRecords = async (): Promise<void> => {
    try {
      setResultsLoading(true);
      setItems_AllProducts([]);
      setIsBrowseMode(true);
      setColumnFilters([]);
      setNextSearchStartRow(undefined);
      setCurrentSearchQuery("");

      const result = await loadRecordsBatchService(activeSp, undefined, 1000);
      setItems_AllProducts(result.items);
      setNextSkipId(result.nextSkipId);
      setHasMoreRecords(result.hasMore);
    } catch (error) {
      console.error("handleLoadAllRecords error:", error);
    } finally {
      setResultsLoading(false);
    }
  };

  const handleRefresh = async (): Promise<void> => {
    try {
      setResultsLoading(true);
      setItems_AllProducts([]);
      setIsBrowseMode(true);
      setColumnFilters([]);
      setNextSearchStartRow(undefined);
      setCurrentSearchQuery("");

      const [recordsResult] = await Promise.all([
        loadRecordsBatchService(activeSp, undefined, 1000),
        loadListFieldMetadataService(activeSp)
          .then(({ formatters, choices }) => {
            setFieldFormatters(formatters);
            setLibraryChoices(choices);
          })
          .catch((err) => console.warn("handleRefresh metadata error:", err)),
      ]);

      setItems_AllProducts(recordsResult.items);
      setNextSkipId(recordsResult.nextSkipId);
      setHasMoreRecords(recordsResult.hasMore);
    } catch (error) {
      console.error("handleRefresh error:", error);
    } finally {
      setResultsLoading(false);
    }
  };

  const handleLoadNextBatch = async (): Promise<void> => {
    if (!hasMoreRecords || isLoadingMore) return;

    try {
      setIsLoadingMore(true);
      if (isBrowseMode) {
        const result = await loadRecordsBatchService(activeSp, nextSkipId, 1000);
        setItems_AllProducts((prev) => {
          const existingIds = new Set(prev.map((p) => p.id));
          const newUnique = result.items.filter((item) => !existingIds.has(item.id));
          return [...prev, ...newUnique];
        });
        setNextSkipId(result.nextSkipId);
        setHasMoreRecords(result.hasMore);
      } else {
        const result = await searchRecordsService(
          activeSp,
          currentSearchQuery,
          nextSearchStartRow ?? 0,
          1000
        );
        setItems_AllProducts((prev) => {
          const existingIds = new Set(prev.map((p) => p.id));
          const newUnique = result.items.filter((item) => !existingIds.has(item.id));
          return [...prev, ...newUnique];
        });
        setNextSearchStartRow(result.nextStartRow);
        setHasMoreRecords(result.hasMore);
      }
    } catch (error) {
      console.error("handleLoadNextBatch error:", error);
    } finally {
      setIsLoadingMore(false);
    }
  };

  const applyColumnFiltersToDialog = (): void => {
    // 1. Clients
    const clientColFilter = activeColumnFilters.find((f) => f.id === "ManufacturerSearchText")?.value;
    if (clientColFilter) {
      const vals = (Array.isArray(clientColFilter) ? clientColFilter : [clientColFilter])
        .map((v) => String(v).trim())
        .filter((v) => v && v.toLowerCase() !== "(empty)" && v !== "-");

      if (vals.length > 0) {
        const seenTitles = new Set<string>();
        const newItems: IClientLookupItem[] = [];
        vals.forEach((v) => {
          const vLower = v.toLowerCase();
          if (!seenTitles.has(vLower)) {
            seenTitles.add(vLower);
            const fromLoaded = clients.find((c) => (c.Title || "").toLowerCase() === vLower);
            newItems.push(fromLoaded || { ID: -Math.floor(Math.random() * 100000), Title: v });
          }
        });
        setSelectedClients(newItems);
      } else {
        setSelectedClients([]);
      }
    } else {
      setSelectedClients([]);
    }
    setClientSearchText("");

    // 2. Products
    const productColFilter = activeColumnFilters.find((f) => f.id === "PIMProductSearchText")?.value;
    if (productColFilter) {
      const vals = (Array.isArray(productColFilter) ? productColFilter : [productColFilter])
        .map((v) => String(v).trim())
        .filter((v) => v && v.toLowerCase() !== "(empty)" && v !== "-");

      if (vals.length > 0) {
        const allKnownProducts: IProductLookupItem[] = [];
        const seenProductKeys = new Set<string>();

        const registerProduct = (item: IProductLookupItem) => {
          const key = `${item.Title || ""} ${item.PIMProductName || ""}`.trim().toLowerCase();
          if (key && !seenProductKeys.has(key)) {
            seenProductKeys.add(key);
            allKnownProducts.push(item);
          }
        };

        items_AllProducts.forEach((rec) => {
          if (Array.isArray(rec.PIMProduct)) {
            rec.PIMProduct.forEach(registerProduct);
          }
        });

        products.forEach(registerProduct);

        const newItems: IProductLookupItem[] = [];
        const seenKeys = new Set<string>();

        vals.forEach((v) => {
          const vLower = v.toLowerCase();

          const matched = allKnownProducts.find((p) => {
            const fullKey = `${p.Title || ""} ${p.PIMProductName || ""}`.trim().toLowerCase();
            return (
              fullKey === vLower ||
              (p.Title && p.Title.toLowerCase() === vLower) ||
              (p.PIMProductName && p.PIMProductName.toLowerCase() === vLower)
            );
          });

          if (matched) {
            const matchedKey = `${matched.Title || ""} ${matched.PIMProductName || ""}`.trim().toLowerCase();
            if (!seenKeys.has(matchedKey)) {
              newItems.push(matched);
              seenKeys.add(matchedKey);
            }
          } else {
            let parsedTitle = v;
            let parsedName = "";

            if (v.includes("|")) {
              const parts = v.split("|").map((p) => p.trim());
              parsedTitle = parts[0] || "";
              parsedName = parts.slice(1).join(" ");
            } else {
              const spaceIdx = v.indexOf(" ");
              if (spaceIdx > 0) {
                parsedTitle = v.substring(0, spaceIdx).trim();
                parsedName = v.substring(spaceIdx + 1).trim();
              }
            }

            const itemKey = `${parsedTitle} ${parsedName}`.trim().toLowerCase();
            if (!seenKeys.has(itemKey)) {
              newItems.push({
                ID: -Math.floor(Math.random() * 100000),
                Title: parsedTitle,
                PIMProductName: parsedName,
              });
              seenKeys.add(itemKey);
            }
          }
        });

        setSelectedProducts(newItems);
      } else {
        setSelectedProducts([]);
      }
    } else {
      setSelectedProducts([]);
    }
    setProductSearchText("");

    // 3. Document Types
    const docTypeColFilter = activeColumnFilters.find((f) => f.id === "DocumentTypeSearchText")?.value;
    if (docTypeColFilter) {
      const vals = (Array.isArray(docTypeColFilter) ? docTypeColFilter : [docTypeColFilter])
        .map((v) => String(v).trim())
        .filter((v) => v && v.toLowerCase() !== "(empty)" && v !== "-");

      setSelectedDocumentTypes(Array.from(new Set(vals)));
    } else {
      setSelectedDocumentTypes([]);
    }

    // 4. Sub Document Types
    const subDocTypeColFilter = activeColumnFilters.find((f) => f.id === "SubDocumentTypeSearchText")?.value;
    if (subDocTypeColFilter) {
      const vals = (Array.isArray(subDocTypeColFilter) ? subDocTypeColFilter : [subDocTypeColFilter])
        .map((v) => String(v).trim())
        .filter((v) => v && v.toLowerCase() !== "(empty)" && v !== "-");

      setSelectedSubDocumentTypes(Array.from(new Set(vals)));
    } else {
      setSelectedSubDocumentTypes([]);
    }

    // 5. Date filters (Document Date, Expiry Date, Next Review Date, Created, Modified)
    const dateColCandidates: { id: string; key: keyof IDateFiltersState }[] = [
      { id: "DocumentDate", key: "DocumentDate" },
      { id: "ExpiryDate", key: "ExpiryDate" },
      { id: "NextReviewDate", key: "NextReviewDate" },
      { id: "Created", key: "Created" },
      { id: "Modified", key: "Modified" },
    ];
    const newDateFilters: IDateFiltersState = {
      DocumentDate: { from: null, to: null },
      ExpiryDate: { from: null, to: null },
      NextReviewDate: { from: null, to: null },
      Created: { from: null, to: null },
      Modified: { from: null, to: null },
    };
    dateColCandidates.forEach((item) => {
      const dateColFilter = activeColumnFilters.find((f) => f.id === item.id)?.value;
      if (dateColFilter) {
        const parsedDate = dayjs(dateColFilter as string | Date);
        if (parsedDate.isValid()) {
          newDateFilters[item.key] = {
            from: parsedDate,
            to: null,
          };
        }
      }
    });
    setDateFilters(newDateFilters);

    let hasMoreFilterActive = false;

    // 6. Business Line
    const blFilter = activeColumnFilters.find((f) => f.id === "BusinessLine")?.value;
    if (blFilter) {
      const vals = (Array.isArray(blFilter) ? blFilter : [blFilter])
        .map((v) => String(v).trim())
        .filter((v) => v && v.toLowerCase() !== "(empty)" && v !== "-");
      const uniqueVals = Array.from(new Set(vals));
      setSelectedBusinessLines(uniqueVals);
      if (uniqueVals.length > 0) hasMoreFilterActive = true;
    } else {
      setSelectedBusinessLines([]);
    }

    // 7. Country Sold To
    const countryFilter = activeColumnFilters.find((f) => f.id === "CountrySoldTo")?.value;
    if (countryFilter) {
      const vals = (Array.isArray(countryFilter) ? countryFilter : [countryFilter])
        .map((v) => String(v).trim())
        .filter((v) => v && v.toLowerCase() !== "(empty)" && v !== "-");
      const uniqueVals = Array.from(new Set(vals));
      setSelectedCountries(uniqueVals);
      if (uniqueVals.length > 0) hasMoreFilterActive = true;
    } else {
      setSelectedCountries([]);
    }

    // 8. Confidentiality
    const confFilter = activeColumnFilters.find((f) => f.id === "Confidentiality")?.value;
    if (confFilter) {
      const vals = (Array.isArray(confFilter) ? confFilter : [confFilter])
        .map((v) => String(v).trim())
        .filter((v) => v && v.toLowerCase() !== "(empty)" && v !== "-");
      const uniqueVals = Array.from(new Set(vals));
      setSelectedConfidentialities(uniqueVals);
      if (uniqueVals.length > 0) hasMoreFilterActive = true;
    } else {
      setSelectedConfidentialities([]);
    }

    if (hasMoreFilterActive) {
      setShowMoreFilters(true);
    }

    // 9. Keyword / File Name
    const filenameFilter = activeColumnFilters.find((f) => f.id === "filename")?.value;
    if (filenameFilter) {
      const textVal =
        typeof filenameFilter === "string"
          ? filenameFilter
          : (filenameFilter as { text?: string })?.text;
      if (textVal && typeof textVal === "string" && textVal.trim()) {
        setKeywordInput(textVal.trim());
      } else {
        setKeywordInput("");
      }
    } else {
      setKeywordInput("");
    }
    setAdditionalKeywords([]);
  };

  const handleOpenSearchDialog = (): void => {
    // Only apply column filters into dialog if the toggle is ON; otherwise retain previous dialog values
    if (applyColumnFilters) {
      applyColumnFiltersToDialog();
    }
    setIsSearchDialogOpen(true);
  };

  const handleToggleApplyColumnFilters = (enabled: boolean): void => {
    setApplyColumnFilters(enabled);
    if (enabled) {
      // Snapshot current dialog control values so they can be restored if user unchecks
      savedDialogValuesRef.current = {
        selectedProducts,
        productSearchText,
        selectedClients,
        clientSearchText,
        selectedDocumentTypes,
        selectedSubDocumentTypes,
        dateFilters,
        additionalKeywords,
        keywordInput,
        selectedBusinessLines,
        selectedCountries,
        selectedConfidentialities,
      };
      applyColumnFiltersToDialog();
    } else {
      // Restore previous dialog values if available
      if (savedDialogValuesRef.current) {
        const prev = savedDialogValuesRef.current;
        setSelectedProducts(prev.selectedProducts);
        setProductSearchText(prev.productSearchText);
        setSelectedClients(prev.selectedClients);
        setClientSearchText(prev.clientSearchText);
        setSelectedDocumentTypes(prev.selectedDocumentTypes);
        setSelectedSubDocumentTypes(prev.selectedSubDocumentTypes);
        setDateFilters(prev.dateFilters);
        setAdditionalKeywords(prev.additionalKeywords);
        setKeywordInput(prev.keywordInput);
        setSelectedBusinessLines(prev.selectedBusinessLines);
        setSelectedCountries(prev.selectedCountries);
        setSelectedConfidentialities(prev.selectedConfidentialities);
        setShowMoreFilters(
          prev.selectedBusinessLines.length > 0 ||
          prev.selectedCountries.length > 0 ||
          prev.selectedConfidentialities.length > 0
        );
      }
    }
  };

  const handleProductInputChange = (
    _event: React.SyntheticEvent,
    newInputValue: string
  ) => {
    setProductSearchText(newInputValue);
  };

  const handleProductSelectionChange = (
    _event: React.SyntheticEvent,
    newValue: IProductLookupItem[]
  ) => {
    setSelectedProducts(newValue);
  };

  const handleClientInputChange = (
    _event: React.SyntheticEvent,
    newInputValue: string
  ) => {
    setClientSearchText(newInputValue);
  };

  const handleClientSelectionChange = (
    _event: React.SyntheticEvent,
    newValue: IClientLookupItem[]
  ) => {
    setSelectedClients(newValue);
  };

  const isSearchFormValid = React.useMemo(() => {
    const hasProduct = selectedProducts.some(
      (item) => Boolean(item.Title?.trim()) || Boolean(item.PIMProductName?.trim())
    );
    const hasClient = selectedClients.some((item) => Boolean(item.Title?.trim()));
    const hasDocType = selectedDocumentTypes.length > 0;
    const hasSubDocType = selectedSubDocumentTypes.length > 0;
    const hasDate = Object.values(dateFilters).some(
      (range) => range.from !== null || range.to !== null
    );
    const hasKeyword = Boolean(compileKeywordsToKql(additionalKeywords, keywordInput));
    const hasBusinessLine = selectedBusinessLines.length > 0;
    const hasCountry = selectedCountries.length > 0;
    const hasConfidentiality = selectedConfidentialities.length > 0;

    return (
      hasProduct ||
      hasClient ||
      hasDocType ||
      hasSubDocType ||
      hasDate ||
      hasKeyword ||
      hasBusinessLine ||
      hasCountry ||
      hasConfidentiality
    );
  }, [
    selectedProducts,
    selectedClients,
    selectedDocumentTypes,
    selectedSubDocumentTypes,
    dateFilters,
    additionalKeywords,
    keywordInput,
    selectedBusinessLines,
    selectedCountries,
    selectedConfidentialities,
  ]);

  const handleClearDialogFilters = (): void => {
    setSelectedProducts([]);
    setProductSearchText("");
    setSelectedClients([]);
    setClientSearchText("");
    setSelectedDocumentTypes([]);
    setSelectedSubDocumentTypes([]);
    setDateFilters(INITIAL_DATE_FILTERS);
    setAdditionalKeywords([]);
    setKeywordInput("");
    setSelectedBusinessLines([]);
    setSelectedCountries([]);
    setSelectedConfidentialities([]);
    setShowMoreFilters(false);
    setApplyColumnFilters(false);
    savedDialogValuesRef.current = null;
  };

  const handleSearch = async () => {
    if (!isSearchFormValid) return;

    try {
      setResultsLoading(true);
      setItems_AllProducts([]);
      setIsSearchDialogOpen(false);
      setIsBrowseMode(false);
      // Retain columnFilters and grouping across searches
      setHasMoreRecords(false);
      setNextSkipId(undefined);
      setNextSearchStartRow(undefined);

      const queryText = buildSearchQuery();
      if (!queryText) {
        await handleLoadAllRecords();
        return;
      }

      setCurrentSearchQuery(queryText);
      const result = await searchRecordsService(activeSp, queryText, 0, 1000);
      setItems_AllProducts(result.items);
      setNextSearchStartRow(result.nextStartRow);
      setHasMoreRecords(result.hasMore);
    } catch (error) {
      console.error("handleSearch error:", error);
    } finally {
      setResultsLoading(false);
    }
  };

  const getItemsFilteredExcluding = React.useCallback(
    (excludedColId: string): doclib_AllProducts[] => {
      if (!columnFilters || columnFilters.length === 0) {
        return items_AllProducts;
      }
      const activeFilters = columnFilters.filter(
        (f) =>
          f.id !== excludedColId &&
          isFilterActive(f.value)
      );
      if (activeFilters.length === 0) {
        return items_AllProducts;
      }
      return items_AllProducts.filter((item) =>
        activeFilters.every((f) => {
          if (f.id === "Alerts") {
            const rowAlerts = getRowAlerts(item);
            const selected = (Array.isArray(f.value) ? f.value : [f.value]).map((v) => String(v).trim());
            if (selected.length === 0) return true;
            if (rowAlerts.length === 0) {
              return selected.some((v) => v.toLowerCase() === "(empty)");
            }
            return selected.some((v) =>
              rowAlerts.some((a) => a.toLowerCase() === v.toLowerCase())
            );
          }
          return itemMatchesFilter(item, f.id, f.value, currentUserEmail, currentUserName);
        })
      );
    },
    [items_AllProducts, columnFilters, getRowAlerts, currentUserEmail, currentUserName]
  );

  const personFilterFn = React.useCallback(
    (
      row: { original: doclib_AllProducts; getValue: (columnId: string) => unknown },
      columnId: string,
      filterValue: unknown
    ): boolean => {
      if (filterValue === undefined || filterValue === null || filterValue === "") return true;
      if (Array.isArray(filterValue) && filterValue.length === 0) return true;

      const selectedValues = (Array.isArray(filterValue) ? filterValue : [filterValue])
        .map((v) => String(v).trim())
        .filter((v) => v !== "");

      if (selectedValues.length === 0) return true;

      const item = row.original;
      let personEmail = "";
      let personTitle = "";

      if (columnId === "AuthorTitle") {
        personEmail = item.AuthorEmail || "";
        personTitle = item.AuthorTitle || "";
      } else if (columnId === "EditorTitle") {
        personEmail = item.EditorEmail || "";
        personTitle = item.EditorTitle || "";
      } else if (columnId === "ReviewedByTitle") {
        personEmail = item.ReviewedByEmail || "";
        personTitle = item.ReviewedByTitle || "";
      } else {
        personTitle = String(row.getValue(columnId) || "");
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
    },
    [currentUserEmail, currentUserName]
  );
  (personFilterFn as { autoRemove?: (val: unknown) => boolean }).autoRemove = (val: unknown) => !isFilterActive(val);

  const alertsFilterFn = React.useCallback(
    (row: { original: doclib_AllProducts }, _columnId: string, filterValue: unknown): boolean => {
      if (filterValue === undefined || filterValue === null || filterValue === "") return true;
      if (Array.isArray(filterValue) && filterValue.length === 0) return true;

      const selectedValues = (Array.isArray(filterValue) ? filterValue : [filterValue])
        .map((v) => String(v).trim())
        .filter((v) => v !== "");

      if (selectedValues.length === 0) return true;

      const rowAlerts = getRowAlerts(row.original);
      if (rowAlerts.length === 0) {
        return selectedValues.some((v) => v.toLowerCase() === "(empty)");
      }

      const regularSelections = selectedValues
        .filter((v) => v.toLowerCase() !== "(empty)")
        .map((v) => v.toLowerCase());

      if (regularSelections.length === 0) return false;

      return regularSelections.some((sel) =>
        rowAlerts.some((a) => {
          const aLower = a.toLowerCase();
          return aLower.includes(sel) || sel.includes(aLower);
        })
      );
    },
    [getRowAlerts]
  );
  (alertsFilterFn as { autoRemove?: (val: unknown) => boolean }).autoRemove = (val: unknown) => !isFilterActive(val);

  const alertsOptions = useMemo(() => {
    const items = getItemsFilteredExcluding("Alerts");
    const unique = new Set<string>();
    let hasEmpty = false;
    items.forEach((item) => {
      const alerts = getRowAlerts(item);
      if (alerts.length === 0) {
        hasEmpty = true;
      } else {
        alerts.forEach((a) => {
          const trimmed = a.trim();
          if (trimmed) unique.add(trimmed);
        });
      }
    });
    const result: string[] = Array.from(unique).sort((a, b) =>
      a.localeCompare(b, undefined, { sensitivity: "base" })
    );
    if (hasEmpty) {
      result.unshift("(Empty)");
    }
    return result;
  }, [getItemsFilteredExcluding, getRowAlerts]);

  const clientOptions = useMemo(
    () => {
      const items = getItemsFilteredExcluding("ManufacturerSearchText");
      const unique = new Set<string>();
      let hasEmpty = false;
      items.forEach((item) => {
        const raw = item.ManufacturerSearchText;
        if (!raw || raw.trim() === "") {
          hasEmpty = true;
        } else {
          raw.split(/[\r\n;,]+/).forEach((val) => {
            const trimmed = val.trim();
            if (trimmed) unique.add(trimmed);
          });
        }
      });
      const result: string[] = Array.from(unique).sort((a, b) =>
        a.localeCompare(b, undefined, { sensitivity: "base" })
      );
      if (hasEmpty) {
        result.unshift("(Empty)");
      }
      return result;
    },
    [getItemsFilteredExcluding]
  );

  const productOptions = useMemo(
    () => {
      const items = getItemsFilteredExcluding("PIMProductSearchText");
      const unique = new Set<string>();
      let hasEmpty = false;
      items.forEach((item) => {
        if (!item.PIMProduct || !Array.isArray(item.PIMProduct) || item.PIMProduct.length === 0) {
          hasEmpty = true;
        } else {
          let hasValid = false;
          item.PIMProduct.forEach((p) => {
            const name = `${p.Title || ""} ${p.PIMProductName || ""}`.trim();
            if (name) {
              unique.add(name);
              hasValid = true;
            }
          });
          if (!hasValid) {
            hasEmpty = true;
          }
        }
      });
      const result: string[] = Array.from(unique).sort((a, b) =>
        a.localeCompare(b, undefined, { sensitivity: "base" })
      );
      if (hasEmpty) {
        result.unshift("(Empty)");
      }
      return result;
    },
    [getItemsFilteredExcluding]
  );

  const businessLineOptions = useMemo(
    () => {
      const items = getItemsFilteredExcluding("BusinessLine");
      const unique = new Set<string>();
      let hasEmpty = false;
      items.forEach((item) => {
        const raw = item.BusinessLine;
        if (!raw || raw.trim() === "") {
          hasEmpty = true;
        } else {
          raw.split(",").forEach((val) => {
            const trimmed = val.trim();
            if (trimmed) unique.add(trimmed);
          });
        }
      });
      const result: string[] = Array.from(unique).sort((a, b) =>
        a.localeCompare(b, undefined, { sensitivity: "base" })
      );
      if (hasEmpty) {
        result.unshift("(Empty)");
      }
      return result;
    },
    [getItemsFilteredExcluding]
  );

  const countryOptions = useMemo(
    () => {
      const items = getItemsFilteredExcluding("CountrySoldTo");
      const unique = new Set<string>();
      let hasEmpty = false;
      items.forEach((item) => {
        const raw = item.CountrySoldTo;
        if (!raw || raw.trim() === "") {
          hasEmpty = true;
        } else {
          raw.split(",").forEach((val) => {
            const trimmed = val.trim();
            if (trimmed) unique.add(trimmed);
          });
        }
      });
      const result: string[] = Array.from(unique).sort((a, b) =>
        a.localeCompare(b, undefined, { sensitivity: "base" })
      );
      if (hasEmpty) {
        result.unshift("(Empty)");
      }
      return result;
    },
    [getItemsFilteredExcluding]
  );

  const documentTypeOptions = useMemo(
    () => {
      const items = getItemsFilteredExcluding("DocumentTypeSearchText");
      const unique = new Set<string>();
      let hasEmpty = false;
      items.forEach((item) => {
        const raw = item.DocumentTypeSearchText;
        if (!raw || raw.trim() === "") {
          hasEmpty = true;
        } else {
          const trimmed = raw.trim();
          if (trimmed) unique.add(trimmed);
        }
      });
      const result: string[] = Array.from(unique).sort((a, b) =>
        a.localeCompare(b, undefined, { sensitivity: "base" })
      );
      if (hasEmpty) {
        result.unshift("(Empty)");
      }
      return result;
    },
    [getItemsFilteredExcluding]
  );

  const subDocumentTypeOptions = useMemo(
    () => {
      const items = getItemsFilteredExcluding("SubDocumentTypeSearchText");
      const unique = new Set<string>();
      let hasEmpty = false;
      items.forEach((item) => {
        const raw = item.SubDocumentTypeSearchText;
        if (!raw || raw.trim() === "") {
          hasEmpty = true;
        } else {
          raw.split(",").forEach((val) => {
            const trimmed = val.trim();
            if (trimmed) unique.add(trimmed);
          });
        }
      });
      const result: string[] = Array.from(unique).sort((a, b) =>
        a.localeCompare(b, undefined, { sensitivity: "base" })
      );
      if (hasEmpty) {
        result.unshift("(Empty)");
      }
      return result;
    },
    [getItemsFilteredExcluding]
  );

  const confidentialityOptions = useMemo(
    () => {
      const items = getItemsFilteredExcluding("Confidentiality");
      const unique = new Set<string>();
      let hasEmpty = false;
      items.forEach((item) => {
        const raw = item.Confidentiality;
        if (!raw || raw.trim() === "") {
          hasEmpty = true;
        } else {
          raw.split(",").forEach((val) => {
            const trimmed = val.trim();
            if (trimmed) unique.add(trimmed);
          });
        }
      });
      const result: string[] = Array.from(unique).sort((a, b) =>
        a.localeCompare(b, undefined, { sensitivity: "base" })
      );
      if (hasEmpty) {
        result.unshift("(Empty)");
      }
      return result;
    },
    [getItemsFilteredExcluding]
  );

  const availableBusinessLines = useMemo(() => {
    const set = new Set([...libraryChoices.businessLine, ...businessLineOptions.filter((o) => o !== "(Empty)")]);
    return Array.from(set)
      .filter(Boolean)
      .sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
  }, [libraryChoices.businessLine, businessLineOptions]);

  const availableCountries = useMemo(() => {
    const set = new Set([...libraryChoices.country, ...countryOptions.filter((o) => o !== "(Empty)")]);
    return Array.from(set)
      .filter(Boolean)
      .sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
  }, [libraryChoices.country, countryOptions]);

  const availableConfidentialities = useMemo(() => {
    const set = new Set([...libraryChoices.confidentiality, ...confidentialityOptions.filter((o) => o !== "(Empty)")]);
    return Array.from(set)
      .filter(Boolean)
      .sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
  }, [libraryChoices.confidentiality, confidentialityOptions]);

  const issuedByOptions = useMemo(() => {
    const items = getItemsFilteredExcluding("IssuedBy");
    const unique = new Set<string>();
    let hasEmpty = false;
    items.forEach((item) => {
      const raw = item.IssuedBy;
      if (!raw || raw.trim() === "") {
        hasEmpty = true;
      } else {
        raw.split(",").forEach((val) => {
          const trimmed = val.trim();
          if (trimmed) unique.add(trimmed);
        });
      }
    });
    const result: string[] = Array.from(unique).sort((a, b) =>
      a.localeCompare(b, undefined, { sensitivity: "base" })
    );
    if (hasEmpty) {
      result.unshift("(Empty)");
    }
    return result;
  }, [getItemsFilteredExcluding]);

  const documentStatusOptions = useMemo(() => {
    const items = getItemsFilteredExcluding("DocumentStatus");
    const unique = new Set<string>();
    let hasEmpty = false;
    items.forEach((item) => {
      const raw = item.DocumentStatus;
      if (!raw || raw.trim() === "") {
        hasEmpty = true;
      } else {
        raw.split(",").forEach((val) => {
          const trimmed = val.trim();
          if (trimmed) unique.add(trimmed);
        });
      }
    });
    const result: string[] = Array.from(unique).sort((a, b) =>
      a.localeCompare(b, undefined, { sensitivity: "base" })
    );
    if (hasEmpty) {
      result.unshift("(Empty)");
    }
    return result;
  }, [getItemsFilteredExcluding]);

  const documentLanguageOptions = useMemo(() => {
    const items = getItemsFilteredExcluding("DocumentLanguage");
    const unique = new Set<string>();
    let hasEmpty = false;
    items.forEach((item) => {
      const raw = item.DocumentLanguage;
      if (!raw || raw.trim() === "") {
        hasEmpty = true;
      } else {
        raw.split(",").forEach((val) => {
          const trimmed = val.trim();
          if (trimmed) unique.add(trimmed);
        });
      }
    });
    const result: string[] = Array.from(unique).sort((a, b) =>
      a.localeCompare(b, undefined, { sensitivity: "base" })
    );
    if (hasEmpty) {
      result.unshift("(Empty)");
    }
    return result;
  }, [getItemsFilteredExcluding]);

  const reviewedByOptions = useMemo(() => {
    const items = getItemsFilteredExcluding("ReviewedByTitle");
    const unique = new Set<string>();
    let hasEmpty = false;
    items.forEach((item) => {
      const raw = item.ReviewedByTitle;
      if (!raw || raw.trim() === "") {
        hasEmpty = true;
      } else {
        raw.split(/[\r\n;,]+/).forEach((val) => {
          const trimmed = val.trim();
          if (trimmed) unique.add(trimmed);
        });
      }
    });
    const result: string[] = Array.from(unique).sort((a, b) =>
      a.localeCompare(b, undefined, { sensitivity: "base" })
    );
    if (hasEmpty) {
      result.unshift("(Empty)");
    }
    result.unshift("Me");
    return result;
  }, [getItemsFilteredExcluding]);

  const authorOptions = useMemo(() => {
    const items = getItemsFilteredExcluding("AuthorTitle");
    const unique = new Set<string>();
    let hasEmpty = false;
    items.forEach((item) => {
      const raw = item.AuthorTitle;
      if (!raw || raw.trim() === "") {
        hasEmpty = true;
      } else {
        const trimmed = raw.trim();
        if (trimmed) unique.add(trimmed);
      }
    });
    const result: string[] = Array.from(unique).sort((a, b) =>
      a.localeCompare(b, undefined, { sensitivity: "base" })
    );
    if (hasEmpty) {
      result.unshift("(Empty)");
    }
    result.unshift("Me");
    return result;
  }, [getItemsFilteredExcluding]);

  const editorOptions = useMemo(() => {
    const items = getItemsFilteredExcluding("EditorTitle");
    const unique = new Set<string>();
    let hasEmpty = false;
    items.forEach((item) => {
      const raw = item.EditorTitle;
      if (!raw || raw.trim() === "") {
        hasEmpty = true;
      } else {
        const trimmed = raw.trim();
        if (trimmed) unique.add(trimmed);
      }
    });
    const result: string[] = Array.from(unique).sort((a, b) =>
      a.localeCompare(b, undefined, { sensitivity: "base" })
    );
    if (hasEmpty) {
      result.unshift("(Empty)");
    }
    result.unshift("Me");
    return result;
  }, [getItemsFilteredExcluding]);

  const columns_AllProducts = useMemo<MRT_ColumnDef<doclib_AllProducts>[]>(
    () => [
      {
        accessorKey: "filename",
        header: "File Name",
        size: 160,
        minSize: 140,
        filterFn: filenameFilterFn,
        Filter: ({ column }) => <FilenameFilterInput column={column} />,
        Header: ({ column }) => {
          const raw = column.getFilterValue();
          const filterVal: { text?: string; fileTypes?: string[] } =
            typeof raw === "string"
              ? { text: raw, fileTypes: [] }
              : (raw as { text?: string; fileTypes?: string[] }) || {};
          const activeTypes = filterVal.fileTypes || [];

          return (
            <Box sx={{ display: "flex", alignItems: "center", gap: 0.5, flexWrap: "nowrap" }}>
              <span>File Name</span>
              {activeTypes.length > 0 && (
                <Chip
                  size="small"
                  label={activeTypes.length === 1 ? activeTypes[0].toUpperCase() : `${activeTypes.length} Types`}
                  onDelete={(e) => {
                    e.stopPropagation();
                    if (!filterVal.text) {
                      column.setFilterValue(undefined);
                    } else {
                      column.setFilterValue({ ...filterVal, fileTypes: [] });
                    }
                  }}
                  sx={{
                    height: "17px",
                    fontSize: "9.5px",
                    fontWeight: 700,
                    bgcolor: "#e3f2fd",
                    color: "#1976d2",
                    "& .MuiChip-deleteIcon": {
                      fontSize: "11px",
                      color: "#1976d2",
                      marginRight: "2px",
                    },
                  }}
                />
              )}
            </Box>
          );
        },
        renderColumnActionsMenuItems: ({ column, table, internalColumnMenuItems, closeMenu }) => {
          const raw = column.getFilterValue();
          const filterVal: { text?: string; fileTypes?: string[] } =
            typeof raw === "string"
              ? { text: raw, fileTypes: [] }
              : (raw as { text?: string; fileTypes?: string[] }) || {};
          const currentTypes = filterVal.fileTypes || [];

          const toggleType = (type: string) => {
            const exists = currentTypes.includes(type);
            const next = exists
              ? currentTypes.filter((t) => t !== type)
              : [...currentTypes, type];
            if (next.length === 0 && !filterVal.text) {
              column.setFilterValue(undefined);
            } else {
              column.setFilterValue({ ...filterVal, fileTypes: next });
            }
          };

          const clearTypes = () => {
            if (!filterVal.text) {
              column.setFilterValue(undefined);
            } else {
              column.setFilterValue({ ...filterVal, fileTypes: [] });
            }
          };

          const fileTypeDefinitions = [
            { id: "word", label: "Word Documents", ext: ".docx, .doc", icon: <WordFileTypeIcon size={20} /> },
            { id: "excel", label: "Excel Spreadsheets", ext: ".xlsx, .csv", icon: <ExcelFileTypeIcon size={20} /> },
            { id: "powerpoint", label: "PowerPoint Presentations", ext: ".pptx, .ppt", icon: <PptFileTypeIcon size={20} /> },
            { id: "pdf", label: "PDF Documents", ext: ".pdf", icon: <PdfFileTypeIcon size={20} /> },
          ];

          const selectedRowCount = table.getSelectedRowModel().rows.length;
          const totalRowCount = table.getPrePaginationRowModel().rows.length;
          const exportRowCount = selectedRowCount > 0 ? selectedRowCount : totalRowCount;

          return [
            // Quick icons bar at top of menu (horizontally stacked)
            <Box key="file-type-section" sx={{ px: 2, pt: 1.25, pb: 1, borderBottom: "1px solid #edebe9" }}>
              <Typography
                variant="caption"
                sx={{
                  display: "block",
                  fontSize: "10.5px",
                  fontWeight: 700,
                  color: "text.secondary",
                  textTransform: "uppercase",
                  letterSpacing: "0.5px",
                  mb: 0.75,
                }}
              >
                Filter by File Type
              </Typography>
              <Box sx={{ display: "flex", gap: 0.75, alignItems: "center" }}>
                <Tooltip title="All Files (Clear filter)" arrow>
                  <Button
                    size="small"
                    variant={currentTypes.length === 0 ? "contained" : "outlined"}
                    onClick={clearTypes}
                    sx={{
                      minWidth: "32px",
                      height: "26px",
                      fontSize: "11px",
                      fontWeight: 600,
                      textTransform: "none",
                      px: 1,
                      py: 0,
                      borderRadius: "6px",
                      boxShadow: "none",
                      bgcolor: currentTypes.length === 0 ? "#424242" : "transparent",
                      color: currentTypes.length === 0 ? "#fff" : "text.secondary",
                      borderColor: "#d0d0d0",
                      "&:hover": {
                        bgcolor: currentTypes.length === 0 ? "#212121" : "#f5f5f5",
                        boxShadow: "none",
                      },
                    }}
                  >
                    All
                  </Button>
                </Tooltip>

                {fileTypeDefinitions.map((ft) => {
                  const isSelected = currentTypes.includes(ft.id);
                  return (
                    <Tooltip key={ft.id} title={`${ft.label} (${ft.ext})`} arrow>
                      <IconButton
                        size="small"
                        onClick={() => toggleType(ft.id)}
                        sx={{
                          p: "3px",
                          borderRadius: "6px",
                          border: isSelected ? "2px solid #1976d2" : "1px solid #e0e0e0",
                          bgcolor: isSelected ? "rgba(25, 118, 210, 0.1)" : "transparent",
                          transition: "all 0.15s ease-in-out",
                          "&:hover": {
                            bgcolor: isSelected ? "rgba(25, 118, 210, 0.18)" : "#f5f5f5",
                            transform: "scale(1.08)",
                          },
                        }}
                      >
                        {ft.icon}
                      </IconButton>
                    </Tooltip>
                  );
                })}
              </Box>
            </Box>,

            // Export to Excel menu item
            <MenuItem
              key="export-excel-action"
              onClick={() => {
                closeMenu();
                const selectedRows = table.getSelectedRowModel().rows;
                const rowsToExport =
                  selectedRows.length > 0
                    ? selectedRows.map((r) => r.original)
                    : table.getPrePaginationRowModel().rows.map((r) => r.original);
                exportToExcelCsv(rowsToExport);
              }}
              sx={{
                fontSize: "12.5px",
                py: 0.9,
                px: 2,
                display: "flex",
                alignItems: "center",
                gap: 1.25,
                color: "#107c41",
                "&:hover": {
                  bgcolor: "rgba(16, 124, 65, 0.08)",
                },
              }}
            >
              <ExcelFileTypeIcon size={18} />
              <ListItemText
                primary={`Export to Excel (${exportRowCount.toLocaleString()} ${selectedRowCount > 0 ? "selected" : "rows"})`}
                secondary="Download Excel-compatible .csv"
                primaryTypographyProps={{ fontSize: "12.5px", fontWeight: 600, color: "#107c41" }}
                secondaryTypographyProps={{ fontSize: "10.5px" }}
              />
            </MenuItem>,

            <Divider key="divider-mrt-actions" sx={{ my: 0.5 }} />,

            internalColumnMenuItems,
          ];
        },
        Cell: ({ row }) => {
          const rawName = row.original.filename || "";
          const displayName =
            rawName.length > 35
              ? `${rawName.slice(0, 32)}...`
              : rawName;

          const fileIcon = getFilenameIcon(rawName, 16);

          return (
            <Box
              component="span"
              onClick={(e: React.MouseEvent<HTMLElement>) => {
                e.preventDefault();
                e.stopPropagation();
                setSelectedFileForAction(row.original);
                setFileMenuAnchorEl(e.currentTarget);
              }}
              title={rawName}
              sx={{
                display: "inline-flex",
                alignItems: "center",
                gap: "8px",
                maxWidth: "100%",
                cursor: "pointer",
                userSelect: "none",
                "&:hover .sp-filename-text": {
                  color: "#0078d4",
                  textDecoration: "underline",
                },
              }}
            >
              <Box sx={{ display: "inline-flex", alignItems: "center", flexShrink: 0 }}>
                {fileIcon}
              </Box>
              <Box
                component="span"
                className="sp-filename-text"
                sx={{
                  color: "#323130",
                  fontWeight: 500,
                  fontSize: "12.5px",
                  overflow: "hidden",
                  textOverflow: "ellipsis",
                  whiteSpace: "nowrap",
                  transition: "color 0.15s ease",
                }}
              >
                {displayName}
              </Box>
            </Box>
          );
        },
      },
      {
        accessorKey: "ManufacturerSearchText",
        header: "Clients",
        filterFn: multiSelectFilterFn,
        Filter: ({ column }) => (
          <MultiSelectAutocompleteFilter
            column={column}
            options={clientOptions}
            placeholder="Select/type client..."
          />
        ),
        size: 220,
        minSize: 180,
        Cell: ({ cell, column }) => {
          const raw = String(cell.getValue() || "").trim();
          if (!raw || raw === "-") return "-";
          const items = raw
            .split(/[\r\n;,]+/)
            .map((item) => item.trim())
            .filter(Boolean);
          return (
            <CollapsibleItemList
              items={items}
              filterValues={column.getFilterValue()}
            />
          );
        },
      },
      {
        accessorKey: "PIMProductSearchText",
        header: "Products",
        filterFn: multiSelectFilterFn,
        Filter: ({ column }) => (
          <MultiSelectAutocompleteFilter
            column={column}
            options={productOptions}
            placeholder="Select/type product..."
          />
        ),
        size: 180,
        minSize: 180,
        Cell: ({ row, column }) => {
          const items = (row.original.PIMProduct || [])
            .map((p: IProductLookupItem) => `${p.Title || ""} ${p.PIMProductName || ""}`.trim())
            .filter(Boolean);
          if (items.length === 0) return "-";
          return (
            <CollapsibleItemList
              items={items}
              filterValues={column.getFilterValue()}
            />
          );
        },
      },
      {
        accessorKey: "DocumentTypeSearchText",
        header: "Document Type",
        filterVariant: "multi-select",
        filterSelectOptions: documentTypeOptions,
        filterFn: multiSelectFilterFn,
        size: 150,
        minSize: 150,
        Cell: ({ cell }) => {
          const raw = String(cell.getValue() || "").trim();
          return <span>{raw || "-"}</span>;
        },
      },
      {
        accessorKey: "SubDocumentTypeSearchText",
        header: "Sub Document Type",
        filterFn: multiSelectFilterFn,
        Filter: ({ column }) => (
          <MultiSelectAutocompleteFilter
            column={column}
            options={subDocumentTypeOptions}
            placeholder="Select/type sub doc type..."
          />
        ),
        size: 175,
        minSize: 175,
        Cell: ({ cell, column }) => {
          const raw = String(cell.getValue() || "").trim();
          if (!raw || raw === "-") return "-";
          const items = raw
            .split(",")
            .map((item) => item.trim())
            .filter(Boolean);
          return (
            <CollapsibleItemList
              items={items}
              filterValues={column.getFilterValue()}
            />
          );
        },
      },
      {
        accessorKey: "BusinessLine",
        header: "Business Line",
        filterVariant: "multi-select",
        filterSelectOptions: businessLineOptions,
        filterFn: multiSelectFilterFn,
        size: 155,
        minSize: 140,
        Cell: ({ cell, column }) => {
          const raw = String(cell.getValue() || "").trim();
          if (!raw || raw === "-") return "-";
          const items = raw
            .split(",")
            .map((v) => v.trim())
            .filter(Boolean);

          return (
            <CollapsibleItemList
              items={items}
              filterValues={column.getFilterValue()}
              renderItem={(item, idx) => {
                const style =
                  fieldFormatters.businessLine[item.toLowerCase()] ||
                  getDynamicChipStyle(item);

                return (
                  <span
                    key={idx}
                    style={{
                      backgroundColor: style.bg,
                      color: style.text,
                      border: `1px solid ${style.border}`,
                      borderRadius: "12px",
                      padding: "1px 8px",
                      fontSize: "11px",
                      fontWeight: 600,
                      lineHeight: "18px",
                      display: "inline-block",
                      whiteSpace: "nowrap",
                    }}
                  >
                    {item}
                  </span>
                );
              }}
            />
          );
        },
      },
      {
        accessorKey: "CountrySoldTo",
        header: "Country Sold To",
        filterVariant: "multi-select",
        filterSelectOptions: countryOptions,
        filterFn: multiSelectFilterFn,
        size: 165,
        minSize: 165,
        Cell: ({ cell, column }) => {
          const raw = String(cell.getValue() || "").trim();
          if (!raw || raw === "-") return "-";
          const items = raw
            .split(",")
            .map((item) => item.trim())
            .filter(Boolean);
          return (
            <CollapsibleItemList
              items={items}
              filterValues={column.getFilterValue()}
            />
          );
        },
      },
      {
        accessorKey: "IssuedBy",
        header: "Issued By",
        filterVariant: "multi-select",
        filterSelectOptions: issuedByOptions,
        filterFn: multiSelectFilterFn,
        size: 140,
        minSize: 130,
        Cell: ({ cell }) => {
          const raw = String(cell.getValue() || "").trim();
          return <span>{raw || "-"}</span>;
        },
      },
      {
        accessorKey: "DocumentStatus",
        header: "Document Status",
        filterVariant: "multi-select",
        filterSelectOptions: documentStatusOptions,
        filterFn: multiSelectFilterFn,
        size: 150,
        minSize: 130,
        Cell: ({ cell, column }) => {
          const raw = String(cell.getValue() || "").trim();
          if (!raw || raw === "-") return "-";
          const items = raw
            .split(",")
            .map((v) => v.trim())
            .filter(Boolean);

          return (
            <CollapsibleItemList
              items={items}
              filterValues={column.getFilterValue()}
              renderItem={(item, idx) => {
                const style =
                  (fieldFormatters.documentStatus &&
                    fieldFormatters.documentStatus[item.toLowerCase()]) ||
                  getSemanticDocumentStatusStyle(item) ||
                  getDynamicChipStyle(item);

                return (
                  <span
                    key={idx}
                    style={{
                      backgroundColor: style.bg,
                      color: style.text,
                      border: `1px solid ${style.border}`,
                      borderRadius: "12px",
                      padding: "1px 8px",
                      fontSize: "11px",
                      fontWeight: 600,
                      lineHeight: "18px",
                      display: "inline-block",
                      whiteSpace: "nowrap",
                    }}
                  >
                    {item}
                  </span>
                );
              }}
            />
          );
        },
      },
      {
        accessorKey: "DocumentLanguage",
        header: "Document Language",
        filterVariant: "multi-select",
        filterSelectOptions: documentLanguageOptions,
        filterFn: multiSelectFilterFn,
        size: 140,
        minSize: 130,
        Cell: ({ cell }) => {
          const raw = String(cell.getValue() || "").trim();
          return <span>{raw || "-"}</span>;
        },
      },
      {
        accessorKey: "Confidentiality",
        header: "Confidentiality",
        filterVariant: "multi-select",
        filterSelectOptions: confidentialityOptions,
        filterFn: multiSelectFilterFn,
        size: 150,
        minSize: 140,
        Cell: ({ cell, column }) => {
          const raw = String(cell.getValue() || "").trim();
          if (!raw || raw === "-") return "-";
          const items = raw
            .split(",")
            .map((v) => v.trim())
            .filter(Boolean);

          return (
            <CollapsibleItemList
              items={items}
              filterValues={column.getFilterValue()}
              renderItem={(item, idx) => {
                const style =
                  fieldFormatters.confidentiality[item.toLowerCase()] ||
                  getSemanticConfidentialityStyle(item) ||
                  getDynamicChipStyle(item);

                return (
                  <span
                    key={idx}
                    style={{
                      backgroundColor: style.bg,
                      color: style.text,
                      border: `1px solid ${style.border}`,
                      borderRadius: "12px",
                      padding: "1px 8px",
                      fontSize: "11px",
                      fontWeight: 600,
                      lineHeight: "18px",
                      display: "inline-block",
                      whiteSpace: "nowrap",
                    }}
                  >
                    {item}
                  </span>
                );
              }}
            />
          );
        },
      },
      {
        accessorKey: "DocumentDate",
        header: "Document Date",
        size: 100,
        minSize: 100,
        filterFn: documentDateFilter,
        Filter: ({ column }) => (
          <DateFilterControl
            column={column}
            label="Min Date"
          />
        ),
        Cell: ({ cell }) => {
          const value = cell.getValue<Date | string | null>();
          if (!value) return "-";

          const d = dayjs(value);
          if (!d.isValid()) return "-";

          return d.format("DD/MM/YYYY");
        },
      },
      {
        accessorKey: "ExpiryDate",
        header: "Expiry Date",
        size: 115,
        minSize: 100,
        filterFn: documentDateFilter,
        Filter: ({ column }) => (
          <DateFilterControl
            column={column}
            label="Min Date"
          />
        ),
        Cell: ({ cell }) => {
          const value = cell.getValue<Date | string | null>();
          if (!value) return "-";

          const d = dayjs(value);
          if (!d.isValid()) return "-";

          const style = evaluateDateCustomFormatter(
            d,
            fieldFormatters.expiryDateCustomFormatter
          );

          return (
            <span
              style={{
                backgroundColor: style.bg,
                color: style.text,
                border: `1px solid ${style.border}`,
                borderRadius: "12px",
                padding: "1px 8px",
                fontSize: "11px",
                fontWeight: 600,
                lineHeight: "18px",
                display: "inline-block",
                whiteSpace: "nowrap",
              }}
            >
              {d.format("DD/MM/YYYY")}
            </span>
          );
        },
      },
      {
        accessorKey: "NextReviewDate",
        header: "Next Review Date",
        size: 130,
        minSize: 110,
        filterFn: documentDateFilter,
        Filter: ({ column }) => (
          <DateFilterControl
            column={column}
            label="Min Date"
          />
        ),
        Cell: ({ cell }) => {
          const value = cell.getValue<Date | string | null>();
          if (!value) return "-";

          const d = dayjs(value);
          if (!d.isValid()) return "-";

          const style = evaluateDateCustomFormatter(
            d,
            fieldFormatters.nextReviewDateCustomFormatter
          );

          return (
            <span
              style={{
                backgroundColor: style.bg,
                color: style.text,
                border: `1px solid ${style.border}`,
                borderRadius: "12px",
                padding: "1px 8px",
                fontSize: "11px",
                fontWeight: 600,
                lineHeight: "18px",
                display: "inline-block",
                whiteSpace: "nowrap",
              }}
            >
              {d.format("DD/MM/YYYY")}
            </span>
          );
        },
      },
      {
        accessorKey: "Created",
        header: "Created",
        size: 110,
        minSize: 100,
        filterFn: exactDateFilter,
        Filter: ({ column }) => (
          <DateFilterControl
            column={column}
            label="Date"
          />
        ),
        Cell: ({ cell }) => {
          const value = cell.getValue<Date | string | null>();
          if (!value) return "-";

          const d = dayjs(value);
          if (!d.isValid()) return "-";

          return d.format("DD/MM/YYYY");
        },
      },
      {
        accessorKey: "AuthorTitle",
        header: "Created by",
        filterVariant: "multi-select",
        filterSelectOptions: authorOptions,
        filterFn: personFilterFn,
        size: 150,
        minSize: 140,
        Cell: ({ cell }) => {
          const raw = String(cell.getValue() || "").trim();
          return <span>{raw || "-"}</span>;
        },
      },
      {
        accessorKey: "Modified",
        header: "Modified",
        size: 110,
        minSize: 100,
        filterFn: exactDateFilter,
        Filter: ({ column }) => (
          <DateFilterControl
            column={column}
            label="Date"
          />
        ),
        Cell: ({ cell }) => {
          const value = cell.getValue<Date | string | null>();
          if (!value) return "-";

          const d = dayjs(value);
          if (!d.isValid()) return "-";

          return d.format("DD/MM/YYYY");
        },
      },
      {
        accessorKey: "EditorTitle",
        header: "Modified by",
        filterVariant: "multi-select",
        filterSelectOptions: editorOptions,
        filterFn: personFilterFn,
        size: 150,
        minSize: 140,
        Cell: ({ cell }) => {
          const raw = String(cell.getValue() || "").trim();
          return <span>{raw || "-"}</span>;
        },
      },
      {
        accessorKey: "ReviewedByTitle",
        header: "Reviewed By",
        filterVariant: "multi-select",
        filterSelectOptions: reviewedByOptions,
        filterFn: personFilterFn,
        size: 150,
        minSize: 140,
        Cell: ({ cell }) => {
          const raw = String(cell.getValue() || "").trim();
          return <span>{raw || "-"}</span>;
        },
      },
      {
        accessorKey: "Alerts",
        header: "Alerts",
        filterFn: alertsFilterFn,
        Filter: ({ column }) => (
          <MultiSelectAutocompleteFilter
            column={column}
            options={alertsOptions}
            placeholder="Select/type alert..."
          />
        ),
        size: 160,
        minSize: 140,
        Cell: ({ row, column }) => {
          const items = getRowAlerts(row.original);
          if (items.length === 0) return "-";

          return (
            <CollapsibleItemList
              items={items}
              filterValues={column.getFilterValue()}
              renderItem={(item, idx) => {
                const rule = fieldFormatters.alerts;
                const style =
                  (rule && rule.text.toLowerCase() === item.toLowerCase() ? rule.style : null) ||
                  getSemanticAlertStyle(item) ||
                  getDynamicChipStyle(item);

                return (
                  <span
                    key={idx}
                    style={{
                      backgroundColor: style.bg,
                      color: style.text,
                      border: `1px solid ${style.border}`,
                      borderRadius: "12px",
                      padding: "1px 8px",
                      fontSize: "11px",
                      fontWeight: 600,
                      lineHeight: "18px",
                      display: "inline-block",
                      whiteSpace: "nowrap",
                    }}
                  >
                    {item}
                  </span>
                );
              }}
            />
          );
        },
      },
      {
        accessorKey: "id",
        header: "ID",
        size: 70,
        minSize: 60,
        filterFn: "equals",
        Cell: ({ cell }) => {
          const val = cell.getValue<number | undefined>();
          return <span>{val !== undefined && val !== null ? val : "-"}</span>;
        },
      },
    ],
    [
      props.urlSite,
      clientOptions,
      productOptions,
      businessLineOptions,
      countryOptions,
      documentTypeOptions,
      subDocumentTypeOptions,
      confidentialityOptions,
      issuedByOptions,
      documentStatusOptions,
      documentLanguageOptions,
      reviewedByOptions,
      authorOptions,
      editorOptions,
      alertsOptions,
      fieldFormatters,
      getRowAlerts,
      personFilterFn,
      alertsFilterFn,
    ]
  );

  const handleTableMouseDown = (e: React.MouseEvent<HTMLDivElement>) => {
    if (e.button !== 0) return; // Only trigger on left-click

    const target = e.target as HTMLElement;

    // Only allow drag-to-scroll if the click originates from the table header (thead)
    // This allows users to freely select and copy text in the table body cells
    if (!target.closest("thead")) {
      return;
    }

    // Don't drag if clicking interactive controls, filter inputs, sort labels, resize handles, or grab handles
    if (
      target.closest(
        'button, input, textarea, select, [role="button"], [role="checkbox"], .MuiInputBase-root, .MuiIconButton-root, .MuiSelect-select, .MuiTableSortLabel-root, .Mui-TableHeadCell-ResizeHandle, .Mui-TableHeadCell-GrabHandle, .Mui-TableHeadCell-ColumnActionsButton'
      )
    ) {
      return;
    }

    const container = e.currentTarget;
    const startClientX = e.clientX;
    const initialScrollLeft = container.scrollLeft;
    let hasDragged = false;

    // Prevent HTML5 native drag on the header cell so our smooth horizontal scroll works
    const preventDragStart = (dragEvent: DragEvent) => {
      dragEvent.preventDefault();
    };
    window.addEventListener("dragstart", preventDragStart, { once: true });

    const handleMouseMove = (moveEvent: MouseEvent) => {
      const dx = moveEvent.clientX - startClientX;
      if (!hasDragged && Math.abs(dx) > 4) {
        hasDragged = true;
        container.style.cursor = "grabbing";
        container.style.userSelect = "none";
        document.body.style.cursor = "grabbing";
        document.body.style.userSelect = "none";
      }

      if (hasDragged) {
        container.scrollLeft = initialScrollLeft - dx;
      }
    };

    const handleMouseUp = () => {
      window.removeEventListener("mousemove", handleMouseMove);
      window.removeEventListener("mouseup", handleMouseUp);
      window.removeEventListener("dragstart", preventDragStart);

      if (hasDragged) {
        container.style.cursor = "";
        container.style.removeProperty("user-select");
        document.body.style.cursor = "";
        document.body.style.removeProperty("user-select");

        const preventClick = (clickEvent: MouseEvent) => {
          clickEvent.stopPropagation();
          clickEvent.preventDefault();
        };
        window.addEventListener("click", preventClick, {
          capture: true,
          once: true,
        });
      }
    };

    window.addEventListener("mousemove", handleMouseMove);
    window.addEventListener("mouseup", handleMouseUp);
  };

  const table = useMaterialReactTable({
    columns: columns_AllProducts,
    data: items_AllProducts,
    enableRowSelection: true,
    enableStickyHeader: true,
    enableFullScreenToggle: false,
    positionToolbarAlertBanner: "none",
    renderTopToolbarCustomActions: ({ table }) => (
      <Box sx={{ display: "flex", gap: "10px", alignItems: "center", flexWrap: "wrap" }}>
        <IconButton
          color="primary"
          disabled={table.getSelectedRowModel().rows.length === 0}
          onClick={() => {
            const selectedRows = table.getSelectedRowModel().rows.map(row => row.original);
            setSelectedRowsData(selectedRows);
            setIsShareDialogOpen(true);
          }}
          title="Share Selected"
        >
          <ShareIcon />
        </IconButton>

        {/* --- SharePoint OOB Style View Switcher --- */}
        <Box sx={{ display: "flex", alignItems: "center" }}>
          <Button
            size="small"
            onClick={(e) => setViewMenuAnchorEl(e.currentTarget)}
            endIcon={<KeyboardArrowDownIcon sx={{ fontSize: 16, color: "action.active" }} />}
            sx={{
              height: "28px",
              fontSize: "12.5px",
              fontWeight: 600,
              textTransform: "none",
              color: "text.primary",
              backgroundColor: Boolean(viewMenuAnchorEl) ? "rgba(0, 0, 0, 0.06)" : "transparent",
              border: "1px solid",
              borderColor: Boolean(viewMenuAnchorEl) ? "#c8c6c4" : "#e1dfdd",
              borderRadius: "4px",
              px: 1.25,
              gap: 0.5,
              "&:hover": {
                backgroundColor: "rgba(0, 0, 0, 0.04)",
                borderColor: "#a19f9d",
              },
            }}
            title="Switch view or save current view"
          >
            {currentView.name}
          </Button>

          <Menu
            anchorEl={viewMenuAnchorEl}
            open={Boolean(viewMenuAnchorEl)}
            onClose={() => setViewMenuAnchorEl(null)}
            anchorOrigin={{ vertical: "bottom", horizontal: "left" }}
            transformOrigin={{ vertical: "top", horizontal: "left" }}
            PaperProps={{
              sx: {
                minWidth: 210,
                boxShadow: "0 4px 16px rgba(0, 0, 0, 0.14)",
                borderRadius: "6px",
                py: 0.5,
                border: "1px solid #edebe9",
              },
            }}
          >
            <Typography
              variant="caption"
              sx={{
                px: 2,
                pt: 1,
                pb: 0.5,
                display: "block",
                fontSize: "10.5px",
                fontWeight: 700,
                color: "text.secondary",
                textTransform: "uppercase",
                letterSpacing: "0.5px",
              }}
            >
              Standard Views
            </Typography>

            {BUILT_IN_VIEWS.map((v) => (
              <MenuItem
                key={v.id}
                selected={selectedViewId === v.id}
                onClick={() => {
                  handleApplyView(v.id);
                  setViewMenuAnchorEl(null);
                }}
                sx={{
                  fontSize: "12.5px",
                  py: 0.75,
                  px: 1.5,
                  display: "flex",
                  alignItems: "center",
                  gap: 1,
                }}
              >
                <Box sx={{ width: 16, display: "flex", alignItems: "center", justifyContent: "center" }}>
                  {selectedViewId === v.id && (
                    <CheckIcon sx={{ fontSize: 16, color: "primary.main" }} />
                  )}
                </Box>
                <ListItemText
                  primary={v.name}
                  primaryTypographyProps={{
                    fontSize: "12.5px",
                    fontWeight: selectedViewId === v.id ? 600 : 400,
                  }}
                />
              </MenuItem>
            ))}

            {customViews.length > 0 && (
              <>
                <Divider sx={{ my: 0.5 }} />
                <Typography
                  variant="caption"
                  sx={{
                    px: 2,
                    pt: 1,
                    pb: 0.5,
                    display: "block",
                    fontSize: "10.5px",
                    fontWeight: 700,
                    color: "text.secondary",
                    textTransform: "uppercase",
                    letterSpacing: "0.5px",
                  }}
                >
                  Saved Views
                </Typography>
                {customViews.map((v) => (
                  <MenuItem
                    key={v.id}
                    selected={selectedViewId === v.id}
                    onClick={() => {
                      handleApplyView(v.id);
                      setViewMenuAnchorEl(null);
                    }}
                    sx={{
                      fontSize: "12.5px",
                      py: 0.75,
                      px: 1.5,
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "space-between",
                    }}
                  >
                    <Box sx={{ display: "flex", alignItems: "center", gap: 1 }}>
                      <Box sx={{ width: 16, display: "flex", alignItems: "center", justifyContent: "center" }}>
                        {selectedViewId === v.id && (
                          <CheckIcon sx={{ fontSize: 16, color: "primary.main" }} />
                        )}
                      </Box>
                      <ListItemText
                        primary={v.name}
                        primaryTypographyProps={{
                          fontSize: "12.5px",
                          fontWeight: selectedViewId === v.id ? 600 : 400,
                        }}
                      />
                    </Box>
                    <IconButton
                      size="small"
                      onClick={(e) => {
                        e.stopPropagation();
                        handleDeleteView(v.id, e);
                      }}
                      sx={{
                        ml: 1,
                        p: 0.25,
                        color: "text.secondary",
                        "&:hover": { color: "error.main" },
                      }}
                      title="Delete view"
                    >
                      <DeleteOutlineIcon sx={{ fontSize: 15 }} />
                    </IconButton>
                  </MenuItem>
                ))}
              </>
            )}

            <Divider sx={{ my: 0.5 }} />

            <MenuItem
              onClick={() => {
                setViewMenuAnchorEl(null);
                setIsSaveViewDialogOpen(true);
              }}
              sx={{
                fontSize: "12.5px",
                py: 0.75,
                px: 1.5,
                color: "primary.main",
                display: "flex",
                alignItems: "center",
                gap: 1,
                "&:hover": {
                  backgroundColor: "rgba(25, 118, 210, 0.08)",
                },
              }}
            >
              <Box sx={{ width: 16, display: "flex", alignItems: "center", justifyContent: "center" }}>
                <BookmarkBorderIcon sx={{ fontSize: 16 }} />
              </Box>
              <ListItemText
                primary="Save view as..."
                primaryTypographyProps={{ fontSize: "12.5px", fontWeight: 500 }}
              />
            </MenuItem>
          </Menu>
        </Box>

        <Box sx={{ display: "flex", gap: "8px", alignItems: "center" }}>
          {!resultsLoading && items_AllProducts.length > 0 && (
            <Typography variant="caption" sx={{ fontSize: "12px", color: "text.secondary" }}>
              {items_AllProducts.length.toLocaleString()} records loaded
            </Typography>
          )}

          {hasMoreRecords && (
            <Button
              variant="outlined"
              size="small"
              onClick={handleLoadNextBatch}
              disabled={isLoadingMore}
              sx={{
                fontSize: "12px",
                textTransform: "none",
                py: 0.25,
                px: 1.5,
                minHeight: "28px",
              }}
            >
              {isLoadingMore ? (
                <Box sx={{ display: "flex", alignItems: "center", gap: "6px" }}>
                  <CircularProgress size={14} color="inherit" />
                  <span>Loading next 1,000...</span>
                </Box>
              ) : (
                "+ Load Next 1,000"
              )}
            </Button>
          )}

          {!hasMoreRecords && !resultsLoading && items_AllProducts.length > 0 && (
            <Typography
              variant="caption"
              sx={{
                fontSize: "11px",
                color: "#2e7d32",
                bgcolor: "#e8f5e9",
                px: 1,
                py: 0.25,
                borderRadius: "10px",
                fontWeight: 500,
              }}
            >
              All records loaded
            </Typography>
          )}

          <Button
            variant="outlined"
            size="small"
            color="primary"
            disabled={activeColumnFilters.length === 0}
            onClick={() => {
              setColumnFilters([]);
              table.resetColumnFilters();
              table.resetGlobalFilter();
            }}
            startIcon={<FilterListOffIcon fontSize="small" />}
            sx={{
              fontSize: "12px",
              textTransform: "none",
              py: 0.25,
              px: 1.5,
              minHeight: "28px",
            }}
            title="Clear all column filters"
          >
            Clear Filters{activeColumnFilters.length > 0 ? ` (${activeColumnFilters.length})` : ""}
          </Button>
        </Box>
      </Box>
    ),
    enableGrouping: true,
    enableColumnDragging: true,
    enableColumnOrdering: true,
    enableColumnResizing: true,
    enableColumnPinning: true,
    enableColumnActions: true,
    columnResizeMode: "onChange",
    layoutMode: "grid-no-grow",
    renderToolbarInternalActions: ({ table }) => (
      <>
        <MRT_ToggleGlobalFilterButton table={table} />
        <MRT_ShowHideColumnsButton table={table} />
        <MRT_ToggleDensePaddingButton table={table} />
      </>
    ),
    initialState: {
      density: "compact",
      columnFilters: [],
      pagination: {
        pageIndex: 0,
        pageSize: 50,
      },
    },
    muiTablePaperProps: {
      ref: paperRef,
      sx: {
        boxShadow: "none",
        border: "1px solid #e1dfdd",
        width: "100%",
        maxWidth: "100%",
        height: tableHeight,
        display: "flex",
        flexDirection: "column",
      },
    },
    muiTableContainerProps: {
      ref: tableContainerRef,
      onMouseDown: handleTableMouseDown,
      sx: {
        flex: 1,
        minHeight: 0,
        height: "100%",
        width: "100%",
        maxWidth: "100%",
      },
    },
    muiTableProps: {
      sx: {
        minWidth: "100%",
        width: "100%",
      },
    },
    muiTableHeadProps: {
      sx: {
        opacity: 1,
      },
    },
    muiTableHeadCellProps: {
      sx: {
        fontSize: "12.5px",
        padding: "7px 10px",
        lineHeight: 1.2,
        whiteSpace: "normal",
        wordBreak: "break-word",
        cursor: "grab",
        "&:active": {
          cursor: "grabbing",
        },
      },
    },
    muiTableBodyCellProps: {
      sx: {
        fontSize: "12.5px",
        padding: "5px 10px",
        lineHeight: 1.25,
        whiteSpace: "normal",
        wordBreak: "normal",
        overflowWrap: "break-word",
        userSelect: "text",
        cursor: "default",
      },
    },
    muiPaginationProps: {
      rowsPerPageOptions: [10, 50, 100, 500, 1000],
      sx: {
        ".MuiTablePagination-selectLabel, .MuiTablePagination-displayedRows": {
          fontSize: "13px",
        },
      },
    },
    onColumnFiltersChange: (updater) => {
      setColumnFilters((prev) => {
        const next = typeof updater === "function" ? updater(prev) : updater;
        return next.filter((f) => isFilterActive(f.value));
      });
    },
    onGroupingChange: setGrouping,
    onSortingChange: setSorting,
    onColumnVisibilityChange: setColumnVisibility,
    onColumnOrderChange: setColumnOrder,
    onColumnPinningChange: setColumnPinning,
    state: {
      isLoading: resultsLoading,
      showProgressBars: isLoadingMore,
      showColumnFilters: true,
      columnFilters,
      columnPinning,
      grouping,
      sorting,
      columnVisibility,
      ...(columnOrder.length > 0 ? { columnOrder } : {}),
    },
  });

  const renderKeywordInput = () => (
    <Autocomplete
      multiple
      freeSolo
      size="small"
      options={[]}
      value={additionalKeywords}
      inputValue={keywordInput}
      onInputChange={(_event, newInputValue, reason) => {
        if (reason === "input") {
          if (newInputValue.includes(",")) {
            const parts = newInputValue
              .split(",")
              .map((p) => p.trim())
              .filter(Boolean);
            setAdditionalKeywords((prev) =>
              Array.from(new Set([...prev, ...parts]))
            );
            setKeywordInput("");
          } else {
            setKeywordInput(newInputValue);
          }
        } else if (reason === "reset" || reason === "clear") {
          setKeywordInput("");
        }
      }}
      onChange={(_event, newValue) => {
        const cleaned: string[] = [];
        (newValue as (string | any)[]).forEach((val) => {
          if (typeof val === "string") {
            val.split(",").forEach((item) => {
              const trimmed = item.trim();
              if (trimmed && cleaned.indexOf(trimmed) === -1) {
                cleaned.push(trimmed);
              }
            });
          }
        });
        setAdditionalKeywords(cleaned);
        setKeywordInput("");
      }}
      renderTags={(value: readonly string[], getTagProps) =>
        value.map((option: string, index: number) => {
          const upper = option.trim().toUpperCase();
          const isOperator = upper === "AND" || upper === "OR" || upper === "NOT";
          return (
            <Chip
              {...getTagProps({ index })}
              key={index}
              label={isOperator ? upper : option}
              size="small"
              sx={{
                height: "24px",
                fontSize: "12px",
                margin: "2px",
                ...(isOperator
                  ? {
                      fontWeight: 700,
                      bgcolor: "#e0e7ff",
                      color: "#3730a3",
                      border: "1px solid #c7d2fe",
                    }
                  : {}),
              }}
            />
          );
        })
      }
      renderInput={(params) => (
        <TextField
          {...params}
          size="small"
          label="Additional Keywords"
          placeholder={
            additionalKeywords.length === 0
              ? "e.g. ALLERG*, cardio OR pulmonary, NOT pediatric"
              : ""
          }
        />
      )}
      sx={{
        width: "100%",
        "& .MuiInputBase-root": {
          minHeight: "40px",
          alignItems: "center",
          flexWrap: "wrap",
        },
      }}
    />
  );

  return (
    <ThemeProvider theme={compactTheme}>
      <style>{`
        .CanvasZoneContainer,
        [data-automation-id="CanvasZoneContainer"],
        .CanvasZoneContainer--read,
        .CanvasZone,
        [data-automation-id="CanvasZone"],
        .CanvasSection,
        [data-automation-id="CanvasSection"],
        .CanvasSection-col,
        [class*="CanvasSection-col"],
        [class*="CanvasSection-xl"],
        .ControlZone,
        .ControlZone--clean,
        [data-automation-id="CanvasControl"],
        [data-automation-id="CanvasLayout"] {
          max-width: none !important;
          width: 100% !important;
          margin-left: 0 !important;
          margin-right: 0 !important;
          padding-left: 0 !important;
          padding-right: 0 !important;
          box-sizing: border-box !important;
        }
        #spPageCanvasContent,
        [data-automation-id="Canvas"] {
          max-width: none !important;
          width: 100% !important;
        }
        @media (max-width: 768px) {
          .filter-row-grid {
            grid-template-columns: 1fr !important;
          }
        }
      `}</style>
      <Box
        sx={{
          width: "100%",
          maxWidth: "100%",
          boxSizing: "border-box",
          px: { xs: 0.5, sm: 1 },
          py: 0.5,
        }}
      >
        {/* Compact, full-width Header & Search Controls Bar */}
        <Box
          sx={{
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            width: "100%",
            mb: 1,
            px: 0.5,
          }}
        >
          <Typography
            variant="h6"
            component="h2"
            sx={{
              fontSize: "18px",
              fontWeight: 600,
              color: "text.primary",
              m: 0,
            }}
          >
            Search Clients & Products
          </Typography>

          <Box sx={{ display: "flex", alignItems: "center", gap: "12px" }}>
            <Button
              variant="contained"
              color="primary"
              size="small"
              disabled={resultsLoading || isLoadingMore}
              onClick={handleOpenSearchDialog}
              sx={{ textTransform: "none", fontWeight: 600, fontSize: "13px", px: 2, height: "32px" }}
            >
              New Search
            </Button>

            <Link
              component="button"
              variant="body2"
              onClick={handleRefresh}
              disabled={resultsLoading || isLoadingMore}
              sx={{
                display: "inline-flex",
                alignItems: "center",
                gap: "4px",
                fontSize: "13px",
                textDecoration: "none",
                cursor: resultsLoading || isLoadingMore ? "not-allowed" : "pointer",
                color: resultsLoading || isLoadingMore ? "text.disabled" : "primary.main",
                "&:hover": {
                  textDecoration: resultsLoading || isLoadingMore ? "none" : "underline",
                },
              }}
            >
              <RefreshIcon fontSize="small" sx={{ fontSize: "16px" }} />
              Refresh
            </Link>
          </Box>
        </Box>

        {/* --- 1. SEARCH FILTERS DIALOG POPUP --- */}
        <Dialog
          open={isSearchDialogOpen}
          onClose={() => setIsSearchDialogOpen(false)}
          fullWidth
          maxWidth="md"
        >
          <DialogTitle
            sx={{
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
              py: 1.5,
              px: 2,
              borderBottom: "1px solid #e0e0e0",
            }}
          >
            <Typography variant="h6" sx={{ fontSize: "16px", fontWeight: 600 }}>
              Search Filters
            </Typography>
            <Box sx={{ display: "flex", alignItems: "center", gap: 1 }}>
              {isSearchFormValid && (
                <Button
                  size="small"
                  color="inherit"
                  onClick={handleClearDialogFilters}
                  startIcon={<FilterListOffIcon fontSize="small" />}
                  sx={{ textTransform: "none", fontSize: "12px", color: "text.secondary" }}
                >
                  Clear Filters
                </Button>
              )}
              <IconButton size="small" onClick={() => setIsSearchDialogOpen(false)}>
                <CloseIcon fontSize="small" />
              </IconButton>
            </Box>
          </DialogTitle>
          <DialogContent sx={{ p: 2.5, pt: 2.5, mt: 1 }}>
            {/* Toggle: Apply Column Filters */}
            <Box
              sx={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                mb: 1.5,
                pb: 1,
                borderBottom: "1px solid #f0f0f0",
              }}
            >
              <FormControlLabel
                control={
                  <Switch
                    size="small"
                    checked={applyColumnFilters}
                    onChange={(e) => handleToggleApplyColumnFilters(e.target.checked)}
                    color="primary"
                  />
                }
                label={
                  <Typography
                    variant="body2"
                    sx={{
                      fontSize: "13px",
                      fontWeight: 500,
                      color: "text.primary",
                      userSelect: "none",
                    }}
                  >
                    Apply Column filters
                  </Typography>
                }
              />
              {activeColumnFilters.length > 0 && (
                <Typography variant="caption" sx={{ fontSize: "11px", color: "text.secondary" }}>
                  {activeColumnFilters.length} column filter{activeColumnFilters.length > 1 ? "s" : ""} active in grid
                </Typography>
              )}
            </Box>

            <div
              className="filter-row-grid"
              style={{
                display: "grid",
                gridTemplateColumns: "repeat(2, minmax(220px, 1fr))",
                gap: "12px",
                marginBottom: "12px",
              }}
            >
              <Autocomplete
                multiple
                fullWidth
                size="small"
                loading={lookupLoading}
                options={products}
                value={selectedProducts}
                ListboxComponent={ProductListbox}
                getOptionLabel={(option: IProductLookupItem) =>
                  option
                    ? `${option.Title ?? ""} ${option.PIMProductName ?? ""}`
                    : ""
                }
                isOptionEqualToValue={(option, value) =>
                  option.ID === value.ID ||
                  (Boolean(option.Title) &&
                    Boolean(value.Title) &&
                    `${option.Title} ${option.PIMProductName ?? ""}`.trim().toLowerCase() ===
                      `${value.Title} ${value.PIMProductName ?? ""}`.trim().toLowerCase())
                }
                onInputChange={handleProductInputChange}
                onChange={handleProductSelectionChange}
                renderOption={(props, option) => (
                  <li {...props} key={option.ID}>
                    <Box
                      sx={{
                        display: "grid",
                        gridTemplateColumns: "2.2fr 1.3fr 1.2fr",
                        gap: 0.5,
                        alignItems: "center",
                        minWidth: "500px",
                        width: "100%",
                        px: 1,
                        py: 0.25,
                        fontSize: "11px",
                      }}
                    >
                      <Typography
                        variant="caption"
                        sx={{
                          fontSize: "10.5px",
                          fontWeight: 600,
                          whiteSpace: "normal",
                          overflow: "visible",
                          textFail: "clip",
                          wordBreak: "break-word",
                        }}
                      >
                        {`${option.Title || ""} ${option.PIMProductName || ""}`.replace(/\|\s*$/g, "").trim() || "-"}
                      </Typography>
                      <Typography
                        variant="caption"
                        sx={{
                          fontSize: "10.5px",
                          whiteSpace: "nowrap",
                          overflow: "hidden",
                          textOverflow: "ellipsis",
                        }}
                      >
                        {option.Manufacturer || "-"}
                      </Typography>
                      <Typography
                        variant="caption"
                        sx={{
                          fontSize: "10.5px",
                          whiteSpace: "nowrap",
                          overflow: "hidden",
                          textOverflow: "ellipsis",
                        }}
                      >
                        {option.BusinessLine || "-"}
                      </Typography>
                    </Box>
                  </li>
                )}
                renderInput={(params) => (
                  <TextField
                    {...params}
                    size="small"
                    label="Product"
                    placeholder="Search product"
                    InputProps={{
                      ...params.InputProps,
                      endAdornment: (
                        <>
                          {lookupLoading ? (
                            <CircularProgress color="inherit" size={20} />
                          ) : null}
                          {params.InputProps.endAdornment}
                        </>
                      ),
                    }}
                  />
                )}
              />

              <Autocomplete
                multiple
                fullWidth
                size="small"
                loading={lookupLoading}
                options={clients}
                value={selectedClients}
                getOptionLabel={(option: IClientLookupItem) =>
                  option ? option.Title : ""
                }
                isOptionEqualToValue={(option, value) =>
                  option.ID === value.ID ||
                  (Boolean(option.Title) &&
                    Boolean(value.Title) &&
                    option.Title.trim().toLowerCase() === value.Title.trim().toLowerCase())
                }
                onInputChange={handleClientInputChange}
                onChange={handleClientSelectionChange}
                renderInput={(params) => (
                  <TextField
                    {...params}
                    size="small"
                    label="Client"
                    placeholder="Search client"
                    InputProps={{
                      ...params.InputProps,
                      endAdornment: (
                        <>
                          {lookupLoading ? (
                            <CircularProgress color="inherit" size={20} />
                          ) : null}
                          {params.InputProps.endAdornment}
                        </>
                      ),
                    }}
                  />
                )}
              />
            </div>

            <div
              className="filter-row-grid"
              style={{
                display: "grid",
                gridTemplateColumns: "repeat(2, minmax(220px, 1fr))",
                gap: "12px",
                marginBottom: "12px",
              }}
            >
              <FormControl fullWidth size="small" disabled={taxonomyLoading}>
                <InputLabel id="document-type-select-label">Document Type</InputLabel>
                <Select
                  labelId="document-type-select-label"
                  multiple
                  value={selectedDocumentTypes}
                  onChange={(event) => {
                    const val = event.target.value;
                    setSelectedDocumentTypes(
                      typeof val === "string" ? val.split(",") : val
                    );
                  }}
                  input={<OutlinedInput label="Document Type" />}
                  renderValue={(selected) => (selected as string[]).join(", ")}
                  size="small"
                >
                  {documentTypes.map((item) => (
                    <MenuItem key={item.ID} value={item.Title}>
                      <Checkbox
                        size="small"
                        checked={selectedDocumentTypes.indexOf(item.Title) > -1}
                      />
                      <ListItemText
                        primary={
                          item.ShortTitle
                            ? `${item.Title} (${item.ShortTitle})`
                            : item.Title
                        }
                      />
                    </MenuItem>
                  ))}
                </Select>
              </FormControl>

              <Autocomplete
                multiple
                fullWidth
                size="small"
                disabled={taxonomyLoading || availableSubDocumentTypes.length === 0}
                options={availableSubDocumentTypes.map((item) => item.Title)}
                value={selectedSubDocumentTypes}
                onChange={(_e, newValue) => {
                  setSelectedSubDocumentTypes(newValue);
                }}
                renderInput={(params) => (
                  <TextField
                    {...params}
                    size="small"
                    label="Sub Document Type"
                    placeholder={
                      selectedSubDocumentTypes.length === 0
                        ? "Select/type sub document type..."
                        : ""
                    }
                  />
                )}
              />
            </div>

            <Box sx={{ mb: 1.5, width: "100%" }}>
              {renderKeywordInput()}
            </Box>

            <LocalizationProvider dateAdapter={AdapterDayjs}>
              {(() => {
                const activeDateFiltersCount = (
                  Object.keys(dateFilters) as (keyof IDateFiltersState)[]
                ).filter((k) => dateFilters[k].from !== null || dateFilters[k].to !== null).length;

                return (
                  <Accordion
                    disableGutters
                    elevation={0}
                    defaultExpanded={false}
                    sx={{
                      mb: 1.5,
                      border: "1px solid #e0e0e0",
                      borderRadius: "6px !important",
                      bgcolor: "#fafafa",
                      "&:before": { display: "none" },
                    }}
                  >
                    <AccordionSummary
                      expandIcon={<ExpandMoreIcon sx={{ fontSize: 20 }} />}
                      sx={{
                        minHeight: "36px",
                        height: "36px",
                        px: 1.5,
                        "& .MuiAccordionSummary-content": {
                          my: 0,
                          alignItems: "center",
                          gap: 1,
                        },
                      }}
                    >
                      <Typography
                        variant="caption"
                        sx={{
                          fontWeight: 700,
                          color: "text.secondary",
                          textTransform: "uppercase",
                          fontSize: "11px",
                          letterSpacing: "0.5px",
                        }}
                      >
                        Date Filters
                      </Typography>
                      {activeDateFiltersCount > 0 && (
                        <Chip
                          size="small"
                          label={`${activeDateFiltersCount} active`}
                          sx={{
                            height: "18px",
                            fontSize: "10px",
                            fontWeight: 600,
                            bgcolor: "#e3f2fd",
                            color: "#1976d2",
                          }}
                        />
                      )}
                    </AccordionSummary>
                    <AccordionDetails
                      sx={{
                        p: 1.5,
                        pt: 0.5,
                        display: "flex",
                        flexDirection: "column",
                        gap: 1.25,
                        borderTop: "1px solid #eee",
                      }}
                    >
                      {/* 1. Document Date */}
                      <Box
                        sx={{
                          display: "grid",
                          gridTemplateColumns: "130px 1fr 1fr",
                          gap: "10px",
                          alignItems: "center",
                        }}
                      >
                        <Typography variant="body2" sx={{ fontSize: "12.5px", fontWeight: 600 }}>
                          Document Date:
                        </Typography>
                        <DatePicker
                          label="From"
                          format="DD/MM/YYYY"
                          value={dateFilters.DocumentDate.from}
                          onChange={(newValue) => handleDateChange("DocumentDate", "from", newValue)}
                          slotProps={{
                            textField: { size: "small", fullWidth: true },
                            field: { clearable: true },
                          }}
                        />
                        <DatePicker
                          label="To"
                          format="DD/MM/YYYY"
                          value={dateFilters.DocumentDate.to}
                          onChange={(newValue) => handleDateChange("DocumentDate", "to", newValue)}
                          minDate={dateFilters.DocumentDate.from ?? undefined}
                          slotProps={{
                            textField: { size: "small", fullWidth: true },
                            field: { clearable: true },
                          }}
                        />
                      </Box>

                      {/* 2. Expiry Date */}
                      <Box
                        sx={{
                          display: "grid",
                          gridTemplateColumns: "130px 1fr 1fr",
                          gap: "10px",
                          alignItems: "center",
                        }}
                      >
                        <Typography variant="body2" sx={{ fontSize: "12.5px", fontWeight: 600 }}>
                          Expiry Date:
                        </Typography>
                        <DatePicker
                          label="From"
                          format="DD/MM/YYYY"
                          value={dateFilters.ExpiryDate.from}
                          onChange={(newValue) => handleDateChange("ExpiryDate", "from", newValue)}
                          slotProps={{
                            textField: { size: "small", fullWidth: true },
                            field: { clearable: true },
                          }}
                        />
                        <DatePicker
                          label="To"
                          format="DD/MM/YYYY"
                          value={dateFilters.ExpiryDate.to}
                          onChange={(newValue) => handleDateChange("ExpiryDate", "to", newValue)}
                          minDate={dateFilters.ExpiryDate.from ?? undefined}
                          slotProps={{
                            textField: { size: "small", fullWidth: true },
                            field: { clearable: true },
                          }}
                        />
                      </Box>

                      {/* 3. Next Review Date */}
                      <Box
                        sx={{
                          display: "grid",
                          gridTemplateColumns: "130px 1fr 1fr",
                          gap: "10px",
                          alignItems: "center",
                        }}
                      >
                        <Typography variant="body2" sx={{ fontSize: "12.5px", fontWeight: 600 }}>
                          Next Review Date:
                        </Typography>
                        <DatePicker
                          label="From"
                          format="DD/MM/YYYY"
                          value={dateFilters.NextReviewDate.from}
                          onChange={(newValue) => handleDateChange("NextReviewDate", "from", newValue)}
                          slotProps={{
                            textField: { size: "small", fullWidth: true },
                            field: { clearable: true },
                          }}
                        />
                        <DatePicker
                          label="To"
                          format="DD/MM/YYYY"
                          value={dateFilters.NextReviewDate.to}
                          onChange={(newValue) => handleDateChange("NextReviewDate", "to", newValue)}
                          minDate={dateFilters.NextReviewDate.from ?? undefined}
                          slotProps={{
                            textField: { size: "small", fullWidth: true },
                            field: { clearable: true },
                          }}
                        />
                      </Box>

                      {/* 4. Created Date */}
                      <Box
                        sx={{
                          display: "grid",
                          gridTemplateColumns: "130px 1fr 1fr",
                          gap: "10px",
                          alignItems: "center",
                        }}
                      >
                        <Typography variant="body2" sx={{ fontSize: "12.5px", fontWeight: 600 }}>
                          Created Date:
                        </Typography>
                        <DatePicker
                          label="From"
                          format="DD/MM/YYYY"
                          value={dateFilters.Created.from}
                          onChange={(newValue) => handleDateChange("Created", "from", newValue)}
                          slotProps={{
                            textField: { size: "small", fullWidth: true },
                            field: { clearable: true },
                          }}
                        />
                        <DatePicker
                          label="To"
                          format="DD/MM/YYYY"
                          value={dateFilters.Created.to}
                          onChange={(newValue) => handleDateChange("Created", "to", newValue)}
                          minDate={dateFilters.Created.from ?? undefined}
                          slotProps={{
                            textField: { size: "small", fullWidth: true },
                            field: { clearable: true },
                          }}
                        />
                      </Box>

                      {/* 4. Modified Date */}
                      <Box
                        sx={{
                          display: "grid",
                          gridTemplateColumns: "130px 1fr 1fr",
                          gap: "10px",
                          alignItems: "center",
                        }}
                      >
                        <Typography variant="body2" sx={{ fontSize: "12.5px", fontWeight: 600 }}>
                          Modified Date:
                        </Typography>
                        <DatePicker
                          label="From"
                          format="DD/MM/YYYY"
                          value={dateFilters.Modified.from}
                          onChange={(newValue) => handleDateChange("Modified", "from", newValue)}
                          slotProps={{
                            textField: { size: "small", fullWidth: true },
                            field: { clearable: true },
                          }}
                        />
                        <DatePicker
                          label="To"
                          format="DD/MM/YYYY"
                          value={dateFilters.Modified.to}
                          onChange={(newValue) => handleDateChange("Modified", "to", newValue)}
                          minDate={dateFilters.Modified.from ?? undefined}
                          slotProps={{
                            textField: { size: "small", fullWidth: true },
                            field: { clearable: true },
                          }}
                        />
                      </Box>
                    </AccordionDetails>
                  </Accordion>
                );
              })()}
            </LocalizationProvider>

            {(() => {
              const activeMoreFiltersCount =
                (selectedBusinessLines.length > 0 ? 1 : 0) +
                (selectedCountries.length > 0 ? 1 : 0) +
                (selectedConfidentialities.length > 0 ? 1 : 0);

              return (
                <Accordion
                  disableGutters
                  elevation={0}
                  expanded={showMoreFilters}
                  onChange={(_e, expanded) => setShowMoreFilters(expanded)}
                  sx={{
                    mb: 1.5,
                    border: "1px solid #e0e0e0",
                    borderRadius: "6px !important",
                    bgcolor: "#fafafa",
                    "&:before": { display: "none" },
                  }}
                >
                  <AccordionSummary
                    expandIcon={<ExpandMoreIcon sx={{ fontSize: 20 }} />}
                    sx={{
                      minHeight: "36px",
                      height: "36px",
                      px: 1.5,
                      "& .MuiAccordionSummary-content": {
                        my: 0,
                        alignItems: "center",
                        gap: 1,
                      },
                    }}
                  >
                    <Typography
                      variant="caption"
                      sx={{
                        fontWeight: 700,
                        color: "text.secondary",
                        textTransform: "uppercase",
                        fontSize: "11px",
                        letterSpacing: "0.5px",
                      }}
                    >
                      More Filters
                    </Typography>
                    {activeMoreFiltersCount > 0 && (
                      <Chip
                        size="small"
                        label={`${activeMoreFiltersCount} active`}
                        sx={{
                          height: "18px",
                          fontSize: "10px",
                          fontWeight: 600,
                          bgcolor: "#e3f2fd",
                          color: "#1976d2",
                        }}
                      />
                    )}
                  </AccordionSummary>
                  <AccordionDetails
                    sx={{
                      p: 1.5,
                      pt: 1,
                      borderTop: "1px solid #eee",
                    }}
                  >
                    <div
                      className="filter-row-grid"
                      style={{
                        display: "grid",
                        gridTemplateColumns: "repeat(3, minmax(180px, 1fr))",
                        gap: "12px",
                        alignItems: "center",
                      }}
                    >
                      <FormControl fullWidth size="small">
                        <InputLabel id="business-line-filter-label">Business Line</InputLabel>
                        <Select
                          labelId="business-line-filter-label"
                          multiple
                          value={selectedBusinessLines}
                          onChange={(event) => {
                            const val = event.target.value;
                            setSelectedBusinessLines(
                              typeof val === "string" ? val.split(",") : val
                            );
                          }}
                          input={<OutlinedInput label="Business Line" />}
                          renderValue={(selected) => (selected as string[]).join(", ")}
                          size="small"
                        >
                          {availableBusinessLines.map((opt) => (
                            <MenuItem key={opt} value={opt}>
                              <Checkbox
                                size="small"
                                checked={selectedBusinessLines.indexOf(opt) > -1}
                              />
                              <ListItemText primary={opt} />
                            </MenuItem>
                          ))}
                        </Select>
                      </FormControl>

                      <FormControl fullWidth size="small">
                        <InputLabel id="country-sold-to-filter-label">Country Sold To</InputLabel>
                        <Select
                          labelId="country-sold-to-filter-label"
                          multiple
                          value={selectedCountries}
                          onChange={(event) => {
                            const val = event.target.value;
                            setSelectedCountries(
                              typeof val === "string" ? val.split(",") : val
                            );
                          }}
                          input={<OutlinedInput label="Country Sold To" />}
                          renderValue={(selected) => (selected as string[]).join(", ")}
                          size="small"
                        >
                          {availableCountries.map((opt) => (
                            <MenuItem key={opt} value={opt}>
                              <Checkbox
                                size="small"
                                checked={selectedCountries.indexOf(opt) > -1}
                              />
                              <ListItemText primary={opt} />
                            </MenuItem>
                          ))}
                        </Select>
                      </FormControl>

                      <FormControl fullWidth size="small">
                        <InputLabel id="confidentiality-filter-label">Confidentiality</InputLabel>
                        <Select
                          labelId="confidentiality-filter-label"
                          multiple
                          value={selectedConfidentialities}
                          onChange={(event) => {
                            const val = event.target.value;
                            setSelectedConfidentialities(
                              typeof val === "string" ? val.split(",") : val
                            );
                          }}
                          input={<OutlinedInput label="Confidentiality" />}
                          renderValue={(selected) => (selected as string[]).join(", ")}
                          size="small"
                        >
                          {availableConfidentialities.map((opt) => (
                            <MenuItem key={opt} value={opt}>
                              <Checkbox
                                size="small"
                                checked={selectedConfidentialities.indexOf(opt) > -1}
                              />
                              <ListItemText primary={opt} />
                            </MenuItem>
                          ))}
                        </Select>
                      </FormControl>
                    </div>
                  </AccordionDetails>
                </Accordion>
              );
            })()}
          </DialogContent>
          <DialogActions sx={{ px: 2.5, pb: 2, pt: 1, borderTop: "1px solid #e0e0e0" }}>
            <Button
              onClick={handleClearDialogFilters}
              size="small"
              color="inherit"
              disabled={!isSearchFormValid}
              startIcon={<FilterListOffIcon fontSize="small" />}
              sx={{
                mr: "auto",
                textTransform: "none",
                fontWeight: 500,
                fontSize: "13px",
                color: isSearchFormValid ? "text.secondary" : "text.disabled",
              }}
            >
              Clear Filters
            </Button>
            <Button onClick={() => setIsSearchDialogOpen(false)} size="small" sx={{ textTransform: "none" }}>
              Cancel
            </Button>
            <Button
              variant="contained"
              color="primary"
              size="small"
              onClick={handleSearch}
              disabled={!isSearchFormValid || resultsLoading}
              sx={{ textTransform: "none", fontWeight: 600 }}
            >
              {resultsLoading ? "Searching..." : "Search"}
            </Button>
          </DialogActions>
        </Dialog>

        {/* --- 2. RESULTS TABLE --- */}
        <Box sx={{ width: "100%", maxWidth: "100%" }}>
          <LocalizationProvider dateAdapter={AdapterDayjs}>
            <MaterialReactTable table={table} />
          </LocalizationProvider>
        </Box>

        <EmailShareDialog
          open={isShareDialogOpen}
          onClose={() => setIsShareDialogOpen(false)}
          selectedItems={selectedRowsData}
          siteUrl={props.urlSite}
          defaultSubject={sharingConfig.subject}
          defaultMessage={sharingConfig.message}
          currentUserEmail={props.context?.pageContext?.user?.email || ""}
        />

        {/* File Actions Popup Menu (View / Edit) */}
        <Menu
          anchorEl={fileMenuAnchorEl}
          open={Boolean(fileMenuAnchorEl)}
          onClose={() => setFileMenuAnchorEl(null)}
          anchorOrigin={{ vertical: "bottom", horizontal: "left" }}
          transformOrigin={{ vertical: "top", horizontal: "left" }}
        >
          <MenuItem
            onClick={() => {
              if (selectedFileForAction) {
                const origin = window.location.origin;
                const fileUrl = selectedFileForAction.fileUrl?.startsWith("http")
                  ? selectedFileForAction.fileUrl
                  : `${origin}${selectedFileForAction.fileUrl || ""}`;
                window.open(fileUrl, "_blank", "noopener,noreferrer");
              }
              setFileMenuAnchorEl(null);
            }}
          >
            View
          </MenuItem>
          <MenuItem
            onClick={() => {
              if (selectedFileForAction) {
                const baseUrl = props.urlSite ? props.urlSite.replace(/\/$/, "") : window.location.origin;
                const editUrl = `${baseUrl}/Products/Forms/EditForm.aspx?ID=${selectedFileForAction.id || ""}`;
                setEditModalUrl(editUrl);
              }
              setFileMenuAnchorEl(null);
            }}
          >
            Edit
          </MenuItem>
        </Menu>

        {/* In-page Modal Dialog for OOB Edit Form */}
        <Dialog
          open={Boolean(editModalUrl)}
          onClose={() => setEditModalUrl(null)}
          fullWidth
          maxWidth="md"
          PaperProps={{
            sx: {
              height: "85vh",
              maxHeight: "850px",
              display: "flex",
              flexDirection: "column",
            },
          }}
        >
          <DialogTitle
            sx={{
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
              py: 1.5,
              px: 2,
              borderBottom: "1px solid #e0e0e0",
            }}
          >
            <Typography variant="h6" sx={{ fontSize: "16px", fontWeight: 600 }}>
              Edit Properties - {selectedFileForAction?.filename || "Document"}
            </Typography>
            <IconButton
              size="small"
              onClick={() => setEditModalUrl(null)}
            >
              <CloseIcon fontSize="small" />
            </IconButton>
          </DialogTitle>
          <DialogContent
            sx={{
              p: 0,
              flex: 1,
              overflow: "auto",
              display: "flex",
              flexDirection: "column",
            }}
          >
            {editModalUrl && (
              <iframe
                src={editModalUrl}
                title="Edit Document Form"
                style={{
                  width: "100%",
                  height: "100%",
                  border: "none",
                  flex: 1,
                  display: "block",
                }}
              />
            )}
          </DialogContent>
        </Dialog>

        {/* Modal Dialog to Name & Save Custom View */}
        <Dialog
          open={isSaveViewDialogOpen}
          onClose={() => setIsSaveViewDialogOpen(false)}
          maxWidth="xs"
          fullWidth
        >
          <DialogTitle sx={{ fontSize: "15px", fontWeight: 600, py: 1.5, px: 2 }}>
            Save Current View
          </DialogTitle>
          <DialogContent sx={{ px: 2, pt: 1 }}>
            <Typography variant="body2" sx={{ fontSize: "12px", color: "text.secondary", mb: 1.5 }}>
              Saves the current column order, visibility, grouping, and sorting.
            </Typography>
            <TextField
              autoFocus
              fullWidth
              size="small"
              label="View Name"
              placeholder="e.g. My Custom View"
              value={newViewName}
              onChange={(e) => setNewViewName(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter" && newViewName.trim()) {
                  handleSaveCurrentView();
                }
              }}
            />
          </DialogContent>
          <DialogActions sx={{ px: 2, pb: 1.5 }}>
            <Button
              onClick={() => setIsSaveViewDialogOpen(false)}
              size="small"
              sx={{ textTransform: "none", fontSize: "12px" }}
            >
              Cancel
            </Button>
            <Button
              variant="contained"
              size="small"
              disabled={!newViewName.trim()}
              onClick={handleSaveCurrentView}
              sx={{ textTransform: "none", fontSize: "12px", fontWeight: 600 }}
            >
              Save View
            </Button>
          </DialogActions>
        </Dialog>
      </Box>
    </ThemeProvider>
  );
};

export default AdvanceSearch;