import * as React from "react";
import { useMemo } from "react";
import { IAdvanceSearchProps } from "./IAdvanceSearchProps";
import { sp } from "../AdvanceSearchWebPart";
import { spfi, SPFx, SPFI } from "@pnp/sp";
import {
  MaterialReactTable,
  useMaterialReactTable,
  type MRT_ColumnDef,
  type MRT_ColumnFiltersState,
  type MRT_GroupingState,
} from "material-react-table";

import Autocomplete from "@mui/material/Autocomplete";
import TextField from "@mui/material/TextField";
import Chip from "@mui/material/Chip";
import CircularProgress from "@mui/material/CircularProgress";
import Button from "@mui/material/Button";
import IconButton from "@mui/material/IconButton";
import ShareIcon from "@mui/icons-material/Share";
import Dialog from "@mui/material/Dialog";
import DialogTitle from "@mui/material/DialogTitle";
import DialogContent from "@mui/material/DialogContent";
import DialogActions from "@mui/material/DialogActions";
import FormControl from "@mui/material/FormControl";
import InputLabel from "@mui/material/InputLabel";
import MenuItem from "@mui/material/MenuItem";
import Menu from "@mui/material/Menu";
import Select from "@mui/material/Select";
import Checkbox from "@mui/material/Checkbox";
import ListItemText from "@mui/material/ListItemText";
import OutlinedInput from "@mui/material/OutlinedInput";
import Box from "@mui/material/Box";
import Typography from "@mui/material/Typography";
import Link from "@mui/material/Link";
import CloseIcon from "@mui/icons-material/Close";
import RefreshIcon from "@mui/icons-material/Refresh";
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
} from "../utils/formatters";
import {
  multiSelectFilterFn,
  documentDateFilter,
  itemMatchesFilter,
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

const DocumentDateFilter: React.FC<{
  column: { getFilterValue: () => unknown; setFilterValue: (value: unknown) => void };
}> = ({ column }) => {
  const filterValue = column.getFilterValue() as string | null;
  const pickerValue = filterValue ? dayjs(filterValue) : null;

  return (
    <DatePicker
      format="DD/MM/YYYY"
      label="Min Date"
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
  const [dateFrom, setDateFrom] = React.useState<Dayjs | null>(null);
  const [dateTo, setDateTo] = React.useState<Dayjs | null>(null);

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

  const [fieldFormatters, setFieldFormatters] = React.useState<IFieldFormatters>({
    businessLine: {},
    confidentiality: {},
  });

  const [columnFilters, setColumnFilters] = React.useState<MRT_ColumnFiltersState>([]);
  const [grouping, setGrouping] = React.useState<MRT_GroupingState>([]);

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

  // Dynamic table container height to prevent double vertical scrollbars
  const [tableMaxHeight, setTableMaxHeight] = React.useState<string>("calc(100vh - 210px)");
  const tableContainerRef = React.useRef<HTMLDivElement | null>(null);

  React.useEffect(() => {
    const updateHeight = (): void => {
      if (tableContainerRef.current) {
        const rect = tableContainerRef.current.getBoundingClientRect();
        // Reserve space for MRT bottom pagination toolbar (~52px) and bottom padding (~16px)
        const availableHeight = window.innerHeight - rect.top - 68;
        setTableMaxHeight(`${Math.max(300, Math.floor(availableHeight))}px`);
      }
    };

    updateHeight();
    window.addEventListener("resize", updateHeight);
    const timer = setTimeout(updateHeight, 300);

    return () => {
      window.removeEventListener("resize", updateHeight);
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

    const allKeywords = [...additionalKeywords];
    if (keywordInput.trim()) {
      keywordInput
        .split(",")
        .map((k) => k.trim())
        .filter(Boolean)
        .forEach((k) => {
          if (!allKeywords.includes(k)) allKeywords.push(k);
        });
    }

    if (allKeywords.length === 1) {
      clauses.push(`"${sanitizeKqlValue(allKeywords[0])}"`);
    } else if (allKeywords.length > 1) {
      clauses.push(
        `(${allKeywords.map((k) => `"${sanitizeKqlValue(k)}"`).join(" AND ")})`
      );
    }

    if (dateFrom && dateTo) {
      clauses.push(
        `DocumentDateOWSTDATE:${dateFrom
          .startOf("day")
          .toISOString()}..${dateTo.endOf("day").toISOString()}`
      );
    } else if (dateFrom) {
      clauses.push(
        `DocumentDateOWSTDATE>=${dateFrom.startOf("day").toISOString()}`
      );
    } else if (dateTo) {
      clauses.push(
        `DocumentDateOWSTDATE<=${dateTo.endOf("day").toISOString()}`
      );
    }

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

  const handleOpenSearchDialog = (): void => {
    // Populate search dialog controls with active column filter values if dialog controls are currently empty
    const clientColFilter = columnFilters.find((f) => f.id === "ManufacturerSearchText")?.value;
    if (clientColFilter) {
      const vals = (Array.isArray(clientColFilter) ? clientColFilter : [clientColFilter])
        .map((v) => String(v).trim())
        .filter((v) => v && v.toLowerCase() !== "(empty)" && v !== "-");

      if (vals.length > 0) {
        setSelectedClients((prev) => {
          const existingTitles = new Set(prev.map((c) => (c.Title || "").toLowerCase()));
          const newItems: IClientLookupItem[] = [];
          vals.forEach((v) => {
            if (!existingTitles.has(v.toLowerCase())) {
              const fromLoaded = clients.find((c) => (c.Title || "").toLowerCase() === v.toLowerCase());
              newItems.push(fromLoaded || { ID: -Math.floor(Math.random() * 100000), Title: v });
              existingTitles.add(v.toLowerCase());
            }
          });
          return newItems.length > 0 ? [...prev, ...newItems] : prev;
        });
      }
    }

    const productColFilter = columnFilters.find((f) => f.id === "PIMProductSearchText")?.value;
    if (productColFilter) {
      const vals = (Array.isArray(productColFilter) ? productColFilter : [productColFilter])
        .map((v) => String(v).trim())
        .filter((v) => v && v.toLowerCase() !== "(empty)" && v !== "-");

      if (vals.length > 0) {
        // Collect all distinct product objects across loaded items
        const allKnownProducts: IProductLookupItem[] = [];
        const seenProductKeys = new Set<string>();

        const registerProduct = (item: IProductLookupItem) => {
          const key = `${item.Title || ""} ${item.PIMProductName || ""}`.trim().toLowerCase();
          if (key && !seenProductKeys.has(key)) {
            seenProductKeys.add(key);
            allKnownProducts.push(item);
          }
        };

        // 1. From loaded records in the grid
        items_AllProducts.forEach((rec) => {
          if (Array.isArray(rec.PIMProduct)) {
            rec.PIMProduct.forEach(registerProduct);
          }
        });

        // 2. From searchProducts lookup cache
        products.forEach(registerProduct);

        setSelectedProducts((prev) => {
          const existingKeys = new Set(
            prev.map((p) => `${p.Title || ""} ${p.PIMProductName || ""}`.trim().toLowerCase())
          );
          const newItems: IProductLookupItem[] = [];

          vals.forEach((v) => {
            const vLower = v.toLowerCase();

            // Match against known product items from grid or lookup
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
              if (!existingKeys.has(matchedKey)) {
                newItems.push(matched);
                existingKeys.add(matchedKey);
              }
            } else {
              // If not found in known products, parse "Title | ProductName" or "Code ProductName"
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
              if (!existingKeys.has(itemKey)) {
                newItems.push({
                  ID: -Math.floor(Math.random() * 100000),
                  Title: parsedTitle,
                  PIMProductName: parsedName,
                });
                existingKeys.add(itemKey);
              }
            }
          });

          return newItems.length > 0 ? [...prev, ...newItems] : prev;
        });
      }
    }

    const docTypeColFilter = columnFilters.find((f) => f.id === "DocumentTypeSearchText")?.value;
    if (docTypeColFilter) {
      const vals = (Array.isArray(docTypeColFilter) ? docTypeColFilter : [docTypeColFilter])
        .map((v) => String(v).trim())
        .filter((v) => v && v.toLowerCase() !== "(empty)" && v !== "-");

      if (vals.length > 0) {
        setSelectedDocumentTypes((prev) => Array.from(new Set([...prev, ...vals])));
      }
    }

    const subDocTypeColFilter = columnFilters.find((f) => f.id === "SubDocumentTypeSearchText")?.value;
    if (subDocTypeColFilter) {
      const vals = (Array.isArray(subDocTypeColFilter) ? subDocTypeColFilter : [subDocTypeColFilter])
        .map((v) => String(v).trim())
        .filter((v) => v && v.toLowerCase() !== "(empty)" && v !== "-");

      if (vals.length > 0) {
        setSelectedSubDocumentTypes((prev) => Array.from(new Set([...prev, ...vals])));
      }
    }

    const docDateColFilter = columnFilters.find((f) => f.id === "DocumentDate")?.value;
    if (docDateColFilter) {
      const parsedDate = dayjs(docDateColFilter as string | Date);
      if (parsedDate.isValid()) {
        setDateFrom(parsedDate);
      }
    }

    setIsSearchDialogOpen(true);
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
    const hasDate = dateFrom !== null || dateTo !== null;
    const hasKeyword = additionalKeywords.length > 0 || Boolean(keywordInput.trim());
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
    dateFrom,
    dateTo,
    additionalKeywords,
    keywordInput,
    selectedBusinessLines,
    selectedCountries,
    selectedConfidentialities,
  ]);

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
          f.value !== undefined &&
          f.value !== null &&
          f.value !== "" &&
          !(Array.isArray(f.value) && f.value.length === 0)
      );
      if (activeFilters.length === 0) {
        return items_AllProducts;
      }
      return items_AllProducts.filter((item) =>
        activeFilters.every((f) => itemMatchesFilter(item, f.id, f.value))
      );
    },
    [items_AllProducts, columnFilters]
  );

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

  const columns_AllProducts = useMemo<MRT_ColumnDef<doclib_AllProducts>[]>(
    () => [
      {
        accessorKey: "filename",
        header: "File Name",
        size: 140,
        minSize: 140,
        filterFn: "contains",
        Cell: ({ row }) => {
          const displayName =
            row.original.filename.length > 20
              ? `${row.original.filename.slice(0, 17)}...`
              : row.original.filename;

          return (
            <span
              onClick={(e) => {
                e.preventDefault();
                e.stopPropagation();
                setSelectedFileForAction(row.original);
                setFileMenuAnchorEl(e.currentTarget);
              }}
              title={row.original.filename}
              style={{
                display: "inline-block",
                maxWidth: "100%",
                overflow: "hidden",
                textOverflow: "ellipsis",
                whiteSpace: "nowrap",
                color: "#1976d2",
                textDecoration: "underline",
                cursor: "pointer",
              }}
            >
              {displayName}
            </span>
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
        accessorKey: "Alerts",
        header: "Alerts",
        size: 160,
        minSize: 140,
        filterFn: "contains",
        Cell: ({ cell }) => {
          const raw = String(cell.getValue() || "").trim();
          if (!raw) return "-";
          return (
            <div
              style={{
                whiteSpace: "pre-wrap",
                wordBreak: "break-word",
                lineHeight: 1.4,
              }}
            >
              {raw}
            </div>
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
          <DocumentDateFilter
            column={column}
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
      fieldFormatters,
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
        'button, input, textarea, select, [role="button"], [role="checkbox"], .MuiInputBase-root, .MuiIconButton-root, .MuiSelect-select, .MuiTableSortLabel-root, .Mui-TableHeadCell-ResizeHandle, .Mui-TableHeadCell-GrabHandle'
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
      <Box sx={{ display: "flex", gap: "12px", alignItems: "center" }}>
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
        </Box>
      </Box>
    ),
    enableGrouping: true,
    enableColumnDragging: true,
    enableColumnResizing: true,
    columnResizeMode: "onChange",
    layoutMode: "grid-no-grow",
    initialState: {
      density: "compact",
    },
    muiTablePaperProps: {
      sx: {
        boxShadow: "none",
        border: "1px solid #e1dfdd",
        width: "100%",
        maxWidth: "100%",
      },
    },
    muiTableContainerProps: {
      ref: tableContainerRef,
      onMouseDown: handleTableMouseDown,
      sx: {
        maxHeight: tableMaxHeight,
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
        fontSize: "13px",
        padding: "10px 12px",
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
        fontSize: "13px",
        padding: "9px 12px",
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
    onColumnFiltersChange: setColumnFilters,
    onGroupingChange: setGrouping,
    state: {
      isLoading: resultsLoading,
      showProgressBars: isLoadingMore,
      showColumnFilters: true,
      columnFilters,
      grouping,
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
        value.map((option: string, index: number) => (
          <Chip
            {...getTagProps({ index })}
            key={index}
            label={option}
            size="small"
            sx={{
              height: "24px",
              fontSize: "12px",
              margin: "2px",
            }}
          />
        ))
      }
      renderInput={(params) => (
        <TextField
          {...params}
          size="small"
          label="Additional Keywords"
          placeholder={
            additionalKeywords.length === 0
              ? "Type keyword & press comma or Enter"
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
            <IconButton size="small" onClick={() => setIsSearchDialogOpen(false)}>
              <CloseIcon fontSize="small" />
            </IconButton>
          </DialogTitle>
          <DialogContent sx={{ p: 2.5, pt: 2.5, mt: 1 }}>
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
                isOptionEqualToValue={(option, value) => option.ID === value.ID}
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
                isOptionEqualToValue={(option, value) => option.ID === value.ID}
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

            <LocalizationProvider dateAdapter={AdapterDayjs}>
              <div
                className="filter-row-grid"
                style={{
                  display: "grid",
                  gridTemplateColumns: "repeat(2, minmax(220px, 1fr))",
                  gap: "12px",
                  marginBottom: "12px",
                }}
              >
                <DatePicker
                  label="Document Date From"
                  format="DD/MM/YYYY"
                  value={dateFrom}
                  onChange={(newValue) => setDateFrom(newValue)}
                  slotProps={{
                    textField: { size: "small", fullWidth: true },
                    field: { clearable: true },
                  }}
                />
                <DatePicker
                  label="Document Date To"
                  format="DD/MM/YYYY"
                  value={dateTo}
                  onChange={(newValue) => setDateTo(newValue)}
                  minDate={dateFrom ?? undefined}
                  slotProps={{
                    textField: { size: "small", fullWidth: true },
                    field: { clearable: true },
                  }}
                />
              </div>
            </LocalizationProvider>

            {!showMoreFilters ? (
              <div
                className="filter-row-grid"
                style={{
                  display: "grid",
                  gridTemplateColumns: "repeat(2, minmax(220px, 1fr))",
                  gap: "12px",
                  marginBottom: "12px",
                  alignItems: "center",
                }}
              >
                {renderKeywordInput()}
                <Button
                  variant="outlined"
                  fullWidth
                  size="small"
                  onClick={() => setShowMoreFilters(true)}
                  sx={{
                    height: "40px",
                    textTransform: "none",
                    fontWeight: 600,
                    fontSize: "13px",
                  }}
                >
                  More Filters
                </Button>
              </div>
            ) : (
              <>
                <div
                  style={{
                    marginBottom: "12px",
                    width: "100%",
                  }}
                >
                  {renderKeywordInput()}
                </div>

                <div
                  className="filter-row-grid"
                  style={{
                    display: "grid",
                    gridTemplateColumns: "repeat(2, minmax(220px, 1fr))",
                    gap: "12px",
                    marginBottom: "12px",
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

                  <Button
                    variant="outlined"
                    fullWidth
                    size="small"
                    onClick={() => setShowMoreFilters(false)}
                    sx={{
                      height: "40px",
                      textTransform: "none",
                      fontWeight: 600,
                      fontSize: "13px",
                    }}
                  >
                    Less Filters
                  </Button>
                </div>
              </>
            )}

            <style>{`
            .CanvasZone,
            [data-automation-id="CanvasZone"],
            .CanvasSection,
            .ControlZone,
            [data-automation-id="CanvasControl"] {
              max-width: none !important;
            }
            @media (max-width: 768px) {
              .filter-row-grid {
                grid-template-columns: 1fr !important;
              }
            }
          `}</style>
          </DialogContent>
          <DialogActions sx={{ px: 2.5, pb: 2, pt: 1, borderTop: "1px solid #e0e0e0" }}>
            <Button onClick={() => setIsSearchDialogOpen(false)} size="small">
              Cancel
            </Button>
            <Button
              variant="contained"
              color="primary"
              size="small"
              onClick={handleSearch}
              disabled={!isSearchFormValid || resultsLoading}
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
            <IconButton size="small" onClick={() => setEditModalUrl(null)}>
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
      </Box>
    </ThemeProvider>
  );
};

export default AdvanceSearch;