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
} from "material-react-table";

import Autocomplete from "@mui/material/Autocomplete";
import TextField from "@mui/material/TextField";
import CircularProgress from "@mui/material/CircularProgress";
import Button from "@mui/material/Button";
import IconButton from "@mui/material/IconButton";
import ShareIcon from "@mui/icons-material/Share";
import Dialog from "@mui/material/Dialog";
import DialogTitle from "@mui/material/DialogTitle";
import DialogContent from "@mui/material/DialogContent";
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
import CloseIcon from "@mui/icons-material/Close";
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
  loadListFieldFormatting as loadFieldFormattingService,
  loadRecordsBatch as loadRecordsBatchService,
  searchRecords as searchRecordsService,
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
  //const [allRecordsCache, setAllRecordsCache] =
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
  const [additionalKeyword, setAdditionalKeyword] = React.useState("");
  const [dateFrom, setDateFrom] = React.useState<Dayjs | null>(null);
  const [dateTo, setDateTo] = React.useState<Dayjs | null>(null);

  const [lookupLoading, setLookupLoading] = React.useState(false);
  const [resultsLoading, setResultsLoading] = React.useState(true);
  const [hasSearched, setHasSearched] = React.useState(true);
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

    const loadFieldFormatting = async (): Promise<void> => {
      try {
        const formatters = await loadFieldFormattingService(activeSp);
        setFieldFormatters(formatters);
      } catch (err) {
        console.warn("loadListFieldFormatting error:", err);
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
    void loadFieldFormatting();
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

    const productValues = selectedProducts
      .map((item) => (item.PIMProductName || item.Title || "").trim())
      .filter(Boolean);

    const clientValues = selectedClients
      .map((item) => (item.Title || "").trim())
      .filter(Boolean);

    const documentTypeValues = selectedDocumentTypes
      .map((item) => item.trim())
      .filter(Boolean);

    const subDocumentTypeValues = selectedSubDocumentTypes
      .map((item) => item.trim())
      .filter(Boolean);

    if (productValues.length > 0) {
      clauses.push(
        `(${productValues
          .map((value) => `"${sanitizeKqlValue(value)}"`)
          .join(" OR ")})`
      );
    }

    if (clientValues.length > 0) {
      clauses.push(
        `(${clientValues
          .map((value) => `"${sanitizeKqlValue(value)}"`)
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

    if (additionalKeyword.trim()) {
      clauses.push(`"${sanitizeKqlValue(additionalKeyword)}"`);
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
      setHasSearched(true);
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

  const handleSearch = async () => {
    const selectedProductValues = selectedProducts
      .map((item) => (item.PIMProductName || item.Title || "").trim())
      .filter(Boolean);

    const selectedClientValues = selectedClients
      .map((item) => (item.Title || "").trim())
      .filter(Boolean);

    const hasAnySelection =
      selectedProductValues.length > 0 ||
      selectedClientValues.length > 0 ||
      selectedDocumentTypes.length > 0 ||
      selectedSubDocumentTypes.length > 0 ||
      dateFrom !== null ||
      dateTo !== null;

    if (!hasAnySelection) {
      await handleLoadAllRecords();
      return;
    }

    try {
      setResultsLoading(true);
      setHasSearched(true);
      setIsBrowseMode(false);
      setColumnFilters([]);
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
        (f) => f.id !== excludedColId && f.value !== undefined && f.value !== null && f.value !== ""
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
      items.forEach((item) => {
        if (item.ManufacturerSearchText) {
          item.ManufacturerSearchText.split(";").forEach((val) => {
            const trimmed = val.trim();
            if (trimmed) unique.add(trimmed);
          });
        }
      });
      const result: string[] = [];
      unique.forEach((val) => result.push(val));
      return result.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
    },
    [getItemsFilteredExcluding]
  );

  const productOptions = useMemo(
    () => {
      const items = getItemsFilteredExcluding("PIMProductSearchText");
      const unique = new Set<string>();
      items.forEach((item) => {
        if (item.PIMProduct && Array.isArray(item.PIMProduct)) {
          item.PIMProduct.forEach((p) => {
            const name = `${p.Title || ""} ${p.PIMProductName || ""}`.trim();
            if (name) {
              unique.add(name);
            }
          });
        }
      });
      const result: string[] = [];
      unique.forEach((val) => result.push(val));
      return result.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
    },
    [getItemsFilteredExcluding]
  );

  const businessLineOptions = useMemo(
    () => {
      const items = getItemsFilteredExcluding("BusinessLine");
      const unique = new Set<string>();
      items.forEach((item) => {
        if (item.BusinessLine) {
          item.BusinessLine.split(",").forEach((val) => {
            const trimmed = val.trim();
            if (trimmed) unique.add(trimmed);
          });
        }
      });
      const result: string[] = [];
      unique.forEach((val) => result.push(val));
      return result.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
    },
    [getItemsFilteredExcluding]
  );

  const countryOptions = useMemo(
    () => {
      const items = getItemsFilteredExcluding("CountrySoldTo");
      const unique = new Set<string>();
      items.forEach((item) => {
        if (item.CountrySoldTo) {
          item.CountrySoldTo.split(",").forEach((val) => {
            const trimmed = val.trim();
            if (trimmed) unique.add(trimmed);
          });
        }
      });
      const result: string[] = [];
      unique.forEach((val) => result.push(val));
      return result.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
    },
    [getItemsFilteredExcluding]
  );

  const documentTypeOptions = useMemo(
    () => {
      const items = getItemsFilteredExcluding("DocumentTypeSearchText");
      const unique = new Set<string>();
      items.forEach((item) => {
        if (item.DocumentTypeSearchText) {
          const trimmed = item.DocumentTypeSearchText.trim();
          if (trimmed) unique.add(trimmed);
        }
      });
      const result: string[] = [];
      unique.forEach((val) => result.push(val));
      return result.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
    },
    [getItemsFilteredExcluding]
  );

  const subDocumentTypeOptions = useMemo(
    () => {
      const items = getItemsFilteredExcluding("SubDocumentTypeSearchText");
      const unique = new Set<string>();
      items.forEach((item) => {
        if (item.SubDocumentTypeSearchText) {
          item.SubDocumentTypeSearchText.split(",").forEach((val) => {
            const trimmed = val.trim();
            if (trimmed) unique.add(trimmed);
          });
        }
      });
      const result: string[] = [];
      unique.forEach((val) => result.push(val));
      return result.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
    },
    [getItemsFilteredExcluding]
  );

  const confidentialityOptions = useMemo(
    () => {
      const items = getItemsFilteredExcluding("Confidentiality");
      const unique = new Set<string>();
      items.forEach((item) => {
        if (item.Confidentiality) {
          item.Confidentiality.split(",").forEach((val) => {
            const trimmed = val.trim();
            if (trimmed) unique.add(trimmed);
          });
        }
      });
      const result: string[] = [];
      unique.forEach((val) => result.push(val));
      return result.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
    },
    [getItemsFilteredExcluding]
  );

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
        filterVariant: "multi-select",
        filterSelectOptions: clientOptions,
        filterFn: multiSelectFilterFn,
        size: 160,
        minSize: 160,
        Cell: ({ cell }) => (
          <>
            {String(cell.getValue() || "")
              .split(";")
              .filter(Boolean)
              .map((item, idx) => (
                <div
                  key={idx}
                  style={{
                    whiteSpace: "normal",
                    wordBreak: "normal",
                    overflowWrap: "break-word",
                    lineHeight: 1.4,
                  }}
                >
                  {item.trim()}
                </div>
              ))}
          </>
        ),
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
        Cell: ({ row }) => (
          <div>
            {row.original.PIMProduct.map((p: IProductLookupItem, idx) => (
              <div key={idx}>
                {p.Title} {p.PIMProductName}
              </div>
            ))}
          </div>
        ),
      },
      {
        accessorKey: "DocumentTypeSearchText",
        header: "Document Type",
        filterVariant: "multi-select",
        filterSelectOptions: documentTypeOptions,
        filterFn: multiSelectFilterFn,
        size: 150,
        minSize: 150,
      },
      {
        accessorKey: "SubDocumentTypeSearchText",
        header: "Sub Document Type",
        filterVariant: "multi-select",
        filterSelectOptions: subDocumentTypeOptions,
        filterFn: multiSelectFilterFn,
        size: 175,
        minSize: 175,
        Cell: ({ cell }) => (
          <div>
            {String(cell.getValue() || "")
              .split(",")
              .filter(Boolean)
              .map((item, idx) => (
                <div key={idx}>{item.trim()}</div>
              ))}
          </div>
        ),
      },
      {
        accessorKey: "BusinessLine",
        header: "Business Line",
        filterVariant: "multi-select",
        filterSelectOptions: businessLineOptions,
        filterFn: multiSelectFilterFn,
        size: 155,
        minSize: 140,
        Cell: ({ cell }) => {
          const raw = String(cell.getValue() || "");
          if (!raw) return "-";

          return (
            <div style={{ display: "flex", flexWrap: "wrap", gap: "4px" }}>
              {raw
                .split(",")
                .map((v) => v.trim())
                .filter(Boolean)
                .map((item, idx) => {
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
                })}
            </div>
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
        Cell: ({ cell }) => (
          <div>
            {String(cell.getValue() || "")
              .split(",")
              .filter(Boolean)
              .map((item, idx) => (
                <div key={idx}>{item.trim()}</div>
              ))}
          </div>
        ),
      },
      {
        accessorKey: "Confidentiality",
        header: "Confidentiality",
        filterVariant: "multi-select",
        filterSelectOptions: confidentialityOptions,
        filterFn: multiSelectFilterFn,
        size: 150,
        minSize: 140,
        Cell: ({ cell }) => {
          const raw = String(cell.getValue() || "");
          if (!raw) return "-";

          return (
            <div style={{ display: "flex", flexWrap: "wrap", gap: "4px" }}>
              {raw
                .split(",")
                .map((v) => v.trim())
                .filter(Boolean)
                .map((item, idx) => {
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
                })}
            </div>
          );
        },
      },
      {
        accessorKey: "Alerts",
        header: "Alerts",
        size: 160,
        minSize: 140,
        filterFn: "contains",
        Cell: ({ cell }) => (
          <div
            style={{
              whiteSpace: "pre-wrap",
              wordBreak: "break-word",
              lineHeight: 1.4,
            }}
          >
            {String(cell.getValue() || "")}
          </div>
        ),
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
    // Don't drag if clicking interactive controls, filter inputs, header cells, or resize handles
    if (
      target.closest(
        'button, input, textarea, select, [role="button"], [role="checkbox"], .MuiInputBase-root, .MuiIconButton-root, .MuiSelect-select, thead, th, .Mui-TableHeadCell-ResizeHandle, .Mui-TableHeadCell-GrabHandle'
      )
    ) {
      return;
    }

    const container = e.currentTarget;
    const startClientX = e.clientX;
    const initialScrollLeft = container.scrollLeft;
    let hasDragged = false;

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
          <Typography variant="caption" sx={{ fontSize: "12px", color: "text.secondary" }}>
            {items_AllProducts.length.toLocaleString()} records loaded
          </Typography>

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

          {!hasMoreRecords && items_AllProducts.length > 0 && (
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
      },
    },
    muiTableContainerProps: {
      onMouseDown: handleTableMouseDown,
      sx: {
        cursor: "grab",
        "&:active": {
          cursor: "grabbing",
        },
      },
    },
    muiTableHeadCellProps: {
      sx: {
        fontSize: "13px",
        padding: "10px 12px",
        lineHeight: 1.2,
        whiteSpace: "normal",
        wordBreak: "break-word",
      },
    },
    muiTableBodyCellProps: {
      sx: {
        fontSize: "13px",
        padding: "9px 12px",
        whiteSpace: "normal",
        wordBreak: "normal",
        overflowWrap: "break-word",
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
    state: {
      isLoading: resultsLoading,
      showProgressBars: isLoadingMore,
      showColumnFilters: true,
      columnFilters,
    },
  });

  
  return (
    <ThemeProvider theme={compactTheme}>
      <div>
        <h2>Search Clients & Products</h2>

        {/* --- 1. SEARCH FILTERS PLACEHOLDER --- */}
        {!hasSearched && (
          <Box sx={{ mb: 2 }}>
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
                    ? `${option.Title ?? ""} | ${option.PIMProductName ?? ""}`
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
                        {`${option.Title || ""} | ${option.PIMProductName || ""}`.replace(/\|\s*$/g, "").trim() || "-"}
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

              <FormControl
                fullWidth
                size="small"
                disabled={taxonomyLoading || availableSubDocumentTypes.length === 0}
              >
                <InputLabel id="sub-document-type-select-label">
                  Sub Document Type
                </InputLabel>
                <Select
                  labelId="sub-document-type-select-label"
                  multiple
                  value={selectedSubDocumentTypes}
                  onChange={(event) => {
                    const val = event.target.value;
                    setSelectedSubDocumentTypes(
                      typeof val === "string" ? val.split(",") : val
                    );
                  }}
                  input={<OutlinedInput label="Sub Document Type" />}
                  renderValue={(selected) => (selected as string[]).join(", ")}
                  size="small"
                >
                  {availableSubDocumentTypes.map((item) => (
                    <MenuItem key={item.ID} value={item.Title}>
                      <Checkbox
                        size="small"
                        checked={selectedSubDocumentTypes.indexOf(item.Title) > -1}
                      />
                      <ListItemText primary={item.Title} />
                    </MenuItem>
                  ))}
                </Select>
              </FormControl>
            </div>

            <LocalizationProvider dateAdapter={AdapterDayjs}>
              <div
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

            <div
              style={{
                marginBottom: "16px",
                width: "100%",
              }}
            >
              <TextField
                fullWidth
                size="small"
                label="Additional Keyword"
                placeholder="Optional keyword"
                value={additionalKeyword}
                onChange={(e) => setAdditionalKeyword(e.target.value)}
              />
            </div>

            <style>{`
            @media (max-width: 768px) {
              .filter-row-grid {
                grid-template-columns: 1fr !important;
              }
            }
          `}</style>

            <div style={{ marginBottom: "16px", display: "flex", gap: "12px", alignItems: "center" }}>
              <Button
                variant="contained"
                color="primary"
                size="small"
                onClick={handleSearch}
                disabled={resultsLoading}
              >
                {resultsLoading ? "Searching..." : "Search"}
              </Button>
            </div>
          </Box>
        )}

        {/* --- 2. RESULTS TABLE PLACEHOLDER --- */}
        {hasSearched && (
          <Box sx={{ mt: 1 }}>
            <Box sx={{ mb: 1 }}>
              <Button
                variant="contained"
                color="primary"
                size="small"
                disabled={resultsLoading || isLoadingMore}
                onClick={() => {
                  setHasSearched(false);
                  setColumnFilters([]);
                  setNextSearchStartRow(undefined);
                  setCurrentSearchQuery("");
                }}
              >
                New Search
              </Button>
            </Box>

            <LocalizationProvider dateAdapter={AdapterDayjs}>
              <MaterialReactTable table={table} />
            </LocalizationProvider>
          </Box>
        )}

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
      </div>
    </ThemeProvider>
  );
};

export default AdvanceSearch;