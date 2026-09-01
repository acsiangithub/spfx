import * as React from "react";
import { IAdvanceSearchProps } from "./IAdvanceSearchProps";
import { sp } from "../AdvanceSearchWebPart";
import { useMemo } from "react";
import {
  MaterialReactTable,
  useMaterialReactTable,
  type MRT_ColumnDef,
} from "material-react-table";

import Autocomplete from "@mui/material/Autocomplete";
import TextField from "@mui/material/TextField";
import CircularProgress from "@mui/material/CircularProgress";
import Button from "@mui/material/Button";
import FormControl from "@mui/material/FormControl";
import InputLabel from "@mui/material/InputLabel";
import MenuItem from "@mui/material/MenuItem";
import Select from "@mui/material/Select";
import Box from "@mui/material/Box";
import Typography from "@mui/material/Typography";
import { ThemeProvider, createTheme } from "@mui/material/styles";
import { DatePicker } from "@mui/x-date-pickers/DatePicker";
import { LocalizationProvider } from "@mui/x-date-pickers/LocalizationProvider";
import { AdapterDayjs } from "@mui/x-date-pickers/AdapterDayjs";
import dayjs, { Dayjs } from "dayjs";

type doclib_AllProducts = {
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
};

type IProductLookupItem = {
  ID: number;
  Title: string;
  PIMProductName: string;
  Manufacturer?: string;
  BusinessLine?: string;
  ManufacturerLookupId?:number;
};

type IClientLookupItem = {
  ID: number;
  Title: string;
};

type IDocumentTypeItem = {
  ID: number;
  Title: string;
  ShortTitle?: string;
};

type ISubDocumentTypeItem = {
  ID: number;
  Title: string;
  DocumentType?: {
    Title?: string;
  } | null;
};

const choiceToString = (value: string | string[] | undefined | null): string => {
  if (!value) return "";
  return Array.isArray(value) ? value.join(", ") : value;
};

const sanitizeKqlValue = (value: string): string =>
  value.replace(/"/g, '\\"').trim();

const documentDateFilter = (
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

const compactTheme = createTheme({
  components: {
    MuiInputBase: {
      styleOverrides: {
        root: {
          fontSize: "14px",
          minHeight: "36px",
        },
        input: {
          paddingTop: "8px",
          paddingBottom: "8px",
        },
      },
    },
    MuiOutlinedInput: {
      styleOverrides: {
        root: {
          fontSize: "14px",
          minHeight: "36px",
        },
        input: {
          padding: "8px 12px",
        },
      },
    },
    MuiInputLabel: {
      styleOverrides: {
        root: {
          fontSize: "14px",
        },
      },
    },
    MuiButton: {
      styleOverrides: {
        root: {
          minHeight: "36px",
          padding: "7px 14px",
          fontSize: "14px",
          lineHeight: 1.3,
        },
      },
    },
    MuiMenuItem: {
      styleOverrides: {
        root: {
          fontSize: "14px",
          minHeight: "36px",
        },
      },
    },
    MuiAutocomplete: {
      styleOverrides: {
        inputRoot: {
          minHeight: "36px",
        },
      },
    },
    MuiTableCell: {
      styleOverrides: {
        root: {
          padding: "8px 10px",
          fontSize: "13px",
        },
      },
    },
    MuiPaper: {
      styleOverrides: {
        root: {
          fontSize: "14px",
        },
      },
    },
  },
  typography: {
    fontSize: 14,
  },
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
  const [subDocumentTypes, setSubDocumentTypes] = React.useState<
    ISubDocumentTypeItem[]
  >([]);
  const [selectedDocumentType, setSelectedDocumentType] =
    React.useState<string>("");
  const [selectedSubDocumentType, setSelectedSubDocumentType] =
    React.useState<string>("");

  const [productSearchText, setProductSearchText] = React.useState("");
  const [clientSearchText, setClientSearchText] = React.useState("");
  const [additionalKeyword, setAdditionalKeyword] = React.useState("");
  const [dateFrom, setDateFrom] = React.useState<Dayjs | null>(null);
  const [dateTo, setDateTo] = React.useState<Dayjs | null>(null);

  const [lookupLoading, setLookupLoading] = React.useState(false);
  const [resultsLoading, setResultsLoading] = React.useState(false);
  const [documentTypeLoading, setDocumentTypeLoading] = React.useState(false);
  const [subDocumentTypeLoading, setSubDocumentTypeLoading] =
    React.useState(false);

  const searchProducts = async (searchText: string) => {
    if (searchText.trim().length < 3) {
      setProducts([]);
      return;
    }

    setLookupLoading(true);

    try {
      const escapedText = searchText.replace(/'/g, "''");

      const results = await sp.web.lists
        .getByTitle("PIM Product")
        .items.select("ID", "Title", "PIMProductName", "Manufacturer", "BusinessLine","ManufacturerLookupId")
        .filter(`substringof('${escapedText}', PIMProductName)`)
        .top(5000)();

      setProducts(results as IProductLookupItem[]);
    } catch (error) {
      console.error("searchProducts error:", error);
    } finally {
      setLookupLoading(false);
    }
  };

  const searchClients = async (searchText: string) => {
    if (searchText.trim().length < 3) {
      setClients([]);
      return;
    }

    setLookupLoading(true);

    try {
      const escapedText = searchText.replace(/'/g, "''");

      const results = await sp.web.lists
        .getByTitle("PIM Global Client")
        .items.select("ID", "Title")
        .filter(`substringof('${escapedText}', Title)`)
        .top(5000)();

      setClients(results as IClientLookupItem[]);
    } catch (error) {
      console.error("searchClients error:", error);
    } finally {
      setLookupLoading(false);
    }
  };

  React.useEffect(() => {
    const timer = setTimeout(() => {
      if (productSearchText.trim().length >= 3) {
        searchProducts(productSearchText.trim());
      } else {
        setProducts([]);
      }
    }, 300);

    return () => clearTimeout(timer);
  }, [productSearchText]);

  React.useEffect(() => {
    const timer = setTimeout(() => {
      if (clientSearchText.trim().length >= 3) {
        searchClients(clientSearchText.trim());
      } else {
        setClients([]);
      }
    }, 300);

    return () => clearTimeout(timer);
  }, [clientSearchText]);

  React.useEffect(() => {
    const loadDocumentTypes = async (): Promise<void> => {
      setDocumentTypeLoading(true);

      try {
        const items = await sp.web.lists
          .getByTitle("Document Type")
          .items.select("ID", "Title", "ShortTitle")
          .orderBy("Title")();

        setDocumentTypes(items as IDocumentTypeItem[]);
      } catch (error) {
        console.error("loadDocumentTypes error:", error);
        setDocumentTypes([]);
      } finally {
        setDocumentTypeLoading(false);
      }
    };

    void loadDocumentTypes();
  }, []);

  React.useEffect(() => {
    const loadSubDocumentTypes = async (): Promise<void> => {
      if (!selectedDocumentType) {
        setSubDocumentTypes([]);
        setSelectedSubDocumentType("");
        return;
      }

      setSubDocumentTypeLoading(true);

      try {
        const escapedValue = selectedDocumentType.replace(/'/g, "''");
        const items = await sp.web.lists
          .getByTitle("Sub Document Type")
          .items.select("ID", "Title", "DocumentType/Title")
          .expand("DocumentType")
          .filter(`DocumentType/Title eq '${escapedValue}'`)
          .orderBy("Title")();

        setSubDocumentTypes(items as ISubDocumentTypeItem[]);

        if (
          selectedSubDocumentType &&
          !items.some((item: any) => item.Title === selectedSubDocumentType)
        ) {
          setSelectedSubDocumentType("");
        }
      } catch (error) {
        console.error("loadSubDocumentTypes error:", error);
        setSubDocumentTypes([]);
      } finally {
        setSubDocumentTypeLoading(false);
      }
    };

    void loadSubDocumentTypes();
  }, [selectedDocumentType, selectedSubDocumentType]);

  const buildSearchQuery = (): string => {
    const clauses: string[] = [];

    const searchPath = `${props.urlSite.replace(/\/$/, "")}/Products/*`;
    clauses.push(`Path:"${searchPath}"`);

    const productValues = selectedProducts
      .map((item) => (item.PIMProductName || item.Title || "").trim())
      .filter(Boolean);

    const clientValues = selectedClients
      .map((item) => (item.Title || "").trim())
      .filter(Boolean);

    const documentTypeValues = selectedDocumentType
      ? [selectedDocumentType.trim()].filter(Boolean)
      : [];

    const subDocumentTypeValues = selectedSubDocumentType
      ? [selectedSubDocumentType.trim()].filter(Boolean)
      : [];

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

  const loadAllRecords = async (): Promise<void> => {
    try {
      setResultsLoading(true);

      const allProducts: any[] = [];
      const pageSize = 5000;
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
            "Document_x0020_Date"
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

      const mappedData: doclib_AllProducts[] = allProducts.map((item: any) => ({
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
      }));

      setItems_AllProducts(mappedData);
      //setAllRecordsCache(mappedData);
    } catch (error) {
      console.error("loadAllRecords error:", error);
    } finally {
      setResultsLoading(false);
    }
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
      selectedDocumentType.trim().length > 0 ||
      selectedSubDocumentType.trim().length > 0 ||
      dateFrom !== null ||
      dateTo !== null;

    if (!hasAnySelection) {
      await loadAllRecords();
      return;
    }

    try {
      setResultsLoading(true);

      let queryText = buildSearchQuery();

      if (!queryText) {
        await loadAllRecords();
        return;
      }

      let allResults: any[] = [];
      let startRow = 0;
      const pageSize = 500;

      const maxResults = 2500;

      while (startRow < maxResults) {
        const results = await sp.search({
          Querytext: queryText,
          RowLimit: Math.min(pageSize, maxResults - startRow),
          StartRow: startRow,
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
            //"PIMProduct*",
            "owstaxIdPIMProductTermSet",  
            "PIMProductCodeOWSTEXT",
          ],
        });

        const currentResults =
          results?.PrimarySearchResults ?? [];

        allResults.push(...currentResults);

        if (
          currentResults.length < pageSize ||
          allResults.length >= maxResults
        ) {
          break;
        }

        startRow += pageSize;
      }

      const ids: number[] = [];

      allResults.forEach((r: any) => {
        const id = Number(r.ListItemID);
        if (!isNaN(id) && ids.indexOf(id) === -1) {
          ids.push(id);
        }
      });

      if (ids.length === 0) {
        setItems_AllProducts([]);
        return;
      }

      const allProducts: any[] = [];

      for (let i = 0; i < ids.length; i += 100) {
        const currentIds = ids.slice(i, i + 100);
        const filter = currentIds.map((id) => `Id eq ${id}`).join(" or ");

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
            "Document_x0020_Date"
          )
          .expand("PIMProductCode")
          .filter(filter)();

        allProducts.push(...items);
      }

      const mappedData: doclib_AllProducts[] = allProducts.map((item: any) => ({
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
      }));

      setItems_AllProducts(mappedData);
    } catch (error) {
      console.error("handleSearch error:", error);
    } finally {
      setResultsLoading(false);
    }
  };

  const businessLineOptions = useMemo(
    () => {
      const unique = new Set<string>();
      items_AllProducts.forEach((item) => {
        if (item.BusinessLine) {
          item.BusinessLine.split(",").forEach((val) => unique.add(val.trim()));
        }
      });
      const result: string[] = [];
      unique.forEach((val) => result.push(val));
      return result.sort();
    },
    [items_AllProducts]
  );

  const countryOptions = useMemo(
    () => {
      const unique = new Set<string>();
      items_AllProducts.forEach((item) => {
        if (item.CountrySoldTo) {
          item.CountrySoldTo.split(",").forEach((val) => unique.add(val.trim()));
        }
      });
      const result: string[] = [];
      unique.forEach((val) => result.push(val));
      return result.sort();
    },
    [items_AllProducts]
  );

  const documentTypeOptions = useMemo(
    () => {
      const unique = new Set<string>();
      items_AllProducts.forEach((item) => {
        if (item.DocumentTypeSearchText) {
          unique.add(item.DocumentTypeSearchText.trim());
        }
      });
      const result: string[] = [];
      unique.forEach((val) => result.push(val));
      return result.sort();
    },
    [items_AllProducts]
  );

  const subDocumentTypeOptions = useMemo(
    () => {
      const unique = new Set<string>();
      items_AllProducts.forEach((item) => {
        if (item.SubDocumentTypeSearchText) {
          item.SubDocumentTypeSearchText.split(",").forEach((val) => unique.add(val.trim()));
        }
      });
      const result: string[] = [];
      unique.forEach((val) => result.push(val));
      return result.sort();
    },
    [items_AllProducts]
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
          const fileUrl = `${props.urlSite}Products/${encodeURIComponent(
            row.original.filename
          )}`;
          const displayName =
            row.original.filename.length > 20
              ? `${row.original.filename.slice(0, 17)}...`
              : row.original.filename;

          return (
            <a
              href={fileUrl}
              target="_blank"
              rel="noopener noreferrer"
              title={row.original.filename}
              style={{
                display: "inline-block",
                maxWidth: "100%",
                overflow: "hidden",
                textOverflow: "ellipsis",
                whiteSpace: "nowrap",
              }}
            >
              {displayName}
            </a>
          );
        },
      },
      {
        accessorKey: "ManufacturerSearchText",
        header: "Clients",
        filterFn: "contains",
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
        filterFn: "contains",
        size: 160,
        minSize: 160,
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
        size: 150,
        minSize: 150,
      },
      {
        accessorKey: "SubDocumentTypeSearchText",
        header: "Sub Document Type",
        filterVariant: "multi-select",
        filterSelectOptions: subDocumentTypeOptions,
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
        size: 140,
        minSize: 140,
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
        accessorKey: "CountrySoldTo",
        header: "Country Sold To",
        filterVariant: "multi-select",
        filterSelectOptions: countryOptions,
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
    [props.urlSite, businessLineOptions, countryOptions, documentTypeOptions, subDocumentTypeOptions]
  );

  const table = useMaterialReactTable({
    columns: columns_AllProducts,
    data: items_AllProducts,
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
      rowsPerPageOptions: [50, 100, 500, 1000],
      sx: {
        ".MuiTablePagination-selectLabel, .MuiTablePagination-displayedRows": {
          fontSize: "13px",
        },
      },
    },
    state: {
      isLoading: resultsLoading,
      showColumnFilters: true,
    },
  });

  
  return (
    <ThemeProvider theme={compactTheme}>
      <div>
        <h2>Search Clients & Products</h2>

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
          <FormControl fullWidth size="small" disabled={documentTypeLoading}>
            <InputLabel id="document-type-select-label">Document Type</InputLabel>
            <Select
              labelId="document-type-select-label"
              value={selectedDocumentType}
              label="Document Type"
              size="small"
              onChange={(event) => setSelectedDocumentType(event.target.value as string)}
            >
              <MenuItem value="">
                <em>All</em>
              </MenuItem>
              {documentTypes.map((item) => (
                <MenuItem key={item.ID} value={item.Title}>
                  {item.ShortTitle ? `${item.Title} (${item.ShortTitle})` : item.Title}
                </MenuItem>
              ))}
            </Select>
          </FormControl>

          <FormControl
            fullWidth
            size="small"
            disabled={!selectedDocumentType || subDocumentTypeLoading}
          >
            <InputLabel id="sub-document-type-select-label">
              Sub Document Type
            </InputLabel>
            <Select
              labelId="sub-document-type-select-label"
              value={selectedSubDocumentType}
              label="Sub Document Type"
              size="small"
              onChange={(event) =>
                setSelectedSubDocumentType(event.target.value as string)
              }
            >
              <MenuItem value="">
                <em>All</em>
              </MenuItem>
              {subDocumentTypes.map((item) => (
                <MenuItem key={item.ID} value={item.Title}>
                  {item.Title}
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

        <div style={{ marginBottom: "16px" }}>
          <Button
            variant="contained"
            color="primary"
            size="small"
            onClick={() => {
              handleSearch();
              table.setIsFullScreen(true);
            }}
            disabled={resultsLoading}
          >
            {resultsLoading ? "Searching..." : "Search"}
          </Button>
        </div>

        <LocalizationProvider dateAdapter={AdapterDayjs}>
          <MaterialReactTable table={table} />
        </LocalizationProvider>

        <div style={{ marginTop: "16px" }}>
          <a
            href="#"
            onClick={(e) => {
              e.preventDefault();
              //setAllRecordsCache([]);
              loadAllRecords().catch(console.error);
            }}
            style={{ textDecoration: "underline", color: "#1976d2" }}
          >
            REFRESH ALL RECORDS
          </a>
        </div>
      </div>
    </ThemeProvider>
  );
};

export default AdvanceSearch;