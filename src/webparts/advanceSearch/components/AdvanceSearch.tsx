import * as React from "react";
import { IAdvanceSearchProps } from "./IAdvanceSearchProps";
import { sp } from "../AdvanceSearchWebPart";
import "@pnp/sp/sharing";
import { SharingRole } from "@pnp/sp/sharing";
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
import CloseIcon from "@mui/icons-material/Close";
import { ThemeProvider, createTheme } from "@mui/material/styles";
import { DatePicker } from "@mui/x-date-pickers/DatePicker";
import { LocalizationProvider } from "@mui/x-date-pickers/LocalizationProvider";
import { AdapterDayjs } from "@mui/x-date-pickers/AdapterDayjs";
import dayjs, { Dayjs } from "dayjs";

// Email Share Dialog Component with isolated state
const EmailShareDialog: React.FC<{
  open: boolean;
  onClose: () => void;
  selectedItems?: any[];
  siteUrl: string;
  defaultSubject?: string;
  defaultMessage?: string;
  currentUserEmail?: string;
}> = ({
  open,
  onClose,
  selectedItems = [],
  siteUrl,
  defaultSubject = "",
  defaultMessage = "",
  currentUserEmail = "",
}) => {
  const [emailFields, setEmailFields] = React.useState({
    to: "",
    cc: currentUserEmail,
    bcc: "",
    subject: defaultSubject,
    message: defaultMessage,
  });
  const [emailErrors, setEmailErrors] = React.useState({ to: "", cc: "", bcc: "" });
  const [shareErrorMessage, setShareErrorMessage] = React.useState<string>("");
  const [isSharing, setIsSharing] = React.useState(false);

  React.useEffect(() => {
    if (open) {
      setEmailFields((prev) => ({
        ...prev,
        cc: prev.cc || currentUserEmail,
        subject: defaultSubject,
        message: defaultMessage,
      }));
    }
  }, [open, defaultSubject, defaultMessage, currentUserEmail]);

  const validateEmail = (email: string) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);

  const validateEmailField = (field: "to" | "cc" | "bcc", value: string): string => {
    const emails = value.split(/[;,]/).map((e) => e.trim()).filter(Boolean);
    let invalidEmail = "";
    for (let i = 0; i < emails.length; i++) {
      if (!validateEmail(emails[i])) {
        invalidEmail = emails[i];
        break;
      }
    }
    return invalidEmail ? `Invalid email: ${invalidEmail}` : "";
  };

  const handleEmailChange = (field: string, value: string) => {
    setEmailFields((prev) => ({ ...prev, [field]: value }));
  };

  const handleEmailBlur = (field: "to" | "cc" | "bcc") => {
    const error = validateEmailField(field, emailFields[field]);
    setEmailErrors((prev) => ({ ...prev, [field]: error }));
  };

  const handleClose = () => {
    setEmailFields({ to: "", cc: currentUserEmail, bcc: "", subject: "", message: "" });
    setEmailErrors({ to: "", cc: "", bcc: "" });
    setShareErrorMessage("");
    onClose();
  };

  const externalShareFiles = async (
    toEmails: string[],
    ccEmails: string[],
    bccEmails: string[]
  ) => {
    const origin = window.location.origin;
    const allRecipients = Array.from(new Set([...toEmails, ...ccEmails, ...bccEmails]));

    // Format custom message (ensuring it stays within SharePoint's 500-char limit)
    let cleanMessage = (emailFields.message || "").replace(/{}/g, "").trim();
    if (cleanMessage.length > 490) {
      cleanMessage = cleanMessage.substring(0, 487) + "...";
    }

    // Share each file using PnPjs shareObject with custom subject and body
    for (let index = 0; index < selectedItems.length; index++) {
      const item = selectedItems[index];
      const itemUrl: string = item.fileUrl || "";
      const fullUrl = itemUrl.startsWith("http") ? itemUrl : `${origin}${itemUrl}`;

      const result = await sp.web.shareObject(
        fullUrl,
        allRecipients,
        SharingRole.View,
        {
          subject: (emailFields.subject || "Shared document").substring(0, 200),
          body: cleanMessage || "Please find the shared document.",
        }
      );

      if (result && result.ErrorMessage) {
        throw new Error(result.ErrorMessage);
      }
    }
  };

  const handleSend = async () => {
    const toError = validateEmailField("to", emailFields.to);
    const ccError = validateEmailField("cc", emailFields.cc);
    const bbcError = validateEmailField("bcc", emailFields.bcc);
    
    if (!toError && !ccError && !bbcError && emailFields.to) {
      setIsSharing(true);
      setShareErrorMessage("");
      try {
        const toEmails = emailFields.to.split(/[;,]/).map((e) => e.trim()).filter(Boolean);
        const ccEmails = emailFields.cc.split(/[;,]/).map((e) => e.trim()).filter(Boolean);
        const bccEmails = emailFields.bcc.split(/[;,]/).map((e) => e.trim()).filter(Boolean);

        await externalShareFiles(toEmails, ccEmails, bccEmails);
        alert(`Successfully shared ${selectedItems.length} file(s)!`);
        handleClose();
      } catch (error: any) {
        console.error("Error during sharing:", error);
        setShareErrorMessage(error?.message || "Failed to share files. Please verify permissions.");
      } finally {
        setIsSharing(false);
      }
    } else {
      setEmailErrors({ to: toError, cc: ccError, bcc: bbcError });
    }
  };

  return (
    <Dialog open={open} onClose={handleClose} fullWidth maxWidth="sm">
      <DialogTitle>Share Selected Items</DialogTitle>
      <DialogContent>
        {shareErrorMessage && (
          <Typography color="error" variant="body2" sx={{ mb: 1, p: 1, backgroundColor: "#ffebee", borderRadius: "4px" }}>
            {shareErrorMessage}
          </Typography>
        )}
        <TextField
          fullWidth
          label="To"
          value={emailFields.to}
          onChange={(e) => handleEmailChange("to", e.target.value)}
          onBlur={() => handleEmailBlur("to")}
          error={!!emailErrors.to}
          helperText={emailErrors.to}
          margin="dense"
        />
        <TextField
          fullWidth
          label="CC"
          value={emailFields.cc}
          onChange={(e) => handleEmailChange("cc", e.target.value)}
          onBlur={() => handleEmailBlur("cc")}
          error={!!emailErrors.cc}
          helperText={emailErrors.cc}
          margin="dense"
        />
        <TextField
          fullWidth
          label="BCC"
          value={emailFields.bcc}
          onChange={(e) => handleEmailChange("bcc", e.target.value)}
          onBlur={() => handleEmailBlur("bcc")}
          error={!!emailErrors.bcc}
          helperText={emailErrors.bcc}
          margin="dense"
        />
        <TextField
          fullWidth
          label="Subject"
          value={emailFields.subject}
          onChange={(e) => handleEmailChange("subject", e.target.value)}
          margin="dense"
        />
        <TextField
          fullWidth
          label="Message"
          value={emailFields.message}
          onChange={(e) => handleEmailChange("message", e.target.value)}
          margin="dense"
          multiline
          rows={4}
        />
      </DialogContent>
      <DialogActions>
        <Button onClick={handleClose}>Cancel</Button>
        <Button
          onClick={handleSend}
          disabled={!emailFields.to || !!emailErrors.to || !!emailErrors.cc || !!emailErrors.bcc || isSharing}
        >
          {isSharing ? "Sharing..." : "Send"}
        </Button>
      </DialogActions>
    </Dialog>
  );
};

type doclib_AllProducts = {
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

const multiSelectFilterFn = (
  row: { getValue: (columnId: string) => unknown },
  columnId: string,
  filterValue: unknown
): boolean => {
  if (!filterValue || (Array.isArray(filterValue) && filterValue.length === 0)) return true;
  const rowValue = String(row.getValue(columnId) || "").toLowerCase();
  const selectedValues = Array.isArray(filterValue) ? filterValue : [filterValue];
  return selectedValues.some((val) => rowValue.includes(String(val).toLowerCase()));
};

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
  const [resultsLoading, setResultsLoading] = React.useState(false);
  const [hasSearched, setHasSearched] = React.useState(false);
  const [taxonomyLoading, setTaxonomyLoading] = React.useState(false);
  const [isShareDialogOpen, setIsShareDialogOpen] = React.useState(false);
  const [selectedRowsData, setSelectedRowsData] = React.useState<any[]>([]);
  const [sharingConfig, setSharingConfig] = React.useState<{ subject: string; message: string }>({
    subject: "",
    message: "",
  });

  // State for File Context Menu & Edit Modal Dialog
  const [fileMenuAnchorEl, setFileMenuAnchorEl] = React.useState<null | HTMLElement>(null);
  const [selectedFileForAction, setSelectedFileForAction] = React.useState<doclib_AllProducts | null>(null);
  const [editModalUrl, setEditModalUrl] = React.useState<string | null>(null);

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
        .orderBy("PIMProductName")
        .top(5000)();

      const sorted = (results as IProductLookupItem[]).sort((a, b) => {
        const nameA = `${a.Title || ""} ${a.PIMProductName || ""}`.trim();
        const nameB = `${b.Title || ""} ${b.PIMProductName || ""}`.trim();
        return nameA.localeCompare(nameB, undefined, { sensitivity: "base" });
      });

      setProducts(sorted);
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
        .orderBy("Title")
        .top(5000)();

      const sorted = (results as IClientLookupItem[]).sort((a, b) =>
        (a.Title || "").localeCompare(b.Title || "", undefined, { sensitivity: "base" })
      );

      setClients(sorted);
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
    const loadTaxonomy = async (): Promise<void> => {
      setTaxonomyLoading(true);

      try {
        const [docTypes, subDocTypes] = await Promise.all([
          sp.web.lists
            .getByTitle("Document Type")
            .items.select("ID", "Title", "ShortTitle")
            .orderBy("Title")(),
          sp.web.lists
            .getByTitle("Sub Document Type")
            .items.select("ID", "Title", "DocumentType/Title")
            .expand("DocumentType")
            .top(2000)
            .orderBy("Title")(),
        ]);

        const sortedDocTypes = (docTypes as IDocumentTypeItem[]).sort((a, b) =>
          (a.Title || "").localeCompare(b.Title || "", undefined, { sensitivity: "base" })
        );
        const sortedSubDocTypes = (subDocTypes as ISubDocumentTypeItem[]).sort((a, b) =>
          (a.Title || "").localeCompare(b.Title || "", undefined, { sensitivity: "base" })
        );

        setDocumentTypes(sortedDocTypes);
        setAllSubDocumentTypes(sortedSubDocTypes);
      } catch (error) {
        console.error("loadTaxonomy error:", error);
        setDocumentTypes([]);
        setAllSubDocumentTypes([]);
      } finally {
        setTaxonomyLoading(false);
      }
    };

    const loadSharingConfiguration = async (): Promise<void> => {
      try {
        const items = await sp.web.lists
          .getByTitle("Configuration")
          .items.select("Title", "Subject", "Message")
          .filter("Title eq 'External Sharing'")
          .top(1)();

        if (items && items.length > 0) {
          const config = items[0];
          setSharingConfig({
            subject: config.Subject ?? "",
            message: config.Message ?? "",
          });
        }
      } catch (err) {
        console.warn("Could not load 'Configuration' list item for 'External Sharing':", err);
      }
    };

    void loadTaxonomy();
    void loadSharingConfiguration();
  }, []);

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

  const loadAllRecords = async (): Promise<void> => {
    try {
      setResultsLoading(true);
      setHasSearched(true);

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

      const mappedData: doclib_AllProducts[] = allProducts.map((item: any) => ({
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
      selectedDocumentTypes.length > 0 ||
      selectedSubDocumentTypes.length > 0 ||
      dateFrom !== null ||
      dateTo !== null;

    if (!hasAnySelection) {
      await loadAllRecords();
      return;
    }

    try {
      setResultsLoading(true);
      setHasSearched(true);

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
            "ConfidentialityOWSCHCS",
            "AlertsOWSMTXT"
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
            "Document_x0020_Date",
            "Alerts",
            "Confidentiality"
          )
          .expand("PIMProductCode")
          .filter(filter)();

        allProducts.push(...items);
      }

      const mappedData: doclib_AllProducts[] = allProducts.map((item: any) => ({
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

      setItems_AllProducts(mappedData);
    } catch (error) {
      console.error("handleSearch error:", error);
    } finally {
      setResultsLoading(false);
    }
  };

  const clientOptions = useMemo(
    () => {
      const unique = new Set<string>();
      items_AllProducts.forEach((item) => {
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
    [items_AllProducts]
  );

  const productOptions = useMemo(
    () => {
      const unique = new Set<string>();
      items_AllProducts.forEach((item) => {
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
    [items_AllProducts]
  );

  const businessLineOptions = useMemo(
    () => {
      const unique = new Set<string>();
      items_AllProducts.forEach((item) => {
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
    [items_AllProducts]
  );

  const countryOptions = useMemo(
    () => {
      const unique = new Set<string>();
      items_AllProducts.forEach((item) => {
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
    [items_AllProducts]
  );

  const documentTypeOptions = useMemo(
    () => {
      const unique = new Set<string>();
      items_AllProducts.forEach((item) => {
        if (item.DocumentTypeSearchText) {
          const trimmed = item.DocumentTypeSearchText.trim();
          if (trimmed) unique.add(trimmed);
        }
      });
      const result: string[] = [];
      unique.forEach((val) => result.push(val));
      return result.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
    },
    [items_AllProducts]
  );

  const subDocumentTypeOptions = useMemo(
    () => {
      const unique = new Set<string>();
      items_AllProducts.forEach((item) => {
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
    [items_AllProducts]
  );

  const confidentialityOptions = useMemo(
    () => {
      const unique = new Set<string>();
      items_AllProducts.forEach((item) => {
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
      <Box sx={{ display: "flex", gap: "8px", alignItems: "center" }}>
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
    state: {
      isLoading: resultsLoading,
      showColumnFilters: true,
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
            <Box sx={{ display: "flex", justifyContent: "space-between", mb: 1, alignItems: "center" }}>
              <Button
                variant="text"
                size="small"
                onClick={() => setHasSearched(false)}
                style={{ textDecoration: "underline", color: "#1976d2", textTransform: "none", fontWeight: 600 }}
              >
                &larr; Back to Search Filters
              </Button>

              <a
                href="#"
                onClick={(e) => {
                  e.preventDefault();
                  loadAllRecords().catch(console.error);
                }}
                style={{ textDecoration: "underline", color: "#1976d2", fontSize: "13px" }}
              >
                REFRESH ALL RECORDS
              </a>
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