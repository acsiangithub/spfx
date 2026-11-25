import * as React from "react";
import { SPFI } from "@pnp/sp";
import Dialog from "@mui/material/Dialog";
import DialogTitle from "@mui/material/DialogTitle";
import DialogContent from "@mui/material/DialogContent";
import DialogActions from "@mui/material/DialogActions";
import Paper, { PaperProps } from "@mui/material/Paper";
import Box from "@mui/material/Box";
import Typography from "@mui/material/Typography";
import Button from "@mui/material/Button";
import IconButton from "@mui/material/IconButton";
import Link from "@mui/material/Link";
import TextField from "@mui/material/TextField";
import Autocomplete from "@mui/material/Autocomplete";
import FormControl from "@mui/material/FormControl";
import Select from "@mui/material/Select";
import MenuItem from "@mui/material/MenuItem";
import OutlinedInput from "@mui/material/OutlinedInput";
import Checkbox from "@mui/material/Checkbox";
import ListItemText from "@mui/material/ListItemText";
import CircularProgress from "@mui/material/CircularProgress";
import LinearProgress from "@mui/material/LinearProgress";
import Alert from "@mui/material/Alert";
import Chip from "@mui/material/Chip";
import Tooltip from "@mui/material/Tooltip";
import CloseIcon from "@mui/icons-material/Close";
import EditNoteIcon from "@mui/icons-material/EditNote";
import DragIndicatorIcon from "@mui/icons-material/DragIndicator";
import ContentCopyIcon from "@mui/icons-material/ContentCopy";
import ContentPasteIcon from "@mui/icons-material/ContentPaste";
import CheckIcon from "@mui/icons-material/Check";
import OpenInNewIcon from "@mui/icons-material/OpenInNew";
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
} from "../types/advanceSearchTypes";
import {
  searchProducts as searchProductsService,
  searchClients as searchClientsService,
  loadItemDetailsForEdit,
  checkUserWritePermissions,
  resolveProductsBatch,
  IEditPropertiesPayload,
  ILibraryColumnChoices,
  getSenderUserAccessMatrix,
  getCountryRegionMapping,
  checkConfidentialDocumentAccess,
} from "../../../services/sharePointService";

const EditProductListbox = React.forwardRef<
  HTMLUListElement,
  React.HTMLAttributes<HTMLElement>
>(function EditProductListbox(props, ref) {
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

const DraggablePaper = React.forwardRef<HTMLDivElement, PaperProps>(function DraggablePaper(props, ref) {
  const paperRef = React.useRef<HTMLDivElement | null>(null);
  const offsetRef = React.useRef<{ x: number; y: number }>({ x: 0, y: 0 });
  const dragInfoRef = React.useRef<{
    isDragging: boolean;
    startX: number;
    startY: number;
    startOffsetX: number;
    startOffsetY: number;
    paperRect: DOMRect | null;
  }>({
    isDragging: false,
    startX: 0,
    startY: 0,
    startOffsetX: 0,
    startOffsetY: 0,
    paperRect: null,
  });

  const setRefs = React.useCallback(
    (node: HTMLDivElement | null) => {
      paperRef.current = node;
      if (typeof ref === "function") {
        ref(node);
      } else if (ref) {
        (ref as React.MutableRefObject<HTMLDivElement | null>).current = node;
      }
    },
    [ref]
  );

  React.useEffect(() => {
    const paper = paperRef.current;
    if (!paper) return;

    const handle = (paper.querySelector("#draggable-edit-dialog-title") ||
      paper.querySelector(".draggable-dialog-handle")) as HTMLElement | null;
    if (!handle) return;

    handle.style.cursor = "grab";
    handle.style.userSelect = "none";

    const onPointerDown = (e: PointerEvent) => {
      if (e.button !== 0) return;

      const target = e.target as HTMLElement | null;
      if (
        target &&
        target.closest(
          'button, input, textarea, select, a, [role="button"], .MuiButtonBase-root, .MuiSwitch-root'
        )
      ) {
        return;
      }

      e.preventDefault();
      const rect = paper.getBoundingClientRect();
      dragInfoRef.current = {
        isDragging: true,
        startX: e.clientX,
        startY: e.clientY,
        startOffsetX: offsetRef.current.x,
        startOffsetY: offsetRef.current.y,
        paperRect: rect,
      };

      handle.style.cursor = "grabbing";
      document.body.style.userSelect = "none";

      const onPointerMove = (moveEvent: PointerEvent) => {
        if (!dragInfoRef.current.isDragging || !dragInfoRef.current.paperRect) return;

        const deltaX = moveEvent.clientX - dragInfoRef.current.startX;
        const deltaY = moveEvent.clientY - dragInfoRef.current.startY;

        const initialRect = dragInfoRef.current.paperRect;
        const minDeltaX = -(initialRect.width - 120) - initialRect.left;
        const maxDeltaX = window.innerWidth - 120 - initialRect.left;
        const minDeltaY = -initialRect.top;
        const maxDeltaY = window.innerHeight - 60 - initialRect.top;

        const clampedDeltaX = Math.min(Math.max(deltaX, minDeltaX), maxDeltaX);
        const clampedDeltaY = Math.min(Math.max(deltaY, minDeltaY), maxDeltaY);

        const newX = dragInfoRef.current.startOffsetX + clampedDeltaX;
        const newY = dragInfoRef.current.startOffsetY + clampedDeltaY;

        offsetRef.current = { x: newX, y: newY };
        if (paperRef.current) {
          paperRef.current.style.transform = `translate(${newX}px, ${newY}px)`;
        }
      };

      const onPointerUp = () => {
        dragInfoRef.current.isDragging = false;
        if (handle) {
          handle.style.cursor = "grab";
        }
        document.body.style.userSelect = "";
        window.removeEventListener("pointermove", onPointerMove);
        window.removeEventListener("pointerup", onPointerUp);
        window.removeEventListener("pointercancel", onPointerUp);
      };

      window.addEventListener("pointermove", onPointerMove);
      window.addEventListener("pointerup", onPointerUp);
      window.addEventListener("pointercancel", onPointerUp);
    };

    handle.addEventListener("pointerdown", onPointerDown);
    return () => {
      handle.removeEventListener("pointerdown", onPointerDown);
      document.body.style.userSelect = "";
    };
  }, []);

  return (
    <Paper
      {...props}
      ref={setRefs}
      style={{
        ...props.style,
        transform: `translate(${offsetRef.current.x}px, ${offsetRef.current.y}px)`,
      }}
    />
  );
});

const copyToClipboard = async (text: string): Promise<boolean> => {
  if (!text) return false;
  if (navigator?.clipboard?.writeText) {
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch {
      // Fall through to textarea execCommand fallback
    }
  }
  try {
    const textArea = document.createElement("textarea");
    textArea.value = text;
    textArea.style.position = "fixed";
    textArea.style.left = "-9999px";
    textArea.style.top = "0";
    document.body.appendChild(textArea);
    textArea.focus();
    textArea.select();
    const success = document.execCommand("copy");
    document.body.removeChild(textArea);
    return success;
  } catch (err) {
    console.warn("Failed to copy value:", err);
    return false;
  }
};

const parseDelimitedText = (text: string): string[] => {
  const clean = (text || "").trim();
  if (!clean) return [];

  // 1. If text contains ';' or newline, split ONLY by semicolon/newline so commas inside product names (e.g. "OIL, REFINED") are NOT split!
  if (clean.includes(";") || clean.includes("\n")) {
    return clean
      .split(/[\r\n;]+/)
      .map((item) => item.trim())
      .filter(Boolean);
  }

  // 2. Only if NO semicolon/newline exists, check if comma is used as delimiter
  if (clean.includes(",")) {
    return clean
      .split(/,+/)
      .map((item) => item.trim())
      .filter(Boolean);
  }

  return [clean];
};

const getRelatedClients = (
  products: IProductLookupItem[],
  existingClients: IClientLookupItem[],
  clientOptions: IClientLookupItem[]
): IClientLookupItem[] => {
  const clientTitles = new Map<string, string>();
  products.forEach((product) => {
    parseDelimitedText(product.Manufacturer || "").forEach((title) => {
      const key = title.toLowerCase();
      if (!clientTitles.has(key)) clientTitles.set(key, title);
    });
  });

  return Array.from(clientTitles.entries()).map(([key, title], index) =>
    existingClients.find((client) => (client.Title || "").trim().toLowerCase() === key) ||
    clientOptions.find((client) => (client.Title || "").trim().toLowerCase() === key) || {
      ID: -(index + 100),
      Title: title,
    }
  );
};

const haveSameClients = (
  first: IClientLookupItem[],
  second: IClientLookupItem[]
): boolean =>
  first.length === second.length &&
  first.every(
    (client, index) =>
      (client.Title || "").trim().toLowerCase() ===
      (second[index].Title || "").trim().toLowerCase()
  );

const parseDelimitedPaste = (
  e: React.ClipboardEvent<HTMLInputElement | HTMLDivElement>
): string[] | null => {
  const pasteData = e.clipboardData?.getData("text");
  if (!pasteData) return null;

  if (pasteData.includes(";") || pasteData.includes(",") || pasteData.includes("\n")) {
    e.preventDefault();
    return parseDelimitedText(pasteData);
  }
  return null;
};

interface ICopyFieldValueButtonProps {
  value: unknown;
  label?: string;
  formatValue?: (val: any) => string;
}

const CopyFieldValueButton: React.FC<ICopyFieldValueButtonProps> = ({
  value,
  label = "value",
  formatValue,
}) => {
  const [copied, setCopied] = React.useState(false);

  const textToCopy = React.useMemo(() => {
    if (formatValue) return formatValue(value);
    if (value === null || value === undefined) return "";
    if (typeof value === "string") return value.trim();
    if (typeof value === "number") return String(value);

    // Dayjs date objects
    if (dayjs.isDayjs(value)) {
      return value.isValid() ? value.format("DD/MM/YYYY") : "";
    }

    // Array (e.g. Products, Clients, SubDocTypes, Languages)
    if (Array.isArray(value)) {
      return value
        .map((item) => {
          if (!item) return "";
          if (typeof item === "string") return item.trim();
          if (item.Title && item.PIMProductName) {
            return `${item.Title} ${item.PIMProductName}`.trim();
          }
          return item.Title || item.name || String(item);
        })
        .filter(Boolean)
        .join("; ");
    }

    // Single object (e.g. DocumentType)
    if (typeof value === "object") {
      const obj = value as any;
      if (obj.Title && obj.PIMProductName) {
        return `${obj.Title} ${obj.PIMProductName}`.trim();
      }
      return obj.Title || obj.name || "";
    }

    return String(value);
  }, [value, formatValue]);

  const hasValue = Boolean(textToCopy && textToCopy.trim().length > 0);

  const handleCopy = async (e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (!hasValue) return;

    const ok = await copyToClipboard(textToCopy);
    if (ok) {
      setCopied(true);
      setTimeout(() => setCopied(false), 1500);
    }
  };

  return (
    <Tooltip
      title={
        !hasValue
          ? `No ${label} to copy`
          : copied
          ? "Copied!"
          : `Copy ${label}`
      }
      arrow
      placement="top"
    >
      <span>
        <IconButton
          size="small"
          onClick={handleCopy}
          disabled={!hasValue}
          sx={{
            p: "2px",
            color: copied ? "success.main" : "text.secondary",
            opacity: hasValue ? 0.75 : 0.3,
            transition: "all 0.15s ease",
            "&:hover": {
              opacity: 1,
              color: copied ? "success.main" : "primary.main",
              backgroundColor: "rgba(0, 120, 212, 0.08)",
            },
          }}
        >
          {copied ? (
            <CheckIcon sx={{ fontSize: 13 }} />
          ) : (
            <ContentCopyIcon sx={{ fontSize: 13 }} />
          )}
        </IconButton>
      </span>
    </Tooltip>
  );
};

interface IPasteFieldValueButtonProps {
  label?: string;
  onPasteText: (text: string) => void | Promise<void>;
  disabled?: boolean;
}

const PasteFieldValueButton: React.FC<IPasteFieldValueButtonProps> = ({
  label = "value",
  onPasteText,
  disabled = false,
}) => {
  const [pasted, setPasted] = React.useState(false);

  const handlePasteClick = async (e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (disabled) return;

    try {
      if (navigator?.clipboard?.readText) {
        const text = await navigator.clipboard.readText();
        if (text && text.trim()) {
          await onPasteText(text.trim());
          setPasted(true);
          setTimeout(() => setPasted(false), 1500);
        }
      }
    } catch (err) {
      console.warn("Failed to read from clipboard:", err);
    }
  };

  return (
    <Tooltip
      title={pasted ? "Pasted!" : `Paste ${label} from clipboard`}
      arrow
      placement="top"
    >
      <span>
        <IconButton
          size="small"
          onClick={handlePasteClick}
          disabled={disabled}
          sx={{
            p: "2px",
            color: pasted ? "success.main" : "text.secondary",
            opacity: 0.75,
            transition: "all 0.15s ease",
            "&:hover": {
              opacity: 1,
              color: pasted ? "success.main" : "primary.main",
              backgroundColor: "rgba(0, 120, 212, 0.08)",
            },
          }}
        >
          {pasted ? (
            <CheckIcon sx={{ fontSize: 13 }} />
          ) : (
            <ContentPasteIcon sx={{ fontSize: 13 }} />
          )}
        </IconButton>
      </span>
    </Tooltip>
  );
};

const parseDateString = (text: string): Dayjs | null => {
  const clean = text.trim();
  if (!clean) return null;
  const formats = ["DD/MM/YYYY", "DD-MM-YYYY", "YYYY-MM-DD", "D/M/YYYY", "YYYY/MM/DD"];
  for (const f of formats) {
    const d = dayjs(clean, f, true);
    if (d.isValid()) return d;
  }
  const fallback = dayjs(clean);
  return fallback.isValid() ? fallback : null;
};

const splitMultiValue = (value?: string | null): string[] => {
  const seen = new Set<string>();
  return (value || "")
    .split(/[\r\n;,]+/)
    .map((part) => part.trim())
    .filter((part) => {
      const key = part.toLowerCase();
      if (!part || seen.has(key)) return false;
      seen.add(key);
      return true;
    });
};

const getProductValues = (
  products: IProductLookupItem[],
  property: "BusinessLine" | "Country"
): string[] => splitMultiValue(products.map((product) => product[property] || "").join("; "));

const getClientValues = (
  clients: IClientLookupItem[],
  property: "BusinessLine" | "Country"
): string[] => splitMultiValue(clients.map((client) => client[property] || "").join("; "));

type FilePreviewType = "pdf" | "office" | null;

const getFilePreviewType = (item: doclib_AllProducts): FilePreviewType => {
  const paths = [item.filename, item.OriginalFilename, item.fileUrl].map((value) =>
    (value || "").split(/[?#]/)[0].toLowerCase()
  );

  if (paths.some((path) => path.endsWith(".pdf"))) return "pdf";
  if (paths.some((path) => /\.(docx|xlsx|pptx)$/.test(path))) return "office";
  return null;
};

export interface IEditPropertiesDialogProps {
  open: boolean;
  onClose: () => void;
  selectedItems: doclib_AllProducts[];
  sp: SPFI;
  documentTypes: IDocumentTypeItem[];
  allSubDocumentTypes: ISubDocumentTypeItem[];
  libraryChoices?: ILibraryColumnChoices;
  siteUrl?: string;
  currentUserEmail?: string;
  onSave: (payload: IEditPropertiesPayload, itemIds: number[]) => Promise<void>;
}

export const EditPropertiesDialog: React.FC<IEditPropertiesDialogProps> = ({
  open,
  onClose,
  selectedItems,
  sp,
  documentTypes,
  allSubDocumentTypes,
  libraryChoices,
  siteUrl,
  currentUserEmail,
  onSave,
}) => {
  const isBulkEdit = selectedItems.length > 1;
  const singleItem = selectedItems.length === 1 ? selectedItems[0] : null;
  const formatMetadataDate = (value?: Date | string | null): string => {
    if (!value) return "-";
    const date = dayjs(value);
    return date.isValid() ? date.format("DD/MM/YYYY") : "-";
  };
  const documentId = singleItem?.OData__dlc_DocId?.trim() || "-";
  const documentVersion = singleItem?.DocVersion !== undefined && singleItem?.DocVersion !== null
    ? String(singleItem.DocVersion)
    : "-";
  const created = formatMetadataDate(singleItem?.Created);
  const createdBy = singleItem?.AuthorTitle?.trim() || "-";
  const modified = formatMetadataDate(singleItem?.Modified);
  const modifiedBy = singleItem?.EditorTitle?.trim() || "-";
  const [selectedPreviewItem, setSelectedPreviewItem] = React.useState<doclib_AllProducts | null>(null);
  const [officePreviewUrl, setOfficePreviewUrl] = React.useState<string>("");
  const [isLoadingOfficePreview, setIsLoadingOfficePreview] = React.useState<boolean>(false);
  const [officePreviewError, setOfficePreviewError] = React.useState<string | null>(null);

  const getAbsoluteFileUrl = React.useCallback(
    (rawUrl?: string): string => {
      if (!rawUrl) return "";
      if (rawUrl.startsWith("http://") || rawUrl.startsWith("https://")) {
        return rawUrl;
      }
      const cleanSiteUrl = (siteUrl || "").trim().replace(/\/+$/, "");
      const origin = cleanSiteUrl
        ? (() => {
            try {
              return new URL(cleanSiteUrl).origin;
            } catch {
              return typeof window !== "undefined" ? window.location.origin : "";
            }
          })()
        : typeof window !== "undefined"
        ? window.location.origin
        : "";

      if (rawUrl.startsWith("/")) {
        return `${origin}${rawUrl}`;
      }

      return cleanSiteUrl ? `${cleanSiteUrl}/${rawUrl}` : `${origin}/${rawUrl}`;
    },
    [siteUrl]
  );

  const getDirectPreviewFileUrl = React.useCallback(
    (item: doclib_AllProducts): string => {
      if ((item.Confidentiality || "").trim().toLowerCase() === "confidential") {
        const cleanFileName = (item.filename || item.OriginalFilename || "")
          .replace(/\.url$/i, "")
          .trim();
        if (!cleanFileName) return "";
        const webUrl = (siteUrl?.trim() || window.location.origin).replace(/\/+$/, "");
        return `${webUrl}/Clients%20%20ProductsConfidential/${encodeURIComponent(cleanFileName)}`;
      }

      return getAbsoluteFileUrl(item.fileUrl);
    },
    [getAbsoluteFileUrl, siteUrl]
  );

  const getOfficeViewUrl = React.useCallback(
    async (item: doclib_AllProducts): Promise<string> => {
      const absoluteFileUrl = getDirectPreviewFileUrl(item);
      if (!absoluteFileUrl) {
        throw new Error("No file URL is available for this Office document.");
      }

      const serverRelativeUrl = decodeURIComponent(new URL(absoluteFileUrl).pathname);
      const file = await sp.web
        .getFileByServerRelativePath(serverRelativeUrl)
        .select("UniqueId")();
      if (!file.UniqueId) {
        throw new Error("SharePoint did not return a unique file ID for this Office document.");
      }

      const webUrl = (siteUrl?.trim() || new URL(absoluteFileUrl).origin).replace(/\/+$/, "");
      const viewerUrl = new URL(`${webUrl}/_layouts/15/WopiFrame.aspx`);
      viewerUrl.searchParams.set("sourcedoc", `{${file.UniqueId.replace(/[{}]/g, "")}}`);
      // Embed view keeps Office for the web in viewing mode instead of opening its editor.
      viewerUrl.searchParams.set("action", "embedview");
      return viewerUrl.toString();
    },
    [getDirectPreviewFileUrl, siteUrl, sp]
  );

  React.useEffect(() => {
    if (!open || selectedItems.length !== 1 || !getFilePreviewType(selectedItems[0])) {
      setSelectedPreviewItem(null);
      return;
    }

    setSelectedPreviewItem(selectedItems[0]);
  }, [open, selectedItems]);

  React.useEffect(() => {
    if (!open || !selectedPreviewItem || getFilePreviewType(selectedPreviewItem) !== "office") {
      setOfficePreviewUrl("");
      setOfficePreviewError(null);
      setIsLoadingOfficePreview(false);
      return;
    }

    let isCancelled = false;
    setOfficePreviewUrl("");
    setOfficePreviewError(null);
    setIsLoadingOfficePreview(true);

    getOfficeViewUrl(selectedPreviewItem)
      .then((url) => {
        if (!isCancelled) setOfficePreviewUrl(url);
      })
      .catch((err) => {
        console.error("Could not prepare Office document preview:", err);
        if (!isCancelled) {
          setOfficePreviewError("Could not load the Office preview. Check your SharePoint access and try again.");
          setIsLoadingOfficePreview(false);
        }
      });

    return () => {
      isCancelled = true;
    };
  }, [getOfficeViewUrl, open, selectedPreviewItem]);

  // Track field values
  const [selectedProducts, setSelectedProducts] = React.useState<IProductLookupItem[]>([]);
  const selectedProductsRef = React.useRef<IProductLookupItem[]>(selectedProducts);
  const updateSelectedProducts = React.useCallback(
    (value: React.SetStateAction<IProductLookupItem[]>) => {
      const nextProducts =
        typeof value === "function" ? value(selectedProductsRef.current) : value;
      selectedProductsRef.current = nextProducts;
      setSelectedProducts(nextProducts);
    },
    []
  );
  const [productsModified, setProductsModified] = React.useState<boolean>(false);
  const [metadataSource, setMetadataSource] = React.useState<"product" | "client" | null>(null);
  const [businessLines, setBusinessLines] = React.useState<string[]>([]);
  const [countrySoldToValues, setCountrySoldToValues] = React.useState<string[]>([]);

  const [selectedClients, setSelectedClients] = React.useState<IClientLookupItem[]>([]);
  const [clientsModified, setClientsModified] = React.useState<boolean>(false);

  const [selectedDocumentType, setSelectedDocumentType] = React.useState<IDocumentTypeItem | null>(null);
  const [documentTypeModified, setDocumentTypeModified] = React.useState<boolean>(false);

  const [selectedSubDocumentTypes, setSelectedSubDocumentTypes] = React.useState<ISubDocumentTypeItem[]>([]);
  const [subDocumentTypesModified, setSubDocumentTypesModified] = React.useState<boolean>(false);

  // Additional 11 Fields
  const [issuedBy, setIssuedBy] = React.useState<string>("");
  const [issuedByModified, setIssuedByModified] = React.useState<boolean>(false);

  const [supplier, setSupplier] = React.useState<string>("");
  const [supplierModified, setSupplierModified] = React.useState<boolean>(false);

  const [supplierEmail, setSupplierEmail] = React.useState<string>("");
  const [supplierEmailModified, setSupplierEmailModified] = React.useState<boolean>(false);

  const [confidentiality, setConfidentiality] = React.useState<string>("");
  const [confidentialityModified, setConfidentialityModified] = React.useState<boolean>(false);

  const [documentDate, setDocumentDate] = React.useState<Dayjs | null>(null);
  const [documentDateModified, setDocumentDateModified] = React.useState<boolean>(false);

  const [expiryDate, setExpiryDate] = React.useState<Dayjs | null>(null);
  const [expiryDateModified, setExpiryDateModified] = React.useState<boolean>(false);

  const [nextReviewDate, setNextReviewDate] = React.useState<Dayjs | null>(null);
  const [nextReviewDateModified, setNextReviewDateModified] = React.useState<boolean>(false);

  const [documentLanguage, setDocumentLanguage] = React.useState<string[]>([]);
  const [documentLanguageModified, setDocumentLanguageModified] = React.useState<boolean>(false);

  const [documentStatus, setDocumentStatus] = React.useState<string>("");
  const [documentStatusModified, setDocumentStatusModified] = React.useState<boolean>(false);

  const [customerName, setCustomerName] = React.useState<string>("");
  const [customerNameModified, setCustomerNameModified] = React.useState<boolean>(false);

  const [batchNumber, setBatchNumber] = React.useState<string>("");
  const [batchNumberModified, setBatchNumberModified] = React.useState<boolean>(false);

  // Type-ahead lookup state
  const [productSearchText, setProductSearchText] = React.useState("");
  const [productOptions, setProductOptions] = React.useState<IProductLookupItem[]>([]);
  const [productLoading, setProductLoading] = React.useState(false);

  const [clientSearchText, setClientSearchText] = React.useState("");
  const [clientOptions, setClientOptions] = React.useState<IClientLookupItem[]>([]);
  const [clientLoading, setClientLoading] = React.useState(false);

  const updateClientsForProducts = React.useCallback(
    (products: IProductLookupItem[]) => {
      const relatedClients = getRelatedClients(products, selectedClients, clientOptions);
      if (!haveSameClients(selectedClients, relatedClients)) {
        setSelectedClients(relatedClients);
        setClientsModified(true);
      }
    },
    [clientOptions, selectedClients]
  );

  // Save execution state
  const [isSaving, setIsSaving] = React.useState<boolean>(false);
  const [isLoadingLiveItem, setIsLoadingLiveItem] = React.useState<boolean>(false);
  const [errorMessage, setErrorMessage] = React.useState<string | null>(null);

  // Permission state
  const [hasWritePermission, setHasWritePermission] = React.useState<boolean>(true);
  const [isCheckingPermissions, setIsCheckingPermissions] = React.useState<boolean>(false);

  // Choices lists
  const availableIssuedBy = libraryChoices?.issuedBy || [];
  const availableConfidentiality = libraryChoices?.confidentiality || [];
  const availableDocStatus = libraryChoices?.documentStatus || [];
  const availableDocLanguage = libraryChoices?.documentLanguage || [];
  const isNextReviewDateReadOnly = Boolean(documentDate?.isValid() || expiryDate?.isValid());
  const isConfidentialityReadOnly = Boolean(
    selectedDocumentType?.Title?.trim() ||
      selectedSubDocumentTypes.some((item) => Boolean(item.Title?.trim()))
  );

  // Reset form
  const resetForm = React.useCallback(() => {
    updateSelectedProducts([]);
    setMetadataSource(null);
    setBusinessLines([]);
    setCountrySoldToValues([]);
    setSelectedClients([]);
    setSelectedDocumentType(null);
    setSelectedSubDocumentTypes([]);
    setIssuedBy("");
    setSupplier("");
    setSupplierEmail("");
    setConfidentiality("");
    setDocumentDate(null);
    setExpiryDate(null);
    setNextReviewDate(null);
    setDocumentLanguage([]);
    setDocumentStatus("");
    setCustomerName("");
    setBatchNumber("");

    setProductsModified(false);
    setClientsModified(false);
    setDocumentTypeModified(false);
    setSubDocumentTypesModified(false);
    setIssuedByModified(false);
    setSupplierModified(false);
    setSupplierEmailModified(false);
    setConfidentialityModified(false);
    setDocumentDateModified(false);
    setExpiryDateModified(false);
    setNextReviewDateModified(false);
    setDocumentLanguageModified(false);
    setDocumentStatusModified(false);
    setCustomerNameModified(false);
    setBatchNumberModified(false);

    setErrorMessage(null);
    setIsSaving(false);
    setIsLoadingLiveItem(false);
    setHasWritePermission(true);
    setIsCheckingPermissions(false);
  }, [updateSelectedProducts]);

  // Initialize form state whenever dialog opens or selectedItems change
  React.useEffect(() => {
    if (!open) {
      resetForm();
      return;
    }

    setErrorMessage(null);
    setIsSaving(false);

    // Explicit check: If more than 1 item is selected, ALWAYS wipe and keep empty!
    if (selectedItems.length > 1) {
      resetForm();
      return;
    }

    if (selectedItems.length === 1 && typeof selectedItems[0]?.id === "number") {
      const itemToEdit = selectedItems[0];
      const itemId = itemToEdit.id as number;
      // 1. Initial fast prepopulate from data table to avoid layout shift
      setMetadataSource(null);
      updateSelectedProducts(itemToEdit.PIMProduct ? [...itemToEdit.PIMProduct] : []);
      setProductsModified(true);
      setBusinessLines(splitMultiValue(itemToEdit.BusinessLine));
      setCountrySoldToValues(splitMultiValue(itemToEdit.CountrySoldTo));

      const clientStr = itemToEdit.ManufacturerSearchText || "";
      if (clientStr.trim()) {
        const parsedClients = clientStr
          .split(/[\r\n;,]+/)
          .map((c) => c.trim())
          .filter(Boolean)
          .map((name, idx) => ({ ID: -(idx + 1), Title: name }));
        setSelectedClients(parsedClients);
      } else {
        setSelectedClients([]);
      }
      setClientsModified(true);

      const docTypeTitle = (itemToEdit.DocumentTypeSearchText || "").trim();
      const matchedDocType = documentTypes.find(
        (d) => (d.Title || "").toLowerCase() === docTypeTitle.toLowerCase()
      ) || (docTypeTitle ? { ID: 0, Title: docTypeTitle } : null);
      setSelectedDocumentType(matchedDocType);
      setDocumentTypeModified(true);

      const subDocStr = itemToEdit.SubDocumentTypeSearchText || "";
      if (subDocStr.trim()) {
        const subTitles = subDocStr
          .split(/[\r\n;,]+/)
          .map((s) => s.trim())
          .filter(Boolean);
        const matchedSubs = subTitles.map((title, idx) => {
          const found = allSubDocumentTypes.find(
            (s) => (s.Title || "").toLowerCase() === title.toLowerCase()
          );
          return found || { ID: -(idx + 1), Title: title };
        });
        setSelectedSubDocumentTypes(matchedSubs);
      } else {
        setSelectedSubDocumentTypes([]);
      }
      setSubDocumentTypesModified(true);

      setIssuedBy(itemToEdit.IssuedBy || "");
      setIssuedByModified(true);

      setSupplier(itemToEdit.Supplier || "");
      setSupplierModified(true);

      setSupplierEmail(itemToEdit.SupplierEmail || "");
      setSupplierEmailModified(true);

      setConfidentiality(itemToEdit.Confidentiality || "");
      setConfidentialityModified(true);

      setDocumentDate(itemToEdit.DocumentDate ? dayjs(itemToEdit.DocumentDate) : null);
      setDocumentDateModified(true);

      setExpiryDate(itemToEdit.ExpiryDate ? dayjs(itemToEdit.ExpiryDate) : null);
      setExpiryDateModified(true);

      setNextReviewDate(itemToEdit.NextReviewDate ? dayjs(itemToEdit.NextReviewDate) : null);
      setNextReviewDateModified(true);

      const langList = itemToEdit.DocumentLanguage
        ? itemToEdit.DocumentLanguage.split(/[\r\n;,]+/).map((s) => s.trim()).filter(Boolean)
        : [];
      setDocumentLanguage(langList);
      setDocumentLanguageModified(true);

      setDocumentStatus(itemToEdit.DocumentStatus || "");
      setDocumentStatusModified(true);

      setCustomerName(itemToEdit.CustomerName || "");
      setCustomerNameModified(true);

      setBatchNumber(itemToEdit.BatchNumber || "");
      setBatchNumberModified(true);

      // 2. Fetch authoritative live metadata directly from SharePoint List
      let isCancelled = false;
      setIsLoadingLiveItem(true);

      loadItemDetailsForEdit(sp, itemId)
        .then((live) => {
          if (isCancelled) return;

          // Products
          if (live.products && live.products.length > 0) {
            updateSelectedProducts(live.products);
          }
          if (live.businessLines !== undefined) setBusinessLines(live.businessLines);
          if (live.countrySoldTo !== undefined) setCountrySoldToValues(live.countrySoldTo);

          // Clients
          if (live.clients && live.clients.length > 0) {
            setSelectedClients(live.clients);
          }

          // Document Type: match with loaded documentTypes list for complete metadata
          if (live.documentType) {
            const found = documentTypes.find(
              (d) =>
                d.ID === live.documentType?.ID ||
                (d.Title || "").toLowerCase() === (live.documentType?.Title || "").toLowerCase()
            );
            setSelectedDocumentType(found || live.documentType);
          } else {
            setSelectedDocumentType(null);
          }

          // Sub Document Types: match with loaded allSubDocumentTypes list
          if (live.subDocumentTypes && live.subDocumentTypes.length > 0) {
            const matchedSubs = live.subDocumentTypes.map((sub) => {
              const found = allSubDocumentTypes.find(
                (s) =>
                  s.ID === sub.ID ||
                  (s.Title || "").toLowerCase() === (sub.Title || "").toLowerCase()
              );
              return found || sub;
            });
            setSelectedSubDocumentTypes(matchedSubs);
          } else {
            setSelectedSubDocumentTypes([]);
          }

          if (live.issuedBy !== undefined) setIssuedBy(live.issuedBy || "");
          if (live.supplier !== undefined) setSupplier(live.supplier || "");
          if (live.supplierEmail !== undefined) setSupplierEmail(live.supplierEmail || "");
          if (live.confidentiality !== undefined) setConfidentiality(live.confidentiality || "");
          if (live.documentDate !== undefined) setDocumentDate(live.documentDate ? dayjs(live.documentDate) : null);
          if (live.expiryDate !== undefined) setExpiryDate(live.expiryDate ? dayjs(live.expiryDate) : null);
          if (live.nextReviewDate !== undefined) setNextReviewDate(live.nextReviewDate ? dayjs(live.nextReviewDate) : null);
          if (live.documentLanguage !== undefined) setDocumentLanguage(live.documentLanguage || []);
          if (live.documentStatus !== undefined) setDocumentStatus(live.documentStatus || "");
          if (live.customerName !== undefined) setCustomerName(live.customerName || "");
          if (live.batchNumber !== undefined) setBatchNumber(live.batchNumber || "");
        })
        .catch((err) => {
          console.warn("Could not load live details, using table data fallback:", err);
        })
        .finally(() => {
          if (!isCancelled) {
            setIsLoadingLiveItem(false);
          }
        });

      return () => {
        isCancelled = true;
      };
    } else {
      resetForm();
    }
  }, [open, selectedItems, sp, documentTypes, allSubDocumentTypes, resetForm, updateSelectedProducts]);

  React.useEffect(() => {
    if (metadataSource === "product") {
      setBusinessLines(getProductValues(selectedProducts, "BusinessLine"));
      setCountrySoldToValues(getProductValues(selectedProducts, "Country"));
    } else if (metadataSource === "client") {
      setBusinessLines(getClientValues(selectedClients, "BusinessLine"));
      setCountrySoldToValues(getClientValues(selectedClients, "Country"));
    }
  }, [metadataSource, selectedProducts, selectedClients]);

  // Check user write permissions whenever dialog opens or selectedItems change
  React.useEffect(() => {
    if (!open || !sp) {
      setHasWritePermission(true);
      setIsCheckingPermissions(false);
      return;
    }

    let isCancelled = false;
    setIsCheckingPermissions(true);

    const itemIds = selectedItems
      .map((item) => item.id)
      .filter((id): id is number => typeof id === "number" && id > 0);

    const confidentialItems = selectedItems.filter(
      (item) => (item?.Confidentiality || "").trim().toLowerCase() === "confidential"
    );

    Promise.all([
      checkUserWritePermissions(sp, "Clients & Products", itemIds),
      confidentialItems.length > 0 && currentUserEmail
        ? Promise.all([
            getSenderUserAccessMatrix(sp, currentUserEmail),
            getCountryRegionMapping(sp),
          ])
        : Promise.resolve(null),
    ])
      .then(([hasLibPermission, matrixData]) => {
        if (isCancelled) return;
        if (!hasLibPermission) {
          setHasWritePermission(false);
          return;
        }

        if (confidentialItems.length > 0 && matrixData) {
          const [matrix, countryMap] = matrixData;
          const allAllowed = matrix
            ? confidentialItems.every((doc) =>
                checkConfidentialDocumentAccess(doc, matrix, countryMap, "edit")
              )
            : false;
          if (!allAllowed) {
            setHasWritePermission(false);
            setErrorMessage(
              "You do not have write permission in the User Access Matrix to edit confidential document(s). The Save button is disabled."
            );
            return;
          }
        }

        setHasWritePermission(true);
      })
      .catch((err) => {
        console.warn("checkUserWritePermissions error in EditPropertiesDialog:", err);
        if (!isCancelled) {
          setHasWritePermission(false);
        }
      })
      .finally(() => {
        if (!isCancelled) {
          setIsCheckingPermissions(false);
        }
      });

    return () => {
      isCancelled = true;
    };
  }, [open, sp, selectedItems, currentUserEmail]);

  // Product Autocomplete type-ahead
  React.useEffect(() => {
    const timer = setTimeout(async () => {
      if (productSearchText.trim().length >= 3 && sp) {
        setProductLoading(true);
        try {
          const results = await searchProductsService(sp, productSearchText.trim());
          setProductOptions(results);
        } catch (error) {
          console.error("EditPropertiesDialog searchProducts error:", error);
        } finally {
          setProductLoading(false);
        }
      } else {
        setProductOptions([]);
      }
    }, 300);

    return () => clearTimeout(timer);
  }, [productSearchText, sp]);

  // Client Autocomplete type-ahead
  React.useEffect(() => {
    const timer = setTimeout(async () => {
      if (clientSearchText.trim().length >= 3 && sp) {
        setClientLoading(true);
        try {
          const results = await searchClientsService(sp, clientSearchText.trim());
          setClientOptions(results);
        } catch (error) {
          console.error("EditPropertiesDialog searchClients error:", error);
        } finally {
          setClientLoading(false);
        }
      } else {
        setClientOptions([]);
      }
    }, 300);

    return () => clearTimeout(timer);
  }, [clientSearchText, sp]);

  // Filter Sub Document Types cascaded on selected Document Type
  const availableSubDocumentTypes = React.useMemo(() => {
    if (!selectedDocumentType || !selectedDocumentType.Title) {
      return allSubDocumentTypes;
    }
    const currentDocType = selectedDocumentType.Title.toLowerCase();
    return allSubDocumentTypes.filter(
      (sub) =>
        sub.DocumentType?.Title &&
        sub.DocumentType.Title.toLowerCase() === currentDocType
    );
  }, [allSubDocumentTypes, selectedDocumentType]);

  // Clean up selected sub document types if parent document type changed and no longer contains them
  const handleDocumentTypeChange = (
    _event: React.SyntheticEvent,
    newValue: IDocumentTypeItem | null
  ) => {
    setSelectedDocumentType(newValue);
    setDocumentTypeModified(true);

    if (newValue && newValue.Title) {
      const validForNewDocType = new Set(
        allSubDocumentTypes
          .filter(
            (s) =>
              s.DocumentType?.Title &&
              s.DocumentType.Title.toLowerCase() === newValue.Title.toLowerCase()
          )
          .map((s) => (s.Title || "").toLowerCase())
      );

      setSelectedSubDocumentTypes((prev) =>
        prev.filter((item) => validForNewDocType.has((item.Title || "").toLowerCase()))
      );
    }
  };

  // Paste handlers for all controls (Autocomplete, Dropdown, Date, Text)
  const handlePasteProducts = React.useCallback(
    async (text: string) => {
      const parts = parseDelimitedText(text);
      if (parts.length === 0) return;

      const currentKeys = new Set(
        selectedProductsRef.current.map(
          (p) => `${p.Title || ""} ${p.PIMProductName || ""}`.trim().toLowerCase()
        )
      );
      const newEntries: IProductLookupItem[] = [];

      parts.forEach((raw) => {
        let parsedTitle = "";
        let parsedName = "";

        const pimMatch = raw.match(/\b(PIM\d+)\b/i);
        if (pimMatch) {
          parsedTitle = pimMatch[1].toUpperCase();
          parsedName = raw.replace(pimMatch[0], "").replace(/^[|\s:-]+|[|\s:-]+$/g, "").trim();
        } else if (raw.includes("|")) {
          const segments = raw.split("|").map((s) => s.trim());
          parsedTitle = segments[0] || "";
          parsedName = segments.slice(1).join(" ").trim();
        } else if (raw.includes(" : ")) {
          const segments = raw.split(" : ").map((s) => s.trim());
          parsedTitle = segments[0] || "";
          parsedName = segments.slice(1).join(" ").trim();
        } else {
          const spaceIdx = raw.indexOf(" ");
          if (spaceIdx > 0 && /^[A-Z0-9_-]+$/.test(raw.substring(0, spaceIdx))) {
            parsedTitle = raw.substring(0, spaceIdx).trim();
            parsedName = raw.substring(spaceIdx + 1).trim();
          } else {
            parsedName = raw.trim();
            parsedTitle = "";
          }
        }

        const key = `${parsedTitle} ${parsedName}`.trim().toLowerCase();
        if (key && !currentKeys.has(key)) {
          currentKeys.add(key);

          const matched = productOptions.find((p) => {
            const optTitle = (p.Title || "").trim().toLowerCase();
            const optName = (p.PIMProductName || "").trim().toLowerCase();
            const targetTitle = parsedTitle.toLowerCase();
            const targetName = parsedName.toLowerCase();

            if (targetTitle && optTitle) {
              return targetTitle === optTitle;
            }
            if (targetName && optName) {
              return targetName === optName;
            }
            return false;
          });

          if (matched) {
            newEntries.push(matched);
          } else {
            newEntries.push({
              ID: -Math.floor(Math.random() * 100000),
              Title: parsedTitle,
              PIMProductName: parsedName,
            });
          }
        }
      });

      if (newEntries.length > 0) {
        const nextProducts = [...selectedProductsRef.current, ...newEntries];
        updateSelectedProducts(nextProducts);
        setProductsModified(true);
        setMetadataSource("product");
        setProductSearchText("");

        if (sp) {
          const unresolved = newEntries.filter((p) => !p.ID || p.ID <= 0 || !p.TermGuid);
          if (unresolved.length > 0) {
            resolveProductsBatch(sp, unresolved)
              .then((resolved) => {
                const updatedProducts = selectedProductsRef.current.map((item) => {
                  const itemTitle = (item.Title || "").trim().toLowerCase();
                  const itemName = (item.PIMProductName || "").trim().toLowerCase();

                  const found = resolved.find((r) => {
                    const rTitle = (r.Title || "").trim().toLowerCase();
                    const rName = (r.PIMProductName || "").trim().toLowerCase();

                    // Priority 1: Match by Product Code (Title)
                    if (itemTitle && rTitle) {
                      return itemTitle === rTitle;
                    }
                    // Priority 2: Only match by ProductName if item had no Product Code
                    if (!itemTitle && itemName && rName) {
                      return itemName === rName;
                    }
                    return false;
                  });
                  return found ? { ...item, ...found } : item;
                });
                updateSelectedProducts(updatedProducts);
                updateClientsForProducts(updatedProducts);
              })
              .catch((err) => {
                console.warn("Could not batch resolve pasted products immediately:", err);
              });
          } else {
            updateClientsForProducts(nextProducts);
          }
        } else {
          updateClientsForProducts(nextProducts);
        }
      }
    },
    [productOptions, sp, updateClientsForProducts, updateSelectedProducts]
  );

  const handlePasteClients = React.useCallback(
    (text: string) => {
      const parts = parseDelimitedText(text);
      if (parts.length === 0) return;

      const currentTitles = new Set(
        selectedClients.map((c) => (c.Title || "").trim().toLowerCase())
      );
      const newEntries: IClientLookupItem[] = [];

      parts.forEach((title) => {
        const lower = title.toLowerCase();
        if (!currentTitles.has(lower)) {
          currentTitles.add(lower);
          const found = clientOptions.find(
            (c) => (c.Title || "").toLowerCase() === lower
          );
          newEntries.push(
            found || { ID: -Math.floor(Math.random() * 100000), Title: title }
          );
        }
      });

      if (newEntries.length > 0) {
        setSelectedClients((prev) => [...prev, ...newEntries]);
        setClientsModified(true);
        setMetadataSource("client");
        setClientSearchText("");
      }
    },
    [selectedClients, clientOptions]
  );

  const handlePasteDocumentType = React.useCallback(
    (text: string) => {
      const parts = parseDelimitedText(text);
      if (parts.length === 0) return;
      const target = parts[0].toLowerCase();
      const found = documentTypes.find(
        (d) =>
          (d.Title || "").toLowerCase() === target ||
          (d.ShortTitle && d.ShortTitle.toLowerCase() === target)
      );
      if (found) {
        handleDocumentTypeChange({} as any, found);
      } else if (parts[0]) {
        handleDocumentTypeChange({} as any, { ID: 0, Title: parts[0] });
      }
    },
    [documentTypes, handleDocumentTypeChange]
  );

  const handlePasteSubDocumentType = React.useCallback(
    (text: string) => {
      const parts = parseDelimitedText(text);
      if (parts.length === 0) return;

      const currentTitles = new Set(
        selectedSubDocumentTypes.map((s) => (s.Title || "").toLowerCase())
      );
      const newEntries: ISubDocumentTypeItem[] = [];

      parts.forEach((title) => {
        const lower = title.toLowerCase();
        if (!currentTitles.has(lower)) {
          currentTitles.add(lower);
          const found = availableSubDocumentTypes.find(
            (s) => (s.Title || "").toLowerCase() === lower
          );
          newEntries.push(
            found || { ID: -Math.floor(Math.random() * 100000), Title: title }
          );
        }
      });

      if (newEntries.length > 0) {
        setSelectedSubDocumentTypes((prev) => [...prev, ...newEntries]);
        setSubDocumentTypesModified(true);
      }
    },
    [selectedSubDocumentTypes, availableSubDocumentTypes]
  );

  const handlePasteIssuedBy = React.useCallback(
    (text: string) => {
      const clean = text.trim();
      const matched = availableIssuedBy.find(
        (opt) => opt.toLowerCase() === clean.toLowerCase()
      );
      setIssuedBy(matched || clean);
      setIssuedByModified(true);
    },
    [availableIssuedBy]
  );

  const handlePasteSupplier = React.useCallback((text: string) => {
    setSupplier(text.trim().toUpperCase());
    setSupplierModified(true);
  }, []);

  const handlePasteSupplierEmail = React.useCallback((text: string) => {
    setSupplierEmail(text.trim());
    setSupplierEmailModified(true);
  }, []);

  const handlePasteConfidentiality = React.useCallback(
    (text: string) => {
      const clean = text.trim();
      const matched = availableConfidentiality.find(
        (opt) => opt.toLowerCase() === clean.toLowerCase()
      );
      setConfidentiality(matched || clean);
      setConfidentialityModified(true);
    },
    [availableConfidentiality]
  );

  const handlePasteDocumentStatus = React.useCallback(
    (text: string) => {
      const clean = text.trim();
      const matched = availableDocStatus.find(
        (opt) => opt.toLowerCase() === clean.toLowerCase()
      );
      setDocumentStatus(matched || clean);
      setDocumentStatusModified(true);
    },
    [availableDocStatus]
  );

  const handlePasteDocumentLanguage = React.useCallback(
    (text: string) => {
      const parts = parseDelimitedText(text);
      if (parts.length === 0) return;

      const currentSet = new Set(documentLanguage.map((l) => l.toLowerCase()));
      const newLangs = [...documentLanguage];

      parts.forEach((p) => {
        const matched = availableDocLanguage.find(
          (opt) => opt.toLowerCase() === p.toLowerCase()
        );
        const toAdd = matched || p;
        if (!currentSet.has(toAdd.toLowerCase())) {
          currentSet.add(toAdd.toLowerCase());
          newLangs.push(toAdd);
        }
      });

      setDocumentLanguage(newLangs);
      setDocumentLanguageModified(true);
    },
    [documentLanguage, availableDocLanguage]
  );

  const handlePasteDocumentDate = React.useCallback((text: string) => {
    const d = parseDateString(text);
    if (d) {
      setDocumentDate(d);
      setDocumentDateModified(true);
    }
  }, []);

  const handlePasteExpiryDate = React.useCallback((text: string) => {
    const d = parseDateString(text);
    if (d) {
      setExpiryDate(d);
      setExpiryDateModified(true);
    }
  }, []);

  const handlePasteNextReviewDate = React.useCallback((text: string) => {
    const d = parseDateString(text);
    if (d) {
      setNextReviewDate(d);
      setNextReviewDateModified(true);
    }
  }, []);

  const handlePasteCustomerName = React.useCallback((text: string) => {
    setCustomerName(text.trim().toUpperCase());
    setCustomerNameModified(true);
  }, []);

  const handlePasteBatchNumber = React.useCallback((text: string) => {
    setBatchNumber(text.trim().toUpperCase());
    setBatchNumberModified(true);
  }, []);

  const handleSave = async (): Promise<void> => {
    if (!hasWritePermission) {
      setErrorMessage("You do not have write permission to update items in this library.");
      return;
    }

    const itemIds = selectedItems.map((item) => item.id).filter((id): id is number => typeof id === "number" && id > 0);
    if (itemIds.length === 0) {
      setErrorMessage("No valid item IDs selected.");
      return;
    }

    if (
      isBulkEdit &&
      !productsModified &&
      !clientsModified &&
      !documentTypeModified &&
      !subDocumentTypesModified &&
      !issuedByModified &&
      !supplierModified &&
      !supplierEmailModified &&
      !confidentialityModified &&
      !documentDateModified &&
      !expiryDateModified &&
      !nextReviewDateModified &&
      !documentLanguageModified &&
      !documentStatusModified &&
      !customerNameModified &&
      !batchNumberModified
    ) {
      setErrorMessage("Please modify at least one field to apply changes to the selected items.");
      return;
    }

    try {
      setIsSaving(true);
      setErrorMessage(null);

      const payload: IEditPropertiesPayload = {
        productsModified,
        selectedProducts: productsModified ? selectedProducts : undefined,

        clientsModified,
        selectedClients: clientsModified ? selectedClients : undefined,

        documentTypeModified,
        selectedDocumentType: documentTypeModified ? selectedDocumentType : undefined,

        subDocumentTypesModified,
        selectedSubDocumentTypes: subDocumentTypesModified ? selectedSubDocumentTypes : undefined,

        issuedByModified,
        issuedBy: issuedByModified ? issuedBy : undefined,

        supplierModified,
        supplier: supplierModified ? (supplier ? supplier.trim().toUpperCase() : "") : undefined,

        supplierEmailModified,
        supplierEmail: supplierEmailModified ? supplierEmail : undefined,

        confidentialityModified,
        confidentiality: confidentialityModified ? confidentiality || undefined : undefined,

        documentDateModified,
        documentDate: documentDateModified ? (documentDate ? documentDate.toDate() : null) : undefined,

        expiryDateModified,
        expiryDate: expiryDateModified ? (expiryDate ? expiryDate.toDate() : null) : undefined,

        nextReviewDateModified,
        nextReviewDate: nextReviewDateModified ? (nextReviewDate ? nextReviewDate.toDate() : null) : undefined,

        documentLanguageModified,
        documentLanguage: documentLanguageModified ? documentLanguage : undefined,

        documentStatusModified,
        documentStatus: documentStatusModified ? documentStatus || undefined : undefined,

        customerNameModified,
        customerName: customerNameModified ? (customerName ? customerName.trim().toUpperCase() : "") : undefined,

        batchNumberModified,
        batchNumber: batchNumberModified ? (batchNumber ? batchNumber.trim().toUpperCase() : "") : undefined,
      };

      await onSave(payload, itemIds);
      onClose();
    } catch (err: any) {
      console.error("Save error in EditPropertiesDialog:", err);
      setErrorMessage(err?.message || "Failed to save changes. Please check permissions and try again.");
    } finally {
      setIsSaving(false);
    }
  };

  const hasModifications = isBulkEdit
    ? productsModified ||
      clientsModified ||
      documentTypeModified ||
      subDocumentTypesModified ||
      issuedByModified ||
      supplierModified ||
      supplierEmailModified ||
      confidentialityModified ||
      documentDateModified ||
      expiryDateModified ||
      nextReviewDateModified ||
      documentLanguageModified ||
      documentStatusModified ||
      customerNameModified ||
      batchNumberModified
    : true;
  const selectedPreviewType = selectedPreviewItem ? getFilePreviewType(selectedPreviewItem) : null;
  const selectedPreviewUrl =
    selectedPreviewItem && selectedPreviewType === "pdf"
      ? getDirectPreviewFileUrl(selectedPreviewItem)
      : selectedPreviewType === "office"
      ? officePreviewUrl
      : "";

  return (
    <Dialog
      open={open}
      onClose={isSaving ? undefined : onClose}
      PaperComponent={DraggablePaper}
      aria-labelledby="draggable-edit-dialog-title"
      fullWidth
      maxWidth={false}
      PaperProps={{
        sx: {
          height: "90vh",
          maxHeight: "95vh",
          width: "calc(100vw - 32px)",
          maxWidth: "none",
          m: "16px",
        },
      }}
    >
      <DialogTitle
        id="draggable-edit-dialog-title"
        sx={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          py: 1.5,
          px: 2.5,
          borderBottom: "1px solid #e0e0e0",
          cursor: "grab",
          userSelect: "none",
          touchAction: "none",
        }}
      >
        <Box sx={{ display: "flex", alignItems: "flex-start", gap: 1, overflow: "hidden", pr: 1 }}>
          <DragIndicatorIcon sx={{ color: "text.secondary", fontSize: 20, opacity: 0.6, cursor: "grab", mt: 0.4 }} />
          <EditNoteIcon color="primary" sx={{ fontSize: 24, mt: 0.2 }} />
          <Box sx={{ overflow: "hidden" }}>
            {isBulkEdit ? (
              <>
                <Typography variant="h6" sx={{ fontSize: "16px", fontWeight: 600 }}>
                  Edit Properties ({selectedItems.length} items selected)
                </Typography>
                <Typography variant="caption" sx={{ fontSize: "11px", color: "text.secondary", display: "block", mb: 0.5 }}>
                  Fields left untouched will retain their existing values on each document.
                </Typography>
                <Box sx={{ display: "flex", flexWrap: "wrap", gap: 0.5, mt: 0.5, maxHeight: "60px", overflowY: "auto" }}>
                  {selectedItems.map((item) => {
                    const fileUrl = getAbsoluteFileUrl(item.fileUrl);
                    const previewType = getFilePreviewType(item);
                    const newTabUrl =
                      previewType === "pdf" ? getDirectPreviewFileUrl(item) : fileUrl;
                    const isPreviewSelected = selectedPreviewItem === item;
                    return (
                      <Chip
                        key={item.id ?? item.filename}
                        size="small"
                        label={item.filename || "Document"}
                        component={fileUrl && !previewType ? "a" : "div"}
                        href={fileUrl && !previewType ? fileUrl : undefined}
                        target={fileUrl && !previewType ? "_blank" : undefined}
                        rel={fileUrl && !previewType ? "noopener noreferrer" : undefined}
                        data-interception="off"
                        clickable={Boolean(fileUrl || previewType)}
                        color={isPreviewSelected ? "primary" : "default"}
                        variant={isPreviewSelected ? "filled" : "outlined"}
                        deleteIcon={previewType === "office" ? undefined : <OpenInNewIcon sx={{ fontSize: 13 }} />}
                        onDelete={
                          newTabUrl && previewType !== "office"
                            ? (e: any) => {
                                if (e?.preventDefault) e.preventDefault();
                                if (e?.stopPropagation) e.stopPropagation();
                                window.open(newTabUrl, "_blank", "noopener,noreferrer");
                              }
                            : undefined
                        }
                        onMouseDown={(e: React.MouseEvent) => e.stopPropagation()}
                        onClick={
                          fileUrl || previewType
                            ? (e: React.MouseEvent) => {
                                e.preventDefault();
                                e.stopPropagation();
                                if (previewType) {
                                  setSelectedPreviewItem(item);
                                } else if (fileUrl) {
                                  window.open(fileUrl, "_blank", "noopener,noreferrer");
                                }
                              }
                            : undefined
                        }
                        sx={{
                          fontSize: "11px",
                          height: "22px",
                          cursor: fileUrl || previewType ? "pointer" : "default",
                          "& .MuiChip-deleteIcon": {
                            fontSize: "13px",
                            color: "primary.main",
                          },
                        }}
                      />
                    );
                  })}
                </Box>
              </>
            ) : (
              <Box sx={{ display: "flex", alignItems: "center", gap: 0.75, flexWrap: "wrap" }}>
                <Typography variant="h6" component="span" sx={{ fontSize: "16px", fontWeight: 600 }}>
                  Edit Properties -
                </Typography>
                {(() => {
                  const fileUrl = getAbsoluteFileUrl(singleItem?.fileUrl);
                  const fileName = singleItem?.filename || "Document";
                  const previewType = singleItem ? getFilePreviewType(singleItem) : null;
                  if (singleItem && previewType) {
                    const viewerUrl =
                      previewType === "pdf" ? getDirectPreviewFileUrl(singleItem) : officePreviewUrl;
                    return (
                      <Box sx={{ display: "inline-flex", alignItems: "center", gap: 0.25, minWidth: 0 }}>
                        <Typography
                          variant="h6"
                          component="span"
                          sx={{ fontSize: "15px", fontWeight: 600, wordBreak: "break-all" }}
                        >
                          {fileName}
                        </Typography>
                        {viewerUrl && (
                          <Tooltip title={`Open ${previewType === "office" ? "Office viewer" : "PDF"} in new tab`}>
                            <IconButton
                              size="small"
                              onMouseDown={(e: React.MouseEvent) => e.stopPropagation()}
                              onClick={(e: React.MouseEvent) => {
                                e.stopPropagation();
                                window.open(viewerUrl, "_blank", "noopener,noreferrer");
                              }}
                            >
                              <OpenInNewIcon sx={{ fontSize: 16 }} />
                            </IconButton>
                          </Tooltip>
                        )}
                      </Box>
                    );
                  }
                  if (fileUrl) {
                    return (
                      <Link
                        href={fileUrl}
                        target="_blank"
                        rel="noopener noreferrer"
                        data-interception="off"
                        onMouseDown={(e: React.MouseEvent) => e.stopPropagation()}
                        onClick={(e: React.MouseEvent) => {
                          e.preventDefault();
                          e.stopPropagation();
                          window.open(fileUrl, "_blank", "noopener,noreferrer");
                        }}
                        sx={{
                          display: "inline-flex",
                          alignItems: "center",
                          gap: 0.5,
                          fontSize: "15px",
                          fontWeight: 600,
                          color: "primary.main",
                          textDecoration: "underline",
                          cursor: "pointer",
                          wordBreak: "break-all",
                          "&:hover": {
                            color: "primary.dark",
                          },
                        }}
                        title="Open document in new tab"
                      >
                        <span>{fileName}</span>
                        <OpenInNewIcon sx={{ fontSize: 16 }} />
                      </Link>
                    );
                  }
                  return (
                    <Typography variant="h6" component="span" sx={{ fontSize: "16px", fontWeight: 600 }}>
                      {fileName}
                    </Typography>
                  );
                })()}
              </Box>
            )}
          </Box>
        </Box>
        <IconButton size="small" onClick={onClose} disabled={isSaving}>
          <CloseIcon fontSize="small" />
        </IconButton>
      </DialogTitle>
      {(isLoadingLiveItem || isCheckingPermissions) && <LinearProgress sx={{ height: "2.5px" }} />}

      <DialogContent sx={{ p: 2, display: "flex", minHeight: 0, overflow: "hidden" }}>
        <Box
          sx={{
            display: "grid",
            gridTemplateColumns: { xs: "minmax(0, 1fr)", md: "minmax(0, 1fr) minmax(0, 1fr)" },
            gridTemplateRows: { xs: "minmax(0, 1fr) minmax(220px, 0.8fr)", md: "minmax(0, 1fr)" },
            gap: 2,
            flex: 1,
            minHeight: 0,
            width: "100%",
          }}
        >
        <Box
          sx={{
            display: "grid",
            gridTemplateColumns: { xs: "minmax(0, 1fr)", lg: "repeat(2, minmax(0, 1fr))" },
            alignContent: "start",
            gap: 2.5,
            minWidth: 0,
            minHeight: 0,
            overflowY: "auto",
            pr: 0.5,
          }}
        >
        {errorMessage && (
          <Alert severity="error" onClose={() => setErrorMessage(null)} sx={{ gridColumn: { xs: "auto", lg: "1 / -1" } }}>
            {errorMessage}
          </Alert>
        )}

        {!hasWritePermission && !isCheckingPermissions && (
          <Alert severity="warning" sx={{ gridColumn: { xs: "auto", lg: "1 / -1" } }}>
            You do not have write permission to edit items in the <strong>Clients &amp; Products</strong> library. The Save button is disabled.
          </Alert>
        )}

        {/* 1. Product (multi value) */}
        <Box sx={{ gridColumn: { xs: "auto", lg: "1 / -1" } }}>
          <Box sx={{ display: "flex", justifyContent: "space-between", alignItems: "center", mb: 0.5 }}>
            <Box sx={{ display: "flex", alignItems: "center", gap: 0.5 }}>
              <Typography variant="subtitle2" sx={{ fontSize: "13px", fontWeight: 600 }}>
                Product
              </Typography>
              <CopyFieldValueButton value={selectedProducts} label="products" />
              <PasteFieldValueButton onPasteText={handlePasteProducts} label="products" />
            </Box>
            {isBulkEdit && productsModified && (
              <Chip
                size="small"
                label="Modified"
                color="primary"
                variant="outlined"
                sx={{ height: "18px", fontSize: "10px" }}
              />
            )}
          </Box>
          <Autocomplete
            multiple
            fullWidth
            size="small"
            loading={productLoading}
            options={productOptions}
            value={selectedProducts}
            ListboxComponent={EditProductListbox}
            getOptionLabel={(option: IProductLookupItem) =>
              option ? `${option.Title ?? ""} ${option.PIMProductName ?? ""}`.trim() : ""
            }
            isOptionEqualToValue={(option, value) =>
              option.ID === value.ID ||
              (Boolean(option.Title) &&
                Boolean(value.Title) &&
                `${option.Title} ${option.PIMProductName ?? ""}`.trim().toLowerCase() ===
                  `${value.Title} ${value.PIMProductName ?? ""}`.trim().toLowerCase())
            }
            onInputChange={(_e, val) => setProductSearchText(val)}
            onChange={(_e, newValue) => {
              updateSelectedProducts(newValue);
              setProductsModified(true);
              setMetadataSource("product");
              updateClientsForProducts(newValue);
            }}
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
                      wordBreak: "break-word",
                    }}
                  >
                    {`${option.Title || ""} ${option.PIMProductName || ""}`.trim() || "-"}
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
                placeholder={isBulkEdit && !productsModified ? "(Unchanged - type to modify)" : "Search product (min 3 chars)..."}
                onPaste={(e) => {
                  const pasteData = e.clipboardData?.getData("text");
                  if (pasteData) {
                    e.preventDefault();
                    void handlePasteProducts(pasteData);
                  }
                }}
                InputProps={{
                  ...params.InputProps,
                  endAdornment: (
                    <>
                      {productLoading ? <CircularProgress color="inherit" size={18} /> : null}
                      {params.InputProps.endAdornment}
                    </>
                  ),
                }}
              />
            )}
          />
        </Box>

        <Box
          sx={{
            display: { xs: "grid", lg: "contents" },
            gridTemplateColumns: "1fr",
            gap: 2,
          }}
        >
          <TextField
            fullWidth
            size="small"
            label="Business Line"
            value={businessLines.join(", ")}
            InputProps={{ readOnly: true }}
          />
          <TextField
            fullWidth
            size="small"
            label="Country Sold To"
            value={countrySoldToValues.join(", ")}
            InputProps={{ readOnly: true }}
          />
        </Box>

        {/* 2. Client (multi value) */}
        <Box sx={{ gridColumn: { xs: "auto", lg: "1 / -1" } }}>
          <Box sx={{ display: "flex", justifyContent: "space-between", alignItems: "center", mb: 0.5 }}>
            <Box sx={{ display: "flex", alignItems: "center", gap: 0.5 }}>
              <Typography variant="subtitle2" sx={{ fontSize: "13px", fontWeight: 600 }}>
                Client
              </Typography>
              <CopyFieldValueButton value={selectedClients} label="clients" />
              <PasteFieldValueButton
                onPasteText={handlePasteClients}
                label="clients"
                disabled={selectedProducts.length > 0}
              />
            </Box>
            {isBulkEdit && clientsModified && (
              <Chip
                size="small"
                label="Modified"
                color="primary"
                variant="outlined"
                sx={{ height: "18px", fontSize: "10px" }}
              />
            )}
          </Box>
          <Autocomplete
            multiple
            fullWidth
            size="small"
            readOnly={selectedProducts.length > 0}
            loading={clientLoading}
            options={clientOptions}
            value={selectedClients}
            getOptionLabel={(option: IClientLookupItem) => (option ? option.Title : "")}
            isOptionEqualToValue={(option, value) =>
              option.ID === value.ID ||
              (Boolean(option.Title) &&
                Boolean(value.Title) &&
                option.Title.trim().toLowerCase() === value.Title.trim().toLowerCase())
            }
            onInputChange={(_e, val) => setClientSearchText(val)}
            onChange={(_e, newValue) => {
              setSelectedClients(newValue);
              setClientsModified(true);
              setMetadataSource("client");
            }}
            renderInput={(params) => (
              <TextField
                {...params}
                size="small"
                placeholder={isBulkEdit && !clientsModified ? "(Unchanged - type to modify)" : "Search client (min 3 chars)..."}
                onPaste={(e) => {
                  const pasteData = e.clipboardData?.getData("text");
                  if (pasteData) {
                    e.preventDefault();
                    handlePasteClients(pasteData);
                  }
                }}
                InputProps={{
                  ...params.InputProps,
                  endAdornment: (
                    <>
                      {clientLoading ? <CircularProgress color="inherit" size={18} /> : null}
                      {params.InputProps.endAdornment}
                    </>
                  ),
                }}
              />
            )}
          />
        </Box>

        {/* 3. Document Type (single value) & 4. Sub Document Type (multi value) */}
        <Box
          sx={{
            display: { xs: "grid", lg: "contents" },
            gridTemplateColumns: "1fr",
            gap: 2,
          }}
        >
          {/* Document Type */}
          <Box>
            <Box sx={{ display: "flex", justifyContent: "space-between", alignItems: "center", mb: 0.5 }}>
              <Box sx={{ display: "flex", alignItems: "center", gap: 0.5 }}>
                <Typography variant="subtitle2" sx={{ fontSize: "13px", fontWeight: 600 }}>
                  Document Type
                </Typography>
                <CopyFieldValueButton value={selectedDocumentType} label="document type" />
                <PasteFieldValueButton onPasteText={handlePasteDocumentType} label="document type" />
              </Box>
              {isBulkEdit && documentTypeModified && (
                <Chip
                  size="small"
                  label="Modified"
                  color="primary"
                  variant="outlined"
                  sx={{ height: "18px", fontSize: "10px" }}
                />
              )}
            </Box>
            <Autocomplete
              fullWidth
              size="small"
              options={documentTypes}
              value={selectedDocumentType}
              getOptionLabel={(option: IDocumentTypeItem) =>
                option ? (option.ShortTitle ? `${option.Title} (${option.ShortTitle})` : option.Title) : ""
              }
              isOptionEqualToValue={(option, value) =>
                option.ID === value.ID ||
                (option.Title || "").toLowerCase() === (value.Title || "").toLowerCase()
              }
              onChange={handleDocumentTypeChange}
              renderInput={(params) => (
                <TextField
                  {...params}
                  size="small"
                  placeholder={isBulkEdit && !documentTypeModified ? "(Unchanged)" : "Select document type..."}
                  onPaste={(e) => {
                    const parts = parseDelimitedPaste(e);
                    if (parts && parts.length > 0) {
                      handlePasteDocumentType(parts.join("; "));
                    }
                  }}
                />
              )}
            />
          </Box>

          {/* Sub Document Type */}
          <Box>
            <Box sx={{ display: "flex", justifyContent: "space-between", alignItems: "center", mb: 0.5 }}>
              <Box sx={{ display: "flex", alignItems: "center", gap: 0.5 }}>
                <Typography variant="subtitle2" sx={{ fontSize: "13px", fontWeight: 600 }}>
                  Sub Document Type
                </Typography>
                <CopyFieldValueButton value={selectedSubDocumentTypes} label="sub document types" />
                <PasteFieldValueButton onPasteText={handlePasteSubDocumentType} label="sub document types" />
              </Box>
              {isBulkEdit && subDocumentTypesModified && (
                <Chip
                  size="small"
                  label="Modified"
                  color="primary"
                  variant="outlined"
                  sx={{ height: "18px", fontSize: "10px" }}
                />
              )}
            </Box>
            <Autocomplete
              multiple
              fullWidth
              size="small"
              disabled={availableSubDocumentTypes.length === 0}
              options={availableSubDocumentTypes}
              value={selectedSubDocumentTypes}
              getOptionLabel={(option: ISubDocumentTypeItem) => (option ? option.Title : "")}
              isOptionEqualToValue={(option, value) =>
                option.ID === value.ID ||
                (option.Title || "").toLowerCase() === (value.Title || "").toLowerCase()
              }
              onChange={(_e, newValue) => {
                setSelectedSubDocumentTypes(newValue);
                setSubDocumentTypesModified(true);
              }}
              renderInput={(params) => (
                <TextField
                  {...params}
                  size="small"
                  placeholder={
                    isBulkEdit && !subDocumentTypesModified
                      ? "(Unchanged)"
                      : availableSubDocumentTypes.length === 0
                      ? "No sub types available"
                      : "Select sub document type(s)..."
                  }
                  onPaste={(e) => {
                    const parts = parseDelimitedPaste(e);
                    if (parts && parts.length > 0) {
                      handlePasteSubDocumentType(parts.join("; "));
                    }
                  }}
                />
              )}
            />
          </Box>
        </Box>

        {/* --- 11 Additional Fields --- */}

        {/* Row 1: Issued by, Issuer Name (Supplier), Document Provider Email */}
        <Box
          sx={{
            display: { xs: "grid", lg: "contents" },
            gridTemplateColumns: "1fr",
            gap: 2,
          }}
        >
          {/* Issued by */}
          <Box>
            <Box sx={{ display: "flex", justifyContent: "space-between", alignItems: "center", mb: 0.5 }}>
              <Box sx={{ display: "flex", alignItems: "center", gap: 0.5 }}>
                <Typography variant="subtitle2" sx={{ fontSize: "13px", fontWeight: 600 }}>
                  Issued by
                </Typography>
                <CopyFieldValueButton value={issuedBy} label="issued by" />
                <PasteFieldValueButton onPasteText={handlePasteIssuedBy} label="issued by" />
              </Box>
              {isBulkEdit && issuedByModified && (
                <Chip size="small" label="Modified" color="primary" variant="outlined" sx={{ height: "18px", fontSize: "10px" }} />
              )}
            </Box>
            <FormControl
              fullWidth
              size="small"
              onPaste={(e) => {
                const t = e.clipboardData?.getData("text");
                if (t) handlePasteIssuedBy(t);
              }}
            >
              <Select
                value={issuedBy}
                displayEmpty
                onChange={(e) => {
                  setIssuedBy(e.target.value);
                  setIssuedByModified(true);
                }}
                input={<OutlinedInput />}
                renderValue={(selected) => {
                  if (!selected) {
                    return <span style={{ color: "#aaa" }}>{isBulkEdit && !issuedByModified ? "(Unchanged)" : "Select..."}</span>;
                  }
                  return selected;
                }}
              >
                <MenuItem value=""><em>None / Clear</em></MenuItem>
                {availableIssuedBy.map((opt) => (
                  <MenuItem key={opt} value={opt}>{opt}</MenuItem>
                ))}
              </Select>
            </FormControl>
          </Box>

          {/* Issuer Name (Supplier) */}
          <Box>
            <Box sx={{ display: "flex", justifyContent: "space-between", alignItems: "center", mb: 0.5 }}>
              <Box sx={{ display: "flex", alignItems: "center", gap: 0.5 }}>
                <Typography variant="subtitle2" sx={{ fontSize: "13px", fontWeight: 600 }}>
                  Issuer Name
                </Typography>
                <CopyFieldValueButton value={supplier} label="issuer name" />
                <PasteFieldValueButton onPasteText={handlePasteSupplier} label="issuer name" />
              </Box>
              {isBulkEdit && supplierModified && (
                <Chip size="small" label="Modified" color="primary" variant="outlined" sx={{ height: "18px", fontSize: "10px" }} />
              )}
            </Box>
            <TextField
              fullWidth
              size="small"
              value={supplier}
              placeholder={isBulkEdit && !supplierModified ? "(Unchanged)" : "Enter issuer name..."}
              inputProps={{ style: { textTransform: "uppercase" } }}
              onChange={(e) => {
                setSupplier(e.target.value);
                setSupplierModified(true);
              }}
              onBlur={() => {
                if (supplier) setSupplier((prev) => prev.toUpperCase());
              }}
            />
          </Box>

          {/* Document Provider Email (Supplier_x0020_Email) */}
          <Box>
            <Box sx={{ display: "flex", justifyContent: "space-between", alignItems: "center", mb: 0.5 }}>
              <Box sx={{ display: "flex", alignItems: "center", gap: 0.5 }}>
                <Typography variant="subtitle2" sx={{ fontSize: "13px", fontWeight: 600 }}>
                  Document Provider Email
                </Typography>
                <CopyFieldValueButton value={supplierEmail} label="provider email" />
                <PasteFieldValueButton onPasteText={handlePasteSupplierEmail} label="provider email" />
              </Box>
              {isBulkEdit && supplierEmailModified && (
                <Chip size="small" label="Modified" color="primary" variant="outlined" sx={{ height: "18px", fontSize: "10px" }} />
              )}
            </Box>
            <TextField
              fullWidth
              size="small"
              value={supplierEmail}
              placeholder={isBulkEdit && !supplierEmailModified ? "(Unchanged)" : "Enter provider email..."}
              onChange={(e) => {
                setSupplierEmail(e.target.value);
                setSupplierEmailModified(true);
              }}
            />
          </Box>
        </Box>

        {/* Row 2: Confidentiality, Document Status, Document Language */}
        <Box
          sx={{
            display: { xs: "grid", lg: "contents" },
            gridTemplateColumns: "1fr",
            gap: 2,
          }}
        >
          {/* Confidentiality */}
          <Box>
            <Box sx={{ display: "flex", justifyContent: "space-between", alignItems: "center", mb: 0.5 }}>
              <Box sx={{ display: "flex", alignItems: "center", gap: 0.5 }}>
                <Typography variant="subtitle2" sx={{ fontSize: "13px", fontWeight: 600 }}>
                  Confidentiality
                </Typography>
                <CopyFieldValueButton value={confidentiality} label="confidentiality" />
                <PasteFieldValueButton
                  onPasteText={handlePasteConfidentiality}
                  label="confidentiality"
                  disabled={isConfidentialityReadOnly}
                />
              </Box>
              {isBulkEdit && confidentialityModified && (
                <Chip size="small" label="Modified" color="primary" variant="outlined" sx={{ height: "18px", fontSize: "10px" }} />
              )}
            </Box>
            <FormControl
              fullWidth
              size="small"
              onPaste={(e) => {
                const t = e.clipboardData?.getData("text");
                if (t && !isConfidentialityReadOnly) handlePasteConfidentiality(t);
              }}
            >
              <Select
                value={confidentiality}
                displayEmpty
                readOnly={isConfidentialityReadOnly}
                onChange={(e) => {
                  setConfidentiality(e.target.value);
                  setConfidentialityModified(true);
                }}
                input={<OutlinedInput />}
                renderValue={(selected) => {
                  if (!selected) {
                    return <span style={{ color: "#aaa" }}>{isBulkEdit && !confidentialityModified ? "(Unchanged)" : "Select..."}</span>;
                  }
                  return selected;
                }}
              >
                <MenuItem value=""><em>None / Clear</em></MenuItem>
                {availableConfidentiality.map((opt) => (
                  <MenuItem key={opt} value={opt}>{opt}</MenuItem>
                ))}
              </Select>
            </FormControl>
          </Box>

          {/* Document Status */}
          <Box>
            <Box sx={{ display: "flex", justifyContent: "space-between", alignItems: "center", mb: 0.5 }}>
              <Box sx={{ display: "flex", alignItems: "center", gap: 0.5 }}>
                <Typography variant="subtitle2" sx={{ fontSize: "13px", fontWeight: 600 }}>
                  Document Status
                </Typography>
                <CopyFieldValueButton value={documentStatus} label="document status" />
                <PasteFieldValueButton onPasteText={handlePasteDocumentStatus} label="document status" />
              </Box>
              {isBulkEdit && documentStatusModified && (
                <Chip size="small" label="Modified" color="primary" variant="outlined" sx={{ height: "18px", fontSize: "10px" }} />
              )}
            </Box>
            <FormControl
              fullWidth
              size="small"
              onPaste={(e) => {
                const t = e.clipboardData?.getData("text");
                if (t) handlePasteDocumentStatus(t);
              }}
            >
              <Select
                value={documentStatus}
                displayEmpty
                onChange={(e) => {
                  setDocumentStatus(e.target.value);
                  setDocumentStatusModified(true);
                }}
                input={<OutlinedInput />}
                renderValue={(selected) => {
                  if (!selected) {
                    return <span style={{ color: "#aaa" }}>{isBulkEdit && !documentStatusModified ? "(Unchanged)" : "Select..."}</span>;
                  }
                  return selected;
                }}
              >
                <MenuItem value=""><em>None / Clear</em></MenuItem>
                {availableDocStatus.map((opt) => (
                  <MenuItem key={opt} value={opt}>{opt}</MenuItem>
                ))}
              </Select>
            </FormControl>
          </Box>

          {/* Document Language (Multi-select) */}
          <Box>
            <Box sx={{ display: "flex", justifyContent: "space-between", alignItems: "center", mb: 0.5 }}>
              <Box sx={{ display: "flex", alignItems: "center", gap: 0.5 }}>
                <Typography variant="subtitle2" sx={{ fontSize: "13px", fontWeight: 600 }}>
                  Document Language
                </Typography>
                <CopyFieldValueButton value={documentLanguage} label="document language" />
                <PasteFieldValueButton onPasteText={handlePasteDocumentLanguage} label="document language" />
              </Box>
              {isBulkEdit && documentLanguageModified && (
                <Chip size="small" label="Modified" color="primary" variant="outlined" sx={{ height: "18px", fontSize: "10px" }} />
              )}
            </Box>
            <FormControl
              fullWidth
              size="small"
              onPaste={(e) => {
                const t = e.clipboardData?.getData("text");
                if (t) handlePasteDocumentLanguage(t);
              }}
            >
              <Select
                multiple
                value={documentLanguage}
                onChange={(e) => {
                  const val = e.target.value;
                  setDocumentLanguage(typeof val === "string" ? val.split(",") : val);
                  setDocumentLanguageModified(true);
                }}
                input={<OutlinedInput />}
                renderValue={(selected) => {
                  if (!selected || (selected as string[]).length === 0) {
                    return <span style={{ color: "#aaa" }}>{isBulkEdit && !documentLanguageModified ? "(Unchanged)" : "Select language(s)..."}</span>;
                  }
                  return (selected as string[]).join(", ");
                }}
              >
                {availableDocLanguage.map((opt) => (
                  <MenuItem key={opt} value={opt}>
                    <Checkbox size="small" checked={documentLanguage.indexOf(opt) > -1} />
                    <ListItemText primary={opt} />
                  </MenuItem>
                ))}
              </Select>
            </FormControl>
          </Box>
        </Box>

        {/* Row 3: Document Date, Expiry Date, Next Review Date */}
        <LocalizationProvider dateAdapter={AdapterDayjs}>
          <Box
            sx={{
              display: { xs: "grid", lg: "contents" },
              gridTemplateColumns: "1fr",
              gap: 2,
            }}
          >
            {/* Document Date */}
            <Box>
              <Box sx={{ display: "flex", justifyContent: "space-between", alignItems: "center", mb: 0.5 }}>
                <Box sx={{ display: "flex", alignItems: "center", gap: 0.5 }}>
                  <Typography variant="subtitle2" sx={{ fontSize: "13px", fontWeight: 600 }}>
                    Document Date
                  </Typography>
                  <CopyFieldValueButton value={documentDate} label="document date" />
                  <PasteFieldValueButton onPasteText={handlePasteDocumentDate} label="document date" />
                </Box>
                {isBulkEdit && documentDateModified && (
                  <Chip size="small" label="Modified" color="primary" variant="outlined" sx={{ height: "18px", fontSize: "10px" }} />
                )}
              </Box>
              <DatePicker
                format="DD/MM/YYYY"
                value={documentDate}
                onChange={(newValue) => {
                  setDocumentDate(newValue);
                  setDocumentDateModified(true);
                }}
                slotProps={{
                  textField: {
                    size: "small",
                    fullWidth: true,
                    placeholder: isBulkEdit && !documentDateModified ? "(Unchanged)" : "DD/MM/YYYY",
                  },
                  field: { clearable: true },
                }}
              />
            </Box>

            {/* Expiry Date */}
            <Box>
              <Box sx={{ display: "flex", justifyContent: "space-between", alignItems: "center", mb: 0.5 }}>
                <Box sx={{ display: "flex", alignItems: "center", gap: 0.5 }}>
                  <Typography variant="subtitle2" sx={{ fontSize: "13px", fontWeight: 600 }}>
                    Expiry Date
                  </Typography>
                  <CopyFieldValueButton value={expiryDate} label="expiry date" />
                  <PasteFieldValueButton onPasteText={handlePasteExpiryDate} label="expiry date" />
                </Box>
                {isBulkEdit && expiryDateModified && (
                  <Chip size="small" label="Modified" color="primary" variant="outlined" sx={{ height: "18px", fontSize: "10px" }} />
                )}
              </Box>
              <DatePicker
                format="DD/MM/YYYY"
                value={expiryDate}
                onChange={(newValue) => {
                  setExpiryDate(newValue);
                  setExpiryDateModified(true);
                }}
                slotProps={{
                  textField: {
                    size: "small",
                    fullWidth: true,
                    placeholder: isBulkEdit && !expiryDateModified ? "(Unchanged)" : "DD/MM/YYYY",
                  },
                  field: { clearable: true },
                }}
              />
            </Box>

            {/* Next Review Date */}
            <Box>
              <Box sx={{ display: "flex", justifyContent: "space-between", alignItems: "center", mb: 0.5 }}>
                <Box sx={{ display: "flex", alignItems: "center", gap: 0.5 }}>
                  <Typography variant="subtitle2" sx={{ fontSize: "13px", fontWeight: 600 }}>
                    Next Review Date
                  </Typography>
                  <CopyFieldValueButton value={nextReviewDate} label="next review date" />
                  <PasteFieldValueButton
                    onPasteText={handlePasteNextReviewDate}
                    label="next review date"
                    disabled={isNextReviewDateReadOnly}
                  />
                </Box>
                {isBulkEdit && nextReviewDateModified && (
                  <Chip size="small" label="Modified" color="primary" variant="outlined" sx={{ height: "18px", fontSize: "10px" }} />
                )}
              </Box>
              <DatePicker
                format="DD/MM/YYYY"
                value={nextReviewDate}
                readOnly={isNextReviewDateReadOnly}
                onChange={(newValue) => {
                  setNextReviewDate(newValue);
                  setNextReviewDateModified(true);
                }}
                slotProps={{
                  textField: {
                    size: "small",
                    fullWidth: true,
                    placeholder: isBulkEdit && !nextReviewDateModified ? "(Unchanged)" : "DD/MM/YYYY",
                  },
                  field: { clearable: true },
                }}
              />
            </Box>
          </Box>
        </LocalizationProvider>

        {/* Row 4: Customer Name, Batch Number */}
        <Box
          sx={{
            display: { xs: "grid", lg: "contents" },
            gridTemplateColumns: "1fr",
            gap: 2,
          }}
        >
          {/* Customer Name */}
          <Box>
            <Box sx={{ display: "flex", justifyContent: "space-between", alignItems: "center", mb: 0.5 }}>
              <Box sx={{ display: "flex", alignItems: "center", gap: 0.5 }}>
                <Typography variant="subtitle2" sx={{ fontSize: "13px", fontWeight: 600 }}>
                  Customer Name
                </Typography>
                <CopyFieldValueButton value={customerName} label="customer name" />
                <PasteFieldValueButton onPasteText={handlePasteCustomerName} label="customer name" />
              </Box>
              {isBulkEdit && customerNameModified && (
                <Chip size="small" label="Modified" color="primary" variant="outlined" sx={{ height: "18px", fontSize: "10px" }} />
              )}
            </Box>
            <TextField
              fullWidth
              size="small"
              value={customerName}
              placeholder={isBulkEdit && !customerNameModified ? "(Unchanged)" : "Enter customer name..."}
              inputProps={{ style: { textTransform: "uppercase" } }}
              onChange={(e) => {
                setCustomerName(e.target.value);
                setCustomerNameModified(true);
              }}
              onBlur={() => {
                if (customerName) setCustomerName((prev) => prev.toUpperCase());
              }}
            />
          </Box>

          {/* Batch Number */}
          <Box>
            <Box sx={{ display: "flex", justifyContent: "space-between", alignItems: "center", mb: 0.5 }}>
              <Box sx={{ display: "flex", alignItems: "center", gap: 0.5 }}>
                <Typography variant="subtitle2" sx={{ fontSize: "13px", fontWeight: 600 }}>
                  Batch Number
                </Typography>
                <CopyFieldValueButton value={batchNumber} label="batch number" />
                <PasteFieldValueButton onPasteText={handlePasteBatchNumber} label="batch number" />
              </Box>
              {isBulkEdit && batchNumberModified && (
                <Chip size="small" label="Modified" color="primary" variant="outlined" sx={{ height: "18px", fontSize: "10px" }} />
              )}
            </Box>
            <TextField
              fullWidth
              size="small"
              value={batchNumber}
              placeholder={isBulkEdit && !batchNumberModified ? "(Unchanged)" : "Enter batch number..."}
              inputProps={{ style: { textTransform: "uppercase" } }}
              onChange={(e) => {
                setBatchNumber(e.target.value);
                setBatchNumberModified(true);
              }}
              onBlur={() => {
                if (batchNumber) setBatchNumber((prev) => prev.toUpperCase());
              }}
            />
          </Box>
        </Box>

        {/* Read-only document metadata: single-item edit only */}
        {!isBulkEdit && (
          <Box
            sx={{
              gridColumn: { xs: "auto", lg: "1 / -1" },
              display: "grid",
              gridTemplateColumns: { xs: "1fr", sm: "repeat(2, minmax(0, 1fr))" },
              gap: 2,
              pt: 0.5,
              borderTop: "1px solid",
              borderColor: "divider",
            }}
          >
            {[
              { label: "Document ID", value: documentId },
              { label: "Document Version", value: documentVersion },
              { label: "Created", value: created },
              { label: "Created By", value: createdBy },
              { label: "Modified", value: modified },
              { label: "Modified By", value: modifiedBy },
            ].map((field) => (
              <Box key={field.label}>
                <Typography variant="subtitle2" sx={{ fontSize: "13px", fontWeight: 600, mb: 0.5 }}>
                  {field.label}
                </Typography>
                <TextField
                  fullWidth
                  size="small"
                  value={field.value}
                  inputProps={{ readOnly: true }}
                />
              </Box>
            ))}
          </Box>
        )}
        </Box>

        <Box
          sx={{
            display: "flex",
            flexDirection: "column",
            minWidth: 0,
            minHeight: 0,
            border: "1px solid",
            borderColor: "divider",
            borderRadius: 1,
            overflow: "hidden",
            backgroundColor: "#f5f5f5",
          }}
        >
          <Box
            sx={{
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              gap: 1,
              px: 1.5,
              py: 1,
              borderBottom: "1px solid",
              borderColor: "divider",
              backgroundColor: "#fff",
            }}
          >
            <Typography variant="subtitle2" sx={{ fontWeight: 600, flexShrink: 0 }}>
              {selectedPreviewType === "office" ? "Office Preview" : "Document Preview"}
            </Typography>
            {selectedPreviewType === "office" && (
              <Chip size="small" label="View only" color="info" variant="outlined" sx={{ height: "20px" }} />
            )}
            {selectedPreviewItem && (
              <Typography
                variant="caption"
                color="text.secondary"
                title={selectedPreviewItem.filename}
                sx={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", flex: 1 }}
              >
                {selectedPreviewItem.filename}
              </Typography>
            )}
            {selectedPreviewUrl && (
              <Tooltip title="Open viewer in new tab">
                <IconButton
                  size="small"
                  onClick={() => window.open(selectedPreviewUrl, "_blank", "noopener,noreferrer")}
                  sx={{ flexShrink: 0 }}
                >
                  <OpenInNewIcon fontSize="small" />
                </IconButton>
              </Tooltip>
            )}
          </Box>
          <Box sx={{ flex: 1, minHeight: 0, position: "relative", display: "flex", alignItems: "center", justifyContent: "center" }}>
            {isLoadingOfficePreview && (
              <Box
                sx={{
                  position: "absolute",
                  inset: 0,
                  zIndex: 1,
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  gap: 1,
                  backgroundColor: "rgba(255,255,255,0.8)",
                }}
              >
                <CircularProgress size={24} />
                <Typography variant="body2" color="text.secondary">
                  Loading Office preview...
                </Typography>
              </Box>
            )}
            {officePreviewError ? (
              <Alert severity="error" sx={{ m: 2 }}>
                {officePreviewError}
              </Alert>
            ) : selectedPreviewUrl ? (
              <iframe
                key={selectedPreviewUrl}
                src={selectedPreviewUrl}
                title={selectedPreviewItem?.filename || "Document preview"}
                onLoad={() => {
                  if (selectedPreviewType === "office") setIsLoadingOfficePreview(false);
                }}
                onError={() => {
                  if (selectedPreviewType === "office") {
                    setOfficePreviewError("The Office viewer could not load this document.");
                    setIsLoadingOfficePreview(false);
                  }
                }}
                style={{ width: "100%", height: "100%", border: "none", backgroundColor: "#fff" }}
              />
            ) : (
              <Typography variant="body2" color="text.secondary" sx={{ p: 2, textAlign: "center" }}>
                {selectedPreviewItem
                  ? selectedPreviewType === "office"
                    ? "The Office preview is not available."
                    : "No file URL is available for this PDF."
                  : isBulkEdit
                  ? "Select a PDF or Office file above to preview it."
                  : "This file type opens in a new tab from its filename."}
              </Typography>
            )}
          </Box>
        </Box>
        </Box>
      </DialogContent>

      <DialogActions sx={{ px: 2.5, py: 1.5, borderTop: "1px solid #e0e0e0" }}>
        <Button onClick={onClose} size="small" disabled={isSaving} sx={{ textTransform: "none" }}>
          Cancel
        </Button>
        <Button
          variant="contained"
          color="primary"
          size="small"
          onClick={handleSave}
          disabled={isSaving || isCheckingPermissions || !hasWritePermission || !hasModifications}
          sx={{ textTransform: "none", fontWeight: 600, minWidth: "90px" }}
        >
          {isSaving ? <CircularProgress size={16} color="inherit" /> : "Save"}
        </Button>
      </DialogActions>
    </Dialog>
  );
};

export default EditPropertiesDialog;
