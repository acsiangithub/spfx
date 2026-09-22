import * as React from "react";
import { SPFI } from "@pnp/sp";
import Dialog from "@mui/material/Dialog";
import DialogTitle from "@mui/material/DialogTitle";
import DialogContent from "@mui/material/DialogContent";
import DialogActions from "@mui/material/DialogActions";
import Box from "@mui/material/Box";
import Typography from "@mui/material/Typography";
import Button from "@mui/material/Button";
import IconButton from "@mui/material/IconButton";
import TextField from "@mui/material/TextField";
import Autocomplete from "@mui/material/Autocomplete";
import CircularProgress from "@mui/material/CircularProgress";
import LinearProgress from "@mui/material/LinearProgress";
import Alert from "@mui/material/Alert";
import Chip from "@mui/material/Chip";
import CloseIcon from "@mui/icons-material/Close";
import EditNoteIcon from "@mui/icons-material/EditNote";

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
  IEditPropertiesPayload,
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

export interface IEditPropertiesDialogProps {
  open: boolean;
  onClose: () => void;
  selectedItems: doclib_AllProducts[];
  sp: SPFI;
  documentTypes: IDocumentTypeItem[];
  allSubDocumentTypes: ISubDocumentTypeItem[];
  onSave: (payload: IEditPropertiesPayload, itemIds: number[]) => Promise<void>;
}

export const EditPropertiesDialog: React.FC<IEditPropertiesDialogProps> = ({
  open,
  onClose,
  selectedItems,
  sp,
  documentTypes,
  allSubDocumentTypes,
  onSave,
}) => {
  const isBulkEdit = selectedItems.length > 1;
  const singleItem = selectedItems.length === 1 ? selectedItems[0] : null;

  // Track field values
  const [selectedProducts, setSelectedProducts] = React.useState<IProductLookupItem[]>([]);
  const [productsModified, setProductsModified] = React.useState<boolean>(false);

  const [selectedClients, setSelectedClients] = React.useState<IClientLookupItem[]>([]);
  const [clientsModified, setClientsModified] = React.useState<boolean>(false);

  const [selectedDocumentType, setSelectedDocumentType] = React.useState<IDocumentTypeItem | null>(null);
  const [documentTypeModified, setDocumentTypeModified] = React.useState<boolean>(false);

  const [selectedSubDocumentTypes, setSelectedSubDocumentTypes] = React.useState<ISubDocumentTypeItem[]>([]);
  const [subDocumentTypesModified, setSubDocumentTypesModified] = React.useState<boolean>(false);

  // Type-ahead lookup state
  const [productSearchText, setProductSearchText] = React.useState("");
  const [productOptions, setProductOptions] = React.useState<IProductLookupItem[]>([]);
  const [productLoading, setProductLoading] = React.useState(false);

  const [clientSearchText, setClientSearchText] = React.useState("");
  const [clientOptions, setClientOptions] = React.useState<IClientLookupItem[]>([]);
  const [clientLoading, setClientLoading] = React.useState(false);

  // Save execution state
  const [isSaving, setIsSaving] = React.useState<boolean>(false);
  const [isLoadingLiveItem, setIsLoadingLiveItem] = React.useState<boolean>(false);
  const [errorMessage, setErrorMessage] = React.useState<string | null>(null);

  // Initialize form state whenever dialog opens or selectedItems change
  React.useEffect(() => {
    if (!open) {
      setErrorMessage(null);
      setIsSaving(false);
      setIsLoadingLiveItem(false);
      return;
    }

    setErrorMessage(null);
    setIsSaving(false);

    if (singleItem && singleItem.id) {
      // 1. Initial fast prepopulate from data table to avoid layout shift
      setSelectedProducts(singleItem.PIMProduct ? [...singleItem.PIMProduct] : []);
      setProductsModified(true);

      const clientStr = singleItem.ManufacturerSearchText || "";
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

      const docTypeTitle = (singleItem.DocumentTypeSearchText || "").trim();
      const matchedDocType = documentTypes.find(
        (d) => (d.Title || "").toLowerCase() === docTypeTitle.toLowerCase()
      ) || (docTypeTitle ? { ID: 0, Title: docTypeTitle } : null);
      setSelectedDocumentType(matchedDocType);
      setDocumentTypeModified(true);

      const subDocStr = singleItem.SubDocumentTypeSearchText || "";
      if (subDocStr.trim()) {
        const subTitles = subDocStr
          .split(",")
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

      // 2. Fetch authoritative live metadata directly from SharePoint List
      let isCancelled = false;
      setIsLoadingLiveItem(true);

      loadItemDetailsForEdit(sp, singleItem.id)
        .then((live) => {
          if (isCancelled) return;

          // Products
          if (live.products && live.products.length > 0) {
            setSelectedProducts(live.products);
          }

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
      // Bulk Edit: Start empty; only flag as modified when user touches a control
      setSelectedProducts([]);
      setProductsModified(false);

      setSelectedClients([]);
      setClientsModified(false);

      setSelectedDocumentType(null);
      setDocumentTypeModified(false);

      setSelectedSubDocumentTypes([]);
      setSubDocumentTypesModified(false);
      setIsLoadingLiveItem(false);
    }
  }, [open, singleItem, sp, documentTypes, allSubDocumentTypes]);

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

  const handleSave = async (): Promise<void> => {
    const itemIds = selectedItems.map((item) => item.id).filter((id): id is number => typeof id === "number" && id > 0);
    if (itemIds.length === 0) {
      setErrorMessage("No valid item IDs selected.");
      return;
    }

    if (isBulkEdit && !productsModified && !clientsModified && !documentTypeModified && !subDocumentTypesModified) {
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
    ? productsModified || clientsModified || documentTypeModified || subDocumentTypesModified
    : true;

  return (
    <Dialog
      open={open}
      onClose={isSaving ? undefined : onClose}
      fullWidth
      maxWidth="md"
      PaperProps={{
        sx: {
          borderRadius: "8px",
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
          px: 2.5,
          borderBottom: "1px solid #e0e0e0",
        }}
      >
        <Box sx={{ display: "flex", alignItems: "center", gap: 1 }}>
          <EditNoteIcon color="primary" sx={{ fontSize: 24 }} />
          <Box>
            <Typography variant="h6" sx={{ fontSize: "16px", fontWeight: 600 }}>
              {isBulkEdit
                ? `Edit Properties (${selectedItems.length} items selected)`
                : `Edit Properties - ${singleItem?.filename || "Document"}`}
            </Typography>
            {isBulkEdit && (
              <Typography variant="caption" sx={{ fontSize: "11px", color: "text.secondary" }}>
                Fields left untouched will retain their existing values on each document.
              </Typography>
            )}
          </Box>
        </Box>
        <IconButton size="small" onClick={onClose} disabled={isSaving}>
          <CloseIcon fontSize="small" />
        </IconButton>
      </DialogTitle>
      {isLoadingLiveItem && <LinearProgress sx={{ height: "2.5px" }} />}

      <DialogContent sx={{ p: 2.5, display: "flex", flexDirection: "column", gap: 2.5 }}>
        {errorMessage && (
          <Alert severity="error" onClose={() => setErrorMessage(null)}>
            {errorMessage}
          </Alert>
        )}

        {/* 1. Product (multi value) */}
        <Box>
          <Box sx={{ display: "flex", justifyContent: "space-between", alignItems: "center", mb: 0.5 }}>
            <Typography variant="subtitle2" sx={{ fontSize: "13px", fontWeight: 600 }}>
              Product
            </Typography>
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
              setSelectedProducts(newValue);
              setProductsModified(true);
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

        {/* 2. Client (multi value) */}
        <Box>
          <Box sx={{ display: "flex", justifyContent: "space-between", alignItems: "center", mb: 0.5 }}>
            <Typography variant="subtitle2" sx={{ fontSize: "13px", fontWeight: 600 }}>
              Client
            </Typography>
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
            }}
            renderInput={(params) => (
              <TextField
                {...params}
                size="small"
                placeholder={isBulkEdit && !clientsModified ? "(Unchanged - type to modify)" : "Search client (min 3 chars)..."}
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
            display: "grid",
            gridTemplateColumns: { xs: "1fr", sm: "1fr 1fr" },
            gap: 2,
          }}
        >
          {/* Document Type */}
          <Box>
            <Box sx={{ display: "flex", justifyContent: "space-between", alignItems: "center", mb: 0.5 }}>
              <Typography variant="subtitle2" sx={{ fontSize: "13px", fontWeight: 600 }}>
                Document Type
              </Typography>
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
                />
              )}
            />
          </Box>

          {/* Sub Document Type */}
          <Box>
            <Box sx={{ display: "flex", justifyContent: "space-between", alignItems: "center", mb: 0.5 }}>
              <Typography variant="subtitle2" sx={{ fontSize: "13px", fontWeight: 600 }}>
                Sub Document Type
              </Typography>
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
                />
              )}
            />
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
          disabled={isSaving || !hasModifications}
          sx={{ textTransform: "none", fontWeight: 600, minWidth: "90px" }}
        >
          {isSaving ? <CircularProgress size={16} color="inherit" /> : "Save"}
        </Button>
      </DialogActions>
    </Dialog>
  );
};

export default EditPropertiesDialog;
