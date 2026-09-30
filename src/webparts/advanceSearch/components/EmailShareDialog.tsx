import * as React from "react";
import { SPFI } from "@pnp/sp";
import { sp as defaultSp, DEFAULT_SHARE_FLOW_URL } from "../AdvanceSearchWebPart";
import {
  postToShareFlow,
  parseRejectedDocumentNames,
  IShareFlowPayload,
  searchSharePointUsers,
  IPeoplePickerUserOption,
} from "../../../services/sharePointService";
import Dialog from "@mui/material/Dialog";
import DialogTitle from "@mui/material/DialogTitle";
import DialogContent from "@mui/material/DialogContent";
import DialogActions from "@mui/material/DialogActions";
import TextField from "@mui/material/TextField";
import Button from "@mui/material/Button";
import Typography from "@mui/material/Typography";
import Autocomplete from "@mui/material/Autocomplete";
import Chip from "@mui/material/Chip";
import Box from "@mui/material/Box";
import Avatar from "@mui/material/Avatar";
import CircularProgress from "@mui/material/CircularProgress";

export interface IEmailShareDialogProps {
  open: boolean;
  onClose: () => void;
  selectedItems?: any[];
  siteUrl: string;
  shareFlowUrl?: string;
  defaultSubject?: string;
  defaultMessage?: string;
  currentUserEmail?: string;
  sp?: SPFI;
}

interface IEmailPeoplePickerInputProps {
  field: "to" | "cc" | "bcc";
  label: string;
  values: string[];
  setValues: React.Dispatch<React.SetStateAction<string[]>>;
  inputValue: string;
  setInputValue: React.Dispatch<React.SetStateAction<string>>;
  placeholder: string;
  errorText: string;
  onFieldBlur: (currentAll: string[]) => void;
  onFieldChange: (field: "to" | "cc" | "bcc") => void;
  validateEmail: (email: string) => boolean;
  sp?: SPFI;
}

const EmailPeoplePickerInput: React.FC<IEmailPeoplePickerInputProps> = ({
  field,
  label,
  values,
  setValues,
  inputValue,
  setInputValue,
  placeholder,
  errorText,
  onFieldBlur,
  onFieldChange,
  validateEmail,
  sp,
}) => {
  const [options, setOptions] = React.useState<IPeoplePickerUserOption[]>([]);
  const [loading, setLoading] = React.useState<boolean>(false);
  const debounceRef = React.useRef<any>(null);

  React.useEffect(() => {
    if (debounceRef.current) {
      clearTimeout(debounceRef.current);
    }

    const query = inputValue.trim();
    if (!sp || query.length < 2 || query.includes(",") || query.includes(";")) {
      setOptions([]);
      setLoading(false);
      return;
    }

    setLoading(true);
    debounceRef.current = setTimeout(async () => {
      try {
        const users = await searchSharePointUsers(sp, query, 7);
        setOptions(users);
      } catch (err) {
        console.warn("Failed to search users:", err);
        setOptions([]);
      } finally {
        setLoading(false);
      }
    }, 250);

    return () => {
      if (debounceRef.current) {
        clearTimeout(debounceRef.current);
      }
    };
  }, [inputValue, sp]);

  return (
    <Autocomplete
      multiple
      freeSolo
      size="small"
      options={options}
      loading={loading}
      filterOptions={(x) => x}
      value={values}
      inputValue={inputValue}
      getOptionLabel={(option: string | IPeoplePickerUserOption) => {
        if (typeof option === "string") {
          return option;
        }
        return option.email || option.displayText || "";
      }}
      isOptionEqualToValue={(option: any, val: any) => {
        const optEmail = typeof option === "string" ? option : option.email || option.displayText;
        const valEmail = typeof val === "string" ? val : val.email || val.displayText;
        return optEmail.toLowerCase() === valEmail.toLowerCase();
      }}
      onInputChange={(_event, newInputValue, reason) => {
        if (reason === "input") {
          if (newInputValue.includes(",") || newInputValue.includes(";")) {
            const parts = newInputValue
              .split(/[,;]/)
              .map((p) => p.trim())
              .filter(Boolean);
            setValues((prev) => Array.from(new Set([...prev, ...parts])));
            setInputValue("");
            setOptions([]);
            onFieldChange(field);
          } else {
            setInputValue(newInputValue);
          }
        } else if (reason === "reset" || reason === "clear") {
          setInputValue("");
        }
      }}
      onChange={(_event, newValue) => {
        const cleaned: string[] = [];
        (newValue as (string | IPeoplePickerUserOption)[]).forEach((val) => {
          if (typeof val === "string") {
            val.split(/[,;]/).forEach((item) => {
              const trimmed = item.trim();
              if (trimmed && cleaned.indexOf(trimmed) === -1) {
                cleaned.push(trimmed);
              }
            });
          } else if (val && typeof val === "object") {
            const email = (val.email || val.displayText || "").trim();
            if (email && cleaned.indexOf(email) === -1) {
              cleaned.push(email);
            }
          }
        });
        setValues(cleaned);
        setInputValue("");
        setOptions([]);
        onFieldChange(field);
      }}
      renderOption={(props, option) => {
        if (typeof option === "string") {
          return (
            <li {...props} key={props.key || option}>
              <Typography variant="body2">{option}</Typography>
            </li>
          );
        }

        const initials = (option.displayText || option.email || "?")
          .split(" ")
          .map((w: string) => w[0])
          .filter(Boolean)
          .slice(0, 2)
          .join("")
          .toUpperCase();

        const subtitle = [option.jobTitle, option.department].filter(Boolean).join(" • ");

        return (
          <li {...props} key={props.key || option.key || option.email}>
            <Box sx={{ display: "flex", alignItems: "center", gap: 1.25, py: 0.5, width: "100%" }}>
              <Avatar
                sx={{
                  width: 30,
                  height: 30,
                  fontSize: 12,
                  bgcolor: "#0078d4",
                  color: "#fff",
                  fontWeight: 600,
                }}
              >
                {initials}
              </Avatar>
              <Box sx={{ display: "flex", flexDirection: "column", minWidth: 0, flex: 1 }}>
                <Box sx={{ display: "flex", alignItems: "baseline", gap: 1 }}>
                  <Typography variant="body2" sx={{ fontWeight: 600, fontSize: 13, lineHeight: 1.2 }}>
                    {option.displayText}
                  </Typography>
                  {subtitle && (
                    <Typography
                      variant="caption"
                      sx={{
                        color: "text.secondary",
                        fontSize: 11,
                        whiteSpace: "nowrap",
                        overflow: "hidden",
                        textOverflow: "ellipsis",
                      }}
                    >
                      {subtitle}
                    </Typography>
                  )}
                </Box>
                {option.email && (
                  <Typography
                    variant="caption"
                    sx={{
                      color: "text.secondary",
                      fontSize: 11,
                      whiteSpace: "nowrap",
                      overflow: "hidden",
                      textOverflow: "ellipsis",
                    }}
                  >
                    {option.email}
                  </Typography>
                )}
              </Box>
            </Box>
          </li>
        );
      }}
      renderTags={(value: readonly string[], getTagProps) =>
        value.map((option: string, index: number) => {
          const isValid = validateEmail(option);
          return (
            <Chip
              {...getTagProps({ index })}
              key={index}
              label={option}
              size="small"
              sx={{
                height: "24px",
                fontSize: "12px",
                margin: "2px",
                ...(!isValid
                  ? {
                      bgcolor: "#ffebee",
                      color: "#d32f2f",
                      border: "1px solid #ffcdd2",
                    }
                  : {
                      bgcolor: "#e8f0fe",
                      color: "#1967d2",
                      border: "1px solid #d2e3fc",
                    }),
              }}
            />
          );
        })
      }
      renderInput={(params) => (
        <TextField
          {...params}
          size="small"
          label={label}
          placeholder={values.length === 0 ? placeholder : ""}
          error={Boolean(errorText)}
          helperText={errorText}
          margin="dense"
          onBlur={() => {
            let currentAll = [...values];
            if (inputValue.trim()) {
              const parts = inputValue
                .split(/[,;]/)
                .map((p) => p.trim())
                .filter(Boolean);
              if (parts.length > 0) {
                currentAll = Array.from(new Set([...values, ...parts]));
                setValues(currentAll);
                setInputValue("");
              }
            }
            onFieldBlur(currentAll);
          }}
          InputProps={{
            ...params.InputProps,
            endAdornment: (
              <React.Fragment>
                {loading ? <CircularProgress color="inherit" size={16} sx={{ mr: 1 }} /> : null}
                {params.InputProps.endAdornment}
              </React.Fragment>
            ),
          }}
        />
      )}
    />
  );
};

export const EmailShareDialog: React.FC<IEmailShareDialogProps> = ({
  open,
  onClose,
  selectedItems = [],
  siteUrl,
  shareFlowUrl,
  defaultSubject = "",
  defaultMessage = "",
  currentUserEmail = "",
  sp,
}) => {
  const activeSp = sp || defaultSp;

  const [toEmails, setToEmails] = React.useState<string[]>([]);
  const [toInput, setToInput] = React.useState<string>("");

  const [ccEmails, setCcEmails] = React.useState<string[]>([]);
  const [ccInput, setCcInput] = React.useState<string>("");

  const [bccEmails, setBccEmails] = React.useState<string[]>([]);
  const [bccInput, setBccInput] = React.useState<string>("");

  const [subject, setSubject] = React.useState<string>(defaultSubject);
  const [message, setMessage] = React.useState<string>(defaultMessage);

  const [emailErrors, setEmailErrors] = React.useState<{ to: string; cc: string; bcc: string }>({
    to: "",
    cc: "",
    bcc: "",
  });
  const [shareErrorMessage, setShareErrorMessage] = React.useState<string>("");
  const [isSharing, setIsSharing] = React.useState(false);

  React.useEffect(() => {
    if (open) {
      setToEmails([]);
      setToInput("");
      const initialCc = currentUserEmail ? [currentUserEmail.trim()] : [];
      setCcEmails(initialCc);
      setCcInput("");
      setBccEmails([]);
      setBccInput("");
      setSubject(defaultSubject);
      setMessage(defaultMessage);
      setEmailErrors({ to: "", cc: "", bcc: "" });
      setShareErrorMessage("");
    }
  }, [open, defaultSubject, defaultMessage, currentUserEmail]);

  const validateEmail = (email: string) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email.trim());

  const isInternalEmail = (email: string): boolean => email.trim().toLowerCase().endsWith("@dksh.com");

  const isInternalDocument = (item: any): boolean =>
    (item?.Confidentiality || "").trim().toLowerCase() === "internal";

  const validateEmailList = (emails: string[]): string => {
    for (const em of emails) {
      if (!validateEmail(em)) {
        return `Invalid email: ${em}`;
      }
    }
    return "";
  };

  const handleClose = () => {
    setToEmails([]);
    setToInput("");
    setCcEmails(currentUserEmail ? [currentUserEmail.trim()] : []);
    setCcInput("");
    setBccEmails([]);
    setBccInput("");
    setSubject("");
    setMessage("");
    setEmailErrors({ to: "", cc: "", bcc: "" });
    setShareErrorMessage("");
    onClose();
  };

  const handleSend = async () => {
    // Flush any pending text input into the email arrays
    const finalTo = [
      ...toEmails,
      ...(toInput.trim() ? toInput.split(/[,;]/).map((e) => e.trim()).filter(Boolean) : []),
    ];
    const finalCc = [
      ...ccEmails,
      ...(ccInput.trim() ? ccInput.split(/[,;]/).map((e) => e.trim()).filter(Boolean) : []),
    ];
    const finalBcc = [
      ...bccEmails,
      ...(bccInput.trim() ? bccInput.split(/[,;]/).map((e) => e.trim()).filter(Boolean) : []),
    ];

    const toError = finalTo.length === 0 ? "At least one recipient is required." : validateEmailList(finalTo);
    const ccError = validateEmailList(finalCc);
    const bccError = validateEmailList(finalBcc);

    if (toError || ccError || bccError) {
      setEmailErrors({ to: toError, cc: ccError, bcc: bccError });
      return;
    }

    const flowTriggerUrl = (shareFlowUrl || DEFAULT_SHARE_FLOW_URL || "").trim();
    if (!flowTriggerUrl) {
      setShareErrorMessage("Sharing Flow URL is not configured. Please configure it in the Web Part properties.");
      return;
    }

    // Recipient & Document Classification (matching teammate logic)
    const allRecipients = Array.from(
      new Set([...finalTo, ...finalCc, ...finalBcc].map((e) => e.trim()).filter(Boolean))
    );
    const externalRecipients = allRecipients.filter((email) => !isInternalEmail(email));
    const internalRecipients = allRecipients.filter((email) => isInternalEmail(email));
    const hasOnlyExternalRecipients = allRecipients.length > 0 && internalRecipients.length === 0;

    const allSelectedDocumentsInternal =
      selectedItems.length > 0 && selectedItems.every((item) => isInternalDocument(item));
    const hasAtLeastOneInternalDocument =
      selectedItems.length > 0 && selectedItems.some((item) => isInternalDocument(item));
    const hasInternalExternalRestriction =
      hasAtLeastOneInternalDocument && !allSelectedDocumentsInternal && externalRecipients.length > 0;

    // Hard block check: 100% of selected documents are Internal and there is an external recipient
    if (allSelectedDocumentsInternal && externalRecipients.length > 0) {
      const uniqueExternalRecipients = externalRecipients.filter((e, idx, arr) => arr.indexOf(e) === idx);
      const uniqueInternalRecipients = internalRecipients.filter((e, idx, arr) => arr.indexOf(e) === idx);

      let blockMessage =
        `This document is marked Internal and cannot be shared with external recipients:\n` +
        `${uniqueExternalRecipients.join("; ")}\n\n`;

      if (uniqueInternalRecipients.length > 0) {
        blockMessage +=
          `To share it with the internal recipients: ${uniqueInternalRecipients.join("; ")}, remove the external recipients and send the email again.\n\n`;
      }

      blockMessage +=
        `If external sharing is required, contact your QA/RA Lead to review the document's confidentiality setting.`;

      setShareErrorMessage(blockMessage);
      return;
    }

    setIsSharing(true);
    setShareErrorMessage("");

    // Build payload documents (filter out internal documents if all recipients are external)
    const documentsForPayload = hasOnlyExternalRecipients
      ? selectedItems.filter((i) => !isInternalDocument(i))
      : selectedItems;

    const documents = documentsForPayload.map((item) => ({
      id: String(item.id ?? item.ID ?? ""),
      name: item.filename || item.OriginalFilename || item.FileLeafRef || "Unnamed document",
      confidentiality: item.Confidentiality || "",
    }));

    const formattedBody = (message || "").replace(/\r?\n/g, "<br/>");
    const payload: IShareFlowPayload = {
      senderEmail: (currentUserEmail || "").trim(),
      to: finalTo.map((e) => e.trim()).filter(Boolean).join(";"),
      cc: finalCc.map((e) => e.trim()).filter(Boolean).join(";"),
      bcc: finalBcc.map((e) => e.trim()).filter(Boolean).join(";"),
      subject: (subject || "").trim(),
      body: formattedBody,
      confirmSend: false,
      documents,
    };

    try {
      // Phase 1: Preflight check with confirmSend: false
      const preflightResult = await postToShareFlow(flowTriggerUrl, payload);

      if (!preflightResult.ok) {
        setShareErrorMessage(preflightResult.errorMessage || "Unable to complete the document sharing request.");
        return;
      }

      const status = preflightResult.data && typeof preflightResult.data === "object"
        ? (preflightResult.data as { status?: unknown }).status
        : undefined;

      // Check if Flow returned a WARNING status with rejected confidential documents
      if (status === "WARNING") {
        const rejectedDocumentNames = parseRejectedDocumentNames(preflightResult.data);
        const documentList = rejectedDocumentNames.length > 0
          ? `\n\nThe following confidential document(s) cannot be shared:\n${rejectedDocumentNames.map((name) => `- ${name}`).join("\n")}`
          : "";

        const warningMessage =
          "Some document(s) cannot be shared because you do not have the required rights to share them." +
          documentList +
          (hasInternalExternalRestriction
            ? '\n\nAt least one document is marked as "Internal" and cannot be shared with External recipient(s).'
            : "") +
          "\n\nIf you continue, only the eligible document(s) will be shared with the recipients." +
          "\n\nAlternatively, remove the restricted document(s) and send the email again." +
          "\n\nSelect OK to continue, or Cancel to edit.";

        if (!window.confirm(warningMessage)) {
          return;
        }
      } else if (hasInternalExternalRestriction) {
        // Mixed documents warning (at least 1 internal, at least 1 non-internal, external recipients present)
        const uniqueExternalRecipients = externalRecipients.filter((e, idx, arr) => arr.indexOf(e) === idx);
        const internalWarningMessage =
          `At least one document is marked as "Internal" and cannot be shared with ` +
          `External recipient(s): ${uniqueExternalRecipients.join("; ")}\n\n` +
          `To share the "Internal" document(s) with internal recipients, remove the external recipient(s) and send the email again.\n\n` +
          `Alternatively, remove the Internal document(s) and send the email again.\n\n` +
          `If you continue, only the "non-Internal" document(s) will be shared with all recipients.\n\n` +
          `Select OK to continue, or Cancel to edit.`;

        if (!window.confirm(internalWarningMessage)) {
          return;
        }
      }

      // Phase 2: Final send execution with confirmSend: true
      const finalResult = await postToShareFlow(flowTriggerUrl, { ...payload, confirmSend: true });
      if (!finalResult.ok) {
        setShareErrorMessage(finalResult.errorMessage || "Unable to complete the document sharing request.");
        return;
      }

      const toRecipients = finalTo.filter((r) => r.trim().length > 0).join("; ");
      alert(`Mail trigger successful. Sent to: ${toRecipients}`);
      handleClose();
    } catch (error: any) {
      console.error("Error during sharing flow:", error);
      setShareErrorMessage("Unable to connect to the document sharing service. Please try again.");
    } finally {
      setIsSharing(false);
    }
  };

  const handleFieldBlur = (field: "to" | "cc" | "bcc", currentAll: string[]) => {
    const err =
      field === "to" && currentAll.length === 0
        ? "At least one recipient is required."
        : validateEmailList(currentAll);
    setEmailErrors((prev) => ({ ...prev, [field]: err }));
  };

  const handleFieldChange = (field: "to" | "cc" | "bcc") => {
    if (emailErrors[field]) {
      setEmailErrors((prev) => ({ ...prev, [field]: "" }));
    }
  };

  const isFormValid =
    (toEmails.length > 0 || toInput.trim().length > 0) &&
    !emailErrors.to &&
    !emailErrors.cc &&
    !emailErrors.bcc;

  return (
    <Dialog open={open} onClose={handleClose} fullWidth maxWidth="sm">
      <DialogTitle>Share Selected Items</DialogTitle>
      <DialogContent sx={{ pt: 1 }}>
        {shareErrorMessage && (
          <Typography
            color="error"
            variant="body2"
            sx={{ mb: 1, p: 1, backgroundColor: "#ffebee", borderRadius: "4px", whiteSpace: "pre-line" }}
          >
            {shareErrorMessage}
          </Typography>
        )}

        <EmailPeoplePickerInput
          field="to"
          label="To"
          values={toEmails}
          setValues={setToEmails}
          inputValue={toInput}
          setInputValue={setToInput}
          placeholder="e.g. user@domain.com or search name..."
          errorText={emailErrors.to}
          onFieldBlur={(all) => handleFieldBlur("to", all)}
          onFieldChange={handleFieldChange}
          validateEmail={validateEmail}
          sp={activeSp}
        />
        <EmailPeoplePickerInput
          field="cc"
          label="CC"
          values={ccEmails}
          setValues={setCcEmails}
          inputValue={ccInput}
          setInputValue={setCcInput}
          placeholder="e.g. user@domain.com or search name..."
          errorText={emailErrors.cc}
          onFieldBlur={(all) => handleFieldBlur("cc", all)}
          onFieldChange={handleFieldChange}
          validateEmail={validateEmail}
          sp={activeSp}
        />
        <EmailPeoplePickerInput
          field="bcc"
          label="BCC"
          values={bccEmails}
          setValues={setBccEmails}
          inputValue={bccInput}
          setInputValue={setBccInput}
          placeholder="e.g. user@domain.com or search name..."
          errorText={emailErrors.bcc}
          onFieldBlur={(all) => handleFieldBlur("bcc", all)}
          onFieldChange={handleFieldChange}
          validateEmail={validateEmail}
          sp={activeSp}
        />

        <TextField
          fullWidth
          size="small"
          label="Subject"
          value={subject}
          onChange={(e) => setSubject(e.target.value)}
          margin="dense"
        />
        <TextField
          fullWidth
          size="small"
          label="Message"
          value={message}
          onChange={(e) => setMessage(e.target.value)}
          margin="dense"
          multiline
          rows={4}
        />
      </DialogContent>
      <DialogActions>
        <Button onClick={handleClose}>Cancel</Button>
        <Button
          onClick={handleSend}
          disabled={!isFormValid || isSharing}
        >
          {isSharing ? "Sharing..." : "Send"}
        </Button>
      </DialogActions>
    </Dialog>
  );
};

export default EmailShareDialog;
