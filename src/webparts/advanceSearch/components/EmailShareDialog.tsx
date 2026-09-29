import * as React from "react";
import { SPFI } from "@pnp/sp";
import { sp as defaultSp } from "../AdvanceSearchWebPart";
import { shareFilesByEmail } from "../../../services/sharePointService";
import Dialog from "@mui/material/Dialog";
import DialogTitle from "@mui/material/DialogTitle";
import DialogContent from "@mui/material/DialogContent";
import DialogActions from "@mui/material/DialogActions";
import TextField from "@mui/material/TextField";
import Button from "@mui/material/Button";
import Typography from "@mui/material/Typography";
import Autocomplete from "@mui/material/Autocomplete";
import Chip from "@mui/material/Chip";

export interface IEmailShareDialogProps {
  open: boolean;
  onClose: () => void;
  selectedItems?: any[];
  siteUrl: string;
  defaultSubject?: string;
  defaultMessage?: string;
  currentUserEmail?: string;
  sp?: SPFI;
}

export const EmailShareDialog: React.FC<IEmailShareDialogProps> = ({
  open,
  onClose,
  selectedItems = [],
  siteUrl,
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

    setIsSharing(true);
    setShareErrorMessage("");
    try {
      await shareFilesByEmail(
        activeSp,
        selectedItems,
        finalTo,
        finalCc,
        finalBcc,
        subject,
        message
      );

      alert(`Successfully shared ${selectedItems.length} file(s)!`);
      handleClose();
    } catch (error: any) {
      console.error("Error during sharing:", error);
      setShareErrorMessage(error?.message || "Failed to share files. Please verify permissions.");
    } finally {
      setIsSharing(false);
    }
  };

  const renderEmailChipInput = (
    field: "to" | "cc" | "bcc",
    label: string,
    values: string[],
    setValues: React.Dispatch<React.SetStateAction<string[]>>,
    inputValue: string,
    setInputValue: React.Dispatch<React.SetStateAction<string>>,
    placeholder: string
  ) => {
    const errorText = emailErrors[field];

    return (
      <Autocomplete
        multiple
        freeSolo
        size="small"
        options={[]}
        value={values}
        inputValue={inputValue}
        onInputChange={(_event, newInputValue, reason) => {
          if (reason === "input") {
            if (newInputValue.includes(",") || newInputValue.includes(";")) {
              const parts = newInputValue
                .split(/[,;]/)
                .map((p) => p.trim())
                .filter(Boolean);
              setValues((prev) => Array.from(new Set([...prev, ...parts])));
              setInputValue("");
              if (emailErrors[field]) {
                setEmailErrors((prev) => ({ ...prev, [field]: "" }));
              }
            } else {
              setInputValue(newInputValue);
            }
          } else if (reason === "reset" || reason === "clear") {
            setInputValue("");
          }
        }}
        onChange={(_event, newValue) => {
          const cleaned: string[] = [];
          (newValue as (string | any)[]).forEach((val) => {
            if (typeof val === "string") {
              val.split(/[,;]/).forEach((item) => {
                const trimmed = item.trim();
                if (trimmed && cleaned.indexOf(trimmed) === -1) {
                  cleaned.push(trimmed);
                }
              });
            }
          });
          setValues(cleaned);
          setInputValue("");
          if (emailErrors[field]) {
            setEmailErrors((prev) => ({ ...prev, [field]: "" }));
          }
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
              if (inputValue.trim()) {
                const parts = inputValue
                  .split(/[,;]/)
                  .map((p) => p.trim())
                  .filter(Boolean);
                if (parts.length > 0) {
                  setValues((prev) => Array.from(new Set([...prev, ...parts])));
                  setInputValue("");
                }
              }
              const currentAll = [
                ...values,
                ...(inputValue.trim() ? inputValue.split(/[,;]/).map((e) => e.trim()).filter(Boolean) : []),
              ];
              const err =
                field === "to" && currentAll.length === 0
                  ? "At least one recipient is required."
                  : validateEmailList(currentAll);
              setEmailErrors((prev) => ({ ...prev, [field]: err }));
            }}
          />
        )}
      />
    );
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
            sx={{ mb: 1, p: 1, backgroundColor: "#ffebee", borderRadius: "4px" }}
          >
            {shareErrorMessage}
          </Typography>
        )}

        {renderEmailChipInput(
          "to",
          "To",
          toEmails,
          setToEmails,
          toInput,
          setToInput,
          "e.g. user@domain.com, user2@domain.com"
        )}
        {renderEmailChipInput(
          "cc",
          "CC",
          ccEmails,
          setCcEmails,
          ccInput,
          setCcInput,
          "e.g. user@domain.com"
        )}
        {renderEmailChipInput(
          "bcc",
          "BCC",
          bccEmails,
          setBccEmails,
          bccInput,
          setBccInput,
          "e.g. user@domain.com"
        )}

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
