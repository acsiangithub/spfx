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

  const handleSend = async () => {
    const toError = validateEmailField("to", emailFields.to);
    const ccError = validateEmailField("cc", emailFields.cc);
    const bccError = validateEmailField("bcc", emailFields.bcc);

    if (!toError && !ccError && !bccError && emailFields.to) {
      setIsSharing(true);
      setShareErrorMessage("");
      try {
        const toEmails = emailFields.to.split(/[;,]/).map((e) => e.trim()).filter(Boolean);
        const ccEmails = emailFields.cc.split(/[;,]/).map((e) => e.trim()).filter(Boolean);
        const bccEmails = emailFields.bcc.split(/[;,]/).map((e) => e.trim()).filter(Boolean);

        await shareFilesByEmail(
          activeSp,
          selectedItems,
          toEmails,
          ccEmails,
          bccEmails,
          emailFields.subject,
          emailFields.message
        );

        alert(`Successfully shared ${selectedItems.length} file(s)!`);
        handleClose();
      } catch (error: any) {
        console.error("Error during sharing:", error);
        setShareErrorMessage(error?.message || "Failed to share files. Please verify permissions.");
      } finally {
        setIsSharing(false);
      }
    } else {
      setEmailErrors({ to: toError, cc: ccError, bcc: bccError });
    }
  };

  return (
    <Dialog open={open} onClose={handleClose} fullWidth maxWidth="sm">
      <DialogTitle>Share Selected Items</DialogTitle>
      <DialogContent>
        {shareErrorMessage && (
          <Typography
            color="error"
            variant="body2"
            sx={{ mb: 1, p: 1, backgroundColor: "#ffebee", borderRadius: "4px" }}
          >
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

export default EmailShareDialog;
