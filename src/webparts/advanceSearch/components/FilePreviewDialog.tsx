import * as React from "react";
import Dialog from "@mui/material/Dialog";
import DialogTitle from "@mui/material/DialogTitle";
import DialogContent from "@mui/material/DialogContent";
import Box from "@mui/material/Box";
import Typography from "@mui/material/Typography";
import IconButton from "@mui/material/IconButton";
import Tooltip from "@mui/material/Tooltip";
import CircularProgress from "@mui/material/CircularProgress";
import CloseIcon from "@mui/icons-material/Close";
import OpenInNewIcon from "@mui/icons-material/OpenInNew";
import RefreshIcon from "@mui/icons-material/Refresh";
import DescriptionIcon from "@mui/icons-material/Description";
import { doclib_AllProducts } from "../types/advanceSearchTypes";

export interface IFilePreviewDialogProps {
  open: boolean;
  onClose: () => void;
  item: doclib_AllProducts | null;
  previewUrl: string;
}

export const FilePreviewDialog: React.FC<IFilePreviewDialogProps> = ({
  open,
  onClose,
  item,
  previewUrl,
}) => {
  const [isLoading, setIsLoading] = React.useState<boolean>(true);
  const [iframeKey, setIframeKey] = React.useState<number>(0);

  // Reset loading state whenever dialog opens or previewUrl changes
  React.useEffect(() => {
    if (open) {
      setIsLoading(true);
      setIframeKey((prev) => prev + 1);
    }
  }, [open, previewUrl]);

  const handleOpenInNewTab = (): void => {
    if (previewUrl) {
      window.open(previewUrl, "_blank", "noopener,noreferrer");
    }
  };

  const handleRefreshIframe = (): void => {
    setIsLoading(true);
    setIframeKey((prev) => prev + 1);
  };

  const filename = item?.filename || "Document Preview";

  return (
    <Dialog
      open={open}
      onClose={onClose}
      fullWidth
      maxWidth="xl"
      PaperProps={{
        sx: {
          height: "92vh",
          maxHeight: "95vh",
          display: "flex",
          flexDirection: "column",
          borderRadius: "8px",
          overflow: "hidden",
        },
      }}
    >
      <DialogTitle
        sx={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          py: 1,
          px: 2,
          borderBottom: "1px solid #e0e0e0",
          backgroundColor: "#fafafa",
          minHeight: "48px",
        }}
      >
        <Box sx={{ display: "flex", alignItems: "center", gap: 1, minWidth: 0, pr: 2 }}>
          <DescriptionIcon color="primary" sx={{ fontSize: 22, flexShrink: 0 }} />
          <Typography
            variant="subtitle1"
            sx={{
              fontSize: "14px",
              fontWeight: 600,
              whiteSpace: "nowrap",
              overflow: "hidden",
              textOverflow: "ellipsis",
            }}
            title={filename}
          >
            {filename}
          </Typography>
        </Box>

        <Box sx={{ display: "flex", alignItems: "center", gap: 0.5, flexShrink: 0 }}>
          <Tooltip title="Reload preview">
            <IconButton size="small" onClick={handleRefreshIframe} sx={{ color: "text.secondary" }}>
              <RefreshIcon fontSize="small" />
            </IconButton>
          </Tooltip>

          <Tooltip title="Open in new window">
            <IconButton size="small" onClick={handleOpenInNewTab} sx={{ color: "text.secondary" }}>
              <OpenInNewIcon fontSize="small" />
            </IconButton>
          </Tooltip>

          <Tooltip title="Close">
            <IconButton size="small" onClick={onClose} sx={{ color: "text.secondary" }}>
              <CloseIcon fontSize="small" />
            </IconButton>
          </Tooltip>
        </Box>
      </DialogTitle>

      <DialogContent
        sx={{
          p: 0,
          position: "relative",
          flex: 1,
          display: "flex",
          flexDirection: "column",
          overflow: "hidden",
          backgroundColor: "#f5f5f5",
        }}
      >
        {isLoading && (
          <Box
            sx={{
              position: "absolute",
              top: 0,
              left: 0,
              right: 0,
              bottom: 0,
              display: "flex",
              flexDirection: "column",
              alignItems: "center",
              justifyContent: "center",
              gap: 1.5,
              backgroundColor: "rgba(255, 255, 255, 0.85)",
              zIndex: 10,
            }}
          >
            <CircularProgress size={36} color="primary" />
            <Typography variant="body2" color="text.secondary" sx={{ fontSize: "13px" }}>
              Loading document preview...
            </Typography>
          </Box>
        )}

        {previewUrl ? (
          <iframe
            key={iframeKey}
            src={previewUrl}
            title={filename}
            onLoad={() => setIsLoading(false)}
            style={{
              width: "100%",
              height: "100%",
              border: "none",
              flex: 1,
              backgroundColor: "#fff",
            }}
            allow="fullscreen"
          />
        ) : (
          <Box
            sx={{
              flex: 1,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              color: "text.secondary",
            }}
          >
            <Typography variant="body2">No preview URL available for this file.</Typography>
          </Box>
        )}
      </DialogContent>
    </Dialog>
  );
};

export default FilePreviewDialog;
