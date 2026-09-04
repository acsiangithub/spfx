import { createTheme } from "@mui/material/styles";

export const compactTheme = createTheme({
  typography: {
    fontFamily:
      '"Segoe UI", "Segoe UI Web (West European)", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif',
    fontSize: 14,
  },
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
          textTransform: "none",
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
});
