import * as React from "react";
import { IHelloWorldProps } from "./IHelloWorldProps";
import { sp } from "../HelloWorldWebPart";
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

type doclib_AllProducts = {
  filename: string;
  PIMProduct: IProductLookupItem[];
  PIMProductSearchText: string;
  BusinessLine: string;
  CountrySoldTo: string;
  ManufacturerSearchText: string;
  DocumentTypeSearchText: string;
  SubDocumentTypeSearchText: string;
  fileUrl: string;
};

type IProductLookupItem = {
  ID: number;
  Title: string;
  PIMProductName: string;
  Manufacturer?: string;
};

type IClientLookupItem = {
  ID: number;
  Title: string;
};

const choiceToString = (value: string | string[] | undefined | null): string => {
  if (!value) return "";
  return Array.isArray(value) ? value.join(", ") : value;
};

const sanitizeKqlValue = (value: string): string =>
  value.replace(/"/g, '\\"').trim();

const HelloWorld: React.FC<IHelloWorldProps> = (props) => {
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

  const [productSearchText, setProductSearchText] = React.useState("");
  const [clientSearchText, setClientSearchText] = React.useState("");
  const [additionalKeyword, setAdditionalKeyword] = React.useState("");

  const [lookupLoading, setLookupLoading] = React.useState(false);
  const [resultsLoading, setResultsLoading] = React.useState(false);

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
        .items.select("ID", "Title", "PIMProductName", "Manufacturer")
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

    if (additionalKeyword.trim()) {
      clauses.push(`"${sanitizeKqlValue(additionalKeyword)}"`);
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
            "Sub_x0020_Document_x0020_Type"
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
      }));

      setItems_AllProducts(mappedData);
      //setAllRecordsCache(mappedData);
    } catch (error) {
      console.error("loadAllRecords error:", error);
    } finally {
      setResultsLoading(false);
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

    if (selectedProductValues.length === 0 && selectedClientValues.length === 0) {
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
            "RefinableString100",
            "PIMProductCode",
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
            "Sub_x0020_Document_x0020_Type"
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
      }));

      setItems_AllProducts(mappedData);
    } catch (error) {
      console.error("handleSearch error:", error);
    } finally {
      setResultsLoading(false);
    }
  };

  const columns_AllProducts = useMemo<MRT_ColumnDef<doclib_AllProducts>[]>(
    () => [
      {
        accessorKey: "filename",
        header: "File Name",
        size: 120,
        filterFn: "contains",
        Cell: ({ row }) => {
          const fileUrl = `${props.urlSite}Products/${encodeURIComponent(
            row.original.filename
          )}`;

          return (
            <a href={fileUrl} target="_blank" rel="noopener noreferrer">
              {row.original.filename}
            </a>
          );
        },
      },
      {
        accessorKey: "ManufacturerSearchText",
        header: "Clients",
        filterFn: "contains",
        size: 300,
        Cell: ({ cell }) => (
          <>
            {String(cell.getValue() || "")
              .split(";")
              .filter(Boolean)
              .map((item, idx) => (
                <div key={idx}>{item.trim()}</div>
              ))}
          </>
        ),
      },
      {
        accessorKey: "PIMProductSearchText",
        header: "Products",
        filterFn: "contains",
        size: 300,
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
        filterFn: "contains",
        size: 220,
      },
      {
        accessorKey: "SubDocumentTypeSearchText",
        header: "Sub Document Type",
        filterFn: "contains",
        size: 220,
        Cell: ({ cell }) => (
          <>
            {String(cell.getValue() || "")
              .split(";")
              .filter(Boolean)
              .map((item, idx) => (
                <div key={idx}>{item.trim()}</div>
              ))}
          </>
        ),
      },
      {
        accessorKey: "BusinessLine",
        header: "Business Line",
        filterFn: "contains",
      },
      {
        accessorKey: "CountrySoldTo",
        header: "Country Sold To",
        filterFn: "contains",
      },
    ],
    [props.urlSite]
  );

  const table = useMaterialReactTable({
    columns: columns_AllProducts,
    data: items_AllProducts,
    enableGrouping: true,
    enableColumnDragging: true,
    muiPaginationProps: {
      rowsPerPageOptions: [50, 100, 500, 1000],
    },
    state: {
      isLoading: resultsLoading,
      showColumnFilters: true,
    },
  });

  return (
    <div>
      <h2>Clients & Products</h2>

      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))",
          gap: "16px",
          marginBottom: "16px",
        }}
      >
        <Autocomplete
          multiple
          fullWidth
          loading={lookupLoading}
          options={products}
          value={selectedProducts}
          getOptionLabel={(option: IProductLookupItem) =>
            option
              ? `${option.Title ?? ""} | ${option.PIMProductName ?? ""} | ${option.Manufacturer ?? ""
              }`
              : ""
          }
          isOptionEqualToValue={(option, value) => option.ID === value.ID}
          onInputChange={handleProductInputChange}
          onChange={handleProductSelectionChange}
          renderInput={(params) => (
            <TextField
              {...params}
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

      <TextField
        fullWidth
        margin="normal"
        label="Additional Keyword"
        placeholder="Optional keyword"
        value={additionalKeyword}
        onChange={(e) => setAdditionalKeyword(e.target.value)}
      />

      <div style={{ marginBottom: "16px" }}>
        <Button
          variant="contained"
          color="primary"
          onClick={handleSearch}
          disabled={resultsLoading}
        >
          {resultsLoading ? "Searching..." : "Search"}
        </Button>
      </div>

      <MaterialReactTable table={table} />

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
  );
};

export default HelloWorld;