import * as React from 'react';
//import { spfi, SPFx } from '@pnp/sp';
//import '@pnp/sp/webs';
//import '@pnp/sp/lists';
//import '@pnp/sp/items';
import { IHelloWorldProps } from './IHelloWorldProps';
import { sp } from '../HelloWorldWebPart';
import { useMemo } from 'react';
import {
  MaterialReactTable,
  useMaterialReactTable,
  type MRT_ColumnDef,
} from 'material-react-table';

import Autocomplete from "@mui/material/Autocomplete";
import TextField from "@mui/material/TextField";
import CircularProgress from "@mui/material/CircularProgress";
import ToggleButton from '@mui/material/ToggleButton';
import ToggleButtonGroup from '@mui/material/ToggleButtonGroup';
import Button from '@mui/material/Button';






type doclib_AllProducts = {
  filename: string;
  PIMProduct: IProductLookupItem[];
  PIMProductSearchText: string;
  BusinessLine: string;
  CountrySoldTo: string;
  Manufacturer: IClientLookupItem[];
  ManufacturerSearchText: string;
  DocumentType?: IDocumentTypeLookupItem;
  DocumentTypeSearchText: string;
  SubDocumentType: ISubDocumentTypeLookupItem[];
  SubDocumentTypeSearchText: string;


};


type IProductLookupItem = {
  ID: number;
  Title: string;
  PIMProductName: string;
  PIM_x0020_Product: string;
};

type IClientLookupItem = {
  ID: number;
  Title: string;
};

type IDocumentTypeLookupItem = {
  ID: number;
  Title: string;
};
type ISubDocumentTypeLookupItem = {
  ID: number;
  Title: string;
};


const choiceToString = (
  value: string | string[] | undefined | null
): string => {
  if (!value) return '';

  return Array.isArray(value)
    ? value.join(', ')
    : value;
};




const HelloWorld: React.FC<IHelloWorldProps> = (props) => {
  const [items_AllProducts, setItems_AllProducts] =
    React.useState<doclib_AllProducts[]>([]);
  //const [loading, setLoading] = React.useState<boolean>(true);

  const [products, setProducts] = React.useState<IProductLookupItem[]>([]);
  const [allRecordsCache, setAllRecordsCache] =
    React.useState<doclib_AllProducts[]>([]);

  const [searchType, setSearchType] =
    React.useState<"product" | "client" | "all">("product");

  const [clients, setClients] = React.useState<IClientLookupItem[]>([]);

  //const [loading, setLoading] = React.useState(false);
  const [lookupLoading, setLookupLoading] = React.useState(false);
  const [resultsLoading, setResultsLoading] = React.useState(false);
  const [selectedProduct, setSelectedProduct] = React.useState<IProductLookupItem | null>(null);
  const [selectedClient, setSelectedClient] = React.useState<IClientLookupItem | null>(null);

  const searchProducts = async (searchText: string) => {
    if (searchText.length < 3) {
      setProducts([]);
      return;
    }
    setLookupLoading(true);

    try {

      const escapedText = searchText.replace(/'/g, "''");

      const results = await sp.web.lists
        .getByTitle("PIM Product")
        .items
        .select(
          "ID",
          "Title",
          "PIMProductName", "PIM_x0020_Product"

        )

        .filter(
          `substringof('${escapedText}', PIMProductName)`
        )
        .top(5000)();


      setProducts(results as IProductLookupItem[]);
    }
    catch (error) {
      console.error(error);
    }
    finally {
      setLookupLoading(false);
    }
  };
  const searchClients = async (
    searchText: string
  ) => {

    if (searchText.length < 3) {
      setClients([]);
      return;
    }
    setLookupLoading(true);

    try {
      const escapedText =
        searchText.replace(/'/g, "''");

      const results = await sp.web.lists
        .getByTitle("PIM Global Client")
        .items
        .select(
          "ID",
          "Title"
        )
        .filter(
          `substringof('${escapedText}', Title)`
        )
        .top(5000)();

      setClients(results as IClientLookupItem[]);
    }
    catch (error) {
      console.error(error);
    }
    finally {
      setLookupLoading(false);
    }
  };

  const [searchText, setSearchText] = React.useState("");

  React.useEffect(() => {

    const timer = setTimeout(() => {


      if (searchText.length >= 3) {

        if (searchType === "product") {
          searchProducts(searchText);
        }
        else {
          searchClients(searchText);
        }

      }


    }, 500);

    return () => clearTimeout(timer);


  }, [searchText, searchType]);




  const columns_AllProducts = useMemo<MRT_ColumnDef<doclib_AllProducts>[]>(
    () => [

      {
        accessorKey: 'filename',
        header: 'File Name',
        size: 100,
        filterFn: 'contains',
      },
      {
        accessorKey: 'ManufacturerSearchText',
        header: 'Clients',
        filterFn: 'contains',
        size: 400,


        Cell: ({ row }) => (
          <div>
            {row.original.Manufacturer.map((p, idx) => (
              <div key={idx}>
                {p.Title}
              </div>
            ))}
          </div>
        ),
      },
      {
        accessorKey: 'PIMProductSearchText',
        header: 'Products',
        filterFn: 'contains',
        size: 400,


        Cell: ({ row }) => (
          <div>
            {row.original.PIMProduct.map((p, idx) => (
              <div key={idx}>
                {p.Title} {p.PIMProductName}
              </div>
            ))}
          </div>
        ),
      },
      {
        accessorKey: 'DocumentTypeSearchText',
        header: 'Document Type',
        filterFn: 'contains',
        size: 400,

        Cell: ({ row }) => (
          <div>
            {row.original.DocumentType?.Title ?? ""}
          </div>
        ),
      },
      {
        accessorKey: 'SubDocumentTypeSearchText',
        header: 'Sub Document Type',
        filterFn: 'contains',
        size: 400,

        Cell: ({ row }) => (
          <div>
            {row.original.SubDocumentType.map((p, idx) => (
              <div key={idx}>
                {p.Title}
              </div>
            ))}
          </div>
        ),
      },



      {
        accessorKey: 'BusinessLine',
        header: 'Business Line',
        filterFn: 'contains',
      },
      {
        accessorKey: 'CountrySoldTo',
        header: 'Country Sold To',
        filterFn: 'contains',
      },


    ],
    [],
  );



  const table = useMaterialReactTable({
    columns: columns_AllProducts,
    data: items_AllProducts,
    enableGrouping: true,
    enableColumnDragging: true,
    initialState: {
      showColumnFilters: true,
    },
    state: {
      isLoading: resultsLoading,
      grouping: ['PIMProductSearchText'],
      expanded: true,
      showColumnFilters: true,
    },
    //enablePagination: false,
  });




  const loadAllRecords = async (): Promise<void> => {

    setResultsLoading(true);

    try {

      const allProducts: any[] = [];
      const pageSize = 5000;

      let lastId = 0;

      while (true) {

        const batch = await sp.web.lists
          .getByTitle("Clients & Products")
          .items
          .select(
            "Id",
            "Title",
            "FileLeafRef",
            "FileRef",
            "Country",
            "Business_x0020_Line",
            "PIMProductCode/Title",
            "PIMProductCode/PIMProductName",
            "ManufacturerLookup/Title",
            "DocumentType/Title",
            "SubDocumentType/Title",
            "PIMProductTermSet",
            "GlobalClientTermSet")
          .expand("PIMProductCode", "ManufacturerLookup", "DocumentType", "SubDocumentType")
          .filter(`Id gt ${lastId}`)
          .orderBy("Id")
          .top(pageSize)();

        if (batch.length === 0) {
          break;
        }

        allProducts.push(...batch);

        console.log(
          `Loaded ${allProducts.length} records`
        );

        lastId = batch[batch.length - 1].Id;

        if (batch.length < pageSize) {
          break;
        }
      }

      const mappedData: doclib_AllProducts[] =
        allProducts.map((item: any) => ({
          filename: item.FileLeafRef ?? "",
          PIMProduct: item.PIMProductCode ?? [],
          PIMProductSearchText:
            (item.PIMProductCode ?? [])
              .map((p: any) =>
                `${p.Title} ${p.PIMProductName}`
              )
              .join(" "),


          Manufacturer: item.ManufacturerLookup ?? [],
          ManufacturerSearchText:
            (item.ManufacturerLookup ?? [])
              .map(
                (m: IClientLookupItem) =>
                  m.Title
              )
              .join(" "),

          BusinessLine: choiceToString(
            item.Business_x0020_Line
          ),
          CountrySoldTo: choiceToString(
            item.Country
          ),
          DocumentType: item.DocumentType ?? undefined,
          DocumentTypeSearchText:
            item.DocumentType?.Title || "",
          SubDocumentType: item.SubDocumentType ?? [],
          SubDocumentTypeSearchText:
            (item.SubDocumentType ?? [])
              .map((sdt: ISubDocumentTypeLookupItem) => sdt.Title)
              .join(" "),

        }));

      setItems_AllProducts(mappedData);
      setAllRecordsCache(mappedData);
    }
    catch (error) {
      console.error(error);
    }
    finally {
      setResultsLoading(false);
    }
  };

  return (
    <div>
      <h2>Clients & Products</h2>
      <ToggleButtonGroup
        color="primary"
        exclusive
        value={searchType}
        onChange={(_, value) => {

          if (!value) {
            return;
          }

          setSearchType(value);

          setSelectedProduct(null);
          setSelectedClient(null);
          setSearchText("");

          setProducts([]);
          setClients([]);

          if (value === "all") {

            if (allRecordsCache.length > 0) {

              console.log(
                `Using cache: ${allRecordsCache.length}`
              );

              setItems_AllProducts(allRecordsCache);

            } else {

              console.log("Loading all records");

              loadAllRecords().catch(console.error);
            }

          } else {

            setItems_AllProducts([]);
          }
        }}
      >
        <ToggleButton value="product">
          Product
        </ToggleButton>
        <ToggleButton value="client">
          Client
        </ToggleButton>

        <ToggleButton value="all">
          All Records
        </ToggleButton>

      </ToggleButtonGroup>
      {searchType !== "all" && (
        <Autocomplete<any>
          fullWidth
          loading={lookupLoading}

          options={
            searchType === "product"
              ? products
              : clients
          }



          value={
            searchType === "product"
              ? selectedProduct
              : selectedClient
          }

          getOptionLabel={(option: any) =>
            searchType === "product"
              ? `${option.Title} | ${option.PIMProductName}`
              : option.Title
          }


          isOptionEqualToValue={(option, value) =>
            option.ID === value.ID
          }

          onInputChange={(_, value) => {
            setSearchText(value);
          }}


          onChange={async (_, selectedOption) => {

            setResultsLoading(true);


            if (selectedOption) {

              try {

                setItems_AllProducts([]);

                let searchCodeName = ""; //selectedOption.Title +':'+ selectedOption.PIMProductName;

                if (searchType === "product") {
                  searchCodeName = selectedOption.PIM_x0020_Product;
                  setSelectedProduct(selectedOption);

                } else {
                  searchCodeName = selectedOption.Title;
                  setSelectedClient(selectedOption);


                }
                const path = props.urlSite + "Products/*";
                /*const results = await sp.search({
                  Querytext:
                    `"${searchCodeName}" AND Path:"${path}"`,
  
                  SelectProperties: [
                    "Title",
                    "ListItemID",
                    "Path",
                  ],
  
                  RowLimit: 2000,
                });*/




                let allResults = [];
                let startRow = 0;
                const pageSize = 500;

                while (startRow < 2500) {

                  const results = await sp.search({
                    Querytext: `"${searchCodeName}" AND Path:"${path}"`,
                    RowLimit: pageSize,
                    StartRow: startRow,
                    SelectProperties: [
                      "Title",
                      "ListItemID",
                      "Path",
                      'RefinableString100',
                      'RefinableString00',
                      'owstaxIdPIMProductTermSet',


                    ],
                  });

                  allResults.push(...results.PrimarySearchResults);

                  if (results.RowCount < pageSize) {
                    break;
                  }

                  startRow += pageSize;
                }

                console.log(allResults);

                const ids: number[] = [];

                allResults.forEach((r: any) => {

                  const id = Number(r.ListItemID);

                  if (
                    !isNaN(id) &&
                    ids.indexOf(id) === -1
                  ) {
                    ids.push(id);
                  }
                });

                console.log("Unique IDs:", ids);

                const allProducts: any[] = [];
                if (ids.length === 0) {
                  setItems_AllProducts([]);
                  return;
                }

                for (let i = 0; i < ids.length; i += 100) {


                  const currentIds = ids.slice(i, i + 100);


                  const filter = currentIds
                    .map((id) => `Id eq ${id}`)
                    .join(" or ");

                  const items = await sp.web.lists
                    .getByTitle("Clients & Products")
                    .items
                    .select(
                      "Id",
                      "Title",
                      "FileLeafRef",
                      "FileRef",
                      "Country",
                      "Business_x0020_Line",
                      "PIMProductCode/Title",
                      "PIMProductCode/PIMProductName",
                      "ManufacturerLookup/Title",
                      "DocumentType/Title",
                      "SubDocumentType/Title",
                      "PIMProductTermSet",
                      "GlobalClientTermSet"

                    )
                    .expand("PIMProductCode", "ManufacturerLookup", "DocumentType", "SubDocumentType")
                    .filter(filter)();

                  allProducts.push(...items);
                }

                const mappedData: doclib_AllProducts[] =
                  allProducts.map((item: any) => ({
                    filename: item.FileLeafRef ?? "",
                    PIMProduct: item.PIMProductCode ?? [],
                    PIMProductSearchText:
                      (item.PIMProductCode ?? [])
                        .map((p: any) =>
                          `${p.Title} ${p.PIMProductName}`
                        )
                        .join(" "),
                    Manufacturer: item.ManufacturerLookup ?? [],
                    ManufacturerSearchText:
                      (item.ManufacturerLookup ?? [])
                        .map(
                          (m: IClientLookupItem) =>
                            m.Title
                        )
                        .join(" "),

                    BusinessLine: choiceToString(
                      item.Business_x0020_Line
                    ),
                    CountrySoldTo: choiceToString(
                      item.Country
                    ),
                    DocumentType: item.DocumentType ?? undefined,
                    DocumentTypeSearchText:
                      item.DocumentType?.Title || "",
                    SubDocumentType: item.SubDocumentType ?? [],
                    SubDocumentTypeSearchText:
                      (item.SubDocumentType ?? [])
                        .map((sdt: ISubDocumentTypeLookupItem) => sdt.Title)
                        .join(" "),

                  }));



                setItems_AllProducts(mappedData);

              }
              catch (error) {
                console.error(error);
              } finally {
                setResultsLoading(false);
              }
            }


          }
          }
          renderInput={(params) => (
            <TextField
              {...params}

              label={
                searchType === "product"
                  ? "PIM Product"
                  : "Global Client"
              }
              placeholder={
                searchType === "product"
                  ? "Enter a Product Name"
                  : "Enter a Global Client Name"
              }
              InputProps={{
                ...params.InputProps,
                endAdornment: (
                  <>
                    {lookupLoading ? (
                      <CircularProgress
                        color="inherit"
                        size={20}
                      />
                    ) : null}

                    {params.InputProps.endAdornment}
                  </>
                ),
              }}

            />
          )}
        />
      )}

      <MaterialReactTable table={table} />
      <Button
        onClick={() => {
          setAllRecordsCache([]);
          loadAllRecords().catch(console.error);
        }}
      >
        Refresh All Records
      </Button>


    </div>
  );
};

export default HelloWorld;