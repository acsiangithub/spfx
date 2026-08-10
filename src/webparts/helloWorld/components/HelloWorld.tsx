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




type PIMProduct = {
  Title: string;
  PIMProductName: string;
  PIMProductSearchText: string;
};


type doclib_AllProducts = {
  filename: string;
  PIMProduct: PIMProduct[];
  PIMProductSearchText: string;
  BusinessLine: string;
  CountrySoldTo: string;
  GlobalClient: string;
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
  const [searchType, setSearchType] =
    React.useState<"product" | "client">("product");

  const [clients, setClients] = React.useState<IClientLookupItem[]>([]);

  const [loading, setLoading] = React.useState(false);
  const [selectedProduct, setSelectedProduct] = React.useState<IProductLookupItem | null>(null);
  const [selectedClient, setSelectedClient] = React.useState<IClientLookupItem | null>(null);

  const searchProducts = async (searchText: string) => {
    if (searchText.length < 3) {
      setProducts([]);
      return;
    }
    setLoading(true);

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
      setLoading(false);
    }
  };
  const searchClients = async (
    searchText: string
  ) => {

    if (searchText.length < 3) {
      setClients([]);
      return;
    }
    setLoading(true);

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
      setLoading(false);
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
      },
      {
        accessorKey: 'PIMProductSearchText',
        header: 'Products',
        size: 400,

        Cell: ({ row }) => (
          <div>
            {row.original.PIMProduct.map((p, idx) => (
              <div key={idx}>
                {p.Title} | {p.PIMProductName}
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
      {
        accessorKey: 'GlobalClient',
        header: 'Global Client',
        filterFn: 'contains',

      },
    ],
    [],
  );



  const table = useMaterialReactTable({
    columns: columns_AllProducts,
    data: items_AllProducts,
    //enablePagination: false,
  });

  React.useEffect(() => {
    const loadItems = async (): Promise<void> => {
      try {
        //setLoading(true);

        /*const items = await sp.web.lists
       .getByTitle("Clients & Products")
       .items
       // .filter(`PIMProductTermSet eq 'PIM000002676 : CAPTEX 8000'`)()
       .top(5000)()
       console.log(items);*/



        /*const globalClient = "ABITEC CORPORATION";
        //const pimProduct = "";//"CAPTEX 8000";
        //${pimProduct} AND
        const path = props.urlSite +   "Products/*";
        const results = await sp.search({
          Querytext: ` ${globalClient} AND Path:${path}`,
          SelectProperties: [
            "Title",
            "ListItemID",
            "RefinableString00"
          ],
          RowLimit: 500,
        });

        console.log(results.PrimarySearchResults);*/





        // Initialize PnP JS with SPFx Context
        //const sp = spfi().using(SPFx(props.context));

        // PnP JS v4 Syntax to get list items with selected fields
        const allproducts = await sp.web.lists
          .getByTitle("Clients & Products")
          .items.select("Id", "Title", "FileLeafRef", "FileRef", "Manufacturer", "Country", "Business_x0020_Line", "PIMProductCode/Title", "PIMProductCode/PIMProductName")
          .expand("PIMProductCode")
          .top(5000)()

        const mappedData: doclib_AllProducts[] =
          allproducts.map((item) => ({
            filename: item.FileLeafRef ?? "",
            PIMProduct: item.PIMProductCode ?? [],
            PIMProductSearchText:
              (item.PIMProductCode ?? [])
                .map((p: PIMProduct) =>
                  `${p.Title} ${p.PIMProductName}`
                )
                .join(" "),
            BusinessLine: choiceToString(item.Business_x0020_Line),
            CountrySoldTo: choiceToString(item.Country),
            GlobalClient: item.Manufacturer ?? "",
          }));

        setItems_AllProducts(mappedData);




        /*const results: IListItem[] = await sp.web.lists
          .getByTitle('PIM Product')
          .items.select('Title') // Note the empty parenthesis () at the end instead of .get()
          .filter("Title eq 'PIM000139469'")
          ();

        /*const titles = doclib_Allproducts
          .map((item) => item.Title || '')
          .filter((title) => title.length > 0);

        setItems(titles);*/
      } catch (error) {
        console.error('Error loading list items:', error);
      } finally {
        //setLoading(false);
      }
    };

    if (props.context) {
      loadItems().catch((error) => {
        console.error('Error loading list items:', error);
      });
    }
  }, [props.context]);

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

          // Reset current selection

          setSelectedProduct(null);
          setSelectedClient(null);
          setSearchText("");

          // Clear options
          setProducts([]);
          setClients([]);
          setItems_AllProducts([]);
        }}
      >
        <ToggleButton value="product">
          Product
        </ToggleButton>

        <ToggleButton value="client">
          Client
        </ToggleButton>
      </ToggleButtonGroup>
      <Autocomplete<any>
        fullWidth
        loading={loading}

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




          if (selectedOption) {
            setLoading(true);
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
              const results = await sp.search({
                Querytext:
                  `"${searchCodeName}" AND Path:"${path}"`,

                SelectProperties: [
                  "Title",
                  "ListItemID",
                  "Path",
                ],

                RowLimit: 500,
              });

              console.log(results.PrimarySearchResults);

              const ids: number[] = [];

              results.PrimarySearchResults.forEach((r: any) => {

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

              for (let i = 0; i < ids.length; i += 30) {


                const currentIds = ids.slice(i, i + 30);


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
                    "Manufacturer",
                    "Country",
                    "Business_x0020_Line",
                    "PIMProductCode/Title",
                    "PIMProductCode/PIMProductName"
                  )
                  .expand("PIMProductCode")
                  .filter(filter)();

                allProducts.push(...items);
              }

              const mappedData: doclib_AllProducts[] =
                allProducts.map((item: any) => ({
                  filename: item.FileLeafRef ?? "",
                  PIMProduct: item.PIMProductCode ?? [],

                  PIMProductSearchText:
                    (item.PIMProductCode ?? [])
                      .map(
                        (p: PIMProduct) =>
                          `${p.Title} ${p.PIMProductName}`
                      )
                      .join(" "),

                  BusinessLine: choiceToString(
                    item.Business_x0020_Line
                  ),

                  CountrySoldTo: choiceToString(
                    item.Country
                  ),

                  GlobalClient:
                    item.Manufacturer ?? "",
                }));

              setItems_AllProducts(mappedData);

            }
            catch (error) {
              console.error(error);
            } finally {
              setLoading(false);
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
                ? "Enter a Product Code or Description"
                : "Enter a Global Client Name"
            }
            InputProps={{
              ...params.InputProps,
              endAdornment: (
                <>
                  {loading ? (
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

      <MaterialReactTable table={table} />


    </div>
  );
};

export default HelloWorld;