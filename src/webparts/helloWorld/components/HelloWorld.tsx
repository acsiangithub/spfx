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


/*type Person = {
  name: {
    firstName: string;
    lastName: string;
  };
  address: string;
  city: string;
  state: string;
};*/
type PIMProduct = {
  Title: string;
  PIMProductName: string;
  PIMProductSearchText: string;
};

type doclib_AllProducts = {
  filename: string;
  PIMProduct: PIMProduct[];
  BusinessLine: string;
  CountrySoldTo: string;
  GlobalClient: string;
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



        const globalClient = "ABITEC CORPORATION";
        //const pimProduct = "";//"CAPTEX 8000";
        //${pimProduct} AND

        const results = await sp.search({
          Querytext: ` ${globalClient} AND Path:"https://dksh.sharepoint.com/sites/FileRestore/Products/*"`,
          SelectProperties: [
            "Title",            
            "ListItemID",           
            "RefinableString00"
          ],
          RowLimit: 500,
        });

        console.log(results.PrimarySearchResults);





        // Initialize PnP JS with SPFx Context
        //const sp = spfi().using(SPFx(props.context));

        // PnP JS v4 Syntax to get list items with selected fields
        const allproducts = await sp.web.lists
          .getByTitle("Clients & Products")
          .items.select("Id", "Title", "FileLeafRef", "FileRef", "Manufacturer", "Country", "Business_x0020_Line", "PIMProductCode/Title", "PIMProductCode/PIMProductName") /*"PIMProductCode/PIMProductName", "PIMProductCode/BusinessLine", "PIMProductCode/Country")*/
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
      <MaterialReactTable table={table} />


    </div>
  );
};

export default HelloWorld;