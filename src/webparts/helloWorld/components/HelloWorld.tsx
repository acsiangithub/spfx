import * as React from 'react';
//import { spfi, SPFx } from '@pnp/sp';
//import '@pnp/sp/webs';
//import '@pnp/sp/lists';
//import '@pnp/sp/items';
import { IHelloWorldProps } from './IHelloWorldProps';
import { sp } from '../HelloWorldWebPart';

interface IListItem {
  Title?: string;
}

const HelloWorld: React.FC<IHelloWorldProps> = (props) => {
  const [items, setItems] = React.useState<string[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);

  React.useEffect(() => {
    const loadItems = async (): Promise<void> => {
      try {
        setLoading(true);

        // Initialize PnP JS with SPFx Context
        //const sp = spfi().using(SPFx(props.context));

        // PnP JS v4 Syntax to get list items with selected fields
        const results: IListItem[] = await sp.web.lists
          .getByTitle('PIM Product')
          .items.select('Title') // Note the empty parenthesis () at the end instead of .get()
          .filter("Title eq 'PIM000139469'")
          ();

        const titles = results
          .map((item) => item.Title || '')
          .filter((title) => title.length > 0);

        setItems(titles);
      } catch (error) {
        console.error('Error loading list items:', error);
      } finally {
        setLoading(false);
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
      <h2>Hello from list1</h2>
      <p>{props.description}</p>

      {loading ? (
        <p>Loading...</p>
      ) : items.length > 0 ? (
        <ul>
          {items.map((title, index) => (
            <li key={`${title}-${index}`}>{title}</li>
          ))}
        </ul>
      ) : (
        <p>No items found.</p>
      )}
    </div>
  );
};

export default HelloWorld;