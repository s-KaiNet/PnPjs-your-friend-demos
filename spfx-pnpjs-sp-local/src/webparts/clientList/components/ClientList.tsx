import * as React from 'react';
import { IClientListProps } from './IClientListProps';
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  IDetailsList,
} from 'office-ui-fabric-react/lib/DetailsList';
import { createRef } from 'office-ui-fabric-react/lib/Utilities';

import { sp } from "@pnp/sp";

const _columns: IColumn[] = [
  {
    key: 'name',
    name: 'Name',
    fieldName: 'name',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'address',
    name: 'Address',
    fieldName: 'address',
    minWidth: 50,
    maxWidth: 150,
    isResizable: true
  },
  {
    key: 'email',
    name: 'Email',
    fieldName: 'email',
    minWidth: 50,
    maxWidth: 200,
    isResizable: true
  },
  {
    key: 'company',
    name: 'Company',
    fieldName: 'company',
    minWidth: 50,
    maxWidth: 150,
    isResizable: true
  }
];

export default class ClientList extends React.Component<IClientListProps, {
  items: {}[];
}> {

  private _detailsList = createRef<IDetailsList>();

  constructor(props: IClientListProps) {
    super(props);

    this.state = {
      items: []
    };
  }

  public componentDidMount(): void {
    sp.web.lists.getByTitle('Clients').items.getAll()
      .then((listItems: any[]) => {
        
        let items = listItems.map(item => {
          return {
            name: item.Title,
            address: item.Address,
            email: item.Email,
            company: item.Company
          };
        });

        this.setState({
          items
        });

      });
  }

  public render(): React.ReactElement<IClientListProps> {
    const { items } = this.state;

    return (
      <div>
        <h2>Your Clients</h2>
        <DetailsList
          componentRef={this._detailsList}
          items={items}
          columns={_columns}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}
          selectionPreservedOnEmptyClick={true}
          ariaLabelForSelectionColumn="Toggle selection"
          ariaLabelForSelectAllCheckbox="Toggle selection for all items"
        />
      </div>
    );
  }
}
