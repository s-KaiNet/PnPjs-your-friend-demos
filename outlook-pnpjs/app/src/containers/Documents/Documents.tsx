import { sp } from '@pnp/sp';
import autobind from 'autobind-decorator';
import * as marked from "marked";
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  IDetailsList,
  Selection,
  SelectionMode
} from 'office-ui-fabric-react/lib/DetailsList';
import * as React from 'react';

import { createRef } from 'office-ui-fabric-react/lib/Utilities';
import { ITemplateItem } from '../../pnp/custom-objects/ITemplateItem';
import { TemplatesWeb } from '../../pnp/custom-objects/TemplatesWeb';
import * as styles from './Documents.css';

interface IState {
  documents: ITemplateItem[];
}

const _columns: IColumn[] = [
  {
    key: 'name',
    name: 'Name',
    fieldName: 'displayName',
    minWidth: 50,
    maxWidth: 150,
    isResizable: true
  }
];

export class Documents extends React.Component<{}, IState> {

  private _detailsList = createRef<IDetailsList>();
  private _selection = new Selection({
    selectionMode: SelectionMode.single
  });

  constructor(props: any) {
    super(props);

    this.state = {
      documents: null
    };
  }
  public componentDidMount(): void {

    sp.web.as(TemplatesWeb).TemplatesList()
      .then(list => {
        return list.getTemplates();
      })
      .then(templates => {
        this.setState({
          documents: templates
        });
      });
  }

  public render(): JSX.Element {
    if (!this.state.documents) {
      return <div>Loading...</div>;
    }

    return (
      <div className={styles.docs}>
        <h2>Select a template:</h2>
        <DetailsList
          componentRef={this._detailsList}
          items={this.state.documents}
          columns={_columns}
          selection={this._selection}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}
          selectionPreservedOnEmptyClick={true}
          ariaLabelForSelectionColumn="Toggle selection"
          ariaLabelForSelectAllCheckbox="Toggle selection for all items"
        />
        <br />
        <DefaultButton
          primary={true}
          allowDisabledFocus={true}
          text="Insert"
          onClick={this.insertText}
        />
      </div>
    );
  }

  @autobind
  private insertText() {
    const selection = this._selection.getSelection();
    if (selection.length === 0) return;
    const item = selection[0] as ITemplateItem;

    Office.context.mailbox.item.body.setSelectedDataAsync(marked(item.template), {
      coercionType: Office.CoercionType.Html
    });
  }
}
