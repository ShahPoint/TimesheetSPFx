import * as React from 'react';
const csvToJson: any = require("csvtojson");

import { sp, Web, Items } from '@pnp/sp';
import { CurrentUser } from '@pnp/sp/src/siteusers';
import DataLayer, { SPFilter, SPFilterTree, SPFlatFilter } from './DataLayer';
import TimesheetUpload, { TimesheetRow } from './TimesheetUpload';
import RowSummary from './RowSummary';

export interface IBurndownProps {
  items: Promise<any[]>;
  dataLayer: DataLayer;
  OnSubmit?: () => void;
  ChangeViewState: (view: string) => void;
}

export default class Burndown extends React.Component<IBurndownProps, any> {

  public state: any = {
    timesheetRows: []
  };

  private GetData() {
    return this.props.items.then((d) => {
      this.setState({ timesheetRows: d }, () => this.forceUpdate());
      return d;
    });
  }

  public componentDidMount() {
    this.GetData();
  }

  public render(): React.ReactElement<IBurndownProps> {

    let summary = (<div>
      Total Hours: {this.state.timesheetRows.reduce((prev, item) => parseFloat(item.Hours) + prev, 0)}
    </div>);

    let collection = (<RowSummary items={this.state.timesheetRows}></RowSummary>);

    return (
      <div>
        <div hidden={this.state.timesheetRows === null}>
          {summary}
          <br />
          {collection}
        </div>
      </div>
    );
  }
}