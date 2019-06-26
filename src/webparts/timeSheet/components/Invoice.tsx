import * as React from 'react';
const csvToJson: any = require("csvtojson");

import { sp, Web, Items } from '@pnp/sp';
import { CurrentUser } from '@pnp/sp/src/siteusers';
import DataLayer, { SPFilter, SPFilterTree, SPFlatFilter } from './DataLayer';
import TimesheetUpload, { TimesheetRow } from './TimesheetUpload';
import RowSummary from './RowSummary';

export interface IInvoiceProps {
  items: Promise<any[]>;
  dataLayer: DataLayer;
  OnSubmit?: () => void;
  ChangeViewState: (view: string) => void;
}

export default class Invoice extends React.Component<IInvoiceProps, any> {

  public state: any = {
    timesheetRows: [],
    errors: []
  };

  private SubmitInvoice() {
    return this.props.dataLayer.web.lists.getByTitle(this.props.dataLayer.config.InvoiceListName)
      .items.add({
        TimesheetEntryIds: { results: this.state.timesheetRows.map(v => v.Id.toString()) },
        InvoiceAmount: this.state.timesheetRows.reduce((prev, item) => parseFloat(item.ProjectCode.BillableRate) * parseFloat(item.Hours) + prev, 0),
        TotalHours: this.state.timesheetRows.reduce((prev, item) => parseFloat(item.Hours) + prev, 0),
        Customer: this.state.timesheetRows[0].ProjectCode.Client
      })
      .then(({ data, item }) => {
        let invoiceId = data.Id;
        let batch = this.props.dataLayer.web.createBatch();
        let items = this.props.dataLayer.web.lists.getByTitle(this.props.dataLayer.config.TimesheetListName).items.inBatch(batch);

        for (let i = 0; i < this.state.timesheetRows.length; i++) {
          let id = this.state.timesheetRows[i].Id;
          items.getById(id).update({
            InvoiceId: invoiceId
          });
        }

        return batch.execute();
      })
      .then(() => {
        if (this.props.OnSubmit)
          this.props.OnSubmit();
      });
  }

  private GetData() {
    return this.props.items.then((d) => {
      let errors = [];
      let numClients = d.map(v => v.ProjectCode.Client).reduce((arr: string[], v: string) => {
        if (arr.indexOf(v) === -1)
          arr.push(v);
        return arr;
      }, []).length;
      
      if (d.length == 0)
        errors.push("Cannot create an empty invoice");
      else if (numClients > 1)
        errors.push(`Cannot invoice more than 1 customer`);
      else if (numClients < 0)
        errors.push(`No clients have been selected`);

      if (d.filter(v => v.Invoice != null).length > 0)
        errors.push("At least one of these entries has already been invoiced");

      this.setState({ timesheetRows: d, errors: errors }, () => this.forceUpdate());
      return d;
    });
  }

  public componentDidMount() {
    this.GetData();
  }

  public render(): React.ReactElement<IInvoiceProps> {

    let summary = (<div>
      Total Hours: {this.state.timesheetRows.reduce((prev, item) => parseFloat(item.Hours) + prev, 0)}
      <br />
      Client Charge: $ {this.state.timesheetRows.reduce((prev, item) => parseFloat(item.ProjectCode.BillableRate) * parseFloat(item.Hours) + prev, 0).toFixed(2)}
    </div>);

    let collection = (<RowSummary items={this.state.timesheetRows}></RowSummary>);

    return (
      <div>
        <div hidden={this.state.timesheetRows === null}>
          {summary}
          {collection}
          <ul hidden={this.state.errors == 0} style={{color: "red"}}>
            {this.state.errors.map(e => <li>{e}</li>)}
          </ul>
        </div>
        <button className={`btn btn-primary`} onClick={() => this.props.ChangeViewState("adminEntries")}>Cancel</button>
        &emsp;
        <button className={`btn btn-primary`} onClick={() => this.SubmitInvoice()} disabled={this.state.errors.length > 0}>Submit Invoice</button>
      </div>
    );
  }
}