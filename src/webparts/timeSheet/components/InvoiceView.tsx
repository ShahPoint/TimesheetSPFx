import * as React from 'react';
const csvToJson: any = require("csvtojson");

import { sp, Web, Items } from '@pnp/sp';
import { CurrentUser } from '@pnp/sp/src/siteusers';
import DataLayer, { SPFilter, SPFilterTree, SPFlatFilter } from './DataLayer';
import TimesheetUpload, { TimesheetRow } from './TimesheetUpload';
import RowSummary from './RowSummary';

export interface IInvoiceProps {
  invoiceId: any;
  dataLayer: DataLayer;
  ChangeViewState: (view: string) => void;
}

export default class InvoiceView extends React.Component<IInvoiceProps, any> {

  public state: any = {
    invoice: {},
    timesheetRows: [],
    errors: []
  };

  private Comment: string;

  private Delete() {
    let confirmed = confirm("Are you sure you want to delete this item?");

    if (confirmed) {
      let web = this.props.dataLayer.Web();
      let batch = web.createBatch();
      let timesheetRows = this.props.dataLayer.List("Timesheet", web).items.inBatch(batch);

      this.state.invoice.TimesheetEntryIds.forEach((id) => {
        timesheetRows.getById(id).update({
          InvoiceId: null
        });
      });

      batch.execute()
        .then(() => {
          return new Promise<void>((resolve) => {
            setTimeout(() => resolve(), 100);
          });
        })
        .then(() => {
          return this.props.dataLayer.List("Invoice")
            .items.getById(this.props.invoiceId).delete()
            .then(() => {
              this.props.ChangeViewState("invoices");
            });
        });
    }
  }

  private GetData() {
    let promise = this.props.dataLayer.web.lists.getByTitle(this.props.dataLayer.config.InvoiceListName).items.getById(this.props.invoiceId).get()
      .then((data) => {
        this.Comment = data.Comment || "";
        return this.props.dataLayer.GetTimesheetEntries(SPFilterTree.RepeatingTree(data.TimesheetEntryIds.map(s => new SPFilter("Id", "eq", `'${s}'`)), "or"))
          .then((rows) => { return { timesheetRows: rows, invoice: data }; });
      })
      .then(data => {
        this.setState(data, () => this.forceUpdate());
        return data;
      });

    return promise;
  }

  private SetStatus(status: "Paid" | "Unpaid") {
    let promise = this.props.dataLayer.web.lists.getByTitle(this.props.dataLayer.config.InvoiceListName).items.getById(this.props.invoiceId)
      .update({
        Status: status,
        Comment: this.Comment
      })
      .then(() => {
        let batch = this.props.dataLayer.web.createBatch();
        let items = this.props.dataLayer.web.lists.getByTitle(this.props.dataLayer.config.TimesheetListName).items.inBatch(batch);

        this.state.timesheetRows.forEach((v) => {
          items.getById(v.Id).update({
            Invoiced: status == "Paid"
          });
        });

        return batch.execute();
      })
      .then(() => {
        this.props.ChangeViewState("invoices");
      });

    return promise;
  }

  public componentDidMount() {
    this.GetData();
  }

  public render(): React.ReactElement<IInvoiceProps> {

    let summary = (<div>
      Payment Status: {this.state.invoice.Status}
      <br />
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
        <br />
        <div>
          <label>Comment</label>
          {this.state.invoice ? <textarea rows={2} className="form-control" onChange={(e) => this.Comment = e.currentTarget.value}>{this.Comment}</textarea> : null}
        </div>
        <br />
        <button className={`btn btn-primary`} onClick={() => this.props.ChangeViewState("invoices")}>Cancel</button>
        &emsp;
        <button className={`btn btn-warning`} onClick={() => this.SetStatus("Unpaid")} disabled={this.state.invoice.Status === "Unpaid"}>Mark as Unpaid</button>
        &emsp;
        <button className={`btn btn-success`} onClick={() => this.SetStatus("Paid")} disabled={this.state.invoice.Status === "Paid"}>Mark as Paid</button>
        {/* <button className={`btn btn-danger pull-right`} onClick={() => this.Delete()} disabled={this.state.invoice.Status === "Paid"}><i className="fa fa-trash-o"></i> Delete</button> */}
      </div>
    );
  }
}