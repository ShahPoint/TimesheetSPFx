import * as React from 'react';
const csvToJson: any = require("csvtojson");

import { sp, Web, Items } from '@pnp/sp';
import { CurrentUser } from '@pnp/sp/src/siteusers';
import DataLayer, { SPFilter, SPFilterTree, SPFlatFilter } from './DataLayer';
import TimesheetUpload, { TimesheetRow } from './TimesheetUpload';
import RowSummary from './RowSummary';

export interface IPaymentProps {
  paymentId: any;
  dataLayer: DataLayer;
  ChangeViewState: (view: string) => void;
}

export default class PaymentView extends React.Component<IPaymentProps, any> {

  public state: any = {
    payment: {},
    timesheetRows: [],
    errors: []
  };

  private Comment: string;

  private GetData() {
    let promise = this.props.dataLayer.web.lists.getByTitle(this.props.dataLayer.config.PaymentsListName).items.getById(this.props.paymentId).get()
      .then((data) => {
        this.Comment = data.Comment || "";
        return this.props.dataLayer.GetTimesheetEntries(SPFilterTree.RepeatingTree(data.TimesheetEntryIds.map(s => new SPFilter("Id", "eq", `'${s}'`)), "or"))
          .then((rows) => { return { timesheetRows: rows, payment: data }; });
      })
      .then(data => {
        this.setState(data, () => this.forceUpdate());
        return data;
      });

    return promise;
  }

  private SetStatus(status: "Paid" | "Unpaid") {
    let promise = this.props.dataLayer.web.lists.getByTitle(this.props.dataLayer.config.PaymentsListName).items.getById(this.props.paymentId)
      .update({
        Status: status,
        Comment: this.Comment
      })
      .then(() => {
        this.props.ChangeViewState("payments");
      });

    return promise;
  }

  public componentDidMount() {
    this.GetData();
  }

  public render(): React.ReactElement<IPaymentProps> {

    let summary = (<div>
      Payment Status: {this.state.payment.Status}
      <br />
      Total Hours: {this.state.timesheetRows.reduce((prev, item) => parseFloat(item.Hours) + prev, 0)}
      <br />
      Payment Amount: $ {this.state.timesheetRows.reduce((prev, item) => parseFloat(item.ProjectCode.ContractorRate) * parseFloat(item.Hours) + prev, 0).toFixed(2)}
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
          {this.state.payment ? <textarea rows={2} className="form-control" onChange={(e) => this.Comment = e.currentTarget.value}>{this.Comment}</textarea> : null}
        </div>
        <br />
        <button className={`btn btn-primary`} onClick={() => this.props.ChangeViewState("payments")}>Cancel</button>
        &emsp;
        <button className={`btn btn-warning`} onClick={() => this.SetStatus("Unpaid")} disabled={this.state.payment.Status === "Unpaid"}>Mark as Unpaid</button>
        &emsp;
        <button className={`btn btn-success`} onClick={() => this.SetStatus("Paid")} disabled={this.state.payment.Status === "Paid"}>Mark as Paid</button>
      </div>
    );
  }
}