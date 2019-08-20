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
    errors: [],
    Comment: "",
    isSavingNote: false
  };

  private GetData() {
    let promise = this.props.dataLayer.web.lists.getByTitle(this.props.dataLayer.config.PaymentsListName).items.getById(this.props.paymentId).get()
      .then((data) => {
        return this.props.dataLayer.GetTimesheetEntries(SPFilterTree.RepeatingTree(data.TimesheetEntryIds.map(s => new SPFilter("Id", "eq", `'${s}'`)), "or"))
          .then((rows) => { return { timesheetRows: rows, payment: data, Comment: data.Comment || "" }; });
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
        Comment: this.state.Comment
      })
      .then(() => {
        this.props.ChangeViewState("payments");
      });

    return promise;
  }

  private SaveComment() {
    if (this.state.isSavingNote)
      return null;

    this.setState({ isSavingNote: true });
    let promise = this.props.dataLayer.web.lists.getByTitle(this.props.dataLayer.config.PaymentsListName).items.getById(this.props.paymentId)
      .update({
        Comment: this.state.Comment
      }).then(() => {
        this.setState({ isSavingNote: false });
      });

    return promise;
  }

  public componentDidMount() {
    this.GetData();
  }

  private groupBy(array: any[], keySelector: (v: any) => any) {
    let dict = {};
    for (let i = 0; i < array.length; i++) {
      let key = keySelector(array[i]);
      dict[key] = dict[key] || [];
      dict[key].push(array[i]);
    }
    return dict;
  }

  private renderPaypalBatch() {
    let payouts = this.groupBy(this.state.timesheetRows, (v) => v.Author.EMail);
    let rawEmails = this.state.timesheetRows.map(v => v.Author.EMail).filter((value, index, self) => self.indexOf(value) === index).join(", ");

    this.props.dataLayer.GetPaypalEmails(rawEmails).then(emails => {
      let paypalBatch = [];
      for (let key in payouts) {
        let total = payouts[key].reduce((a, v) => v.ProjectCode.ContractorRate * parseFloat(v.Hours) + a, 0).toFixed(2);
        let email = emails[key] || key;
        let description = payouts[key].map(v => v.ProjectCode.Client).filter((value, index, self) => self.indexOf(value) === index).join(", ");

        paypalBatch.push(`${email},${total},USD,"${description}"`); 
      }

      this.setState({ paypalBatchData: paypalBatch.join("\r\n") });
    })
  }

  public render(): React.ReactElement<IPaymentProps> {
    if (this.state.timesheetRows)
      this.renderPaypalBatch();

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
          <label>Comment <button className="btn btn-sm btn-outline-secondary" onClick={() => this.SaveComment()}>Save <i className="fa fa-spinner fa-spin" hidden={!this.state.isSavingNote}></i></button></label>
          {this.state.payment ? <textarea rows={2} className="form-control" value={this.state.Comment} onChange={(e) => this.setState({ Comment: e.currentTarget.value }) }></textarea> : null}
        </div>
        <br />
        <div>
          <button className="btn btn-sm btn-outline-secondary" hidden={this.state.showPaypalBatch} onClick={() => this.setState({ showPaypalBatch: true })}>Show Paypal Payout CSV</button>
          <div className="form-control" hidden={!this.state.showPaypalBatch} style={{ whiteSpace: "pre-line", height: "auto" }}>{this.state.paypalBatchData || <i className="fa fa-spinner fa-spin"></i>}</div>
        </div>
        <br />
        <button className={`btn btn-primary`} onClick={() => this.props.ChangeViewState("payments")}>Cancel</button>
        &emsp;
        <button className={`btn btn-warning`} onClick={() => this.SetStatus("Unpaid")} disabled={this.state.payment.Status === "Unpaid"}>{this.state.payment.Status === "Unaid" ? "Unpaid" : "Mark as Unpaid"}</button>
        &emsp;
        <button className={`btn btn-success`} onClick={() => this.SetStatus("Paid")} disabled={this.state.payment.Status === "Paid"}>{this.state.payment.Status === "Paid" ? "Paid" : "Mark as Paid"}</button>
      </div>
    );
  }
}