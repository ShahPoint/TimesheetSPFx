import * as React from 'react';
const csvToJson: any = require("csvtojson");

import { sp, Web, Items } from '@pnp/sp';
import { CurrentUser } from '@pnp/sp/src/siteusers';
import DataLayer, { SPFilter, SPFilterTree, SPFlatFilter } from './DataLayer';
import TimesheetUpload, { TimesheetRow } from './TimesheetUpload';
import RowSummary from './RowSummary';
import Modal from './modal';
import * as $ from 'jquery';
import 'toastr/build/toastr.css';

import * as toastr from 'toastr';
toastr.options.positionClass = "toast-bottom-right";

export interface IPaymentModalProps {
  dataLayer: DataLayer;
  OnSubmit?: () => void;
  OnMount: (Show: (data: Promise<any[]>) => void, Hide: () => void) => void;
}

export default class PaymentModal extends React.Component<IPaymentModalProps, any> {

  private $modal: any = null
  public state: any = {
    timesheetRows: [],
    errors: []
  };

  private SubmitPayment() {
    return this.props.dataLayer.web.lists.getByTitle(this.props.dataLayer.config.PaymentsListName)
      .items.add({
        TimesheetEntryIds: { results: this.state.timesheetRows.map(v => v.Id.toString()) },
        PaymentAmount: this.state.timesheetRows.reduce((prev, item) => parseFloat(item.ProjectCode.ContractorRate) * parseFloat(item.Hours) + prev, 0),
        TotalHours: this.state.timesheetRows.reduce((prev, item) => parseFloat(item.Hours) + prev, 0),
        Contractor: this.state.timesheetRows[0].Contractor
      })
      .then(({ data, item }) => {
        let paymentId = data.Id;
        let batch = this.props.dataLayer.web.createBatch();
        let items = this.props.dataLayer.web.lists.getByTitle(this.props.dataLayer.config.TimesheetListName).items.inBatch(batch);

        for (let i = 0; i < this.state.timesheetRows.length; i++) {
          let id = this.state.timesheetRows[i].Id;
          items.getById(id).update({
            PaymentId: paymentId
          });
        }

        return batch.execute();
      })
      .then(() => {
        if (this.props.OnSubmit)
          this.props.OnSubmit();
      });
  }

  private GetData(items: Promise<any[]>) {
    return items.then((d) => {
      let errors = [];
      let numClients = d.map(v => v.Contractor).reduce((arr: string[], v: string) => {
        if (arr.indexOf(v) === -1)
          arr.push(v);
        return arr;
      }, []).length;
      
      if (d.length == 0)
        errors.push("Cannot create an empty payment");
      else if (numClients > 1)
        errors.push(`Cannot pay more than 1 contractor`);
      else if (numClients < 0)
        errors.push(`No contractors have been selected`);

      if (d.filter(v => v.Payment != null).length > 0)
        errors.push("At least one of these entries has already been paid");

      this.setState({ timesheetRows: d, errors: errors }, () => this.forceUpdate());
      return d;
    });
  }

  public render(): React.ReactElement<IPaymentModalProps> {

    let summary = (<div>
      Total Hours: {this.state.timesheetRows.reduce((prev, item) => parseFloat(item.Hours) + prev, 0)}
      <br />
      Payment Amount: $ {this.state.timesheetRows.reduce((prev, item) => parseFloat(item.ProjectCode.ContractorRate) * parseFloat(item.Hours) + prev, 0).toFixed(2)}
    </div>);

    let collection = (<RowSummary items={this.state.timesheetRows}></RowSummary>);

    return (
      <Modal
        titleContent="Contractor Payment"
        size="large"
        onMount={($m) => {
          this.$modal = $m;
          this.props.OnMount((data) => { $m.modal("show"); this.GetData(data)}, () => $m.modal("hide"));
        }}
        buttons={[
          {
            closeModal: true,
            text: "Cancel",
            type: "default"
          },
          {
            closeModal: false,
            text: "Submit Payment",
            type: "primary",
            onClick: () => {
              $(toastr.info("Submitting... Please wait")).css("background-color", "dodgerblue");
              this.SubmitPayment().then(() => {
                $(toastr.success("Payment Submitted")).css("background-color", "green");
              }, () => {
                $(toastr.warning("Payment failed to submit")).css("background-color", "darkorange");
              });
            }
          }
        ]}
        bodyContent={<div hidden={this.state.timesheetRows === null}>
          {summary}
          {collection}
          <ul hidden={this.state.errors == 0} style={{color: "red"}}>
            {this.state.errors.map(e => <li>{e}</li>)}
          </ul>
        </div>}
      />
    );
  }
}