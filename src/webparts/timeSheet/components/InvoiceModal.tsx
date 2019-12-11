import * as React from 'react';
const csvToJson: any = require("csvtojson");

import { sp, Web, Items } from '@pnp/sp';
import { CurrentUser } from '@pnp/sp/src/siteusers';
import DataLayer, { SPFilter, SPFilterTree, SPFlatFilter } from './DataLayer';
import TimesheetUpload, { TimesheetRow } from './TimesheetUpload';
import RowSummary from './RowSummary';
import Modal from './modal';
import * as $ from 'jquery';

export interface IInvoiceModalProps {
  dataLayer: DataLayer;
  OnSubmit?: () => void;
  OnMount: (Show: (data: Promise<any[]>) => void, Hide: () => void) => void;
}

export default class InvoiceModal extends React.Component<IInvoiceModalProps, any> {

  private $modal: any = null
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

  private GetData(items: Promise<any[]>) {
    return items.then((d) => {
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

  public render(): React.ReactElement<IInvoiceModalProps> {

    let summary = (<div>
      Total Hours: {this.state.timesheetRows.reduce((prev, item) => parseFloat(item.Hours) + prev, 0)}
      <br />
      Client Charge: $ {this.state.timesheetRows.reduce((prev, item) => parseFloat(item.ProjectCode.BillableRate) * parseFloat(item.Hours) + prev, 0).toFixed(2)}
    </div>);

    let collection = (<RowSummary items={this.state.timesheetRows}></RowSummary>);

    return (
      <Modal
        titleContent="Invoice"
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
            text: "Submit Invoice",
            type: "primary",
            onClick: () => {
              this.SubmitInvoice();
            }
          }
        ]}
        bodyContent={<div>
          <div hidden={this.state.timesheetRows === null}>
            {summary}
            {collection}
            <ul hidden={this.state.errors == 0} style={{color: "red"}}>
              {this.state.errors.map(e => <li>{e}</li>)}
            </ul>
          </div>
        </div>}
      />
    );
  }
}