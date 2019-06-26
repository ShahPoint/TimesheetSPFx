import * as React from 'react';
import DataLayer, { SPFilter } from "./DataLayer";
import TimesheetUpload, { TimesheetRow } from './TimesheetUpload';
import { Web } from '@pnp/sp';
import RowSummary from './RowSummary';

const csvToJson: any = require("csvtojson");

export interface IQuickViewProps {
    dataLayer: DataLayer;
    uploadId: any;
    admin: boolean;

    OnSubmit?: () => void;
    ChangeViewState: (view: any) => void;
}

export default class QuickView extends React.Component<IQuickViewProps, any> {

    public state: {
      json: TimesheetRow[],
      upload: TimesheetUpload
    } = {
      json: null,
      upload: null
    };

    private Approve(status: "Approved" | "Denied" | "Pending") {
      let web = new Web(this.props.dataLayer.config.SiteUrl);
      let promise = this.props.dataLayer.GetCurrentUser().then((user) => {
        return web.lists.getByTitle(this.props.dataLayer.config.UploadsListName)
        .items.getById(this.props.uploadId)
        .update({ 
          ApprovedDate: status == "Approved" ? new Date(Date.now()) : null,
          ApprovedById: status == "Approved" ? user.Id : null,
          Status: status
        })
        .then(({ item }) => item.get())
        .then((item) => {
          return this.props.dataLayer.GetTimesheetEntries(new SPFilter("UploadId", "eq", `${item.Id}`));
        })
        .then((items: any[]) => {
          let batch = web.createBatch();
          let listItems = web.lists.getByTitle(this.props.dataLayer.config.TimesheetListName).items.inBatch(batch);
          items.forEach((item) => {
            listItems.getById(item.Id).update({
              Approved: status === "Approved"
            });
          });

          return batch.execute();
        })
        .then(() => {
          if (this.props.OnSubmit)
            this.props.OnSubmit();
        });
      });
    }
  
    private GetData() {
      let upload: TimesheetUpload;
      let item = this.props.dataLayer.GetUploadEntry(this.props.uploadId);
      let promise = item.get()
        .then((data) => {
          return item.attachmentFiles.getByName(data.FileName).getText().then((f) => { data.File = f; return data; });
        })
        .then((data) => {
          upload = TimesheetUpload.CreateFromSP({
            fileName: data.FileName,
            dataLayer: this.props.dataLayer,
            fileData: data.File,
          });
          return this.props.dataLayer.GetTimesheetEntries(new SPFilter("UploadId", "eq", this.props.uploadId));
        })
        .then((timesheetData) => {
          this.setState({ json: timesheetData, upload: upload }, () => this.forceUpdate());
        });
    }

    public componentDidMount() {
      this.GetData();
    }
  
    public render(): React.ReactElement<IQuickViewProps> {
      return (
        <div>
          <div hidden={this.state.json === null}>
            {this.state.json != null ? (<RowSummary items={this.state.json}></RowSummary>) : null}
            <div hidden={!this.props.admin}>
              <button className="btn btn-primary" onClick={() => { this.props.ChangeViewState("display"); }}>Cancel</button>
              &emsp;
              <button className="btn btn-warning" onClick={() => { this.Approve("Denied"); }}>Deny</button>
              &emsp;
              <button className="btn btn-success" onClick={() => { this.Approve("Approved"); }}>Approve</button>
            </div>
          </div>
        </div>
      );
    }
  }
  