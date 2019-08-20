import * as React from 'react';
const csvToJson: any = require("csvtojson");

import { sp, Web } from '@pnp/sp';
import { CurrentUser } from '@pnp/sp/src/siteusers';
import DataLayer, { SPFilter, SPFilterTree, SPFlatFilter } from './DataLayer';
import TimesheetUpload, { TimesheetRow } from './TimesheetUpload';

export interface IImportProps {
  // commit: (json: any[]) => void;
  siteUrl: string;
  timesheetListName: string;
  projectInfoListName: string;
  dataLayer: DataLayer;
  OnSubmit?: () => void;
  ChangeViewState: (view: string) => void;
}

export default class Import extends React.Component<IImportProps, any> {

  public state: any = {
    json: null,
    upload: null,
    additionalValidations: [],
    isSaving: false
  };

  private ProcessUpload(files: FileList) {
    let file = files.length > 0 ? files[0] : null;
    if (file) {
      this.ProcessFile(file);
    }
  }

  private ProcessFile(file: File) {
    let uniqueProjectCodes: string[] = [];

    let upload = TimesheetUpload.Create({
      source: file,
      dataLayer: this.props.dataLayer,
      MutateData: (data) => {
        if (data.ProjectCode && uniqueProjectCodes.indexOf(data.ProjectCode) == -1)
          uniqueProjectCodes.push(data.ProjectCode);
        return data;
      }
    });

    let promise = upload.GetData()
      .then((json) => {
        let filter = new SPFlatFilter(uniqueProjectCodes.map(v => new SPFilter("ProjectCode", "eq", `'${v}'`)), "or");
        return this.props.dataLayer.GetProjectInfoEntries(filter)
          .then((projects) => { return { projects: projects, json: json }; });
      })
      .then(({ json, projects }) => {
        return this.props.dataLayer.GetCurrentUser()
          .then(user => { return { json: json, projects: projects, user: user }; });
      })
      .then((input: any) => {
        return this.props.dataLayer.GetUploadEntries(new SPFilter("FileName", "eq", `'${upload.FileName}'`))
          .then((result) => {
            input.isDuplicate = result.length > 0;
            return input;
          });
      })
      .then(({ json, projects, user, isDuplicate }) => {
        // console.log(json);
        for (let i = 0; i < json.length; i++) {
          let csvRow = i + 2;
          let item = json[i];
          let valid = true;

          let itemValidation = {
            errors: [],
            csvRow: csvRow
          };

          if (isNaN(item.Date.getTime()) || isNaN(item.StartTime.getTime()) || isNaN(item.EndTime.getTime())) {
            if (isNaN(item.Date.getTime()))
              itemValidation.errors.push({ message: "Date is required", level: 0 });
            if (isNaN(item.StartTime.getTime()))
              itemValidation.errors.push({ message: "Start Time is required", level: 0 });
            if (isNaN(item.EndTime.getTime()))
              itemValidation.errors.push({ message: "End Time is required", level: 0 });
          } else if (item.EndTime < item.StartTime)
            itemValidation.errors.push({ message: "End Time cannot be before Start Time. If this was an overnight log, create a new entry for the overflow into the next day", level: 0 });
          if (!item.ProjectCode)
            itemValidation.errors.push({ message: "Project Code is required", level: 0 });
          else {
            let project = projects.filter(p => p.ProjectCode == item.ProjectCode)[0];
            item.Project = project;

            if (!project)
              itemValidation.errors.push({ message: `Invalid project code '${item.ProjectCode}'`, level: 0 });
            else if (project.Contractor.EMail != user.Email)
              itemValidation.errors.push({ message: "You are not the assigned contractor for this project", level: 10 });
          }
          if (!item.Details)
            itemValidation.errors.push({ message: "Details are required", level: 0 });
          if (item.Hours > 12)
            itemValidation.errors.push({ message: "Cannot log more than 12 hours in a single record - split this line into two entries", level: 0 });

          valid = itemValidation.errors.length == 0;
          item.Validation = itemValidation;
          item.Valid = valid;
        }

        let additionalValidations = [];
        if (isDuplicate)
          additionalValidations.push("There is already a record with this filename. Please rename the file and reupload");

        this.setState({ json: json, upload: upload, additionalValidations: additionalValidations });
        return json;
      });
  }

  private SubmitData(list: TimesheetRow[], upload: TimesheetUpload) {

    if (this.state.isSaving)
      return null;
    this.setState({ isSaving: true });

    let web = new Web(this.props.siteUrl);
    // let batch = web.createBatch();
    let contractor = list[0].Project.Contractor.FirstName + " " + list[0].Project.Contractor.LastName;

    let promise = web.lists.getByTitle(this.props.dataLayer.config.UploadsListName)
      .items.add({
        FileName: upload.FileName,
        EntryRangeStart: list.sort((a, b) => a.Date < b.Date ? -1 : 1)[0].Date,
        EntryRangeEnd: list.sort((a, b) => a.Date > b.Date ? -1 : 1)[0].Date,
        TotalHours: list.map(v => v.Hours).reduce((pv, cv) => pv + cv),
        ContractorPay: list.reduce((pv, v) => pv + parseFloat(v.Project.ContractorRate) * v.Hours, 0),
        Contractor: contractor
      })
      .then(({ data, item }) => {
        return item.attachmentFiles.add(upload.FileName, upload.FileData).then(() => data.Id);
      })
      .then((uploadId) => {

        let promises = [];

        list.forEach((item) => {
          promises.push(web.lists.getByTitle(this.props.timesheetListName)
            .items
            .add({
              StartTime: item.StartTime.toLocaleString(),
              EndTime: item.EndTime.toLocaleString(),
              Date: item.Date.toLocaleString(),
              ProjectCodeId: item.Project.Id,
              Details: item.Details,
              InternalNotes: item.InternalNotes,
              UploadId: uploadId,
              Contractor: item.Project.Contractor.FirstName + " " + item.Project.Contractor.LastName
            }));
        });

        return Promise.all(promises);
      });

    promise.then((v) => {
      if (this.props.OnSubmit)
        this.props.OnSubmit();

      this.setState({ isSaving: false });

      return v;
    }, (e) => {
      this.setState({ isSaving: false });
    });

    return promise;
  }

  private ErrorColor(level: number) {
    if (level < 10) return "red";
    else if (level < 20) return "orange";
    else if (level < 30) return "dodgerblue";
    else return "black";
  }

  public render(): React.ReactElement<IImportProps> {
    let data = {};
    let controls = [];

    if (this.state.json) {
      this.state.json.forEach((v) => {
        let key = !v.Project ? v.ProjectCode : v.Project.Client + " - " + v.Project.Project + ` (${v.ProjectCode})`;
        data[key] = data[key] || [];
        data[key].push(v);
      });
    }

    for (let key in data) {
      data[key] = data[key].sort((a, b) => a.StartTime < b.StartTime ? -1 : 1);
      let entries = {};
      data[key].forEach((v) => {
        let k = v.Date ? v.Date.toLocaleDateString() : "";
        entries[k] = entries[k] || [];
        entries[k].push(v);
      });

      let projectTime = 0;

      let entryData = [];
      for (let k in entries) {
        let ec = 0;
        entries[k].forEach(v => ec += v.Hours);
        projectTime += ec;

        entryData.push((<div>
          <b>{k || "[Missing Date]"} ({isNaN(ec) ? "---" : ec} hrs)</b>
          <ul>
            {entries[k].map(v => (
              <li style={{ color: v.Valid === false ? "inherit" : "inherit", cursor: "pointer" }}>
                {isNaN(v.Hours) ? "---" : v.Hours} hrs - ({v.StartTime ? v.StartTime.toLocaleTimeString() : "Invalid Time"} - {v.EndTime ? v.EndTime.toLocaleTimeString() : "Invalid Time"}) - {v.Details}
                <span style={{ display: "block" }} hidden={v.Valid}>
                  {`Row ${v.Validation.csvRow}`}
                  <ul className="fa-ul">
                    {v.Validation.errors.map(m => <li style={{ color: this.ErrorColor(m.level) }}><i className={`fa fa-li fa-exclamation-triangle`} style={{ fontSize: "small", top: ".75em", left: "-2.55em" }}></i>{m.message}</li>)}
                  </ul>
                </span>
              </li>
            ))}
          </ul>
        </div>));
      }

      controls.push((<div style={{ display: "block", marginTop: "20px" }}>
        <b style={{ fontSize: "larger" }}>{key || "[Missing Project Code]"} ({isNaN(projectTime) ? "---" : projectTime} hrs)</b>
        {entryData.map(v => (<div style={{ display: "block", marginTop: "10px", marginLeft: "20px" }}>{v}</div>))}
      </div>));
    }

    let isValid = this.state.json != null && [].concat.apply([], this.state.json.map(j => j.Validation.errors)).filter(e => e.level < 10).length == 0 && this.state.additionalValidations.length == 0;

    return (
      <div>
        <div>
          <input className={`form-control-`} accept=".csv" type="file" onChange={(e) => this.ProcessUpload(e.target.files)} />
        </div>
        <br />
        <div hidden={this.state.json === null}>
          {controls}
        </div>
        <ul hidden={this.state.additionalValidations.length == 0} style={{ color: "red" }}>
          {this.state.additionalValidations.map(v => (<li>{v}</li>))}
        </ul>
        <button className={`btn btn-primary`} disabled={this.state.isSaving} onClick={() => this.props.ChangeViewState("userDisplay")}>Cancel</button>
        &emsp;
          <button className={`btn btn-primary`} disabled={!isValid || this.state.isSaving} onClick={() => this.SubmitData(this.state.json, this.state.upload)}>Submit Hours <i className="fa fa-spinner fa-spin" hidden={!this.state.isSaving}></i></button>
      </div>
    );
  }
}
