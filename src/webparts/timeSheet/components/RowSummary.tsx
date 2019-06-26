import * as React from 'react';
const csvToJson: any = require("csvtojson");

import { sp, Web, Items } from '@pnp/sp';
import { CurrentUser } from '@pnp/sp/src/siteusers';
import DataLayer, { SPFilter, SPFilterTree, SPFlatFilter } from './DataLayer';
import TimesheetUpload, { TimesheetRow } from './TimesheetUpload';

export interface IRowSummaryProps {
  items: TimesheetRow[];
}

export default class RowSummary extends React.Component<IRowSummaryProps, any> {

  public render(): React.ReactElement<IRowSummaryProps> {

    let data = GroupBy(this.props.items || [], (x) => x.ProjectCode.Client);
    let controls = [];
    for (let client in data) {
      let clientEntries = data[client].filter(v => v.ProjectCode.Client == client);
      let clientData = GroupBy(clientEntries, (x) => x.ProjectCode.Project);
      let clientControls = [];
      let clientHours = clientEntries.reduce((pv, v) => pv + parseFloat(v.Hours), 0);

      for (let project in clientData) {
        let projectEntries = clientData[project].filter(v => v.ProjectCode.Project == project);
        let projectData = GroupBy(projectEntries, (x) => new Date(x.Date).toLocaleDateString());
        let projectControls = [];
        let projectHours = projectEntries.reduce((pv, v) => pv + parseFloat(v.Hours), 0);

        for (let date in projectData) {
          let dateEntries = projectData[date].filter(v => new Date(v.Date).toLocaleDateString() == date);
          let dateControls = [];
          let dateHours = dateEntries.reduce((pv, v) => pv + parseFloat(v.Hours), 0);

          dateControls = dateEntries.map(v => (
            <li style={{ cursor: "pointer" }}>
              <span>{parseFloat(v.Hours)} hrs - ({new Date(v.StartTime).toLocaleTimeString()} - {new Date(v.EndTime).toLocaleTimeString()}) - {v.Details}</span>
            </li>
          ));

          projectControls.push(<li style={{ display: "block" }}>
            <h6>{date} - {dateHours} hrs</h6>
            <ul>
              {dateControls}
            </ul>
          </li>);
        }

        clientControls.push(<li style={{ display: "block" }}>
          <h5>{project} - {projectHours} hrs</h5>
          <ul style={{paddingLeft: "10px"}}>
            {projectControls}
          </ul>
        </li>);
      }

      controls.push(<li style={{ display: "block" }}>
        <h3>{client} - {clientHours} hrs</h3>
        <ul style={{paddingLeft: "10px"}}>
          {clientControls}
        </ul>
      </li>);
    }

    return (<ul style={{paddingLeft: "0px"}}>
      {controls}
    </ul>);
  }
}

window["GroupBy"] = GroupBy;
function GroupBy(xs: any, key: string | ((d: any) => string)) {
  return xs.reduce((rv, x) => {
    let k = typeof key === "string" ? key : key(x);
    (rv[k] = rv[k] || []).push(x);
    return rv;
  }, {});
}