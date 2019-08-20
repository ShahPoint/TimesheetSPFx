import * as React from 'react';
const csvToJson: any = require("csvtojson");

import { sp, Web, Items } from '@pnp/sp';
import { CurrentUser } from '@pnp/sp/src/siteusers';
import DataLayer, { SPFilter, SPFilterTree, SPFlatFilter } from './DataLayer';
import TimesheetUpload, { TimesheetRow } from './TimesheetUpload';
import RowSummary from './RowSummary';

export interface ILoadingProps {

}

export default class LoadingIndicator extends React.Component<ILoadingProps, any> {

  public render(): React.ReactElement<ILoadingProps> {
    return (<div>
      <div style={{width: "50px", height: "50px", backgroundColor: "#00000080"}}></div>
      <i className="fa fa-spinner fa-spin"></i>
      Some Text Here
    </div>);
  }
}