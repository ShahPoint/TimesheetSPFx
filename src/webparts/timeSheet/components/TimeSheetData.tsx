import * as React from 'react';
import styles from './TimeSheet.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import TimeSheetTable, { ITableButton } from './TimeSheetTable';

import { sp, Web, Items } from '@pnp/sp';

export interface ITimeSheetDataProps {
  items: Promise<any[]>;
  customButtons?: ITableButton[];
  summary?: any;
  OnCustomExport?: (d: any) => void;
  OnInitialized?: (component: any, refresh: (items: Promise<any[]>) => void) => void;
}

export default class TimeSheet extends React.Component<ITimeSheetDataProps, any> {

  public state: any = {
    file: null,
    data: [],
    initialLoad: false
  };

  public render(): React.ReactElement<ITimeSheetDataProps> {
    return (
      <TimeSheetTable 
        OnInitialized={this.props.OnInitialized}
        items={this.props.items}
        customButtons={this.props.customButtons}
        selection={{
          allowSelectAll: true,
          mode: "multiple",
          showCheckBoxesMode: "always"
        }}
        OnCustomExport={this.props.OnCustomExport}
        summary={this.props.summary || {}}
        // summary={
        //   {
        //     totalItems: [
        //       { name: "custom", summaryType: "custom" },
        //     ],
        //     calculateCustomSummary: (options) => {
        //       // Calculating "customSummary1"
        //       if (options.name == "custom") {
        //         switch (options.summaryProcess) {
        //           case "start":
        //             options.totalValue = {
        //               records: 0,
        //               hours: 0,
        //               maxDate: null,
        //               minDate: null
        //             };
        //             break;
        //           case "calculate":
        //             options.totalValue.records++;
        //             options.totalValue.hours += parseFloat(options.value.Hours);
        //             if (options.totalValue.maxDate == null || options.totalValue.maxDate < Date.parse(options.value.Date))
        //               options.totalValue.maxDate = Date.parse(options.value.Date);
        //             if (options.totalValue.minDate == null || options.totalValue.minDate > Date.parse(options.value.Date))
        //               options.totalValue.minDate = Date.parse(options.value.Date);
        //             break;
        //           case "finalize":
        //             let daySpan = (options.totalValue.maxDate - options.totalValue.minDate) / 3600000 / 24;
        //             options.totalValue = `# Records: ${options.totalValue.records} | Total Hours: ${options.totalValue.hours} | ${daySpan} day range`;
        //             // Assigning the final value to "totalValue" here
        //             break;
        //         }
        //       }
        //     }
        //   }
        // }
        >
          {this.props.children}
        </TimeSheetTable>
      );
  }

  private CustomExport(rows: any[]) {
    const lineFormat: string = `{new Date(source.Date).toLocaleDateString()} {source.ProjectCode.Client}: {new Date(source.StartTime).toLocaleTimeString()}-{new Date(source.EndTime).toLocaleTimeString()} - {source.Details} | {source.ProjectCode.Project}: {(Math.round(source.Hours * 100) / 100)}`;
    const joinPhrase: string = `\r\n`;

    let result = rows.map((source) => Format(lineFormat, source)).join(joinPhrase);

    alert(result);

    }
}

// function Format(format: string, ...params) : string {
//   return format.replace(
//     /\{(\d+)\}/g, 
//     (str, ...captures) => {
//       let i = parseInt(captures[0]);
//       return params[i];
//     }
//   );
// }

window["Format"] = Format;
function Format(format: string, source: any) : string {
  return format.replace(
    /\{((?:[^\{\}]+|\"[^\"]*\"|\'[^\']*\'))\}/g, 
    (str, ...captures) => {
      let i = (captures[0]);
      if (i && i.length > 1 && i[0] == "'" || i[0] == '"')
        i = i.substr(1, i.length - 2);

      let val = eval(i);

      return val !== undefined && val !== null ? val.toString() : val;
    }
  );
}