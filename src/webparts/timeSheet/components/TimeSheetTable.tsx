import * as React from 'react';
import * as ReactDom from 'react-dom';
import styles from './TimeSheet.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';

import DataGrid, {
  Column,
  Grouping,
  GroupPanel,
  Pager,
  Paging,
  SearchPanel,
  ValueFormat,
  OperationDescriptions
} from 'devextreme-react/data-grid';

import dxDatagGrid from 'devextreme/ui/data_grid';

import { sp, Web } from '@pnp/sp';

export interface ITableButton {
  options: any;
  onClick: (selectedData: any[]) => void;
}

export interface ITimeSheetTableProps {
  items: Promise<any[]>;
  customButtons?: ITableButton[];
  OnGetRows?: (rows: any[]) => void;
  OnCustomExport?: (selectedData: any[]) => void;
  OnInitialized?: (gridInstance: dxDatagGrid) => void;
}

export default class TimeSheetTable extends React.Component<ITimeSheetTableProps & any, any> {

  private instance: dxDatagGrid = null;

  public state: any = {
    data: []
  };

  private GetData() {
    this.instance.beginCustomLoading("Retrieving Data...");
    return this.props.items.then((data) => {
      this.instance.endCustomLoading();
      this.setState({ data: data }, () => this.forceUpdate());
      return data;
    });
  }

  public componentDidMount() {
    this.GetData();
  }

  public render(): React.ReactElement<ITimeSheetTableProps> {
    let props = {
      dataSource: this.state.data,
      allowColumnReordering: true,
      allowColumnResizing: true,
      paging: { enabled: false, pageSize: 10 },
      filterRow: { 
        visible: true,
        showOperationChooser: true
      },
      headerFilter: {
        visible: true
      },
      groupPanel: { visible: true },
      export: {
        enabled: true,
        allowExportSelectedData: true
      },
      searchPanel: { visible: true },
      columnChooser: { enabled: true },
      onExporting: (e) => {
        window["excelData"] = e;
      },
      onExported: () => {
  
        let e = window["excelData"];
        e.cancel = true;
        delete window["excelData"];
  
        setTimeout(() => {
          let url = window.URL.createObjectURL(e.data);
          let a = document.createElement("a");
          a.href = url;
          a.download = e.fileName;
          a.click();
          window.URL.revokeObjectURL(url);
        }, 1);
      },
      onToolbarPreparing: (e) => {
        let toolbarItems = e.toolbarOptions.items;

        let buttons = [];
        if (this.props.customButtons)
          buttons = buttons.concat(this.props.customButtons);
        if (this.props.OnCustomExport)
          buttons.unshift({
            onClick: this.props.OnCustomExport,
            icon: "variable"
          });

        let getRows = () => {
          let flatten = (arr: any[]) => {
            return [].concat.apply([], (arr || []).map(v => {
              let val;
              if (v.items !== undefined && v.key !== undefined)
                val = flatten(v.items);
              else 
                val = v;
              return val;
            }));
          };
          let rows = this.instance.getSelectedRowsData() as any[];
          if (rows.length == 0) {
            rows = flatten(this.instance.getDataSource().items());
          }
          return rows;
        };
        // Adds a new item
        toolbarItems.unshift.apply(toolbarItems, buttons.map(b => {
          return {
              location: "after",
              widget: "dxButton",
              options: Object["assign"]({}, b.options, {
                onClick: () => {
                  b.onClick(getRows());
                }
              })
            };
          }));
      }
    };

    return (
      <DataGrid
        id={`datagrid-${Math.round(Math.random() * 1000)}`}
        {...Object["assign"]({}, props, this.props)}
        onInitialized={(dx) => {
          this.instance = dx.component;
          if (typeof this.props.OnInitialized === "function")
            this.props.OnInitialized(dx.component);
        }}
      >
        {this.renderColumns()}
      </DataGrid>
    );
  }

  private renderColumns() {
    return this.props.children || this.defaultColumns().props.children;
  }

  private defaultColumns() {
    return (<div>
      <Column
        dataField="Date"
        dataType="date"
        sortOrder="desc"
      ></Column>
      <Column
        caption="Project Code"
        dataField="ProjectCode.ProjectCode"
        dataType="string"
        visible={false}
      ></Column>
      <Column
        caption="Client"
        dataField="ProjectCode.Client"
        groupIndex={0}
        dataType="string"
      ></Column>
      <Column
        caption="Project"
        dataField="ProjectCode.Project"
        groupIndex={1}
        dataType="string"
      ></Column>
      <Column
        dataField="Details"
        dataType="text"
      ></Column>
      <Column
        dataField="InternalNotes"
        dataType="string"
      ></Column>
      <Column
        dataField="Created"
        dataType="date"
        visible={false}
        ></Column>
      <Column
        dataField="Author.EMail"
        caption="Email"
        dataType="string"
        visible={false}
        ></Column>
      <Column
        caption="Contractor"
        dataType="string"
        calculateCellValue={({ Author, Contractor }) => {
          return Contractor || (Author.FirstName + " " + Author.LastName);
        }}
        calculateDisplayValue={function (data) {
          this.calculateCellValue(data);
        }}
        calculateSortValue={function (data) {
          this.calculateCellValue(data);
        }}
        filterOperations={[ "contains", "notcontains", "startswith", "endswith", "=", "<>" ]}
        cellTemplate={($container, { data }) => {
          ReactDom.render((<span>
            {data.Contractor || (data.Author.FirstName + " " + data.Author.LastName)}
          </span>), $container);
        }}
        allowFiltering={true}
        allowSearch={true}
      ></Column>
      <Column
        dataField="Hours"
        dataType="number"
      ></Column>
      <Column
        caption="Contractor Rate"
        dataType="number"
        dataField="ProjectCode.ContractorRate"
        cellTemplate={($container, { value }) => {
          ReactDom.render((<span>
            $ {value.toFixed(2)}
          </span>), $container);
        }}
      ></Column>
      <Column
        caption="Contractor Pay"
        dataType="number"
        calculateCellValue={rowData => {
          return rowData.Hours * rowData.ProjectCode.ContractorRate;
        }}
        cellTemplate={($container, { value }) => {
          ReactDom.render((<span>
            $ {value.toFixed(2)}
          </span>), $container);
        }}
      ></Column>
      <Column
        caption="Billable Rate"
        dataType="number"
        dataField="ProjectCode.BillableRate"
        cellTemplate={($container, { value }) => {
          ReactDom.render((<span>
            $ {value.toFixed(2)}
          </span>), $container);
        }}
      ></Column>
      <Column
        caption="Client Invoice Amt."
        dataType="number"
        calculateCellValue={rowData => {
          return rowData.Hours * rowData.ProjectCode.BillableRate;
        }}
        cellTemplate={($container, { value }) => {
          ReactDom.render((<span>
            $ {value.toFixed(2)}
          </span>), $container);
        }}
      ></Column>
      <Column
        visible={false}
        dataField="Approved"
        dataType="boolean"
        cellTemplate={($container, { data }) => {
          ReactDom.render((<span>
            {data.Approved ? "Yes" : "No"}
          </span>), $container);
        }}
      ></Column>
      <Column
        caption="Invoiced"
        dataField="Invoice"
        dataType="boolean"
        filterValue={false}
        calculateCellValue={({ Invoice }) => {
          return Invoice != null;
        }}
        cellTemplate={($container, { value }) => {
          ReactDom.render((<span>
            {value ? "Yes" : "No"}
          </span>), $container);
        }}
      ></Column>
      <Column
        visible={false}
        caption={"Invoice Paid"}
        dataField="Invoiced"
        dataType="boolean"
        cellTemplate={($container, { value }) => {
          ReactDom.render((<span>
            {value ? "Yes" : "No"}
          </span>), $container);
        }}
      ></Column>
      <Column
        caption="Contractor Payment Status"
        dataField="Payment.Status"
        dataType="text"
      ></Column>
    </div>);
  }
}
