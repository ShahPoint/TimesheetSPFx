import * as React from 'react';
import * as ReactDom from 'react-dom';
import styles from './TimeSheet.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import TimeSheetData from './TimeSheetData';
import Import from './Import';
import InvoiceView from './InvoiceView';
import Invoice from './Invoice';
import InvoiceModal from './InvoiceModal';
import DataLayer, { SPFilter, IDataLayerInput, SPFilterTree } from './DataLayer';
import QuickView from './TimesheetQuickView';
import Payment from './ContractorPayment';
import PaymentModal from './PaymentModal';
import PaymentView from './PaymentView';
import PaymentViewModal from './PaymentViewModal';

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
import TimeSheetTable from './TimeSheetTable';
import { getGUID } from '@pnp/common';

import * as $ from 'jquery';

export interface ITimeSheetProps {
  admin: boolean;
  devMode: boolean;
}

/*

REFERENCES

  LISTS
    Invoices: https://byrdttoli.sharepoint.com/sites/Sandbox1/Lists/Invoices/AllItems.aspx
    Contractor Payments: https://byrdttoli.sharepoint.com/sites/Sandbox1/Lists/ContractorPayments/AllItems.aspx
    Timesheet Uploads: https://byrdttoli.sharepoint.com/sites/Sandbox1/Lists/TimesheetUploads/AllItems.aspx
    Timesheet Line Items: https://byrdttoli.sharepoint.com/sites/Sandbox1/Lists/Timesheet/AllItems.aspx
    Project Definitions: https://byrdttoli.sharepoint.com/sites/Sandbox1/Lists/ProjectInformation/AllItems.aspx

  DEV RESOURCES
    Workbench: https://byrdttoli.sharepoint.com/sites/Sandbox1/_layouts/15/workbench.aspx
    Live Site: https://byrdttoli.sharepoint.com/sites/Sandbox1/SitePages/Timesheet.aspx
    App Admin: https://byrdttoli.sharepoint.com/sites/Apps/AppCatalog/Forms/AllItems.aspx#InplviewHashae320ad2-cbf0-4530-82f3-a8996b7ebdcc=WebPartID%3D%7BAE320AD2--CBF0--4530--82F3--A8996B7EBDCC%7D

  LIBRARY RESOURCES
    Dx Data Grid: https://js.devexpress.com/Documentation/ApiReference/UI_Widgets/dxDataGrid/
    Dx UI Button: https://js.devexpress.com/Documentation/ApiReference/UI_Widgets/dxButton/
    Dx UI Icons: https://js.devexpress.com/Documentation/Guide/Themes_and_Styles/Icons/
    Font Awesome Icons: https://fontawesome.com/v4.7.0/icons/
    Bootstrap: https://getbootstrap.com/docs/4.3/getting-started/introduction/

*/

export default class TimeSheet extends React.Component<ITimeSheetProps, any> {

  private static Config: IDataLayerInput = {
    SiteUrl: "https://byrdttoli.sharepoint.com/sites/Sandbox1",
    TimesheetListName: "Timesheet",
    ProjectInfoListName: "ProjectInformation",
    UploadsListName: "TimesheetUploads",
    InvoiceListName: "Invoices",
    PaymentsListName: "ContractorPayments"
  };
  private static DataLayer: DataLayer = new DataLayer(TimeSheet.Config);

  public state: any = {
    viewState: "default",
    adminSub: "payouts"
  };

  public ChangeViewState(
    newState:
      "" | "default" | "import" | "display" | "adminQuickView" | "userDisplay" | "userQuickView" | "adminEntries" | "userEntries" |
      "createInvoice" | "invoices" | "invoiceView" | "createPayment" | "payments" | "paymentView" | "userPaymentsView",
    stateVars: any = {}
  ) {
    this.setState(Object["assign"]({}, { viewState: newState }, stateVars), () => { this.forceUpdate(); });
  }

  public render(): React.ReactElement<ITimeSheetProps> {
    let view;

    switch (this.state.viewState) {
      case "import":
        view = this.renderImport();
        break;
      case "display":
        view = this.renderDisplay();
        break;
      case "adminQuickView":
        view = this.renderAdminQuickView();
        break;
      case "userPaymentsView":
        view = this.renderUserPayments();
        break;
      case "userQuickView":
        view = this.renderUserQuickView();
        break;
      case "userDisplay":
        view = this.renderUserDisplay();
        break;
      case "adminEntries":
        view = this.renderAdminEntryView();
        break;
      case "userEntries":
        view = this.renderUserEntryView();
        break;
      // case "createInvoice":
      //   view = this.renderCreateInvoice();
      //   break;
      case "invoices":
        view = this.renderInvoices();
        break;
      case "invoiceView":
        view = this.renderViewInvoice();
        break;
      // case "createPayment":
      //   view = this.renderCreatePayment();
      //   break;
      case "payments":
        view = this.renderPayments();
        break;
      case "paymentView":
        view = this.renderViewPayment();
        break;
      default:
        this.state.viewState = this.props.admin === true ? "display" : "userDisplay";
        view = this.props.admin === true ? this.renderDisplay() : this.renderUserDisplay();
        break;
    }

    let $modal = null;
    
    return <div>{view}</div>;
  }

  private renderInvoices(): React.ReactElement<ITimeSheetProps> {

    let DeleteInvoice = (id) => {
      if (confirm("Are you sure you want to delete this item?")) {
        TimeSheet.DataLayer.DeleteInvoiceEntry(id).then(() => {
          this.setState({ viewState: "adminQuickView" }, () => { this.setState({ viewState: "invoices" }); });
        });
      }
    };

    return (
      <div>
        <div>
          {this.renderAdminTabs()}
          <TimeSheetTable items={TimeSheet.DataLayer.GetInvoiceEntries()}
            summary={
              {
              totalItems: [
                {
                  column: "InvoiceAmount",
                  summaryType: "sum",
                  customizeText: ({ value, valueText }) => `$ ${value ? value.toFixed(2) : "0.00"}`
              },
              {
                  column: "TotalHours",
                  summaryType: "sum",
                  customizeText: ({ value, valueText }) => value
              }
              ]
            }}
          >
            <Column
              caption="Actions"
              dataField="Id"
              cellTemplate={($container, { data, value }) => {
                ReactDom.render((<span>
                  <button className="btn btn-sm btn-outline-secondary" onClick={() => this.ChangeViewState("invoiceView", { invoiceId: data.Id })}>View</button>
                  <span hidden={data.Status === "Paid"}>&nbsp;</span>
                  <button hidden={data.Status === "Paid"} className="btn btn-sm btn-outline-warning" onClick={() => DeleteInvoice(value)}>Delete</button>
                </span>), $container);
              }}
            ></Column>
            <Column
              dataField="Status"
              dataType="text"
            ></Column>
            <Column
              dataField="Customer"
              dataType="text"
            ></Column>
            <Column
              dataField="Created"
              dataType="datetime"
              sortOrder="desc"
            ></Column>
            <Column
              dataField="InvoiceAmount"
              dataType="number"
              cellTemplate={($container, { value }) => {
                ReactDom.render((<span>
                  $ {value ? value.toFixed(2) : "0.00"}
                </span>), $container);
              }}
            ></Column>
            <Column
              dataField="TotalHours"
              dataType="number"
            ></Column>
          </TimeSheetTable>
        </div>
      </div>
    );
  }

  private showPaymentViewModal = null;
  private hidePaymentViewModal = null;

  private renderPayments(): React.ReactElement<ITimeSheetProps> {

    let DeletePayment = (id) => {
      if (confirm("Are you sure you want to delete this item?")) {
        TimeSheet.DataLayer.DeletePaymentEntry(id).then(() => {
          this.setState({ viewState: "adminQuickView" }, () => { this.setState({ viewState: "payments" }); });
        });
      }
    };

    return (
      <div>
        <div>
          <div>
          {this.renderAdminTabs()}
          <PaymentViewModal
            OnMount={(show, hide) => {
              this.showPaymentViewModal = show;
              this.hidePaymentViewModal = hide;
            }}
            dataLayer={TimeSheet.DataLayer}
          />
          <TimeSheetTable items={TimeSheet.DataLayer.GetPaymentEntries()}
            summary={
              {
              totalItems: [
                {
                  column: "PaymentAmount",
                  summaryType: "sum",
                  customizeText: ({ value, valueText }) => `$ ${value ? value.toFixed(2) : "0.00"}`
              },
              {
                  column: "TotalHours",
                  summaryType: "sum",
                  customizeText: ({ value, valueText }) => value
              }
              ]
            }}
          >
            <Column
              caption="Actions"
              dataField="Id"
              cellTemplate={($container, { data, value }) => {
                ReactDom.render((<span style={{"textAlign": "left", "width": "100%"}}>
                  {/* <button className="btn btn-sm btn-outline-secondary" onClick={() => this.ChangeViewState("paymentView", { invoiceId: data.Id })}>View</button> */}
                  <button className="btn btn-sm btn-outline-secondary" onClick={() => this.showPaymentViewModal(data.Id)}>View</button>
                  <span hidden={data.Status === "Paid"}>&nbsp;</span>
                  <button hidden={data.Status === "Paid"} className="btn btn-sm btn-outline-warning" onClick={() => DeletePayment(value)}>Delete</button>
                </span>), $container);
              }}
            ></Column>
            <Column
              dataField="Status"
              dataType="text"
            ></Column>
            <Column
              dataField="Contractor"
              dataType="string"
              calculateCellValue={({ Author, Contractor }) => {
                return Contractor;
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
                  {data.Contractor}
                </span>), $container);
              }}
              allowFiltering={true}
              allowSearch={true}
            ></Column>
            <Column
              dataField="Created"
              dataType="datetime"
              sortOrder="desc"
              visible={false}
            ></Column>
            <Column
              dataField="PaymentAmount"
              dataType="number"
              cellTemplate={($container, { value }) => {
                ReactDom.render((<span>
                  $ {value ? value.toFixed(2) : "0.00"}
                </span>), $container);
              }}
            ></Column>
            <Column
              dataField="TotalHours"
              dataType="number"
            ></Column>
            <Column
              dataField="Clients"
            ></Column>
            <Column
              dataField="EntryRangeStart"
              caption="Entry Range Start"
              dataType="date"
            ></Column>
            <Column
              dataField="EntryRangeEnd"
              caption="Entry Range End"
              dataType="date"
            ></Column>
          </TimeSheetTable>
        </div>
        </div>
      </div>
    );
  }

  private renderUserPayments(): React.ReactElement<ITimeSheetProps> {

    let payments = TimeSheet.DataLayer.GetCurrentUser().then((u) => TimeSheet.DataLayer.GetPaymentEntries(new SPFilter("Contractor", "eq", `'${u.Title}'`)));

    return (
      <div>
        <div>
          <div>
          {this.renderUserTabs()}
          <TimeSheetTable items={payments}
            summary={
              {
              totalItems: [
                {
                  column: "PaymentAmount",
                  summaryType: "sum",
                  customizeText: ({ value, valueText }) => `$ ${value ? value.toFixed(2) : "0.00"}`
              },
              {
                  column: "TotalHours",
                  summaryType: "sum",
                  customizeText: ({ value, valueText }) => value
              }
              ]
            }}
          >
            {/* <Column
              caption="Actions"
              dataField="Id"
              cellTemplate={($container, { data, value }) => {
                ReactDom.render((<span>
                  <button className="btn btn-sm btn-outline-secondary" onClick={()F => this.ChangeViewState("paymentView", { invoiceId: data.Id })}>View</button>
                </span>), $container);
              }}
            ></Column> */}
            <Column
              dataField="Status"
              dataType="text"
            ></Column>
            {/* <Column
              dataField="Contractor"
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
            ></Column> */}
            <Column
              dataField="Created"
              dataType="datetime"
              sortOrder="desc"
            ></Column>
            <Column
              dataField="PaymentAmount"
              dataType="number"
              cellTemplate={($container, { value }) => {
                ReactDom.render((<span>
                  $ {value ? value.toFixed(2) : "0.00"}
                </span>), $container);
              }}
            ></Column>
            <Column
              dataField="TotalHours"
              dataType="number"
            ></Column>
          </TimeSheetTable>
        </div>
        </div>
      </div>
    );
  }

  private renderDisplay(): React.ReactElement<ITimeSheetProps> {

    let DeleteEntry = (id) => {
      if (confirm("Are you sure you want to delete this item?")) {
        TimeSheet.DataLayer.DeleteTimesheetUpload(id).then(() => {
          this.setState({ viewState: "adminQuickView" }, () => { this.setState({ viewState: "display" }); });
        });
      }
    };

    return (
      <div>
        {this.renderAdminTabs()}
        <TimeSheetTable items={TimeSheet.DataLayer.GetUploadEntries()}
          summary={
            {
            totalItems: [
            //   {
            //     column: "Contractor Pay",
            //     summaryType: "sum",
            //     customizeText: ({ value, valueText }) => `$ ${value ? value.toFixed(2) : "0.00"}`
            // },
            {
                column: "TotalHours",
                summaryType: "sum",
                customizeText: ({ value, valueText }) => value
            }
            ]
          }}
        >
          <Column
            caption="Actions"
            dataField="Id"
            cellTemplate={($container, { data, value }) => {
              ReactDom.render((<span>
                <button className="btn btn-sm btn-outline-secondary" onClick={() => this.ChangeViewState("adminQuickView", { uploadId: data.Id })}>View</button>
                <span hidden={data.Status === "Approved"}>&nbsp;</span>
                  <button hidden={data.Status === "Approved"} className="btn btn-sm btn-outline-warning" onClick={() => DeleteEntry(value)}>Delete</button>
              </span>), $container);
            }}
          ></Column>
          <Column
            dataField="Status"
            dataType="text"
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
            caption="Uploaded By"
            dataType="text"
            visible={false}
            cellTemplate={($container, { data }) => {
              ReactDom.render((<span>
                {data.Author.FirstName + " " + data.Author.LastName}
              </span>), $container);
            }}
          ></Column>
          <Column
            caption="Uploaded"
            dataField="Created"
            dataType="datetime"
            sortOrder="desc"
          ></Column>
          <Column
            dataField="TotalHours"
            dataType="number"
          ></Column>
          <Column
            caption="Range Start"
            dataField="EntryRangeStart"
            dataType="date"
          ></Column>
          <Column
            caption="Range End"
            dataField="EntryRangeEnd"
            dataType="date"
          ></Column>
        </TimeSheetTable>
        {/* <TimeSheetData listName={TimeSheet.Config.TimesheetListName} siteUrl={TimeSheet.Config.SiteUrl}></TimeSheetData> */}
      </div>
    );
  }

  private renderUserDisplay(): React.ReactElement<ITimeSheetProps> {
    return (
      <div>
        {this.renderUserTabs()}
        <hr />
        <button className={`btn btn-sm btn-secondary`} onClick={() => this.ChangeViewState("import")}><i className="fa fa-fw fa-plus"></i> New</button>
        <br />
        <br />
        <TimeSheetTable items={TimeSheet.DataLayer.GetCurrentUser().then((u) => TimeSheet.DataLayer.GetUploadEntries(new SPFilter("Author/EMail", "eq", `'${u.Email}'`)))}
          summary={
            {
            totalItems: [
            //   {
            //     column: "Contractor Pay",
            //     summaryType: "sum",
            //     customizeText: ({ value, valueText }) => `$ ${value ? value.toFixed(2) : "0.00"}`
            // },
            {
                column: "TotalHours",
                summaryType: "sum",
                customizeText: ({ value, valueText }) => value
            }
            ]
          }}
        >
          {/* <Column
            caption="Actions"
            cellTemplate={($container, { data }) => {
              ReactDom.render((<span>
                <button className="btn btn-sm btn-default" onClick={() => this.ChangeViewState("userQuickView", { uploadId: data.Id })}>View</button>
              </span>), $container);
            }}
          ></Column> */}
          <Column
            dataField="Status"
            dataType="text"
          ></Column>
          {/* <Column
            caption="Contractor"
            dataType="text"
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
            visible={false}
          ></Column> */}
          <Column
            caption="Uploaded"
            dataField="Created"
            dataType="datetime"
            sortOrder="desc"
          ></Column>
          <Column
            dataField="TotalHours"
            dataType="number"
          ></Column>
          <Column
            caption="Range Start"
            dataField="EntryRangeStart"
            dataType="date"
          ></Column>
          <Column
            caption="Range End"
            dataField="EntryRangeEnd"
            dataType="date"
          ></Column>
          <Column
            caption="Owed"
            dataField="ContractorPay"
            dataType="number"
          ></Column>
        </TimeSheetTable>
        {/* <TimeSheetData listName={TimeSheet.Config.TimesheetListName} siteUrl={TimeSheet.Config.SiteUrl}></TimeSheetData> */}
      </div>
    );
  }

  private showInvoiceModal = null;
  private hideInvoiceModal = null;
  private showPaymentModal = null;
  private hidePaymentModal = null;
  private adminEntryRefresh = null;

  private renderAdminEntryView(): React.ReactElement<ITimeSheetProps> {
    let baseFilter = new SPFilter("Approved", "eq", "1");
    // let additionalFilter = this.state.adminSub != "all" ? (this.state.adminSub == "payouts" ? new SPFilter("Payment/Id", "le", "0") : new SPFilter("Invoice/Id", "le", "0")) : null;
    // let finalFilter = additionalFilter != null ? new SPFilterTree(baseFilter, "and", additionalFilter) : baseFilter;
    let items = TimeSheet.DataLayer.GetTimesheetEntries(baseFilter);

    return (
      <div>
        {this.renderAdminTabs()}
        <hr />
        <span className={`btn-group`}>
          <button className={"btn btn-sm btn-secondary " + (this.state.adminSub === "all" ? "active" : "")} onClick={() => { this.setState({ adminSub: "all" }, () => this.forceUpdate()); }}>All Entries</button>
          <button className={"btn btn-sm btn-secondary " + (this.state.adminSub === "payouts" ? "active" : "")} onClick={() => { this.setState({ adminSub: "payouts" }, () => this.forceUpdate()); }}>Ready for Payout</button>
          <button className={"btn btn-sm btn-secondary " + (this.state.adminSub === "invoices" ? "active" : "")} onClick={() => { this.setState({ adminSub: "invoices" }, () => this.forceUpdate()); }}>Ready for Invoice</button>
        </span>
        <br />
        <br />
        <InvoiceModal
          dataLayer={TimeSheet.DataLayer}
          OnSubmit={() => this.hideInvoiceModal()}
          OnMount={(show, hide) => {
            this.showInvoiceModal = show;
            this.hideInvoiceModal = hide;
          }} />
        <PaymentModal
          dataLayer={TimeSheet.DataLayer}
          OnSubmit={() => this.hidePaymentModal()}
          OnMount={(show, hide) => {
            this.showPaymentModal = show;
            this.hidePaymentModal = hide;
          }} />
        <TimeSheetData
          OnInitialized={(component, refresh) => {
            this.adminEntryRefresh = refresh;
          }}
          summary={
            {
              totalItems: [
                {
                  column: "Contractor Pay",
                  summaryType: "sum",
                  customizeText: ({ value, valueText }) => `$ ${value ? value.toFixed(2) : "0.00"}`
                },
                {
                  column: "Client Invoice Amt.",
                  summaryType: "sum",
                  customizeText: ({ value, valueText }) => `$ ${value ? value.toFixed(2) : "0.00"}`
                },
                {
                  column: "Hours",
                  summaryType: "sum",
                  customizeText: ({ value, valueText }) => value
                }
              ]
            }}
          items={items}
          customButtons={[
            {
              options: {
                icon: "fa fa-refresh",
                hint: "Reload Data"
              },
              onClick: () => {
                this.adminEntryRefresh(TimeSheet.DataLayer.GetTimesheetEntries(baseFilter));
              }
            },
            {
              options: {
                icon: "fa fa-file-text-o",
                hint: "Compile Invoice"
              },
              onClick: (data: any[]) => {
                this.showInvoiceModal(new Promise<any[]>((resolve, reject) => resolve(data)));
              }
            },
            {
              options: {
                icon: "fa fa-usd",
                hint: "Contractor Payout"
              },
              onClick: (data: any[]) => {
                this.showPaymentModal(new Promise<any[]>((resolve, reject) => resolve(data)));
              }
            }
          ]}
        >
          <Column
            dataField="Date"
            dataType="date"
            sortOrder="desc"
          ></Column>
          <Column
            dataField="StartTime"
            dataType="datetime"
            format="hh:mm aa"
            visible={false}
          ></Column>
          <Column
            dataField="EndTime"
            dataType="datetime"
            format="hh:mm aa"
            visible={false}
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
                $ {value ? value.toFixed(2) : "0.00"}
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
                $ {value ? value.toFixed(2) : "0.00"}
              </span>), $container);
            }}
          ></Column>
          <Column
            caption="Billable Rate"
            dataType="number"
            dataField="ProjectCode.BillableRate"
            cellTemplate={($container, { value }) => {
              ReactDom.render((<span>
                $ {value ? value.toFixed(2) : "0.00"}
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
                $ {value ? value.toFixed(2) : "0.00"}
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
            filterValue={this.state.adminSub == "invoices" ? false : undefined}
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
            dataType="string"
            calculateCellValue={({ Payment }) => {
              if (Payment && Payment.Status)
                return Payment.Status;
              return "N/A";
            }}
            selectedFilterOperation={"contains"}
            filterOperations={[ "contains", "notcontains", "startswith", "endswith", "=", "<>" ]}
            filterValue={this.state.adminSub == "payouts" ? "N/A" : ""}
            lookup={{
              dataSource: [{ val: "Pending" }, { val: "Unpaid" }, { val: "Paid" }, { val: "N/A" }],
              displayExpr: "val",
              valueExpr: "val"
            }}
          ></Column>
        </TimeSheetData>
      </div>
    );
  }

  private renderUserEntryView(): React.ReactElement<ITimeSheetProps> {
    let items = TimeSheet.DataLayer.GetCurrentUser().then((u) => TimeSheet.DataLayer.GetTimesheetEntries(new SPFilter("Author/EMail", "eq", `'${u.Email}'`)));
    return (
      <div>
        {this.renderUserTabs()}
        <br />
        <TimeSheetData
          summary={
            {
            totalItems: [
              {
                column: "Contractor Pay",
                summaryType: "sum",
                customizeText: ({ value, valueText }) => `$ ${value ? value.toFixed(2) : "0.00"}`
            },
            {
                column: "Hours",
                summaryType: "sum",
                customizeText: ({ value, valueText }) => value
            }
            ]
          }}
          items={items}
        >
          <Column
            dataField="Date"
            dataType="date"
            sortOrder="desc"
          ></Column>
          <Column
            dataField="StartTime"
            dataType="datetime"
            format="hh:mm aa"
            visible={false}
          ></Column>
          <Column
            dataField="EndTime"
            dataType="datetime"
            format="hh:mm aa"
            visible={false}
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
          {/* <Column
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
          ></Column> */}
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
                $ {value ? value.toFixed(2) : "0.00"}
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
                $ {value ? value.toFixed(2) : "0.00"}
              </span>), $container);
            }}
          ></Column>
          {/* <Column
            caption="Billable Rate"
            dataType="number"
            dataField="ProjectCode.BillableRate"
            cellTemplate={($container, { value }) => {
              ReactDom.render((<span>
                $ {value ? value.toFixed(2) : "0.00"}
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
                $ {value ? value.toFixed(2) : "0.00"}
              </span>), $container);
            }}
          ></Column> */}
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
            filterValue={undefined}
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
            dataType="string"
            calculateCellValue={({ Payment }) => {
              if (Payment && Payment.Status)
                return Payment.Status;
              return "N/A";
            }}
            selectedFilterOperation={"contains"}
            filterOperations={[ "contains", "notcontains", "startswith", "endswith", "=", "<>" ]}
            filterValue={""}
            lookup={{
              dataSource: [{ val: "Pending" }, { val: "Unpaid" }, { val: "Paid" }, { val: "N/A" }],
              displayExpr: "val",
              valueExpr: "val"
            }}
          ></Column>
        </TimeSheetData>
      </div>
    );
  }

  private renderImport(): React.ReactElement<ITimeSheetProps> {
    return (
      <div>
        <button className={`btn btn-primary`} onClick={() => this.ChangeViewState("userDisplay")}>Cancel</button>
        <br />
        <br />
        <Import dataLayer={TimeSheet.DataLayer} projectInfoListName={TimeSheet.Config.ProjectInfoListName} ChangeViewState={(s: any) => this.ChangeViewState(s)} timesheetListName={TimeSheet.Config.TimesheetListName} siteUrl={TimeSheet.Config.SiteUrl} OnSubmit={() => this.ChangeViewState("userDisplay")}></Import>
      </div>
    );
  }

  // private renderCreateInvoice(): React.ReactElement<ITimeSheetProps> {
  //   return (
  //     <div>
  //       <button className={`btn btn-primary`} onClick={() => this.ChangeViewState("invoices")}>Cancel</button>
  //       <Invoice items={this.state.invoiceItems} dataLayer={TimeSheet.DataLayer} ChangeViewState={(s: any) => this.ChangeViewState(s)} OnSubmit={() => this.ChangeViewState("adminEntries")}></Invoice>
  //     </div>
  //   );
  // }

  private renderViewInvoice(): React.ReactElement<ITimeSheetProps> {
    return (
      <div>
        <button className={`btn btn-primary`} onClick={() => this.ChangeViewState("adminEntries")}>Cancel</button>
        <InvoiceView invoiceId={this.state.invoiceId} dataLayer={TimeSheet.DataLayer} ChangeViewState={(s: any) => this.ChangeViewState(s)}></InvoiceView>
      </div>
    );
  }

  // private renderCreatePayment(): React.ReactElement<ITimeSheetProps> {
  //   return (
  //     <div>
  //       <button className={`btn btn-primary`} onClick={() => this.ChangeViewState("payments")}>Cancel</button>
  //       <Payment items={this.state.paymentItems} dataLayer={TimeSheet.DataLayer} ChangeViewState={(s: any) => this.ChangeViewState(s)} OnSubmit={() => this.ChangeViewState("adminEntries")}></Payment>
  //     </div>
  //   );
  // }

  private renderViewPayment(): React.ReactElement<ITimeSheetProps> {
    return (
      <div>
        <button className={`btn btn-primary`} onClick={() => this.ChangeViewState("payments")}>Cancel</button>
        <PaymentView paymentId={this.state.invoiceId} dataLayer={TimeSheet.DataLayer} ChangeViewState={(s: any) => this.ChangeViewState(s)}></PaymentView>
      </div>
    );
  }

  private renderAdminQuickView(): React.ReactElement<ITimeSheetProps> {
    return (
      <div>
        <button className={`btn btn-primary`} onClick={() => this.ChangeViewState("display")}>Cancel</button>
        <QuickView dataLayer={TimeSheet.DataLayer} uploadId={this.state.uploadId} ChangeViewState={(s: any) => this.ChangeViewState(s)} OnSubmit={() => this.ChangeViewState("display")} admin={true}></QuickView>
      </div>
    );
  }

  private renderUserQuickView(): React.ReactElement<ITimeSheetProps> {
    return (
      <div>
        <button className={`btn btn-primary`} onClick={() => this.ChangeViewState("userDisplay")}>Back</button>
        <QuickView dataLayer={TimeSheet.DataLayer} uploadId={this.state.uploadId} ChangeViewState={(s: any) => this.ChangeViewState(s)} OnSubmit={() => this.ChangeViewState("userDisplay")} admin={false}></QuickView>
      </div>
    );
  }

  private renderAdminTabs(): React.ReactElement<ITimeSheetProps> {
    return (<div>
      <span className={`btn-group`}>
        <button className={`btn btn-primary ${this.state.viewState === "display" ? "active" : ""}`} onClick={() => this.ChangeViewState("display")}>Uploads</button>
        <button className={`btn btn-primary ${this.state.viewState === "adminEntries" ? "active" : ""}`} onClick={() => this.ChangeViewState("adminEntries")}>Approved Entries</button>
        <button className={`btn btn-primary ${this.state.viewState === "invoices" ? "active" : ""}`} onClick={() => this.ChangeViewState("invoices")}>Invoices</button>
        <button className={`btn btn-primary ${this.state.viewState === "payments" ? "active" : ""}`} onClick={() => this.ChangeViewState("payments")}>Payouts</button>
      </span>
      <span hidden={!this.props.devMode}>
        &emsp;
        <button className={`btn btn-primary`} onClick={() => this.ChangeViewState("userDisplay")}>User View <small>(Testing)</small></button>
      </span>
    </div>);
  }

  private renderUserTabs(): React.ReactElement<ITimeSheetProps> {
    return (<div>
      <span className={`btn-group`}>
        <button className={`btn btn-primary ${this.state.viewState === "userDisplay" ? "active" : ""}`} onClick={() => this.ChangeViewState("userDisplay")}>Uploads</button>
        <button className={`btn btn-primary ${this.state.viewState === "userEntries" ? "active" : ""}`} onClick={() => this.ChangeViewState("userEntries")}>Entries</button>
        <button className={`btn btn-primary ${this.state.viewState === "userPaymentsView" ? "active" : ""}`} onClick={() => this.ChangeViewState("userPaymentsView")}>Payouts</button>
      </span>
      <span hidden={!this.props.devMode || !this.props.admin}>
        &emsp;
        <button className={`btn btn-primary`} onClick={() => this.ChangeViewState("display")}>Admin View <small>(Testing)</small></button>
      </span>
    </div>);
  }
}
