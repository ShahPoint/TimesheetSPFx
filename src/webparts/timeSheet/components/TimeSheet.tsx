import * as React from 'react';
import * as ReactDom from 'react-dom';
import styles from './TimeSheet.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import TimeSheetData from './TimeSheetData';
import Import from './Import';
import InvoiceView from './InvoiceView';
import Invoice from './Invoice';
import DataLayer, { SPFilter, IDataLayerInput } from './DataLayer';
import QuickView from './TimesheetQuickView';
import Payment from './ContractorPayment';
import PaymentView from './PaymentView';

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
    viewState: "default"
  };

  public ChangeViewState(
    newState:
      "" | "default" | "import" | "display" | "adminQuickView" | "userDisplay" | "userQuickView" | "adminEntries" | "userEntries" |
      "createInvoice" | "invoices" | "invoiceView" | "createPayment" | "payments" | "paymentView",
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
      case "userQuickView":
        view = this.renderUserQuickView();
        break;
      case "userDisplay":
        view = this.renderUserDisplay();
        break;
      case "adminEntries":
        view = this.renderAdminEntryView();
        break;
      case "createInvoice":
        view = this.renderCreateInvoice();
        break;
      case "invoices":
        view = this.renderInvoices();
        break;
      case "invoiceView":
        view = this.renderViewInvoice();
        break;
      case "createPayment":
        view = this.renderCreatePayment();
        break;
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


    return <div>{view}</div>;
  }

  private renderInvoices(): React.ReactElement<ITimeSheetProps> {
    return (
      <div>
        <div>
          {this.renderAdminTabs()}
          <TimeSheetTable items={TimeSheet.DataLayer.GetInvoiceEntries()}>
            <Column
              caption="Actions"
              cellTemplate={($container, { data }) => {
                ReactDom.render((<span>
                  <button className="btn btn-sm btn-default" onClick={() => this.ChangeViewState("invoiceView", { invoiceId: data.Id })}>View</button>
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
            ></Column>
            <Column
              dataField="InvoiceAmount"
              dataType="number"
              cellTemplate={($container, { value }) => {
                ReactDom.render((<span>
                  $ {value.toFixed(2)}
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

  private renderPayments(): React.ReactElement<ITimeSheetProps> {
    return (
      <div>
        <div>
          <div>
          {this.renderAdminTabs()}
          <TimeSheetTable items={TimeSheet.DataLayer.GetPaymentEntries()}>
            <Column
              caption="Actions"
              cellTemplate={($container, { data }) => {
                ReactDom.render((<span>
                  <button className="btn btn-sm btn-default" onClick={() => this.ChangeViewState("paymentView", { invoiceId: data.Id })}>View</button>
                </span>), $container);
              }}
            ></Column>
            <Column
              dataField="Status"
              dataType="text"
            ></Column>
            <Column
              dataField="Contractor"
              dataType="text"
            ></Column>
            <Column
              dataField="Created"
              dataType="datetime"
            ></Column>
            <Column
              dataField="PaymentAmount"
              dataType="number"
              cellTemplate={($container, { value }) => {
                ReactDom.render((<span>
                  $ {value.toFixed(2)}
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
    return (
      <div>
        {this.renderAdminTabs()}
        <TimeSheetTable items={TimeSheet.DataLayer.GetUploadEntries()}>
          <Column
            caption="Actions"
            cellTemplate={($container, { data }) => {
              ReactDom.render((<span>
                <button className="btn btn-sm btn-default" onClick={() => this.ChangeViewState("adminQuickView", { uploadId: data.Id })}>View</button>
              </span>), $container);
            }}
          ></Column>
          <Column
            dataField="Status"
            dataType="text"
          ></Column>
          <Column
            caption="Contractor"
            dataType="text"
            cellTemplate={($container, { data }) => {
              ReactDom.render((<span>
                {data.Author.FirstName} {data.Author.LastName}
              </span>), $container);
            }}
          ></Column>
          <Column
            caption="Uploaded"
            dataField="Created"
            dataType="datetime"
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
        <button className={`btn btn-primary`} onClick={() => this.ChangeViewState("import")}>New</button>
        &emsp;
        {this.renderUserTabs()}
        <TimeSheetTable items={TimeSheet.DataLayer.GetCurrentUser().then((u) => TimeSheet.DataLayer.GetUploadEntries(new SPFilter("Author/EMail", "eq", `'${u.Email}'`)))}>
          <Column
            caption="Actions"
            cellTemplate={($container, { data }) => {
              ReactDom.render((<span>
                <button className="btn btn-sm btn-default" onClick={() => this.ChangeViewState("userQuickView", { uploadId: data.Id })}>View</button>
              </span>), $container);
            }}
          ></Column>
          <Column
            dataField="Status"
            dataType="text"
          ></Column>
          <Column
            caption="Contractor"
            dataType="text"
            cellTemplate={($container, { data }) => {
              ReactDom.render((<span>
                {data.Author.FirstName} {data.Author.LastName}
              </span>), $container);
            }}
            visible={false}
          ></Column>
          <Column
            caption="Uploaded"
            dataField="Created"
            dataType="datetime"
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

  private renderAdminEntryView(): React.ReactElement<ITimeSheetProps> {
    return (
      <div>
        {this.renderAdminTabs()}
        <TimeSheetData
          items={TimeSheet.DataLayer.GetTimesheetEntries(new SPFilter("Approved", "eq", "1"))}
          customButtons={[
            {
              options: {
                icon: "fa fa-file-text-o",
                hint: "Compile Invoice"
              },
              onClick: (data: any[]) => {
                this.ChangeViewState("createInvoice", { invoiceItems: new Promise<any[]>((resolve, reject) => resolve(data)) });
              }
            },
            {
              options: {
                icon: "fa fa-usd",
                hint: "Contractor Payout"
              },
              onClick: (data: any[]) => {
                this.ChangeViewState("createPayment", { paymentItems: new Promise<any[]>((resolve, reject) => resolve(data)) });
              }
            }
          ]}
        >
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

  private renderCreateInvoice(): React.ReactElement<ITimeSheetProps> {
    return (
      <div>
        <button className={`btn btn-primary`} onClick={() => this.ChangeViewState("invoices")}>Cancel</button>
        <Invoice items={this.state.invoiceItems} dataLayer={TimeSheet.DataLayer} ChangeViewState={(s: any) => this.ChangeViewState(s)} OnSubmit={() => this.ChangeViewState("adminEntries")}></Invoice>
      </div>
    );
  }

  private renderViewInvoice(): React.ReactElement<ITimeSheetProps> {
    return (
      <div>
        <button className={`btn btn-primary`} onClick={() => this.ChangeViewState("adminEntries")}>Cancel</button>
        <InvoiceView invoiceId={this.state.invoiceId} dataLayer={TimeSheet.DataLayer} ChangeViewState={(s: any) => this.ChangeViewState(s)}></InvoiceView>
      </div>
    );
  }

  private renderCreatePayment(): React.ReactElement<ITimeSheetProps> {
    return (
      <div>
        <button className={`btn btn-primary`} onClick={() => this.ChangeViewState("payments")}>Cancel</button>
        <Payment items={this.state.paymentItems} dataLayer={TimeSheet.DataLayer} ChangeViewState={(s: any) => this.ChangeViewState(s)} OnSubmit={() => this.ChangeViewState("adminEntries")}></Payment>
      </div>
    );
  }

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
        <button className={`btn btn-primary ${this.state.viewState === "payments" ? "active" : ""}`} onClick={() => this.ChangeViewState("payments")}>Payments</button>
      </span>
      <span hidden={!this.props.devMode}>
        &emsp;
        <button className={`btn btn-primary`} onClick={() => this.ChangeViewState("userDisplay")}>User View <small>(Testing)</small></button>
      </span>
    </div>);
  }

  private renderUserTabs(): React.ReactElement<ITimeSheetProps> {
    return (<div>
      {/* <span className={`btn-group`}>
        <button className={`btn btn-primary ${this.state.viewState === "userDisplay" ? "active" : ""}`} onClick={() => this.ChangeViewState("userDisplay")}>View Uploads</button>
        <button className={`btn btn-primary ${this.state.viewState === "import" ? "active" : ""}`} onClick={() => this.ChangeViewState("import")}>Upload Hours</button>
      </span> */}
      <span hidden={!this.props.devMode || !this.props.admin}>
        <button className={`btn btn-primary`} onClick={() => this.ChangeViewState("display")}>Admin View <small>(Testing)</small></button>
      </span>
    </div>);
  }
}
