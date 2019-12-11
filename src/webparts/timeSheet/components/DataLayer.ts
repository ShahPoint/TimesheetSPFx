import { sp, Web, Item } from '@pnp/sp';
import { string } from 'prop-types';

export type TimesheetListName = "Timesheet" | "ProjectInfo" | "Uploads" | "Invoice" | "Payments";
export interface IDataLayerInput {
    SiteUrl: string;
    TimesheetListName: string;
    ProjectInfoListName: string;
    UploadsListName: string;
    InvoiceListName: string;
    PaymentsListName: string;
}

// #region SP Query Operations
export interface ISPListFieldDictionary {
    [key: string]: string;
}

export interface ISPListField {
    fieldId: string;
    displayName: string;
}

export interface ISPFilter {
    Resolve: () => string;
}

type TreeOperator = "and" | "or";
type FilterOperator = "eq" | "ne" | "gt" | "le";

export class SPFlatFilter implements ISPFilter {
    public constructor (private elements: ISPFilter[], private operator: TreeOperator) { }

    public Resolve(): string {
        return this.elements.map(v => v.Resolve()).join(` ${this.operator} `);
    }
}

export class SPFilterTree implements ISPFilter {

    public constructor(private leftSide: ISPFilter, private operator: TreeOperator, private rightSide: ISPFilter) { }

    public Resolve(): string {
        return `(${this.leftSide.Resolve()} ${this.operator} ${this.rightSide.Resolve()})`;
    }

    public static RepeatingTree(list: ISPFilter[], operator: TreeOperator): ISPFilter {
        if (list.length == 0)
            return null;
        if (list.length == 1)
            return list[0];

        let tree = new SPFilterTree(list[0], operator, list[1]);

        return SPFilterTree.RepeatingTree(([tree] as ISPFilter[]).concat(list.filter((v, i) => i > 1)), operator);
    }
}

export class SPFilter implements ISPFilter {

    public constructor(private fieldId: string, private operator: FilterOperator, private fieldValue: string) { }

    public Resolve(): string {
        return `${this.fieldId} ${this.operator} ${this.fieldValue}`;
    }
}
// #endregion

export default class DataLayer {

    private static TimesheetColumns: ISPListField[] = [
        { displayName: "Date", fieldId: "Date" },
        { displayName: "StartTime", fieldId: "StartTime" },
        { displayName: "EndTime", fieldId: "EndTime" },
        { displayName: "Hours", fieldId: "Hours" },
        { displayName: "Upload", fieldId: "Upload" },
        { displayName: "Upload Id", fieldId: "Upload/Id" },
        { displayName: "Invoice", fieldId: "Invoice" },
        { displayName: "Invoice Id", fieldId: "Invoice/Id" },
        { displayName: "Invoiced", fieldId: "Invoiced" },
        { displayName: "Approved", fieldId: "Approved" },
        { displayName: "ProjectCode", fieldId: "ProjectCode" },
        { displayName: "ProjectCode/ProjectCode", fieldId: "ProjectCode/ProjectCode" },
        { displayName: "ProjectCode/ContractorRate", fieldId: "ProjectCode/ContractorRate" },
        { displayName: "ProjectCode/Project", fieldId: "ProjectCode/Project" },
        { displayName: "ProjectCode/Client", fieldId: "ProjectCode/Client" },
        { displayName: "ProjectCode/ContractorRate", fieldId: "ProjectCode/ContractorRate" },
        { displayName: "ProjectCode/BillableRate", fieldId: "ProjectCode/BillableRate" },
        { displayName: "Contractor", fieldId: "Contractor" },
        { displayName: "Details", fieldId: "Details" },
        { displayName: "InternalNotes", fieldId: "InternalNotes" },
        { displayName: "Created", fieldId: "Created" },
        { displayName: "Author/ID", fieldId: "Author/ID" },
        { displayName: "Author/FirstName", fieldId: "Author/FirstName" },
        { displayName: "Author/LastName", fieldId: "Author/LastName" },
        { displayName: "Author/EMail", fieldId: "Author/EMail" },
        { displayName: "Id", fieldId: "Id" },
        { displayName: "Payment", fieldId: "Payment" },
        { displayName: "Payment Status", fieldId: "Payment/Status" }
    ];

    private static ProjectInfoColumns: ISPListField[] = [
        { displayName: "ProjectCode", fieldId: "ProjectCode" },
        { displayName: "ContractorRate", fieldId: "ContractorRate" },
        { displayName: "Client", fieldId: "Client" },
        { displayName: "Id", fieldId: "Id" },
        { displayName: "Project", fieldId: "Project" },
        { displayName: "Contractor/EMail", fieldId: "Contractor/EMail" },
        { displayName: "Contractor/FirstName", fieldId: "Contractor/FirstName" },
        { displayName: "Contractor/LastName", fieldId: "Contractor/LastName" }
    ];

    private static UploadColumns: ISPListField[] = [
        { displayName: "Total Hours", fieldId: "TotalHours" },
        { displayName: "Creator First Name", fieldId: "Author/FirstName" },
        { displayName: "Creator Last Name", fieldId: "Author/LastName" },
        { displayName: "Uploaded", fieldId: "Created" },
        { displayName: "File Name", fieldId: "FileName" },
        { displayName: "Range Start", fieldId: "EntryRangeStart" },
        { displayName: "Range End", fieldId: "EntryRangeEnd" },
        { displayName: "Status", fieldId: "Status" },
        { displayName: "Approver First Name", fieldId: "ApprovedBy/FirstName" },
        { displayName: "Approver Last Name", fieldId: "ApprovedBy/LastName" },
        { displayName: "Approved Time", fieldId: "ApprovedDate" },
        { displayName: "Id", fieldId: "Id" },
        { displayName: "ContractorPay", fieldId: "ContractorPay" },
        { displayName: "Contractor", fieldId: "Contractor" },
    ];

    private static InvoiceColumns: ISPListField[] = [
        { displayName: "Id", fieldId: "Id" },
        { displayName: "Status", fieldId: "Status" },
        { displayName: "Timesheet IDs", fieldId: "TimesheetEntryIds" },
        { displayName: "Invoice Amount", fieldId: "InvoiceAmount" },
        { displayName: "Uploaded", fieldId: "Created" },
        { displayName: "Total Hours", fieldId: "TotalHours" },
        { displayName: "Customer", fieldId: "Customer" },
        { displayName: "Comment", fieldId: "Comment" },
    ];

    private static PaymentColumns: ISPListField[] = [
        { displayName: "Id", fieldId: "Id" },
        { displayName: "Status", fieldId: "Status" },
        { displayName: "Timesheet IDs", fieldId: "TimesheetEntryIds" },
        { displayName: "Payment Amount", fieldId: "PaymentAmount" },
        { displayName: "Uploaded", fieldId: "Created" },
        { displayName: "Total Hours", fieldId: "TotalHours" },
        { displayName: "Contractor", fieldId: "Contractor" },
        { displayName: "Comment", fieldId: "Comment" },
        { displayName: "Entry Range Start", fieldId: "EntryRangeStart" },
        { displayName: "Entry Range End", fieldId: "EntryRangeEnd" },
        { displayName: "Clients", fieldId: "Clients" },
    ];

    private static PaypalEmailColumns: ISPListField[] = [
        { displayName: "Paypal Email", fieldId: "Email" },
        { displayName: "Email", fieldId: "Author/EMail" }
    ];

    public web: Web;

    public constructor(public config: IDataLayerInput) {
        this.web = this.Web();
    }

    /**
     * Creates a new web object based on site configs
     */
    public Web() {
        return new Web(this.config.SiteUrl);
    }

    public List(list: TimesheetListName, web: Web = null) {
        return (web || this.Web()).lists.getByTitle(this.config[`${list}ListName`]);
    }

    public GetCurrentUser(): Promise<any> {
        return this.web.currentUser.get();
    }

    private GetListData(listTitle: string, columns: string[] = null, filter: ISPFilter = null): Promise<any> {
        let list = this.web.lists.getByTitle(listTitle).items;

        if (columns != null) {
            list = list.select(...columns);
            // assuming only one level of expansion
            let expansions = columns.filter(v => v.indexOf("/") != -1).map(v => v.split("/")[0]);
            if (expansions.length > 0)
                list = list.expand(...expansions);
        }

        if (filter != null) {
            list = list.filter(filter.Resolve());
        }

        return list.getAll();
    }

    private GetItem(listTitle: string, id: any, columns: string[] = null): Item {
        let list = this.web.lists.getByTitle(listTitle).items;

        if (columns != null) {
            list = list.select(...columns);
            // assuming only one level of expansion
            let expansions = columns.filter(v => v.indexOf("/") != -1).map(v => v.split("/")[0]);
            if (expansions.length > 0)
                list = list.expand(...expansions);
        }

        return list.getById(id);
    }

    public GetTimesheetEntries(filter: ISPFilter = null, columns: ISPListField[] = DataLayer.TimesheetColumns): Promise<any> {
        return this.GetListData(
            this.config.TimesheetListName, 
            columns.map(v => v.fieldId),
            filter
        ).then((d) => {
            return d;
        });
    }

    public GetProjectInfoEntries(filter: ISPFilter = null, columns: ISPListField[] = DataLayer.ProjectInfoColumns): Promise<any> {
        return this.GetListData(
            this.config.ProjectInfoListName, 
            columns.map(v => v.fieldId),
            filter
        );
    }

    public GetUploadEntries(filter: ISPFilter = null, columns: ISPListField[] = DataLayer.UploadColumns): Promise<any> {
        return this.GetListData(
            this.config.UploadsListName, 
            columns.map(v => v.fieldId),
            filter
        );
    }

    public GetInvoiceEntries(filter: ISPFilter = null, columns: ISPListField[] = DataLayer.InvoiceColumns): Promise<any> {
        return this.GetListData(
            this.config.InvoiceListName, 
            columns.map(v => v.fieldId),
            filter
        );
    }

    public GetPaymentEntries(filter: ISPFilter = null, columns: ISPListField[] = DataLayer.PaymentColumns): Promise<any> {
        return this.GetListData(
            this.config.PaymentsListName, 
            columns.map(v => v.fieldId),
            filter
        );
    }

    public GetUploadEntry(id: any, columns: ISPListField[] = DataLayer.UploadColumns): Item {
        return this.GetItem(
            this.config.UploadsListName,
            id,
            columns.map(v => v.fieldId)
        );
    }

    public DeletePaymentEntry(id: number) {
        return this.GetTimesheetEntries(new SPFilter("Payment/Id", "eq", id.toString())).then((entries) => {
            let items = this.web.lists.getByTitle(this.config.TimesheetListName).items;
            let promises = [];
            for (let i = 0; i < entries.length; i++) {
                let entry = entries[i];
                promises.push(items.getById(entry.Id).update({
                    PaymentId: null
                }));
            }

            return Promise.all(promises);
        })
        .then(() => {
            return this.web.lists.getByTitle(this.config.PaymentsListName).items.getById(id).delete();
        });
    }

    public DeleteInvoiceEntry(id: number) {
        return this.GetTimesheetEntries(new SPFilter("Invoice/Id", "eq", id.toString())).then((entries) => {
            let items = this.web.lists.getByTitle(this.config.TimesheetListName).items;
            let promises = [];
            for (let i = 0; i < entries.length; i++) {
                let entry = entries[i];
                promises.push(items.getById(entry.Id).update({
                    InvoiceId: null
                }));
            }

            return Promise.all(promises);
        })
        .then(() => {
            return this.web.lists.getByTitle(this.config.InvoiceListName).items.getById(id).delete();
        });
    }

    public DeleteTimesheetUpload(id: number) {
        return this.GetTimesheetEntries(new SPFilter("Upload/Id", "eq", id.toString())).then((entries) => {
            let items = this.web.lists.getByTitle(this.config.TimesheetListName).items;
            let promises = [];
            for (let i = 0; i < entries.length; i++) {
                let entry = entries[i];
                promises.push(items.getById(entry.Id).delete());
            }

            return Promise.all(promises);
        })
        .then(() => {
            return this.web.lists.getByTitle(this.config.UploadsListName).items.getById(id).delete();
        });
    }

    public GetPaypalEmails(emails: string[]) {
        return this.GetListData("PaypalEmails", DataLayer.PaypalEmailColumns.map(v => v.fieldId))
            .then((items) => {
                let dict = {};
                for (let i = 0; i < emails.length; i++) 
                    dict[emails[i]] = emails[i];

                for (let i = 0; i < items.length; i++) 
                    dict[items[i].Author.EMail] = items[i].Email;
                return dict;
            });
    }

}