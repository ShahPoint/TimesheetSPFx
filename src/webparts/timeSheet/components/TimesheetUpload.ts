import DataLayer, { SPFilter } from "./DataLayer";

const csvToJson: any = require("csvtojson");

export interface IUploadProps {
    source: File;
    dataLayer: DataLayer;

    MutateHeader?: (header: any[]) => any[];
    MutateData?: (data: TimesheetRow) => any;
}

export class TimesheetRow {
    public Date: Date;
    public StartTime: Date;
    public EndTime: Date;
    public ProjectCode: string;
    public Details: string;
    public InternalNotes: string;
    public Hours: number;
    public CsvRow: number;

    public UserEmail: any;
    public Project: any;
    public Valid: boolean;
    public Validation: any;

    public constructor(data: any) {
        this.Date = new Date(data.Date);
        this.StartTime = new Date(`${data.Date} ${data["Start Time"]}`);
        this.EndTime = new Date(`${data.Date} ${data["End Time"]}`);
        this.ProjectCode = data["Project Code"];
        this.Details = data.Details;
        this.InternalNotes = data["Internal Notes"];
        this.CsvRow = data.CsvRow;

        this.Hours = (this.EndTime.getTime() - this.StartTime.getTime()) / 3600000;
    }
}

export default class TimesheetUpload {

    public FileName: string;
    public FileData: string | ArrayBuffer;

    private constructor(private sourceFile: File, private sourcePromise: Promise<TimesheetRow[]>) {
        this.FileName = sourceFile.name;
    }

    public GetData(): Promise<TimesheetRow[]> {
        return this.sourcePromise;
    }

    private static MutateHeaderBase(header: any[]) {
        return header;
    }
    private static MutateDataBase(body: any): TimesheetRow {
        return new TimesheetRow(body);
    }

    public static Create({ source, dataLayer, MutateHeader, MutateData }: IUploadProps): TimesheetUpload {
        let upload: TimesheetUpload;

        let csvParsePromise = new Promise<TimesheetRow[]>((resolve, reject) => {
            let row = 2;
            let reader = new FileReader();
            reader.onload = (e) => {
                upload.FileData = reader.result;
                // https://www.npmjs.com/package/csvtojson#api
                csvToJson({}, { objectMode: true }).fromString(reader.result)
                    .on("header", (header: any[]) => {
                        let sub = this.MutateHeaderBase(header);
                        sub = MutateHeader ? MutateHeader(header) : sub;
                        // do a thing?
                    })
                    .on("data", (data: any) => {
                        data.CsvRow = row++;
                        let sub = this.MutateDataBase(data);
                        data["___TimesheetData"] = MutateData ? MutateData(sub) : sub;
                    })
                    .then((json: any[]) => resolve(json.map(d => d["___TimesheetData"] as TimesheetRow)), (error: any) => reject(error));
            };
            reader.readAsText(source);
        });

        upload = new TimesheetUpload(source, csvParsePromise);
        return upload;
    }

    public static CreateFromSP({ fileName, fileData, dataLayer, MutateHeader = null, MutateData = null }): TimesheetUpload {
        let upload: TimesheetUpload;

        let csvParsePromise = new Promise<TimesheetRow[]>((resolve, reject) => {
            let row = 2;
            // https://www.npmjs.com/package/csvtojson#api
            csvToJson({}, { objectMode: true }).fromString(fileData)
                .on("header", (header: any[]) => {
                    let sub = this.MutateHeaderBase(header);
                    sub = MutateHeader ? MutateHeader(header) : sub;
                    // do a thing?
                })
                .on("data", (data: any) => {
                    data.CsvRow = row++;
                    let sub = this.MutateDataBase(data);
                    data["___TimesheetData"] = MutateData ? MutateData(sub) : sub;
                })
                .then((json: any[]) => resolve(json.map(d => d["___TimesheetData"] as TimesheetRow)), (e: any) => reject(e));
        });

        upload = new TimesheetUpload(new File([], fileName), csvParsePromise);
        upload.FileData = fileData;
        return upload;
    }

}