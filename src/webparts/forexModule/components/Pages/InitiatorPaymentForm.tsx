import React, { useEffect, useState } from "react";
import "../Pages/Css/NewRequest.scss";
import { IForexModuleProps } from "../IForexModuleProps";
import { useHistory, useParams } from "react-router-dom";
import SPCRUDOPS from "../../service/BAL/spcrud";
import { SPHttpClient } from "@microsoft/sp-http";
import { number } from "yup";
/* ---------- Layout Helpers ---------- */

const Section = ({ title, children }: any) => (
    <div className="form-section">
        {title && <h3>{title}</h3>}
        {children}
    </div>
);

const Grid = ({ children }: any) => (
    <div className="form-grid">{children}</div>
);

const Field = ({ label, children, full }: any) => (
    <div className={full ? "form-field full" : "form-field"}>
        <label>{label}</label>
        {children}
    </div>
);

interface InvoiceRow {
    invoiceNo: string;
    invoiceDate: string;
    boeNo: string;
    boeDate: string;
    mrnNo: string;
    mrnDate?: string;
    blNo: string;
    blDate: string;
    invoiceAmount: string;
    invoiceAmountnew?: string;
}
interface InvoiceRownew {
    invoiceNo: string;
    invoiceDate: string;
    boeNo: string;
    boeDate: string;
    mrnNo: string;
    mrnDate?: string;
    blNo: string;
    blDate: string;
    invoiceAmountnew?: string;
}
const CollapsibleSection = ({ title, children }: any) => {
    const [open, setOpen] = React.useState(false);

    return (
        <div className="form-section collapsible">
            <div
                className="form-section-header"
                onClick={() => setOpen(!open)}
            >
                <span>{title}</span>
                <i className={`fas fa-chevron-${open ? "up" : "down"}`}></i>
            </div>

            {open && (
                <div className="form-section-body">
                    {children}
                </div>
            )}
        </div>
    );
};

const TrackerForm = (props: IForexModuleProps) => {
    const { Id } = useParams<{ Id: string }>();
    const history = useHistory();
    const spCrudOps = SPCRUDOPS();

    const [paymentType, setPaymentType] = useState("");
    const [employee, setEmployee] = useState<any>({});
    const [vendor, setVendor] = useState<any>({});

    const [requestNumber, setRequestNumber] = useState("");
    const [currency, setCurrency] = useState("");
    const [requestedOn, setRequestedOn] = useState("");
    const [totalAmount, setTotalAmount] = useState("");
    const [foreignBankCharges, setForeignBankCharges] = useState("");
    const [invoiceAttachments, setInvoiceAttachments] = useState<any>({});
    const [otherAttachments, setOtherAttachments] = useState<any>({});
    const [poAttachments, setPoAttachments] = useState<any>({});
    const [piAttachments, setPiAttachments] = useState<any>({});
    const [boeAttachments, setBoeAttachments] = useState<any>({});
    const [blAttachments, setBlAttachments] = useState<any>({});

    const [otherAttachmentsadvance, setOtherAttachmentsadvance] = useState<any>({});
    const [rows, setRows] = useState<InvoiceRow[]>([
        {
            invoiceNo: "",
            invoiceDate: "",
            boeNo: "",
            boeDate: "",
            mrnNo: "",
            blNo: "",
            blDate: "",
            invoiceAmount: "",
        },
    ]);
    const [ShowInvoiceData, serShowInvoiceData] = useState<InvoiceRownew[]>([
        {
            invoiceNo: "",
            invoiceDate: "",
            boeNo: "",
            boeDate: "",
            mrnNo: "",
            blNo: "",
            blDate: "",
            invoiceAmountnew: "",
        },
    ]);
    const [workflowHistory, setWorkflowHistory] = useState<any[]>([]);

    useEffect(() => {
        if (Id) loadForexData(Id);
    }, [Id]);


    const getVendorData = async (vendorCode: string) => {

        const sp = await spCrudOps;

        const res = await sp.getData(
            "VendorMaster",
            "VendorCode,VendorName,VendorNameLegal,VendorShortName,VendorType,City/Title,State/Title,Country/Title,Currency/Title,PostalCode,ContactPersonName,EmailId,PhoneNumber,AlternateContact,BeneficiaryName,BankName,AccountNumberIBAN,SWIFTBICCode,RoutingNumberABA,IFSCCode,IntermediaryBank,IntermediarySWIFTCode,NatureOfPayment/Title,PurposeCodeRBI,BankCountry,BankAddress,VendorAddress",
            "NatureOfPayment,City,State,Country,Currency",
            `VendorCode eq '${vendorCode}'`,
            { column: "ID", isAscending: true },
            1,
            props
        );

        if (res.length > 0) {

            const v = res[0];

            setVendor({
                VendorCode: v.VendorCode,
                VendorName: v.VendorName,
                VendorAddress: v.VendorAddress,
                City: v.City.Title,
                Country: v.Country.Title,
                PostalCode: v.PostalCode,
                BankName: v.BankName,
                BankCountry: v.BankCountry,
                SWIFTBICCode: v.SWIFTBICCode,
                BankAddress: v.BankAddress,
                AccountNumberIBAN: v.AccountNumberIBAN
            });

        }

    };

    const loadForexData = async (id: string) => {
        const sp = await spCrudOps;

        const parent = await sp.getData(
            "ForexRequest",
            "*,RM/Title,HOD/Title,Currency/Title",
            "RM,HOD,Currency",
            `ID eq ${id}`,
            { column: "ID", isAscending: true },
            1,
            props
        );

        if (parent.length > 0) {
            const data = parent[0];
            if (data.WorkFlowHistory) {
                try {
                    const parsed = JSON.parse(data.WorkFlowHistory);
                    setWorkflowHistory(parsed);
                } catch {
                    setWorkflowHistory([]);
                }
            }

            setPaymentType(data.ForexType);
            setRequestNumber(data.ForexNumber);
            setRequestedOn(data.RequestedOn?.split("T")[0]);
            setCurrency(data.Currency.Title);
            setTotalAmount(data.TotalAmount);
            setForeignBankCharges(data.ForeignBankCharges);

            setEmployee({
                EmployeeCode: data.EmployeeCode || "",
                EmployeeName: data.EmployeeName || "",
                Division: data.Division || "",
                Location: data.Location || "",
                RM: data.RM?.Title || "",
                HOD: data.HOD?.Title || "",
                ContactNo: data.ContactNo || "",
                EmployeeStatus: data.EmployeeStatus || "",
                Email: data.Email || ""
            });

            // setVendor({
            //     VendorCode: data.VendorCode,
            //     VendorName: data.VendorName,
            //     VendorAddress: data.VendorAddress,
            //     City: data.City,
            //     Country: data.Country,
            //     PostalCode: data.PostalCode,
            //     BankName: data.BankName,
            //     BankCountry: data.BankCountry,
            //     SWIFTBICCode: data.BankSwiftCode,
            //     BankAddress: data.BankAddress,
            //     AccountNumberIBAN: data.BankAccNo,
            // });
            getVendorData(data.VendorCode);

            // Load child invoice rows

            const child = await sp.getData(
                "ForexServicesBillPayment",
                "*,AttachmentFiles",
                "AttachmentFiles",
                `ForexIDId eq ${id}`,
                { column: "ID", isAscending: true },
                5000,
                props
            );

            if (child.length > 0) {

                const formattedRows = child.map((item: any) => ({
                    invoiceNo: item.InvoiceNumber || "",
                    invoiceDate: item.InvoiceDate?.split("T")[0] || "",
                    invoiceAmountnew: item.InvoiceAmount || "",
                    mrnNo: item.MRNNumber || "",
                    mrnDate: item.MRNDate?.split("T")[0] || "",
                    blNo: item.BillofLandingNo || "",
                    blDate: item.BillOfLandingdate?.split("T")[0] || "",
                    boeNo: item.BOENo || "",
                    boeDate: item.BOEDate?.split("T")[0] || ""
                }));

                serShowInvoiceData(formattedRows);

                const invoiceAttachmentMap: any = {};
                const otherAttachmentMap: any = {};
                const poAttachmentMap: any = {};
                const piAttachmentMap: any = {};

                child.forEach((item: any, index: number) => {

                    const files = item.AttachmentFiles || [];

                    invoiceAttachmentMap[index] = files.filter((f: any) =>
                        f.FileName?.startsWith("INV_")
                    );

                    otherAttachmentMap[index] = files.filter((f: any) =>
                        f.FileName?.startsWith("DOC_")
                    );

                    poAttachmentMap[index] = files.filter((f: any) =>
                        f.FileName?.startsWith("PO_")
                    );

                    piAttachmentMap[index] = files.filter((f: any) =>
                        f.FileName?.startsWith("PI_")
                    );
                });

                setInvoiceAttachments(invoiceAttachmentMap);
                setOtherAttachments(otherAttachmentMap);
                setPoAttachments(poAttachmentMap);
                setPiAttachments(piAttachmentMap);
            }
        }
    };

    /* ---------- Rows ---------- */

    const addRow = () => {
        setRows([
            ...rows,
            {
                invoiceNo: "",
                invoiceDate: "",
                boeNo: "",
                boeDate: "",
                mrnNo: "",
                blNo: "",
                blDate: "",
                invoiceAmount: "",
            },
        ]);
    };

    const deleteRow = (index: number) => {
        const updated = rows.filter((_, i) => i !== index);
        setRows(updated);
    };

    const handleChange = (
        index: number,
        field: keyof InvoiceRow,
        value: any
    ) => {
        const updated = [...rows];
        updated[index][field] = value;
        setRows(updated);
    };

    const totalInvoiceAmount = rows.reduce(
        (sum, r) => sum + (parseFloat(r.invoiceAmount) || 0),
        0
    );
    const totalInvoiceAmountNew = ShowInvoiceData.reduce(
        (sum, r) => sum + (parseFloat(r.invoiceAmountnew) || 0),
        0
    );

    const uniqueBoeNumbers = Array.from(
        new Set(rows.map((r) => r.boeNo).filter(Boolean))
    );

    const uniqueBlNumbers = Array.from(
        new Set(rows.map((r) => r.blNo).filter(Boolean))
    );

    const getTreasuryPaymentApprover = async () => {

        const sp = await spCrudOps;

        const res = await sp.getData(
            "ForexApprovalMatrix",
            "*,Approver/Id,Approver/Title",
            "Approver",
            `Role eq 'TreasuryPayment' and RequestType eq 'Advance Payment' and Status eq 'Active'`,
            { column: "ID", isAscending: true },
            1,
            props
        );

        if (res.length > 0) {
            return res[0].ApproverId;
        }

        return null;
    };
const validateForm = () => {

    for (let i = 0; i < rows.length; i++) {

        const row = rows[i];

        // COMMON
        if (!row.invoiceNo) {
            alert(`Invoice No required in row ${i + 1}`);
            return false;
        }

        if (!row.invoiceDate) {
            alert(`Invoice Date required in row ${i + 1}`);
            return false;
        }

        if (!row.invoiceAmount) {
            alert(`Invoice Amount required in row ${i + 1}`);
            return false;
        }

        // ================= GOODS ADVANCE =================
        if (paymentType === "Goods-Advance Payment") {

            if (!row.boeNo) {
                alert(`BOE No required in row ${i + 1}`);
                return false;
            }

            if (!row.boeDate) {
                alert(`BOE Date required in row ${i + 1}`);
                return false;
            }

            if (!row.blNo) {
                alert(`Bill of Lading No required in row ${i + 1}`);
                return false;
            }

            if (!row.blDate) {
                alert(`Bill of Lading Date required in row ${i + 1}`);
                return false;
            }

            // Invoice file
            if (!invoiceAttachments[i] || invoiceAttachments[i].length === 0) {
                alert(`Invoice attachment required in row ${i + 1}`);
                return false;
            }
        }

        // ================= SERVICE ADVANCE =================
        if (paymentType === "Service-Advance Payment") {

            if (!row.mrnNo) {
                alert(`MRN No required in row ${i + 1}`);
                return false;
            }

            if (!row.mrnDate) {
                alert(`MRN Date required in row ${i + 1}`);
                return false;
            }

            if (!invoiceAttachments[i] || invoiceAttachments[i].length === 0) {
                alert(`Invoice attachment required in row ${i + 1}`);
                return false;
            }
        }
    }

    // ================= GOODS ADVANCE UNIQUE FILES =================
    if (paymentType === "Goods-Advance Payment") {

        for (let i = 0; i < uniqueBoeNumbers.length; i++) {
            if (!boeAttachments[i] || boeAttachments[i].length === 0) {
                alert(`BOE document required for BOE No: ${uniqueBoeNumbers[i]}`);
                return false;
            }
        }

        for (let i = 0; i < uniqueBlNumbers.length; i++) {
            if (!blAttachments[i] || blAttachments[i].length === 0) {
                alert(`BL document required for Bill of Lading: ${uniqueBlNumbers[i]}`);
                return false;
            }
        }
    }

    return true;
};
    const onSubmit = async () => {

        const sp = await spCrudOps;

        try {
            if (!validateForm()) return;

            for (let i = 0; i < rows.length; i++) {

                const row = rows[i];

                const item = await sp.insertData(
                    "ForexAdvanceBillPayment",
                    {
                        ForexIDId: Number(Id),
                        InvoiceNumber: row.invoiceNo,
                        InvoiceDate: row.invoiceDate ? row.invoiceDate : null,
                        InvoiceAmount: '' + row.invoiceAmount,
                        MRNNumber: '' + row.mrnNo,
                        MRNDate: row.mrnDate ? row.mrnDate : null,
                        BOENo: '' + row.boeNo,
                        BOEDate: row.boeDate ? row.boeDate : null,
                        BillofLandingNo: row.blNo,
                        BillOfLandingdate: row.blDate ? row.blDate : null
                    },
                    props
                );

                const itemId = item.data.ID;

                const uploadFiles = async (files: any[]) => {

                    for (const file of files) {

                        const response = await fetch(file.ServerRelativeUrl);
                        const blob = await response.blob();

                        await props.context.spHttpClient.post(
                            `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ForexAdvanceBillPayment')/items(${itemId})/AttachmentFiles/add(FileName='${file.FileName}')`,
                            SPHttpClient.configurations.v1,
                            { body: blob }
                        );

                    }

                };

                if (poAttachments[i]) await uploadFiles(poAttachments[i]);
                if (piAttachments[i]) await uploadFiles(piAttachments[i]);
                if (otherAttachmentsadvance[i]) await uploadFiles(otherAttachmentsadvance[i]);
                if (boeAttachments[i]) {

                    for (const file of boeAttachments[i]) {

                        const fileName = `BOE_${row.boeNo}_${file.name}`;

                        await props.context.spHttpClient.post(
                            `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ForexAdvanceBillPayment')/items(${itemId})/AttachmentFiles/add(FileName='${fileName}')`,
                            SPHttpClient.configurations.v1,
                            { body: file }
                        );
                    }
                }

                /* ---------- BL Attachments ---------- */

                if (blAttachments[i]) {

                    for (const file of blAttachments[i]) {

                        const fileName = `BL_${row.blNo}_${file.name}`;

                        await props.context.spHttpClient.post(
                            `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ForexAdvanceBillPayment')/items(${itemId})/AttachmentFiles/add(FileName='${fileName}')`,
                            SPHttpClient.configurations.v1,
                            { body: file }
                        );
                    }
                }

            }
            const treasuryUserId = await getTreasuryPaymentApprover();

            // 🔥 FETCH EXISTING WORKFLOW HISTORY
            const existingItem = await sp.getData(
                "ForexRequest",
                "WorkFlowHistory",
                "",
                `ID eq ${Id}`,
                { column: "ID", isAscending: true },
                1,
                props
            );

            let existingHistory: any[] = [];

            if (existingItem.length > 0 && existingItem[0].WorkFlowHistory) {
                try {
                    existingHistory = JSON.parse(existingItem[0].WorkFlowHistory);
                } catch {
                    existingHistory = [];
                }
            }
            let updatedHistory = [...existingHistory];

            updatedHistory.push({
                CurrentApprover: props.context.pageContext.user.displayName,   // or logged-in user
                Role: "",
                ActionTaken: "Tracker Submitted",
                Comment: "",
                Date: new Date().toISOString(),
                CurrentStatus: "Paid and Pending for Settlement"
            });
            await sp.updateData(
                "ForexRequest",
                Number(Id),
                {
                    Status: 'Paid and Pending for Settlement',
                    WorkFlowHistory: JSON.stringify(updatedHistory),
                    CurrentApproverId: treasuryUserId
                },
                props
            );


            alert("Tracker saved successfully");

            history.push("/");

        } catch (error) {

            console.error(error);
            alert("Error saving tracker");

        }

    };


    const handleInvoiceFile = (index: number, files: FileList | null) => {
        if (!files) return;

        const updated = { ...invoiceAttachments };
        updated[index] = Array.from(files);
        setInvoiceAttachments(updated);
    };

    const handleOtherFile = (index: number, files: FileList | null) => {
        if (!files) return;

        const updated = { ...otherAttachments };
        updated[index] = Array.from(files);
        setOtherAttachmentsadvance(updated);
    };

    const handleBoeUpload = (index: number, files: FileList | null) => {
        if (!files) return;

        const updated = { ...boeAttachments };
        updated[index] = Array.from(files);
        setBoeAttachments(updated);
    };

    const handleBlUpload = (index: number, files: FileList | null) => {
        if (!files) return;

        const updated = { ...blAttachments };
        updated[index] = Array.from(files);
        setBlAttachments(updated);
    };

    return (
        <div className="forex-wrapper">

            <div className="forex-header">
                <h2>Forex Payment - Advance Payment Tracker</h2>
            </div>

            <div className="forex-card">

                {/* Requestor */}

                <Section title="Requestor Information">
                    <Grid>
                        <Field label="Type"><input value={paymentType} readOnly /></Field>
                        <Field label="Employee Code"><input value={employee.EmployeeCode} readOnly /></Field>
                        <Field label="Employee Name"><input value={employee.EmployeeName} readOnly /></Field>
                        <Field label="Division"><input value={employee.Division} readOnly /></Field>
                        <Field label="Location"><input value={employee.Location} readOnly /></Field>
                        <Field label="RM"><input value={employee.RM} readOnly /></Field>
                        <Field label="HOD"><input value={employee.HOD} readOnly /></Field>
                        <Field label="Contact No"><input value={employee.ContactNo} readOnly /></Field>
                        <Field label="Employee Status"><input value={employee.EmployeeStatus} readOnly /></Field>
                        <Field label="Email" full><input value={employee.Email} readOnly /></Field>
                    </Grid>
                </Section>

                {/* Vendor */}

                <CollapsibleSection title="Vendor - Beneficiary Details">
                    <Grid>
                        <Field label="Vendor Code"><input value={vendor.VendorCode} readOnly /></Field>
                        <Field label="Vendor Name"><input value={vendor.VendorName} readOnly /></Field>
                        <Field label="Address" full><input value={vendor.VendorAddress} readOnly /></Field>
                        <Field label="City"><input value={vendor.City} readOnly /></Field>
                        <Field label="Country"><input value={vendor.Country} readOnly /></Field>
                        <Field label="Pincode"><input value={vendor.PostalCode} readOnly /></Field>
                        <Field label="Bank Name"><input value={vendor.BankName} readOnly /></Field>
                        <Field label="Bank Country"><input value={vendor.BankCountry} readOnly /></Field>
                        <Field label="Bank Swift Code"><input value={vendor.SWIFTBICCode} readOnly /></Field>
                        <Field label="Bank Branch"><input value={vendor.BankAddress} readOnly /></Field>
                        <Field label="Bank IBAN / Account No" full>
                            <input value={vendor.AccountNumberIBAN} readOnly />
                        </Field>
                    </Grid>
                </CollapsibleSection>

                {/* Advance Payment */}

                <Section title="Advance Payment Request Details">
                    <Grid>
                        <Field label="Request Number"><input value={requestNumber} readOnly /></Field>
                        <Field label="Currency"><input value={currency} readOnly /></Field>
                        <Field label="Total Amount"><input value={totalAmount} readOnly /></Field>
                        <Field label="Foreign Bank Charges"><input value={foreignBankCharges} readOnly /></Field>
                        <Field label="Requested On"><input type="date" value={requestedOn} readOnly /></Field>
                    </Grid>
                </Section>
                <table className="data-table" style={{ marginTop: "10px" }}>

                    <thead>
                        <tr>
                            <th>Sr.No.</th>
                            <th>Performa Invoice No</th>
                            <th>Performa Invoice Date</th>
                            <th>Performa Invoice Amount</th>
                            <th>Attach PO</th>
                            <th>Attach PI</th>
                            <th>Attach Other</th>
                        </tr>
                    </thead>

                    <tbody>

                        {ShowInvoiceData.map((row, index) => (
                            <tr key={index}>

                                <td>{index + 1}</td>

                                <td>{row.invoiceNo}</td>

                                <td>{row.invoiceDate}</td>

                                <td>{row.invoiceAmountnew}</td>

                                {/* PO Attachments */}
                                <td>
                                    {poAttachments[index]?.length > 0 ? (
                                        poAttachments[index].map((file: any, i: number) => (
                                            <div key={i}>
                                                <a
                                                    href={file.ServerRelativeUrl}
                                                    target="_blank"
                                                    rel="noopener noreferrer"
                                                    className="upload-link"
                                                >
                                                    {file.FileName}
                                                </a>
                                            </div>
                                        ))
                                    ) : (
                                        <span>-</span>
                                    )}
                                </td>

                                {/* PI Attachments */}
                                <td>
                                    {piAttachments[index]?.length > 0 ? (
                                        piAttachments[index].map((file: any, i: number) => (
                                            <div key={i}>
                                                <a
                                                    href={file.ServerRelativeUrl}
                                                    target="_blank"
                                                    rel="noopener noreferrer"
                                                    className="upload-link"
                                                >
                                                    {file.FileName}
                                                </a>
                                            </div>
                                        ))
                                    ) : (
                                        <span>-</span>
                                    )}
                                </td>

                                {/* Other Attachments */}
                                <td>
                                    {otherAttachments[index]?.length > 0 ? (
                                        otherAttachments[index].map((file: any, i: number) => (
                                            <div key={i}>
                                                <a
                                                    href={file.ServerRelativeUrl}
                                                    target="_blank"
                                                    rel="noopener noreferrer"
                                                    className="upload-link"
                                                >
                                                    {file.FileName}
                                                </a>
                                            </div>
                                        ))
                                    ) : (
                                        <span>-</span>
                                    )}
                                </td>

                            </tr>
                        ))}

                    </tbody>

                    <tfoot>
                        <tr>

                            <td colSpan={2}></td>

                            <td style={{ fontWeight: "bold" }}>Total Amount</td>

                            <td>{totalInvoiceAmountNew.toFixed(2)}</td>

                            <td colSpan={3}></td>

                        </tr>
                    </tfoot>

                </table>

                <CollapsibleSection title="Workflow History">

                    {workflowHistory.length > 0 ? (
                        <table className="data-table">
                            <thead>
                                <tr>
                                    <th>Action By</th>
                                    {/* <th>Role</th> */}
                                    <th>Action</th>
                                    <th>Remark</th>
                                    <th>Date</th>
                                    <th>Status</th>
                                </tr>
                            </thead>

                            <tbody>
                                {workflowHistory.map((item: any, index: number) => (
                                    <tr key={index}>
                                        <td>{item.CurrentApprover}</td>
                                        {/* <td>{item.Role || "-"}</td> */}
                                        <td>{item.ActionTaken}</td>
                                        <td>{item.Comment || "-"}</td>
                                        <td>
                                            {item.Date
                                                ? new Date(item.Date).toLocaleString("en-GB")
                                                : ""}
                                        </td>
                                        <td>{item.CurrentStatus}</td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    ) : (
                        <p>No workflow history available</p>
                    )}

                </CollapsibleSection>

                {paymentType === "Goods-Advance Payment" && (
                    <>
                        <Section title="Bill Payment Details (For Goods Bill Payment)">

                            {/* <p style={{ color: "red", fontSize: "12px" }}>
                                User will enter manually, multiple invoice can be entered
                            </p> */}

                            <table className="data-table">

                                <thead>
                                    <tr>
                                        <th>Sr.No. </th>
                                        <th>Invoice Number <span className="required" style={{ color: "red" }}>*</span></th>
                                        <th>Invoice Date <span className="required" style={{ color: "red" }}>*</span></th>
                                        <th>BOE Number <span className="required" style={{ color: "red" }}>*</span></th>
                                        <th>BOE Date <span className="required" style={{ color: "red" }}>*</span></th>
                                        <th>MRN Number <span className="required" style={{ color: "red" }}>*</span></th>
                                        <th>Bill of Lading Number <span className="required" style={{ color: "red" }}>*</span></th>
                                        <th>Bill of Lading Date <span className="required" style={{ color: "red" }}>*</span></th>
                                        <th>Invoice Amount <span className="required" style={{ color: "red" }}>*</span></th>
                                        <th>Attach Invoice <span className="required" style={{ color: "red" }}>*</span></th>
                                        <th>Attach Other </th>
                                        <th>Add/Delete entry</th>
                                    </tr>
                                </thead>

                                <tbody>

                                    {rows.map((row, index) => (
                                        <tr key={index}>

                                            <td>{index + 1}</td>

                                            <td>
                                                <input
                                                    value={row.invoiceNo}
                                                    onChange={(e) => handleChange(index, "invoiceNo", e.target.value)}
                                                />
                                            </td>

                                            <td>
                                                <input
                                                    type="date"
                                                    value={row.invoiceDate}
                                                    onChange={(e) => handleChange(index, "invoiceDate", e.target.value)}
                                                />
                                            </td>

                                            <td>
                                                <input
                                                    value={row.boeNo}
                                                    onChange={(e) => handleChange(index, "boeNo", e.target.value)}
                                                />
                                            </td>

                                            <td>
                                                <input
                                                    type="date"
                                                    value={row.boeDate}
                                                    onChange={(e) => handleChange(index, "boeDate", e.target.value)}
                                                />
                                            </td>

                                            <td>
                                                <input
                                                    value={row.mrnNo}
                                                    onChange={(e) => handleChange(index, "mrnNo", e.target.value)}
                                                />
                                            </td>

                                            <td>
                                                <input
                                                    value={row.blNo}
                                                    onChange={(e) => handleChange(index, "blNo", e.target.value)}
                                                />
                                            </td>

                                            <td>
                                                <input
                                                    type="date"
                                                    value={row.blDate}
                                                    onChange={(e) => handleChange(index, "blDate", e.target.value)}
                                                />
                                            </td>

                                            <td>
                                                <input
                                                    type="number"
                                                    value={row.invoiceAmount}
                                                    onChange={(e) => handleChange(index, "invoiceAmount", e.target.value)}
                                                />
                                            </td>

                                            {/* Invoice Upload */}
                                            <td>
                                                <input
                                                    type="file"
                                                    onChange={(e) => handleInvoiceFile(index, e.target.files)}
                                                />
                                            </td>

                                            {/* Other Upload */}
                                            <td>
                                                <input
                                                    type="file"
                                                    onChange={(e) => handleOtherFile(index, e.target.files)}
                                                />
                                            </td>

                                            <td style={{ textAlign: "center" }}>

                                                {index === rows.length - 1 && (
                                                    <span style={{ cursor: "pointer" }} onClick={addRow}>+</span>
                                                )}

                                                {rows.length > 1 && (
                                                    <span
                                                        style={{ cursor: "pointer", marginLeft: "8px" }}
                                                        onClick={() => deleteRow(index)}
                                                    >✖</span>
                                                )}

                                            </td>

                                        </tr>
                                    ))}

                                </tbody>

                                <tfoot>
                                    <tr>
                                        <td colSpan={7}></td>
                                        <td style={{ fontWeight: "bold" }}>Total Amount</td>
                                        <td>{totalInvoiceAmount.toFixed(2)}</td>
                                        <td colSpan={3}></td>
                                    </tr>
                                </tfoot>

                            </table>

                        </Section>



                        <div style={{ display: "flex", gap: "40px", marginTop: "20px" }}>

                            {/* BOE TABLE */}

                            <div>

                                <p style={{ color: "red", fontSize: "12px" }}>
                                    Unique BOE no will be listed below
                                </p>

                                <table className="data-table">

                                    <thead>
                                        <tr>
                                            <th>BOE No</th>
                                            <th>Attach Documents</th>
                                        </tr>
                                    </thead>

                                    <tbody>

                                        {uniqueBoeNumbers.map((boe, index) => (

                                            <tr key={index}>

                                                <td>{boe}</td>

                                                <td>

                                                    <input
                                                        type="file"
                                                        onChange={(e) => handleBoeUpload(index, e.target.files)}
                                                    />

                                                    {boeAttachments[index]?.map((file: any, i: number) => (
                                                        <div key={i} style={{ fontSize: "12px" }}>
                                                            {file.name}
                                                        </div>
                                                    ))}

                                                </td>

                                            </tr>

                                        ))}

                                    </tbody>

                                </table>

                            </div>


                            {/* BILL OF LADING TABLE */}

                            <div>

                                <p style={{ color: "red", fontSize: "12px" }}>
                                    Unique Bill of lading no will be listed below
                                </p>

                                <table className="data-table">

                                    <thead>
                                        <tr>
                                            <th>Bill of Lading</th>
                                            <th>Attach Documents</th>
                                        </tr>
                                    </thead>

                                    <tbody>

                                        {uniqueBlNumbers.map((bl, index) => (

                                            <tr key={index}>

                                                <td>{bl}</td>

                                                <td>

                                                    <input
                                                        type="file"
                                                        onChange={(e) => handleBlUpload(index, e.target.files)}
                                                    />

                                                    {blAttachments[index]?.map((file: any, i: number) => (
                                                        <div key={i} style={{ fontSize: "12px" }}>
                                                            {file.name}
                                                        </div>
                                                    ))}

                                                </td>

                                            </tr>

                                        ))}

                                    </tbody>

                                </table>

                            </div>

                        </div>
                    </>
                )}



                {paymentType === "Service-Advance Payment" && (

                    <div style={{ marginTop: "30px" }}>

                        {/* <div style={{ background: "#fff8b3", padding: "4px", fontSize: "12px" }}>
                            <b>Only for Service: Advance Payment from row 65-72</b>
                        </div> */}

                        <p>
                            <b>Bill Payment Details (for Service Bill Payment)</b>
                            {/* <span style={{ color: "red" }}>
                                (User will enter manually, multiple invoice can be entered)
                            </span> */}
                        </p>

                        <table className="data-table">

                            <thead>
                                <tr>
                                    <th>Sr.No.</th>
                                    <th>Invoice Number <span className="required" style={{ color: "red" }}>*</span></th>
                                    <th>Invoice Date <span className="required" style={{ color: "red" }}>*</span></th>
                                    <th>Invoice Amount <span className="required" style={{ color: "red" }}>*</span></th>
                                    <th>MRN Number <span className="required" style={{ color: "red" }}>*</span></th>
                                    <th>MRN Date <span className="required" style={{ color: "red" }}>*</span></th>
                                    <th>Attach Invoice <span className="required" style={{ color: "red" }}>*</span></th>
                                    <th>Attach Other Document</th>
                                    <th>Add/Delete entry</th>
                                </tr>
                            </thead>

                            <tbody>

                                {rows.map((row, index) => (
                                    <tr key={index}>

                                        <td>{index + 1}</td>

                                        <td>
                                            <input
                                                value={row.invoiceNo}
                                                onChange={(e) => handleChange(index, "invoiceNo", e.target.value)}
                                            />
                                        </td>

                                        <td>
                                            <input
                                                type="date"
                                                value={row.invoiceDate}
                                                onChange={(e) => handleChange(index, "invoiceDate", e.target.value)}
                                            />
                                        </td>

                                        <td>
                                            <input
                                                value={row.invoiceAmount}
                                                onChange={(e) => handleChange(index, "invoiceAmount", e.target.value)}
                                            />
                                        </td>

                                        <td>
                                            <input
                                                value={row.mrnNo}
                                                onChange={(e) => handleChange(index, "mrnNo", e.target.value)}
                                            />
                                        </td>

                                        <td>
                                            <input
                                                type="date"
                                                value={row.mrnDate || ""}
                                                onChange={(e) => handleChange(index, "mrnDate", e.target.value)}
                                            />
                                        </td>

                                        <td>
                                            <input
                                                type="file"
                                                onChange={(e) => handleInvoiceFile(index, e.target.files)}
                                            />
                                        </td>

                                        <td>
                                            <input
                                                type="file"
                                                onChange={(e) => handleOtherFile(index, e.target.files)}
                                            />
                                        </td>

                                        <td>

                                            {rows.length > 1 && (
                                                <span onClick={() => deleteRow(index)}>✖</span>
                                            )}

                                            {index === rows.length - 1 && (
                                                <span style={{ marginLeft: "8px" }} onClick={addRow}>+</span>
                                            )}

                                        </td>

                                    </tr>
                                ))}

                            </tbody>
                            <tfoot>
                                    <tr>
                                        <td colSpan={2}></td>
                                        <td style={{ fontWeight: "bold" }}>Total Amount</td>
                                        <td>{totalInvoiceAmount.toFixed(2)}</td>
                                        <td colSpan={5}></td>
                                    </tr>
                                </tfoot>
                        </table>

                    </div>

                )}

                {/* Approver Remarks */}

                {/* <div style={{ marginTop: "20px" }}>

                    <label><b>Approver's Remarks:</b></label>

                    <textarea
                        style={{ width: "70%", height: "40px", marginLeft: "10px" }}
                    />

                </div> */}

                <div className="button-row">
                    <button className="btn-submit" onClick={onSubmit}>Submit</button>
                    <button className="btn-exit" onClick={() => history.push("/")}>Exit</button>
                </div>

            </div>
        </div>
    );
};

export default TrackerForm;