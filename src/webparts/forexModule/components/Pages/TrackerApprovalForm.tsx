import React, { useEffect, useState } from "react";
import "../Pages/Css/NewRequest.scss";
import { IForexModuleProps } from "../IForexModuleProps";
import { useHistory, useParams } from "react-router-dom";
import SPCRUDOPS from "../../service/BAL/spcrud";
import { SPHttpClient } from "@microsoft/sp-http";

import logo from "../../assets/sona-comstarlogo.png";
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

const TrackerApprovalForm = (props: IForexModuleProps) => {
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
    const [poAttachmentsAdvance, setPoAttachmentsAddvance] = useState<any>({});
    const [piAttachmentsAddvance, setPiAttachmentsAddvance] = useState<any>({});
    const [isClosedWithAD, setIsClosedWithAD] = useState(false);
    const [referenceNumber, setReferenceNumber] = useState("");

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
    const [approverRemark, setApproverRemark] = useState("");
    const [poContractNo, setPoContractNo] = useState("");
    const [poDate, setPoDate] = useState("");
    const [expectedSettlementDate, setExpectedSettlementDate] = useState("");

    useEffect(() => {
        if (Id) loadForexData(Id);
    }, [Id]);


    const getVendorData = async (vendorCode: string) => {

        const sp = await spCrudOps;

        const res = await sp.getData(
            "VendorMaster",
            "Pincode,VendorCode,VendorName,VendorNameLegal,VendorShortName,VendorType,City/Title,City/City,State/Title,Country/Country,Currency/Title,PostalCode,ContactPersonName,EmailId,PhoneNumber,AlternateContact,BeneficiaryName,BankName,AccountNumberIBAN,SWIFTBICCode,RoutingNumberABA,IFSCCode,IntermediaryBank,IntermediarySWIFTCode,NatureOfPayment/Title,PurposeCodeRBI,BankCountry,BankAddress,VendorAddress,BalanceEligibleAmount,ApprovedAmountPaidAmount,EligibleAmountWithoutWHT,TaxDocumentAvailable,DTAAApplicable",
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
                City: v.City.City,
                Country: v.Country.Country,
                PostalCode: v.PostalCode,
                BankName: v.BankName,
                BankCountry: v.BankCountry,
                SWIFTBICCode: v.SWIFTBICCode,
                BankAddress: v.BankAddress,
                AccountNumberIBAN: v.AccountNumberIBAN,
                Pincode: v.Pincode,
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
            setPoContractNo(data.poContractNo || "");
            setPoDate(data.poDate?.split("T")[0] || "");
            setExpectedSettlementDate(data.expectedSettlementDate?.split("T")[0] || "");
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
            await loadTrackerData(Number(id));

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
    const totalInvoiceAmountnew = ShowInvoiceData.reduce(
        (sum, r) => sum + (parseFloat(r.invoiceAmountnew || "0") || 0),
        0
    );

    const uniqueBoeNumbers = Array.from(
        new Set(rows.map((r) => r.boeNo).filter(Boolean))
    );

    const uniqueBlNumbers = Array.from(
        new Set(rows.map((r) => r.blNo).filter(Boolean))
    );


    const loadTrackerData = async (forexId: number) => {

        const sp = await spCrudOps;

        const tracker = await sp.getData(
            "ForexAdvanceBillPayment",
            "*,AttachmentFiles",
            "AttachmentFiles",
            `ForexIDId eq ${forexId}`,
            { column: "ID", isAscending: true },
            5000,
            props
        );

        if (tracker.length > 0) {

            const formattedRows = tracker.map((item: any) => ({
                invoiceNo: item.InvoiceNumber || "",
                invoiceDate: item.InvoiceDate?.split("T")[0] || "",
                invoiceAmount: item.InvoiceAmount || "",
                mrnNo: item.MRNNumber || "",
                mrnDate: item.MRNDate?.split("T")[0] || "",
                blNo: item.BillofLandingNo || "",
                blDate: item.BillOfLandingdate?.split("T")[0] || "",
                boeNo: item.BOENo || "",
                boeDate: item.BOEDate?.split("T")[0] || ""
            }));

            setRows(formattedRows);

            const attachmentMap: any = {};
            const poAttachmentMap: any = {};
            const piAttachmentMap: any = {};
            const otherAttachmentMap: any = {};
            const boeAttachmentMap: any = {};
            const blAttachmentMap: any = {};

            tracker.forEach((item: any, index: number) => {

                const files = item.AttachmentFiles || [];

                poAttachmentMap[index] = files.filter((f: any) =>
                    f.FileName?.startsWith("INV_")
                );

                piAttachmentMap[index] = files.filter((f: any) =>
                    f.FileName?.startsWith("DOC_")
                );

                otherAttachmentMap[index] = files.filter((f: any) =>
                    f.FileName?.startsWith("DOC_")
                );

                boeAttachmentMap[index] = files.filter((f: any) =>
                    f.FileName?.startsWith("BOE_")
                );

                blAttachmentMap[index] = files.filter((f: any) =>
                    f.FileName?.startsWith("BL_")
                );

            });

            setPoAttachmentsAddvance(poAttachmentMap);
            setPiAttachmentsAddvance(piAttachmentMap);
            setOtherAttachments(otherAttachmentMap);
            setBoeAttachments(boeAttachmentMap);
            setBlAttachments(blAttachmentMap);
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
        setOtherAttachments(updated);
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
    const rejectRequest = async () => {

        const sp = await spCrudOps;

        if (!approverRemark || approverRemark.trim() === "") {
            alert("Please enter remark before rejecting");
            return;
        }

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

        const user = props.context.pageContext.user;

        let updatedHistory = [...existingHistory];

        updatedHistory.push({
            CurrentApprover: user.displayName,
            Role: "TreasuryApproval",
            ActionTaken: "Rejected",
            Comment: approverRemark,
            Date: new Date().toISOString(),
            CurrentStatus: "Rejected"
        });

        await sp.updateData(
            "ForexRequest",
            Number(Id),
            {
                Status: "Rejected",
                ApproverRemark: approverRemark,
                WorkFlowHistory: JSON.stringify(updatedHistory) // 🔥 IMPORTANT
            },
            props
        );

        alert("Request Rejected");

        history.push("/");
    };
    const approveRequest = async () => {

        const sp = await spCrudOps;

        if (!approverRemark || approverRemark.trim() === "") {
            alert("Please enter remark before approving");
            return;
        }

        // 🔥 Fetch latest history
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

        const user = props.context.pageContext.user;

        let updatedHistory = [...existingHistory];

        updatedHistory.push({
            CurrentApprover: user.displayName,
            Role: "TreasuryApproval",
            ActionTaken: "Approved",
            Comment: approverRemark,
            Date: new Date().toISOString(),
            CurrentStatus: "Paid & Closed"
        });

        await sp.updateData(
            "ForexRequest",
            Number(Id),
            {
                Status: "Paid & Closed",
                TreasuryApproverRemark: approverRemark,
                ClosedWithADBank: isClosedWithAD ? "Yes" : "No",
                ReferenceNumber: referenceNumber,
                WorkFlowHistory: JSON.stringify(updatedHistory), // 🔥 IMPORTANT
                CurrentApproverId: null
            },
            props
        );

        alert("Request Approved");

        history.push("/");
    };

    return (
        <>

            <div className='MainUplodForm' style={{ margin: "5px 0px" }}>
                <div className='row'>
                    <div className='col-md-12'>
                        <div className='Main-Boxpoup'>
                            <div className="bordered">
                                <a><img src={logo} /></a>
                                <h1>Forex Payment - Advance Payment Tracker Approval Form</h1>
                            </div>
                            <div className='borderedbox'>
                                {/* <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Requestor Information</label>
                                </div> */}
                                <CollapsibleSection title="Requestor Information" style={{ marginTop: "10px" }}>

                                    <div className='main-formcontainer'>

                                        <div className='row mb-20'>
                                            <div className='col-md-4'>
                                                <label className='font'>Employee Code</label>
                                                <input type="text" value={employee.EmployeeCode} className="form-control readonly" />
                                            </div>
                                            <div className='col-md-4'>
                                                <label className="font">Employee Name</label>
                                                <input type="text" value={employee.EmployeeName} className="form-control readonly" />
                                            </div>
                                            <div className="col-md-4">
                                                <label className="font">Division</label>
                                                <input type="text" value={employee.Division} className="form-control readonly" />
                                            </div>
                                        </div>
                                        <div className='row mb-20'>
                                            <div className='col-md-4'>
                                                <label className='font'>Location</label>
                                                <input type="text" value={employee.Location} className="form-control readonly" />
                                            </div>
                                            <div className='col-md-4'>
                                                <label className="font">RM</label>
                                                <input type="text" value={employee.RM} className="form-control readonly" />
                                            </div>
                                            <div className="col-md-4">
                                                <label className="font">HOD</label>
                                                <input type="text" value={employee.HOD} className="form-control readonly" />
                                            </div>
                                        </div>
                                        <div className='row mb-20'>
                                            <div className='col-md-4'>
                                                <label className='font'>Contact No</label>
                                                <input type="text" value={employee.ContactNo} className="form-control readonly" />
                                            </div>
                                            <div className='col-md-4'>
                                                <label className="font">Employee Status</label>
                                                <input type="text" value={employee.EmployeeStatus} className="form-control readonly" />
                                            </div>
                                            <div className="col-md-4">
                                                <label className="font">Email</label>
                                                <input type="text" value={employee.Email} className="form-control readonly" />
                                            </div>
                                        </div>
                                    </div>
                                </CollapsibleSection>
                                <CollapsibleSection title="Vendor / Beneficiary Details" style={{ marginTop: "10px" }}>
                                    <div className="heading1" style={{ marginTop: "10px" }}>
                                        <label>Vendor / Beneficiary Details</label>
                                    </div>
                                    <div className='main-formcontainer'>
                                        <div className='row mb-20'>
                                            <div className='col-md-4'>
                                                <label className='font'>Vendor Code</label>
                                                <input value={vendor.VendorCode} onChange={(e) => {
                                                    const code = e.target.value;
                                                    setVendor({ ...vendor, VendorCode: code });
                                                }}
                                                    onBlur={(e) => getVendorData(e.target.value)}
                                                    className="form-control readonly"
                                                />
                                            </div>
                                            <div className='col-md-4'>
                                                <label className="font">Vendor Name</label>
                                                <input type="text" value={vendor.VendorName} className="form-control readonly" />
                                            </div>
                                            <div className="col-md-4">
                                                <label className="font">City</label>
                                                <input type="text" value={vendor.City} className="form-control readonly" />
                                            </div>
                                        </div>
                                        <div className='row mb-20'>
                                            <div className='col-md-4'>
                                                <label className="font">Country</label>
                                                <input type="text" value={vendor.Country} className="form-control readonly" />
                                            </div>
                                            <div className='col-md-4'>
                                                <label className="font">Pincode</label>
                                                <input type="text" value={vendor.Pincode} className="form-control readonly" />
                                            </div>
                                            <div className='col-md-4'>
                                                <label className="font">Bank Name</label>
                                                <input type="text" value={vendor.BankName} className="form-control readonly" />
                                            </div>
                                        </div>
                                        <div className='row mb-20'>
                                            <div className='col-md-4'>
                                                <label className="font">Bank Country</label>
                                                <input type="text" value={vendor.BankCountry} className="form-control readonly" />
                                            </div>
                                            <div className='col-md-4'>
                                                <label className="font">Bank Swift Code</label>
                                                <input type="text" value={vendor.SWIFTBICCode} className="form-control readonly" />
                                            </div>
                                            <div className='col-md-4'>
                                                <label className="font">Bank Branch Address</label>
                                                <input type="text" value={vendor.BankAddress} className="form-control readonly" />
                                            </div>
                                        </div>
                                        <div className='row mb-20'>
                                            <div className='col-md-4'>
                                                <label className="font">Bank IBAN / Account No</label>
                                                <input type="text" value={vendor.AccountNumberIBAN} className="form-control readonly" />
                                            </div>
                                            <div className='col-md-4'>
                                                <label className="font">Address</label>
                                                <input type="text" value={vendor.VendorAddress} className="form-control textbox readonly" />
                                            </div>

                                        </div>
                                    </div>
                                </CollapsibleSection>
                                <CollapsibleSection title="Workflow History" style={{ marginTop: "10px" }}>

                                    {workflowHistory.length > 0 ? (
                                        <table className="custom-table">
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

                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Advance Payment Request Details</label>
                                </div>
                                <div className='main-formcontainer'>
                                    <div className='row mb-20'>
                                        <div className='col-md-4'>
                                            <label className='font'>Type</label>
                                            <input type="text" value={paymentType} className="form-control readonly" />
                                        </div>
                                    </div>
                                    <div className='row mb-20'>
                                       <div className='col-md-4'>
                                            <label className='font'>Request Number</label>
                                            <input type="text" value={requestNumber} className="form-control readonly" />
                                        </div>
                                        <div className='col-md-4'>
                                            <label className="font">Requested On</label>
                                            <input type="text" value={requestedOn} className="form-control readonly" />
                                        </div>
                                        <div className="col-md-4">
                                            <label className="font">Currency</label>
                                            <input type="text" value={currency} className="form-control readonly" />
                                        </div>
                                    </div>
                                    <div className='row mb-20'>
                                        <div className='col-md-4'>
                                            <label className='font'>Total Amount</label>
                                            <input type="text" value={totalAmount} className="form-control readonly" />
                                        </div>
                                        <div className="col-md-4">
                                            <label className="font">Foreign Bank Charges </label>
                                            <input type="text" value={foreignBankCharges} className="form-control readonly" />
                                        </div>
                                        <div className="col-md-4">
                                            <label className="font">PO/Contract No </label>
                                            <input type="text" value={poContractNo} className="form-control readonly" />
                                        </div>
                                    </div>
                                     <div className='row mb-20'>
                                        <div className='col-md-4'>
                                            <label className='font'>PO Date</label>
                                            <input type="date" value={poDate} className="form-control readonly" />
                                        </div>
                                        <div className="col-md-4">
                                            <label className="font">Expected Settlement Date </label>
                                            <input type="date" value={expectedSettlementDate} className="form-control readonly" />
                                        </div>

                                    </div>
                                </div>

                                <div style={{ overflowX: "auto" }}>
                                    <table className="custom-table" style={{ marginTop: "10px" }}>

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

                                                <td>{totalInvoiceAmountnew.toFixed(2)}</td>

                                                <td colSpan={3}></td>

                                            </tr>
                                        </tfoot>

                                    </table>
                                </div>

                                {paymentType === "Goods-Advance Payment" && (
                                    <>
                                        <div className="heading1" style={{ marginTop: "10px" }}>
                                            <label>Bill Payment Details (For Goods Bill Payment)</label>
                                        </div>
                                        <div style={{ overflowX: "auto" }}>
                                            <table className="custom-table">
                                                <thead>
                                                    <tr>
                                                        <th>Sr.No.</th>
                                                        <th>Invoice Number</th>
                                                        <th>Invoice Date</th>
                                                        <th>BOE Number</th>
                                                        <th>BOE Date</th>
                                                        <th>MRN Number</th>
                                                        <th>Bill of Lading Number</th>
                                                        <th>Bill of Lading Date</th>
                                                        <th>Invoice Amount</th>
                                                        <th>Attach Invoice</th>
                                                        <th>Attach Other</th>
                                                    </tr>
                                                </thead>

                                                <tbody>

                                                    {rows.map((row, index) => (
                                                        <tr key={index}>

                                                            <td>{index + 1}</td>

                                                            <td><span>{row.invoiceNo}</span></td>

                                                            <td><span>{row.invoiceDate}</span></td>

                                                            <td><span>{row.boeNo}</span></td>

                                                            <td><span>{row.boeDate}</span></td>

                                                            <td><span>{row.mrnNo}</span></td>

                                                            <td><span>{row.blNo}</span></td>

                                                            <td><span>{row.blDate}</span></td>

                                                            <td><span>{row.invoiceAmount}</span></td>

                                                            {/* Invoice Upload */}
                                                            <td>

                                                                {poAttachmentsAdvance[index]?.length > 0 ? (

                                                                    poAttachmentsAdvance[index].map((file: any, i: number) => (

                                                                        <div key={i}>

                                                                            <a
                                                                                href={file.ServerRelativeUrl}
                                                                                download
                                                                                target="_blank"
                                                                                rel="noopener noreferrer"
                                                                                className="upload-link"
                                                                            >
                                                                                📥 {file.FileName || file.name}
                                                                            </a>

                                                                        </div>

                                                                    ))

                                                                ) : (

                                                                    <span>-</span>

                                                                )}

                                                            </td>

                                                            {/* Other Upload */}
                                                            <td>

                                                                {piAttachmentsAddvance[index]?.length > 0 ? (

                                                                    piAttachmentsAddvance[index].map((file: any, i: number) => (

                                                                        <div key={i}>

                                                                            <a
                                                                                href={file.ServerRelativeUrl}
                                                                                download
                                                                                target="_blank"
                                                                                rel="noopener noreferrer"
                                                                                className="upload-link"
                                                                            >
                                                                                📥 {file.FileName || file.name}
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
                                                        <td colSpan={7}></td>
                                                        <td style={{ fontWeight: "bold" }}>Total Amount</td>
                                                        <td>{totalInvoiceAmount.toFixed(2)}</td>
                                                        <td colSpan={2}></td>
                                                    </tr>
                                                </tfoot>

                                            </table>
                                        </div>
                                        <div style={{ display: "flex", gap: "40px", marginTop: "20px" }}>

                                            {/* BOE TABLE */}

                                            <div>

                                                <p style={{ color: "red", fontSize: "12px" }}>
                                                    Unique BOE no listed below
                                                </p>

                                                <table className="custom-table">

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

                                                                    {boeAttachments[index]?.length > 0 ? (

                                                                        boeAttachments[index].map((file: any, i: number) => (
                                                                            <div key={i} style={{ fontSize: "12px" }}>

                                                                                <a
                                                                                    href={file.ServerRelativeUrl}
                                                                                    target="_blank"
                                                                                    download
                                                                                    rel="noopener noreferrer"
                                                                                    className="upload-link"
                                                                                >
                                                                                    📄 {file.FileName}
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

                                                </table>

                                            </div>


                                            {/* BL TABLE */}

                                            <div>

                                                <p style={{ color: "red", fontSize: "12px" }}>
                                                    Unique Bill of Lading listed below
                                                </p>

                                                <table className="custom-table">

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

                                                                    {/* <input
                                                        type="file"
                                                        onChange={(e) => handleBlUpload(index, e.target.files)}
                                                    />

                                                    {blAttachments[index]?.map((file: any, i: number) => (
                                                        <div key={i} style={{ fontSize: "12px" }}>
                                                            {file.name}
                                                        </div>
                                                    ))} */}
                                                                    {blAttachments[index]?.length > 0 ? (

                                                                        blAttachments[index].map((file: any, i: number) => (
                                                                            <div key={i} style={{ fontSize: "12px" }}>

                                                                                <a
                                                                                    href={file.ServerRelativeUrl}
                                                                                    target="_blank"
                                                                                    download
                                                                                    rel="noopener noreferrer"
                                                                                    className="upload-link"
                                                                                >
                                                                                    📄 {file.FileName}
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

                                                </table>

                                            </div>

                                        </div>
                                    </>
                                )}

                                {paymentType === "Service-Advance Payment" && (
                                    <>
                                        <p>
                                            <b>Bill Payment Details (for Service Bill Payment)</b>
                                        </p>
                                        <div style={{ overflowX: "auto" }}>
                                            <table className="custom-table">

                                                <thead>
                                                    <tr>
                                                        <th>Sr.No.</th>
                                                        <th>Invoice Number</th>
                                                        <th>Invoice Date</th>
                                                        <th>Invoice Amount</th>
                                                        <th>MRN Number</th>
                                                        <th>MRN Date</th>
                                                        <th>Invoice Document</th>
                                                        <th>Other Document</th>
                                                    </tr>
                                                </thead>

                                                <tbody>

                                                    {rows.map((row, index) => (
                                                        <tr key={index}>

                                                            <td>{index + 1}</td>

                                                            <td>
                                                                <span>{row.invoiceNo}</span>
                                                            </td>

                                                            <td>
                                                                <span>{row.invoiceDate}</span>
                                                            </td>

                                                            <td>
                                                                <span>{row.invoiceAmount}</span>
                                                            </td>

                                                            <td>
                                                                <span>{row.mrnNo}</span>
                                                            </td>

                                                            <td>
                                                                <span>{row.mrnDate}</span>
                                                            </td>

                                                            {/* Invoice Attachments */}

                                                            <td>

                                                                {poAttachmentsAdvance[index]?.length > 0 ? (

                                                                    poAttachmentsAdvance[index].map((file: any, i: number) => (

                                                                        <div key={i}>

                                                                            <a
                                                                                href={file.ServerRelativeUrl}
                                                                                download
                                                                                target="_blank"
                                                                                rel="noopener noreferrer"
                                                                                className="upload-link"
                                                                            >
                                                                                📥 {file.FileName || file.name}
                                                                            </a>

                                                                        </div>

                                                                    ))

                                                                ) : (

                                                                    <span>-</span>

                                                                )}

                                                            </td>

                                                            {/* Other Attachments */}

                                                            <td>

                                                                {piAttachmentsAddvance[index]?.length > 0 ? (

                                                                    piAttachmentsAddvance[index].map((file: any, i: number) => (

                                                                        <div key={i}>

                                                                            <a
                                                                                href={file.ServerRelativeUrl}
                                                                                download
                                                                                target="_blank"
                                                                                rel="noopener noreferrer"
                                                                                className="upload-link"
                                                                            >
                                                                                📥 {file.FileName || file.name}
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
                                                        <td>{totalInvoiceAmount.toFixed(2)}</td>
                                                        <td colSpan={5}></td>
                                                    </tr>
                                                </tfoot>
                                            </table>
                                        </div>
                                    </>
                                )}





                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Bank Closure Details</label>
                                </div>
                                <div className='main-formcontainer'>

                                    <div className='row mb-20'>
                                        <div className='col-md-4'>
                                            <label className="font">Closed with AD bank</label>
                                            <input
                                                type="checkbox"
                                                checked={isClosedWithAD}
                                                onChange={(e) => {
                                                    setIsClosedWithAD(e.target.checked);

                                                    // Clear reference if unchecked
                                                    if (!e.target.checked) {
                                                        setReferenceNumber("");
                                                    }
                                                }}
                                            />
                                        </div>
                                        <div className='col-md-4'>
                                            <label className="font">Reference Number (on selection of checkbox)</label>
                                            <input
                                                type="text"
                                                value={referenceNumber}
                                                disabled={!isClosedWithAD}
                                                onChange={(e) => setReferenceNumber(e.target.value)}
                                                className="form-control"
                                            />
                                        </div>
                                        {/* <div className='col-md-4'>
                                            <label className="font">Total Amount</label>
                                            <input type="text" value={totalAmount} className="form-control readonly" />
                                        </div> */}
                                    </div>
                                    {/* <div className='row mb-20'>
                                        <div className='col-md-4'>
                                            <label className="font">Foreign Bank Charges</label>
                                            <input type="text" value={foreignBankCharges} className="form-control readonly" />
                                        </div>
                                        <div className='col-md-4'>
                                            <label className="font">Requested On</label>
                                            <input type="text" value={requestedOn} className="form-control readonly" />
                                        </div>
                                    </div> */}
                                </div>
                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Remarks Section</label>
                                </div>
                                <div className='main-formcontainer'>
                                    <div className='col-md-4'>
                                        <label><b>Remarks:</b></label>
                                        <textarea
                                            style={{ width: "70%", height: "40px", marginLeft: "10px" }}
                                            value={approverRemark}
                                            onChange={(e) => setApproverRemark(e.target.value)}
                                        />

                                    </div>
                                </div>

                                <div className='row my-3'>
                                    <div className='col-md-12'>
                                        <div style={{ display: "flex", justifyContent: "center", gap: "5px" }}>
                                            <button onClick={approveRequest} className="Submit-btn">
                                                Approve
                                            </button>
                                            <button onClick={rejectRequest} className="Reject-btn">
                                                Reject
                                            </button>

                                            <button onClick={() => history.push("/")} className="Exit-btn">
                                                Exit
                                            </button>
                                        </div>
                                    </div>
                                </div>


                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </>
    );
};

export default TrackerApprovalForm;