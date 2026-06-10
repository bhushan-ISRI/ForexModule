import React, { useEffect, useState } from "react";
import "../Pages/Css/NewRequest.scss";
import { IForexModuleProps } from "../IForexModuleProps";
import { useHistory, useParams } from "react-router-dom";
import SPCRUDOPS from "../../service/BAL/spcrud";
import { SPHttpClient } from "@microsoft/sp-http";

import logo from "../../assets/sona-comstarlogo.png";
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
    const [invoiceAttachmentsadvance, setInvoiceAttachmentsadvance] = useState<any>({});
    const [boeAttachmentsadvance, setBoeAttachmentsadvance] = useState<any>({});
    const [blAttachmentsadvance, setBlAttachmentsadvance] = useState<any>({});

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
                Pincode: v.Pincode
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

            const trackerData = await sp.getData(
                "ForexAdvanceBillPayment",
                "*,AttachmentFiles",
                "AttachmentFiles",
                `ForexIDId eq ${id}`,
                { column: "ID", isAscending: true },
                5000,
                props
            );

            if (trackerData.length > 0) {

                const trackerInvoiceAttachments: any = {};
                const trackerOtherAttachments: any = {};
                const trackerBoeAttachments: any = {};
                const trackerBlAttachments: any = {};

                const trackerRows = trackerData.map((item: any, index: number) => {

                    const files = item.AttachmentFiles || [];

                    trackerInvoiceAttachments[index] = files.filter((f: any) =>
                        f.FileName?.startsWith("INV_")
                    );

                    trackerOtherAttachments[index] = files.filter((f: any) =>
                        f.FileName?.startsWith("DOC_")
                    );

                    trackerBoeAttachments[index] = files.filter((f: any) =>
                        f.FileName?.startsWith("BOE_")
                    );

                    trackerBlAttachments[index] = files.filter((f: any) =>
                        f.FileName?.startsWith("BL_")
                    );

                    return {
                        invoiceNo: item.InvoiceNumber || "",
                        invoiceDate: item.InvoiceDate?.split("T")[0] || "",
                        boeNo: item.BOENo || "",
                        boeDate: item.BOEDate?.split("T")[0] || "",
                        mrnNo: item.MRNNumber || "",
                        mrnDate: item.MRNDate?.split("T")[0] || "",
                        blNo: item.BillofLandingNo || "",
                        blDate: item.BillOfLandingdate?.split("T")[0] || "",
                        invoiceAmount: item.InvoiceAmount || "",
                    };
                });

                setInvoiceAttachmentsadvance(trackerInvoiceAttachments);
                setOtherAttachmentsadvance(trackerOtherAttachments);
                setBoeAttachments(trackerBoeAttachments);
                setBlAttachments(trackerBlAttachments);

                const savedAmount = trackerRows.reduce(
                    (sum: number, r: { invoiceAmount: any; }) => sum + (parseFloat(r.invoiceAmount || "0")),
                    0
                );

                let finalRows = [...trackerRows];

                if (savedAmount < parseFloat(data.TotalAmount || "0")) {

                    finalRows.push({
                        invoiceNo: "",
                        invoiceDate: "",
                        boeNo: "",
                        boeDate: "",
                        mrnNo: "",
                        mrnDate: "",
                        blNo: "",
                        blDate: "",
                        invoiceAmount: "",
                    });

                }

                setRows(finalRows);
            }
        }
        // 🔹 Load already saved tracker data


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
        (sum, r) => sum + (parseFloat(r.invoiceAmountnew || "0") || 0),
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
            if (totalInvoiceAmount > parseFloat(totalAmount)) {
                alert("Total Invoice Amount cannot be greater than Total Performa Invoice Amount");
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
                if (!row.mrnNo) {
                    alert(`MRN No required in row ${i + 1}`);
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
                if (!invoiceAttachmentsadvance[i] || invoiceAttachmentsadvance[i].length === 0) {
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

                if (!invoiceAttachmentsadvance[i] || invoiceAttachmentsadvance[i].length === 0) {
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
    // const onSubmit = async () => {

    //     const sp = await spCrudOps;

    //     try {
    //         if (!validateForm()) return;
    //         const trackerData = await sp.getData(
    //             "ForexAdvanceBillPayment",
    //             "*",
    //             "",
    //             `ForexIDId eq ${Id}`,
    //             { column: "ID", isAscending: true },
    //             5000,
    //             props
    //         );
    //         const newRows = rows.filter(
    //             (r) =>
    //                 r.invoiceNo &&
    //                 r.invoiceDate &&
    //                 r.invoiceAmount &&
    //                 !trackerData.some(
    //                     (t: any) =>
    //                         t.InvoiceNumber === r.invoiceNo &&
    //                         t.InvoiceAmount == r.invoiceAmount
    //                 )
    //         );

    //         for (let i = 0; i < newRows.length; i++) {

    //             const row = newRows[i];

    //             const item = await sp.insertData(
    //                 "ForexAdvanceBillPayment",
    //                 {
    //                     ForexIDId: Number(Id),
    //                     InvoiceNumber: row.invoiceNo,
    //                     InvoiceDate: row.invoiceDate ? row.invoiceDate : null,
    //                     InvoiceAmount: '' + row.invoiceAmount,
    //                     MRNNumber: '' + row.mrnNo,
    //                     MRNDate: row.mrnDate ? row.mrnDate : null,
    //                     BOENo: '' + row.boeNo,
    //                     BOEDate: row.boeDate ? row.boeDate : null,
    //                     BillofLandingNo: row.blNo,
    //                     BillOfLandingdate: row.blDate ? row.blDate : null
    //                 },
    //                 props
    //             );

    //             const itemId = item.data.ID;

    //             const uploadFiles = async (files: any[]) => {

    //                 for (const file of files) {

    //                     const response = await fetch(file.ServerRelativeUrl);
    //                     const blob = await response.blob();

    //                     await props.context.spHttpClient.post(
    //                         `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ForexAdvanceBillPayment')/items(${itemId})/AttachmentFiles/add(FileName='${file.FileName}')`,
    //                         SPHttpClient.configurations.v1,
    //                         { body: blob }
    //                     );

    //                 }

    //             };

    //             // if (poAttachments[i]) await uploadFiles(poAttachments[i]);
    //             // if (piAttachments[i]) await uploadFiles(piAttachments[i]);
    //             // if (invoiceAttachmentsadvance[i]) await uploadFiles(invoiceAttachmentsadvance[i]);
    //             for (const file of invoiceAttachmentsadvance[i]) {

    //                 const fileName = `INV_${row.boeNo}_${file.name}`;

    //                 await props.context.spHttpClient.post(
    //                     `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ForexAdvanceBillPayment')/items(${itemId})/AttachmentFiles/add(FileName='${fileName}')`,
    //                     SPHttpClient.configurations.v1,
    //                     { body: file }
    //                 );
    //             }
    //             // if (otherAttachmentsadvance[i]) await uploadFiles(otherAttachmentsadvance[i]);
    //             for (const file of otherAttachmentsadvance[i]) {

    //                 const fileName = `DOC_${row.boeNo}_${file.name}`;

    //                 await props.context.spHttpClient.post(
    //                     `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ForexAdvanceBillPayment')/items(${itemId})/AttachmentFiles/add(FileName='${fileName}')`,
    //                     SPHttpClient.configurations.v1,
    //                     { body: file }
    //                 );
    //             }

    //             if (boeAttachments[i]) {

    //                 for (const file of boeAttachments[i]) {

    //                     const fileName = `BOE_${row.boeNo}_${file.name}`;

    //                     await props.context.spHttpClient.post(
    //                         `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ForexAdvanceBillPayment')/items(${itemId})/AttachmentFiles/add(FileName='${fileName}')`,
    //                         SPHttpClient.configurations.v1,
    //                         { body: file }
    //                     );
    //                 }
    //             }

    //             /* ---------- BL Attachments ---------- */

    //             if (blAttachments[i]) {

    //                 for (const file of blAttachments[i]) {

    //                     const fileName = `BL_${row.blNo}_${file.name}`;

    //                     await props.context.spHttpClient.post(
    //                         `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ForexAdvanceBillPayment')/items(${itemId})/AttachmentFiles/add(FileName='${fileName}')`,
    //                         SPHttpClient.configurations.v1,
    //                         { body: file }
    //                     );
    //                 }
    //             }

    //         }
    //         const treasuryUserId = await getTreasuryPaymentApprover();

    //         // 🔥 FETCH EXISTING WORKFLOW HISTORY
    //         const existingItem = await sp.getData(
    //             "ForexRequest",
    //             "WorkFlowHistory",
    //             "",
    //             `ID eq ${Id}`,
    //             { column: "ID", isAscending: true },
    //             1,
    //             props
    //         );

    //         let existingHistory: any[] = [];

    //         if (existingItem.length > 0 && existingItem[0].WorkFlowHistory) {
    //             try {
    //                 existingHistory = JSON.parse(existingItem[0].WorkFlowHistory);
    //             } catch {
    //                 existingHistory = [];
    //             }
    //         }
    //         let updatedHistory = [...existingHistory];

    //         const advanceAmount = parseFloat(totalAmount || "0");
    //         const settledAmount = parseFloat(totalInvoiceAmount.toFixed(2) || "0");

    //         let finalStatus = "";
    //         let balanceAmount = 0;

    //         if (advanceAmount === settledAmount) {

    //             finalStatus = "Paid and Pending for Settlement";
    //             balanceAmount = 0;

    //         } else {

    //             finalStatus = "Paid";
    //             balanceAmount = advanceAmount - settledAmount;

    //         }

    //         updatedHistory.push({
    //             CurrentApprover: props.context.pageContext.user.displayName,
    //             Role: "",
    //             ActionTaken: "Tracker Submitted",
    //             Comment: "",
    //             Date: new Date().toISOString(),
    //             CurrentStatus: finalStatus
    //         });

    //         await sp.updateData(
    //             "ForexRequest",
    //             Number(Id),
    //             {
    //                 Status: finalStatus,

    //                 // 🔹 New fields
    //                 // TotalPerformaInvoice: advanceAmount,
    //                 // AdvancePaid: advanceAmount,
    //                 // AmountSettled: settledAmount,
    //                 // Balance: balanceAmount,

    //                 WorkFlowHistory: JSON.stringify(updatedHistory),

    //                 CurrentApproverId:
    //                     finalStatus === "Paid and Pending for Settlement"
    //                         ? treasuryUserId
    //                         : null,
    //                 BalenceAmount: "" + balanceAmount,
    //                 SettlementAmount: "" + totalInvoiceAmount
    //             },
    //             props
    //         );


    //         alert("Tracker saved successfully");

    //         history.push("/");

    //     } catch (error) {

    //         console.error(error);
    //         alert("Error saving tracker");

    //     }

    // };
    const onSubmit = async () => {

        const sp = await spCrudOps;

        try {

            if (!validateForm()) return;

            // =====================================================
            // ✅ FETCH EXISTING TRACKER DATA
            // =====================================================

            const trackerData = await sp.getData(
                "ForexAdvanceBillPayment",
                "*",
                "",
                `ForexIDId eq ${Id}`,
                { column: "ID", isAscending: true },
                5000,
                props
            );

            // =====================================================
            // ✅ LOOP ALL ROWS
            // =====================================================

            for (let i = 0; i < rows.length; i++) {

                const row = rows[i];

                if (
                    !row.invoiceNo ||
                    !row.invoiceDate ||
                    !row.invoiceAmount
                ) continue;

                // =====================================================
                // ✅ CHECK EXISTING RECORD
                // =====================================================

                const existingRecord = trackerData.find(
                    (t: any) =>
                        t.InvoiceNumber === row.invoiceNo
                );

                let itemId = 0;

                // =====================================================
                // ✅ UPDATE EXISTING RECORD
                // =====================================================

                if (existingRecord) {

                    itemId = existingRecord.ID;

                    await sp.updateData(
                        "ForexAdvanceBillPayment",
                        itemId,
                        {
                            ForexIDId: Number(Id),

                            InvoiceNumber: row.invoiceNo,

                            InvoiceDate: row.invoiceDate
                                ? row.invoiceDate
                                : null,

                            InvoiceAmount: "" + row.invoiceAmount,

                            MRNNumber: "" + row.mrnNo,

                            MRNDate: row.mrnDate
                                ? row.mrnDate
                                : null,

                            BOENo: "" + row.boeNo,

                            BOEDate: row.boeDate
                                ? row.boeDate
                                : null,

                            BillofLandingNo: row.blNo,

                            BillOfLandingdate: row.blDate
                                ? row.blDate
                                : null
                        },
                        props
                    );

                } else {

                    // =====================================================
                    // ✅ INSERT NEW RECORD
                    // =====================================================

                    const item = await sp.insertData(
                        "ForexAdvanceBillPayment",
                        {
                            ForexIDId: Number(Id),

                            InvoiceNumber: row.invoiceNo,

                            InvoiceDate: row.invoiceDate
                                ? row.invoiceDate
                                : null,

                            InvoiceAmount: "" + row.invoiceAmount,

                            MRNNumber: "" + row.mrnNo,

                            MRNDate: row.mrnDate
                                ? row.mrnDate
                                : null,

                            BOENo: "" + row.boeNo,

                            BOEDate: row.boeDate
                                ? row.boeDate
                                : null,

                            BillofLandingNo: row.blNo,

                            BillOfLandingdate: row.blDate
                                ? row.blDate
                                : null
                        },
                        props
                    );

                    itemId = item.data.ID;
                }

                // =====================================================
                // ✅ INVOICE ATTACHMENTS
                // =====================================================

                if (invoiceAttachmentsadvance[i]) {

                    for (const file of invoiceAttachmentsadvance[i]) {

                        const fileName = `INV_${row.boeNo}_${file.name}`;

                        await props.context.spHttpClient.post(
                            `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ForexAdvanceBillPayment')/items(${itemId})/AttachmentFiles/add(FileName='${fileName}')`,
                            SPHttpClient.configurations.v1,
                            { body: file }
                        );
                    }
                }

                // =====================================================
                // ✅ OTHER ATTACHMENTS
                // =====================================================

                if (otherAttachmentsadvance[i]) {

                    for (const file of otherAttachmentsadvance[i]) {

                        const fileName = `DOC_${row.boeNo}_${file.name}`;

                        await props.context.spHttpClient.post(
                            `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ForexAdvanceBillPayment')/items(${itemId})/AttachmentFiles/add(FileName='${fileName}')`,
                            SPHttpClient.configurations.v1,
                            { body: file }
                        );
                    }
                }

                // =====================================================
                // ✅ BOE ATTACHMENTS
                // =====================================================

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

                // =====================================================
                // ✅ BL ATTACHMENTS
                // =====================================================

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

            // =====================================================
            // ✅ TREASURY USER
            // =====================================================

            const treasuryUserId = await getTreasuryPaymentApprover();

            // =====================================================
            // ✅ FETCH EXISTING WORKFLOW HISTORY
            // =====================================================

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

            if (
                existingItem.length > 0 &&
                existingItem[0].WorkFlowHistory
            ) {

                try {

                    existingHistory = JSON.parse(
                        existingItem[0].WorkFlowHistory
                    );

                } catch {

                    existingHistory = [];
                }
            }

            let updatedHistory = [...existingHistory];

            const advanceAmount = parseFloat(totalAmount || "0");

            const settledAmount = parseFloat(
                totalInvoiceAmount.toFixed(2) || "0"
            );

            let finalStatus = "";

            let balanceAmount = 0;

            if (advanceAmount === settledAmount) {

                finalStatus = "Paid and Pending for Settlement";

                balanceAmount = 0;

            } else {

                finalStatus = "Paid";

                balanceAmount = advanceAmount - settledAmount;
            }

            updatedHistory.push({
                CurrentApprover:
                    props.context.pageContext.user.displayName,

                Role: "",

                ActionTaken: "Tracker Submitted",

                Comment: "",

                Date: new Date().toISOString(),

                CurrentStatus: finalStatus
            });

            // =====================================================
            // ✅ UPDATE FOREX REQUEST
            // =====================================================

            await sp.updateData(
                "ForexRequest",
                Number(Id),
                {
                    Status: finalStatus,

                    WorkFlowHistory: JSON.stringify(updatedHistory),

                    CurrentApproverId:
                        finalStatus ===
                            "Paid and Pending for Settlement"
                            ? treasuryUserId
                            : null,

                    BalenceAmount: "" + balanceAmount,

                    SettlementAmount: "" + totalInvoiceAmount
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
    const handleInvoiceFile = (
        index: number,
        files: FileList | null
    ) => {

        if (!files) return;

        const existingFiles =
            invoiceAttachmentsadvance[index] || [];

        const newFiles = Array.from(files);

        const updated = {
            ...invoiceAttachmentsadvance
        };

        updated[index] = [
            ...existingFiles,
            ...newFiles
        ];

        setInvoiceAttachmentsadvance(updated);
    };

    const handleOtherFile = (
        index: number,
        files: FileList | null
    ) => {

        if (!files) return;

        const existingFiles =
            otherAttachmentsadvance[index] || [];

        const newFiles = Array.from(files);

        const updated = {
            ...otherAttachmentsadvance
        };

        updated[index] = [
            ...existingFiles,
            ...newFiles
        ];

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
        <>

            <div className='MainUplodForm' style={{ margin: "5px 0px" }}>
                <div className='row'>
                    <div className='col-md-12'>
                        <div className='Main-Boxpoup'>
                            <div className="bordered">
                                <a><img src={logo} /></a>
                                <h1>Forex Payment - Advance Payment Tracker</h1>
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
                                <CollapsibleSection title="Vendor - Beneficiary Details" style={{ marginTop: "10px" }}>
                                    <div className='main-formcontainer'>
                                        <div className='row mb-20'>
                                            <div className='col-md-4'>
                                                <label className='font'>Vendor Code</label>
                                                <input value={vendor.VendorCode} className="form-control readonly" />
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
                                            {/* <div className='col-md-4'>
                                                <label className="font">Bank Country</label>
                                                <input type="text" value={vendor.BankCountry} className="form-control readonly" />
                                            </div> */}
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
                                                    {/* <th>Status</th> */}
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
                                                        {/* <td>{item.CurrentStatus}</td> */}
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

                                        {/* <tfoot>
                        <tr>

                            <td colSpan={2}></td>

                            <td style={{ fontWeight: "bold" }}>Total Amount</td>

                            <td>{totalInvoiceAmountNew.toFixed(2)}</td>

                            <td colSpan={3}></td>

                        </tr>
                    </tfoot> */}

                                    </table>
                                </div>


                                {paymentType === "Goods-Advance Payment" && (
                                    <>
                                        <div className="heading1" style={{ marginTop: "10px" }}>
                                            <label>Bill Payment Details (For Goods Bill Payment)</label>
                                        </div>
                                        <div className='main-formcontainer'>
                                            <div className='row mb-20'>
                                                <div className="col-md-12" style={{ overflow: 'auto' }}>
                                                    <table className="custom-table">

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
                                                                        {invoiceAttachmentsadvance[index]?.length > 0 && (
                                                                            <div style={{ marginTop: "5px" }}>
                                                                                {invoiceAttachmentsadvance[index].map((file: any, i: number) => (
                                                                                    <div key={i}>
                                                                                        <a
                                                                                            href={file.ServerRelativeUrl}
                                                                                            target="_blank"
                                                                                            rel="noopener noreferrer"
                                                                                            style={{ fontSize: "12px" }}
                                                                                        >
                                                                                            {file.FileName}
                                                                                        </a>
                                                                                    </div>
                                                                                ))}
                                                                            </div>
                                                                        )}
                                                                    </td>

                                                                    {/* Other Upload */}
                                                                    <td>
                                                                        <input
                                                                            type="file"
                                                                            onChange={(e) => handleOtherFile(index, e.target.files)}
                                                                        />
                                                                        {otherAttachmentsadvance[index]?.length > 0 && (
                                                                            <div style={{ marginTop: "5px" }}>
                                                                                {otherAttachmentsadvance[index].map((file: any, i: number) => (
                                                                                    <div key={i}>
                                                                                        <a
                                                                                            href={file.ServerRelativeUrl}
                                                                                            target="_blank"
                                                                                            rel="noopener noreferrer"
                                                                                            style={{ fontSize: "12px" }}
                                                                                        >
                                                                                            {file.FileName}
                                                                                        </a>
                                                                                    </div>
                                                                                ))}
                                                                            </div>
                                                                        )}

                                                                    </td>

                                                                    <td style={{ textAlign: "center" }}>

                                                                        {index === rows.length - 1 && (
                                                                            <span    style={{
                                                                            background: "#28a745",
                                                                            color: "white",
                                                                            marginRight: "5px",
                                                                            border: "none",
                                                                            padding: "5px 10px",
                                                                            cursor: "pointer",
                                                                            borderRadius: "4px"
                                                                        }} onClick={addRow}>+</span>
                                                                        )}

                                                                        {rows.length > 1 && (
                                                                            <span
                                                                                style={{
                                                                            background: "#dc3545",
                                                                            color: "white",
                                                                            border: "none",
                                                                            padding: "5px 10px",
                                                                            cursor: "pointer",
                                                                            borderRadius: "4px"
                                                                        }}
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
                                                </div>
                                            </div>
                                            <div className="row mb-20">
                                                <div className="col-md-12">
                                                    <div style={{ display: "flex", gap: "40px", marginTop: "20px" }}>

                                                        {/* BOE TABLE */}

                                                        <div>

                                                            <p style={{ color: "red", fontSize: "12px" }}>
                                                                Unique BOE no will be listed below
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

                                                                                <input
                                                                                    type="file"
                                                                                    onChange={(e) => handleBoeUpload(index, e.target.files)}
                                                                                />

                                                                                {boeAttachments[index]?.map((file: any, i: number) => (
                                                                                    <div key={i}>

                                                                                        <a
                                                                                            href={file.ServerRelativeUrl || "#"}
                                                                                            target="_blank"
                                                                                            rel="noopener noreferrer"
                                                                                            style={{ fontSize: "12px" }}
                                                                                        >
                                                                                            {file.FileName || file.name}
                                                                                        </a>

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

                                                                                <input
                                                                                    type="file"
                                                                                    onChange={(e) => handleBlUpload(index, e.target.files)}
                                                                                />

                                                                                {blAttachments[index]?.map((file: any, i: number) => (
                                                                                    <div key={i}>

                                                                                        <a
                                                                                            href={file.ServerRelativeUrl || "#"}
                                                                                            target="_blank"
                                                                                            rel="noopener noreferrer"
                                                                                            style={{ fontSize: "12px" }}
                                                                                        >
                                                                                            {file.FileName || file.name}
                                                                                        </a>

                                                                                    </div>
                                                                                ))}
                                                                            </td>

                                                                        </tr>

                                                                    ))}

                                                                </tbody>

                                                            </table>

                                                        </div>

                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    </>
                                )}

                                {paymentType === "Service-Advance Payment" && (
                                    <>
                                        <div className="heading1" style={{ marginTop: "10px" }}>
                                            <label>Bill Payment Details (For Goods Bill Payment)</label>
                                        </div>                                        <div className='main-formcontainer'>
                                            <div className='row mb-20'>
                                                <div className="col-md-12" style={{ overflow: 'auto' }}>
                                                    <table className="custom-table">

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
                                                                        {invoiceAttachmentsadvance[index]?.length > 0 && (
                                                                            <div style={{ marginTop: "5px" }}>
                                                                                {invoiceAttachmentsadvance[index].map((file: any, i: number) => (
                                                                                    <div key={i}>
                                                                                        <a
                                                                                            href={file.ServerRelativeUrl}
                                                                                            target="_blank"
                                                                                            rel="noopener noreferrer"
                                                                                            style={{ fontSize: "12px" }}
                                                                                        >
                                                                                            {file.FileName}
                                                                                        </a>
                                                                                    </div>
                                                                                ))}
                                                                            </div>
                                                                        )}
                                                                    </td>

                                                                    <td>
                                                                        <input
                                                                            type="file"
                                                                            onChange={(e) => handleOtherFile(index, e.target.files)}
                                                                        />
                                                                        {otherAttachmentsadvance[index]?.length > 0 && (
                                                                            <div style={{ marginTop: "5px" }}>
                                                                                {otherAttachmentsadvance[index].map((file: any, i: number) => (
                                                                                    <div key={i}>
                                                                                        <a
                                                                                            href={file.ServerRelativeUrl}
                                                                                            target="_blank"
                                                                                            rel="noopener noreferrer"
                                                                                            style={{ fontSize: "12px" }}
                                                                                        >
                                                                                            {file.FileName}
                                                                                        </a>
                                                                                    </div>
                                                                                ))}
                                                                            </div>
                                                                        )}

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
                                            </div>
                                        </div>
                                    </>
                                )}

                                <div className='row my-3'>
                                    <div className='col-md-12'>
                                        <div style={{ display: "flex", justifyContent: "center", gap: "5px" }}>
                                            <button onClick={onSubmit} className="Submit-btn">
                                                Submit
                                            </button>

                                            <button onClick={() => history.goBack()} className="Exit-btn">
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

export default TrackerForm;