import React, { useEffect, useState } from "react";
import "../Pages/Css/NewRequest.scss";
import { IForexModuleProps } from "../IForexModuleProps";
import { useHistory, useParams } from 'react-router-dom';
import { Dropdown, IDropdownOption } from '@fluentui/react';
import SPCRUDOPS from "../../service/BAL/spcrud";
import { Attachment } from "@pnp/sp/attachments";
import { View } from "@pnp/sp/views";
//import USESPCRUD from "../../service/BAL/spcrud";

import logo from "../../assets/sona-comstarlogo.png";

interface InvoiceRow {
    invoiceNo: string;
    invoiceDate: string;
    boeNo: string;
    boeDate: string;
    mrnNo: string;
    blNo: string;
    blDate: string;
    invoiceAmount: string;
    mrnDate?: string; // Added for Service-Bill Payment
}
/* Helper Components */
const Section = ({ title, children }: any) => (
    <div className="form-section">
        <h3>{title}</h3>
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

interface IApproverDetails {
    Id: number;
    Name: string;
    Role: string;
    Level: number;
    Status?: string;
    Comment?: string;
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

const ViewRequestForm = (props: IForexModuleProps) => {
    const { Id } = useParams<{ Id: string }>();
    const history = useHistory();
    const spCrudOps = SPCRUDOPS();
    const [paymentType, setPaymentType] = useState("Goods-Bill Payment");
    const [taxDocumentView, setTaxDocumentView] = useState("Yes");
    const [paymenttypeDropdownValue, setPaymentTypeDropdownValue] = useState<IDropdownOption>();
    const [fromdate, setFromDate] = useState("");
    const [todate, setToDate] = useState("");
    const [dTAAApplicable, setDTAAApplicable] = useState("");
    const [showVendorPopup, setShowVendorPopup] = React.useState(false);
    const [bankaccountno, setBankAccountNo] = useState("");
    const [bankname, setBankName] = useState("");
    const [bankswiftcode, setBankSwiftCode] = useState("");
    const [remarks, setRemarks] = useState("");

    const [requestNumber, setRequestNumber] = useState("");
    const [requestedOn, setRequestedOn] = useState("");
    const [currency, setCurrency] = useState("");
    const [totalAmount, setTotalAmount] = useState("");
    const [foreignBankCharges, setForeignBankCharges] = useState("");
    const [poContractNo, setPoContractNo] = useState("");
    const [poDate, setPoDate] = useState("");
    const [expectedSettlementDate, setExpectedSettlementDate] = useState("");

    const [eligibleAmountWithWHT, setEligibleAmountWithWHT] = useState("");
    const [paidAmount, setPaidAmount] = useState("");
    const [ballenceEligibleAmount, setBallenceEligibleAmount] = useState("");

    const [invoiceAttachments, setInvoiceAttachments] = useState<any>({});
    const [otherAttachments, setOtherAttachments] = useState<any>({});
    const [poAttachments, setPoAttachments] = useState<any>({});
    const [piAttachments, setPiAttachments] = useState<any>({});

    const [boeLibraryFiles, setBoeLibraryFiles] = useState<any[]>([]);
    const [bolLibraryFiles, setBolLibraryFiles] = useState<any[]>([]);
    const [approverDetails, setApproverDetails] = useState<IApproverDetails[]>([]);
    const [currentApproverId, setCurrentApproverId] = useState<number | null>(null);
    const [approvers, setApprovers] = useState<any[]>([]);
    const [foreignCurrency, setForeignCurrency] = useState("");
    const [foreignCurrencyAmount, setForeignCurrencyAmount] = useState("");
    const [exchangeRate, setExchangeRate] = useState("");
    const [inrAmount, setInrAmount] = useState("");
    const [paymentDate, setPaymentDate] = useState("");
    const [paymentReferenceNumber, setPaymentReferenceNumber] = useState("");
    const [workflowHistory, setWorkflowHistory] = useState<any[]>([]);
    const [poAttachmentsAdvance, setPoAttachmentsAddvance] = useState<any>({});
    const [piAttachmentsAddvance, setPiAttachmentsAddvance] = useState<any>({});
    const [invoiceAttachmentsTrackeradvance, setInvoiceAttachmentsTrackeradvance] = useState<any>({});
    const [otherAttachmentsTrackeradvance, setOtherAttachmentsTrackeradvance] = useState<any>({});
    const [poAttachmentsTrackeradvance, setPoAttachmentsTrackeradvance] = useState<any>({});
    const [piAttachmentsTrackeradvance, setPiAttachmentsTrackeradvance] = useState<any>({});
    const [boeAttachments, setBoeAttachments] = useState<any>({});
    const [blAttachments, setBlAttachments] = useState<any>({});
    const [boeAttachmentstrackeradvance, setBoeAttachmentstrackeradvance] = useState<any>({});
    const [blAttachmentstrackeradvance, setBlAttachmentstrackeradvance] = useState<any>({});
    const [rowsAdvance, setRowsAdvance] = useState<InvoiceRow[]>([
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

    const [employee, setEmployee] = React.useState({
        EmployeeCode: "",
        EmployeeName: "",
        Division: "",
        Location: "",
        RM: "",
        HOD: "",
        ContactNo: "",
        EmployeeStatus: "",
        Email: "",
        RMId: 0,
        HODId: 0
    });
    const [vendor, setVendor] = React.useState({
        VendorCode: "",
        VendorName: "",
        VendorNameLegal: "",
        VendorShortName: "",
        VendorType: "",
        City: "",
        State: "",
        Country: "",
        Currency: "",
        PostalCode: "",
        ContactPersonName: "",
        EmailId: "",
        PhoneNumber: "",
        AlternateContact: "",
        BeneficiaryName: "",
        BankName: "",
        AccountNumberIBAN: "",
        SWIFTBICCode: "",
        RoutingNumberABA: "",
        IFSCCode: "",
        IntermediaryBank: "",
        IntermediarySWIFTCode: "",
        NatureOfPayment: "",
        PurposeCodeRBI: "",
        BankCountry: "",
        BankAddress: "",
        VendorAddress: "",
        Pincode: ""
    });

    const [permanentEstablishmentDeclaration, setPermanentEstablishmentDeclaration] = useState({
        DocumentAvailable: "",
        DocumentNumber: "",
        DocumentDate: "",
        SEPClause: "",
        ValidityStartDate: "",
        ValidityEndDate: "",
        Attachmenturl: "",
        Attachmentfilename: ""
    });

    const [taxResidencyCertificate, setTaxResidencyCertificate] = useState({
        DocumentAvailable: "",
        DocumentNumber: "",
        DocumentDate: "",
        ValidityStartDate: "",
        ValidityEndDate: "",
        TaxIdentificationNumber: "",
        CountryOfTaxResidence: "",
        Attachmenturl: "",
        Attachmentfilename: ""
    });

    const [form10F, setForm10F] = useState({
        DocumentAvailable: "",
        DocumentNumber: "",
        DocumentDate: "",
        ValidityStartDate: "",
        ValidityEndDate: "",
        AcknowledgmentNumber: "",
        Attachmenturl: "",
        Attachmentfilename: ""
    });

    const [rows, setRows] = React.useState([
        {
            invoiceNo: "",
            invoiceDate: "",
            boeNo: "",
            boeDate: "",
            mrnNo: "",
            blNo: "",
            blDate: "",
            invoiceAmount: "",
            mrnDate: "" // Added for Service-Bill Payment
        }
    ]);
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
                mrnDate: "" // Added for Service-Bill Payment
            }
        ]);
    };

    const deleteRow = (index: number) => {
        const updatedRows = rows.filter((_, i) => i !== index);
        setRows(updatedRows);
    };

    const handleChange = (
        index: number,
        field: keyof InvoiceRow,
        value: string
    ) => {
        const updatedRows = [...rows];
        updatedRows[index][field] = value;
        setRows(updatedRows);
    };

    const [swiftCopy, setSwiftCopy] = useState<any>({});
    const [form15CA, setForm15CA] = useState<any>({});
    const [form15CB, setForm15CB] = useState<any>({});
    const [foreignCurrencypayment, setForeignCurrencypayment] = useState("");
    const [foreignAmountpayment, setForeignAmountpayment] = useState("");
    const [exchangeRatepayment, setExchangeRatepayment] = useState("");
    const [inrAmountpayment, setInrAmountpayment] = useState("");
    const [paymentDatepayment, setPaymentDatepayment] = useState("");
    const [paymentReferencepayment, setPaymentReferencepayment] = useState("");
    const isServicePayment = paymentType.includes("Service");
    const isGoodsPayment = paymentType.includes("Goods");
    const [validationDateshow, setValidationDateshow] = useState("");
    const [voucherNumbershow, setVoucherNumbershow] = useState("");

    const totalInvoiceAmount = rows.reduce((sum, row) => {
        return sum + (parseFloat(row.invoiceAmount) || 0);
    }, 0);
    const uniqueBoeNumbers = rows
        .map(r => r.boeNo)
        .filter((value, index, self) =>
            value && self.indexOf(value) === index
        );

    const uniqueBlNumbers = rows
        .map(r => r.blNo)
        .filter((value, index, self) =>
            value && self.indexOf(value) === index
        );

    useEffect(() => {
        getFinancialYearStart();
        if (Id) {
            loadForexData(Id);
            loadTrackerData(parseInt(Id));
        }
    }, [Id])
    //----------------------- load data for edit------------------------//
    const loadForexData = async (forexId: string) => {
        const sp = await spCrudOps;

        try {

            // 🔹 Load Parent
            const parent = await sp.getData(
                "ForexRequest",
                "*,AttachmentFiles,AllApprovers,RM/Title,HOD/Title,Author/Id,Currency/Title,Currency/Id,CurrentApprover/Id,CurrentApprover/Title,NextApprovers/Id,NextApprovers/Title,PrevApprovers/Id,PrevApprovers/Title",
                "AttachmentFiles,RM,HOD,Author,Currency,CurrentApprover,NextApprovers,PrevApprovers",
                `ID eq ${forexId}`,
                { column: "ID", isAscending: true },
                1,
                props
            );
            const swiftCopyMap: any = {};
            const form15CAMap: any = {};
            const form15CBMap: any = {};
            if (parent.length > 0) {
                const data = parent[0];

                setPaymentType(data.ForexType || "");
                setRequestNumber(data.ForexNumber || "");
                setRequestedOn(data.RequestedOn?.split("T")[0] || "");
                setCurrency(data.Currency?.Title || "");
                setTotalAmount(data.TotalAmount || "");
                setForeignBankCharges(data.ForeignBankCharges || "");
                setPoContractNo(data.poContractNo || "");
                setPoDate(data.poDate?.split("T")[0] || "");
                setExpectedSettlementDate(data.expectedSettlementDate?.split("T")[0] || "");
                setBankName(data.BankName || "");
                setBankAccountNo(data.BankAccNo || "");
                setRemarks(data.Remarks || "");
                setDTAAApplicable(data.DTAAApplicable || "");
                setTaxDocumentView(data.DocumentIsAvailable || "");
                setBankSwiftCode(data.BankSwiftCode || "");
                setEligibleAmountWithWHT(data.EligibleAmountWithWHT || "");
                setPaidAmount(data.PaidAmount || "");
                setBallenceEligibleAmount(data.BallenceEligibleAmount || "");
                setForeignCurrency(data.ForeignCurrency || "");
                setForeignCurrencyAmount(data.ForeignCurrencyAmount || "");
                setExchangeRate(data.ExchangeRate || "");
                setInrAmount(data.INRAmount || "");
                setValidationDateshow(data.ValidationDate?.split("T")[0] || "");
                setVoucherNumbershow(data.VoucherNumber || "");
                setPaymentDate(data.PaymentDate?.split("T")[0] || "");
                setPaymentReferenceNumber(data.PaymentReferenceNumber || "");
                // 🔥 LOAD WORKFLOW HISTORY
                if (data.WorkFlowHistory) {
                    try {
                        const parsed = JSON.parse(data.WorkFlowHistory);
                        setWorkflowHistory(parsed);
                    } catch {
                        setWorkflowHistory([]);
                    }
                }
                setVendor({
                    ...vendor,
                    VendorCode: data.VendorCode,
                    VendorName: data.VendorName
                });
                setEmployee({
                    EmployeeCode: data.EmployeeCode || "",
                    EmployeeName: data.EmployeeName || "",
                    Division: data.Division || "",
                    Location: data.Location || "",
                    RM: data.RM?.Title || "",
                    HOD: data.HOD?.Title || "",
                    ContactNo: data.ContactNo || "",
                    EmployeeStatus: data.EmployeeStatus || "",
                    Email: data.Email || "",
                    RMId: data.RM?.Id || 0,
                    HODId: data.HOD?.Id || 0
                });
                getVendorData(data.VendorCode);
                setCurrentApproverId(data.CurrentApprover?.Id || null);

                const approverList: IApproverDetails[] = [];

                if (data.AllApprovers) {

                    const parsed = JSON.parse(data.AllApprovers);

                    parsed.forEach((a: any) => {

                        let status = "pending";

                        if (a.status === "Approved") {
                            status = "approved";
                        }

                        if (data.CurrentApprover?.Id === a.Id) {
                            status = "current";
                        }

                        approverList.push({
                            Id: a.Id,
                            Name: a.Name,
                            Role: a.Role,
                            Level: a.Level,
                            Status: status,
                            Comment: a.Remarks
                        });

                    });

                }


                setApproverDetails(approverList);

                setForeignAmountpayment(data.ForeignCurrencyAmount || "");
                setPaymentReferencepayment(data.PaymentReferenceNumber || "");
                setPaymentDatepayment(data.PaymentDate?.split("T")[0] || "");
                setInrAmountpayment(data.INRAmount || "");
                setExchangeRatepayment(data.ExchangeRate || "");
                setForeignCurrencypayment(data.ForeignCurrency || "");
                const parentFiles = data.AttachmentFiles || {};

                const swiftFiles = parentFiles.filter((file: any) =>
                    file.FileName?.startsWith("SwiftCopy_")
                );

                const form15CAFiles = parentFiles.filter((file: any) =>
                    file.FileName?.startsWith("15CA_")
                );

                const form15CBFiles = parentFiles.filter((file: any) =>
                    file.FileName?.startsWith("15CB_")
                );

                setSwiftCopy(swiftFiles);
                setForm15CA(form15CAFiles);
                setForm15CB(form15CBFiles);

            }


            const child = await sp.getData(
                "ForexServicesBillPayment",
                "*,AttachmentFiles",
                "AttachmentFiles",
                `ForexIDId eq ${forexId}`,
                { column: "ID", isAscending: true },
                5000,
                props
            );

            if (child.length > 0) {

                const formattedRows = child.map((item: any) => ({
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
                const invoiceAttachmentMap: any = {};
                const otherAttachmentMap: any = {};
                const poAttachmentMap: any = {};
                const piAttachmentMap: any = {};

                child.forEach((item: any, index: number) => {

                    const allFiles = item.AttachmentFiles || [];

                    invoiceAttachmentMap[index] = allFiles.filter((file: any) =>
                        file.FileName?.startsWith("INV_")
                    );

                    otherAttachmentMap[index] = allFiles.filter((file: any) =>
                        file.FileName?.startsWith("DOC_")
                    );

                    poAttachmentMap[index] = allFiles.filter((file: any) =>
                        file.FileName?.startsWith("PO_")
                    );

                    piAttachmentMap[index] = allFiles.filter((file: any) =>
                        file.FileName?.startsWith("PI_")
                    );
                });

                setInvoiceAttachments(invoiceAttachmentMap);
                setOtherAttachments(otherAttachmentMap);
                setPoAttachments(poAttachmentMap);
                setPiAttachments(piAttachmentMap);
            }

            const bolFiles = await sp.getData(
                "BillOfLandingAttachment",
                "FileLeafRef,FileRef,BOLNo,ReqeuestId",
                "",
                `ReqeuestId eq '${Id}'`,
                { column: "ID", isAscending: true },
                5000,
                props
            );

            const boeFiles = await sp.getData(
                "BOEAttachments",
                "FileLeafRef,FileRef,BOENo,ReqeuestId",
                "",
                `ReqeuestId eq '${Id}'`,
                { column: "ID", isAscending: true },
                5000,
                props
            );
            setBoeLibraryFiles(boeFiles);

            setBolLibraryFiles(bolFiles);

        } catch (error) {
            console.error("Error loading edit data:", error);
        }
    };
    //=----------------------- userdata------------------------//

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

            setRowsAdvance(formattedRows);

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

            setPoAttachmentsTrackeradvance(poAttachmentMap);
            setPiAttachmentsTrackeradvance(piAttachmentMap);
            setOtherAttachmentsTrackeradvance(otherAttachmentMap);
            setBoeAttachmentstrackeradvance(boeAttachmentMap);
            setBlAttachmentstrackeradvance(blAttachmentMap);
        }
    };
    const uniqueBoeNumbersadvancetracker = Array.from(
        new Set(rowsAdvance.map((r) => r.boeNo).filter(Boolean))
    );

    const uniqueBlNumbersadvancetracker = Array.from(
        new Set(rowsAdvance.map((r) => r.blNo).filter(Boolean))
    );
    //----------------------VendorData-------------------------//
    const getVendorData = async (vendorCode: string) => {

        if (!vendorCode) return;

        (await spCrudOps).getData(
            "VendorMaster",
            "Pincode,VendorCode,VendorName,VendorNameLegal,VendorShortName,VendorType,City/Title,City/City,State/Title,Country/Country,Currency/Title,PostalCode,ContactPersonName,EmailId,PhoneNumber,AlternateContact,BeneficiaryName,BankName,AccountNumberIBAN,SWIFTBICCode,RoutingNumberABA,IFSCCode,IntermediaryBank,IntermediarySWIFTCode,NatureOfPayment/Title,PurposeCodeRBI,BankCountry,BankAddress,VendorAddress,BalanceEligibleAmount,ApprovedAmountPaidAmount,EligibleAmountWithoutWHT,TaxDocumentAvailable,DTAAApplicable,city0",
            "NatureOfPayment,City,State,Country,Currency",
            `VendorCode eq '${vendorCode}'`,
            { column: "ID", isAscending: true },
            1,
            props
        )
            .then((res: any[]) => {

                if (res.length > 0) {

                    const v = res[0];

                    setVendor({
                        VendorCode: v.VendorCode || "",
                        VendorName: v.VendorName || "",
                        VendorNameLegal: v.VendorNameLegal || "",
                        VendorShortName: v.VendorShortName || "",
                        VendorType: v.VendorType || "",
                        City: v.City?.City || "",
                        State: v.State?.Title || "",
                        Country: v.Country?.Country || "",
                        Currency: v.Currency?.Title || "",
                        PostalCode: v.PostalCode || "",
                        ContactPersonName: v.ContactPersonName || "",
                        EmailId: v.EmailId || "",
                        PhoneNumber: v.PhoneNumber || "",
                        AlternateContact: v.AlternateContact || "",
                        BeneficiaryName: v.BeneficiaryName || "",
                        BankName: v.BankName || "",
                        AccountNumberIBAN: v.AccountNumberIBAN || "",
                        SWIFTBICCode: v.SWIFTBICCode || "",
                        RoutingNumberABA: v.RoutingNumberABA || "",
                        IFSCCode: v.IFSCCode || "",
                        IntermediaryBank: v.IntermediaryBank || "",
                        IntermediarySWIFTCode: v.IntermediarySWIFTCode || "",
                        NatureOfPayment: v.NatureOfPayment?.Title || "",
                        PurposeCodeRBI: v.PurposeCodeRBI || "",
                        BankAddress: v.BankAddress || "",
                        BankCountry: v.BankCountry || "",
                        VendorAddress: v.VendorAddress || "",
                        Pincode: v.Pincode || ""
                    });
                    getTaxDeclarationdata(v.VendorCode);
                    setBankAccountNo(v.AccountNumberIBAN || "");
                    setBankName(v.BankName || "");
                    setBankSwiftCode(v.SWIFTBICCode || "");
                    setBallenceEligibleAmount(v.BalanceEligibleAmount || "");
                    setPaidAmount(v.ApprovedAmountPaidAmount || "");
                    setEligibleAmountWithWHT(v.EligibleAmountWithoutWHT || "");
                    setDTAAApplicable(v.DTAAApplicable || "");
                    setTaxDocumentView(v.TaxDocumentAvailable || "");


                } else {
                    //alert("Vendor not found");
              //      setShowVendorPopup(true);
                }

            })
            .catch((error: any) => {
                console.error("Error fetching vendor data:", error);
            });
    };
    const getTaxDeclarationdata = async (vendorCode: string) => {

        if (!vendorCode) return;

        (await spCrudOps).getData(
            "VendorTaxDeclaration",
            "VendorMasterId/ID,VendorCode/VendorCode,VendorName/VendorName,DeclarationType,DocumentAvailable,DocumentNumber,DocumentDate,SEPClause,ValidityStartDate,ValidityEndDate,TaxIdentificationNumber,CountryOfTaxResidence,AcknowledgmentNumber,AttachmentFiles/FileName,AttachmentFiles/ServerRelativeUrl",
            "VendorMasterId,VendorCode,VendorName,AttachmentFiles",
            `VendorCode/VendorCode eq '${vendorCode}'`,
            { column: "ID", isAscending: true },
            5000,
            props
        )
            .then((res: any[]) => {

                const dates: string[] = [];
                const permanent = res.find(
                    item => item.DeclarationType === "Permanent Establishment"
                );

                if (permanent) {
                    const endDate = permanent.ValidityEndDate?.split("T")[0];
                    if (endDate) dates.push(endDate);

                    setPermanentEstablishmentDeclaration({
                        DocumentAvailable: permanent.DocumentAvailable || "",
                        DocumentNumber: permanent.DocumentNumber || "",
                        DocumentDate: permanent.DocumentDate?.split("T")[0] || "",
                        SEPClause: permanent.SEPClause || "",
                        ValidityStartDate: permanent.ValidityStartDate?.split("T")[0] || "",
                        ValidityEndDate: permanent.ValidityEndDate?.split("T")[0] || "",
                        Attachmenturl: permanent.AttachmentFiles?.[0]?.ServerRelativeUrl || "",
                        Attachmentfilename: permanent.AttachmentFiles?.[0]?.FileName || ""
                    });
                }
                const taxRes = res.find(
                    item => item.DeclarationType === "TAX Residency Certificate"
                );

                if (taxRes) {
                    const endDate = permanent.ValidityEndDate?.split("T")[0];
                    if (endDate) dates.push(endDate);
                    setTaxResidencyCertificate({
                        DocumentAvailable: taxRes.DocumentAvailable || "",
                        DocumentNumber: taxRes.DocumentNumber || "",
                        DocumentDate: taxRes.DocumentDate?.split("T")[0] || "",
                        ValidityStartDate: taxRes.ValidityStartDate?.split("T")[0] || "",
                        ValidityEndDate: taxRes.ValidityEndDate?.split("T")[0] || "",
                        TaxIdentificationNumber: taxRes.TaxIdentificationNumber || "",
                        CountryOfTaxResidence: taxRes.CountryOfTaxResidence || "",
                        Attachmenturl: taxRes.AttachmentFiles?.[0]?.ServerRelativeUrl || "",
                        Attachmentfilename: taxRes.AttachmentFiles?.[0]?.FileName || ""
                    });
                }

                // ✅ Form 10F
                const form10 = res.find(
                    item => item.DeclarationType === "Form 10 F"
                );

                if (form10) {
                    const endDate = permanent.ValidityEndDate?.split("T")[0];
                    if (endDate) dates.push(endDate);

                    setForm10F({
                        DocumentAvailable: form10.DocumentAvailable || "",
                        DocumentNumber: form10.DocumentNumber || "",
                        DocumentDate: form10.DocumentDate?.split("T")[0] || "",
                        ValidityStartDate: form10.ValidityStartDate?.split("T")[0] || "",
                        ValidityEndDate: form10.ValidityEndDate?.split("T")[0] || "",
                        AcknowledgmentNumber: form10.AcknowledgmentNumber || "",
                        Attachmenturl: form10.AttachmentFiles?.[0]?.ServerRelativeUrl || "",
                        Attachmentfilename: form10.AttachmentFiles?.[0]?.FileName || ""
                    });
                }
                if (dates.length > 0) {
                    const minDate = dates.sort()[0];   // YYYY-MM-DD format works for sorting
                    setToDate(minDate);                // store in state
                }
            })
            .catch((error: any) => {
                console.error("Error fetching tax declaration data:", error);
            });
    };

    const getFinancialYearStart = () => {
        const today = new Date();
        const year = today.getFullYear();
        const month = today.getMonth(); // 0 = Jan

        const fyStartYear = month < 3 ? year - 1 : year;

        setFromDate(`${fyStartYear}-04-01`);
    };

    const hasValue = (value: any) =>
        value !== null &&
        value !== undefined &&
        value !== "" &&
        value !== "0";

    const hasAdvanceData = rowsAdvance.some(
        row =>
            row.invoiceNo ||
            row.invoiceDate ||
            row.boeNo ||
            row.boeDate ||
            row.mrnNo ||
            row.blNo ||
            row.blDate ||
            row.invoiceAmount
    );
    const hasPaymentData =
    foreignCurrency ||
    foreignAmountpayment ||
    exchangeRatepayment ||
    inrAmountpayment ||
    paymentDatepayment ||
    paymentReferencepayment ||
    swiftCopy?.length > 0 ||
    form15CA?.length > 0 ||
    form15CB?.length > 0;
    return (

        <>

            <div className='MainUplodForm' style={{ margin: "5px 0px" }}>
                <div className='row'>
                    <div className='col-md-12'>
                        <div className='Main-Boxpoup'>
                            <div className="bordered">
                                <a><img src={logo} /></a>
                                <h1>Forex Payment View Form</h1>
                            </div>
                            <div className="approval-ribbon">

                                <div className="ribbon-step initiator">
                                    {employee.EmployeeName}
                                </div>

                                {approverDetails.map((approver, index) => {

                                    let stepClass = approver.Status || "pending";

                                    return (
                                        <div key={index} className={`ribbon-step ${stepClass}`}>
                                            {approver.Role} - {approver.Name}
                                        </div>
                                    );

                                })}

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
                                {/* <div className="heading1" style={{ marginTop: "10px" }}>
                                        <label>Vendor / Beneficiary Details</label>
                                    </div> */}
                                <CollapsibleSection title="Vendor / Beneficiary Details" style={{ marginTop: "10px" }}>
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
                                <CollapsibleSection title="Tax & Regulatory Information" style={{ marginTop: "10px" }}>

                                    <div className="heading1" style={{ marginTop: "10px" }}>
                                        <label>Tax & Regulatory Information</label>
                                    </div>
                                    <div className='main-formcontainer'>
                                        <div className='row mb-20'>
                                            {/* <div className='col-md-4'>
                                                <label className='font'>Nature of Payment</label>
                                                <input type="text" value={paymentType} className="form-control readonly" />
                                            </div> */}
                                            <div className='col-md-4'>
                                                <label className='font fontblock'>Tax Document Available?</label>
                                                <select onChange={(e) => { setTaxDocumentView(e.target.value) }} disabled
                                                    style={{
                                                        color: "black",
                                                        backgroundColor: "white",
                                                        opacity: 1,
                                                        WebkitTextFillColor: "black" // important for Chrome
                                                    }} className="form-controltext readonly" >
                                                    <option>Yes</option>
                                                    <option>No</option>
                                                </select>
                                                {taxDocumentView === "No" && (
                                                    <Field >
                                                        <span style={{ color: "red" }}>
                                                            (if No, withholding tax will be applicable)
                                                        </span>
                                                    </Field>
                                                )}
                                            </div>
                                            {taxDocumentView === "Yes" && (
                                                <div className="col-md-4">
                                                    <label className="font fontblock">DTAA Applicable?</label>
                                                    <select value={dTAAApplicable} onChange={(e) => setDTAAApplicable(e.target.value)} disabled
                                                        style={{
                                                            color: "black",
                                                            backgroundColor: "white",
                                                            opacity: 1,
                                                            WebkitTextFillColor: "black" // important for Chrome
                                                        }} className="form-controltext readonly" >
                                                        <option value="">Select</option>
                                                        <option value="Yes">Yes</option>
                                                        <option value="No">No</option>
                                                    </select>
                                                </div>
                                            )}
                                        </div>
                                    </div>
                                    {taxDocumentView === "Yes" && (
                                        <>
                                            <div className="heading1" style={{ marginTop: "10px" }}>
                                                <label>Permanent Establishment Declaration</label>
                                            </div>
                                            <div className='main-formcontainer'>
                                                <div className='row mb-20'>
                                                    <div className="col-md-4">
                                                        <label className="font fontblock">Document Available</label>
                                                        <select
                                                            value={permanentEstablishmentDeclaration.DocumentAvailable || ""}
                                                            className="form-controltext readonly">
                                                            <option value="">Select</option>
                                                            <option value="Yes">Yes</option>
                                                            <option value="No">No</option>
                                                        </select>
                                                    </div>
                                                    <div className='col-md-4'>
                                                        <label className='font'>Document Number</label>
                                                        <input type="text" value={permanentEstablishmentDeclaration.DocumentNumber || ""} className="form-control readonly" />
                                                    </div>
                                                    <div className='col-md-4'>
                                                        <label className='font'>Document Date</label>
                                                        <input type="date" value={permanentEstablishmentDeclaration.DocumentDate || ""} className="form-control readonly" />
                                                    </div>
                                                </div>
                                                <div className='row mb-20'>
                                                    <div className="col-md-4">
                                                        <label className="font">Validity Start Date</label>
                                                        <input type="date" value={permanentEstablishmentDeclaration.ValidityStartDate || ""} className="form-control readonly" />
                                                    </div>
                                                    <div className='col-md-4'>
                                                        <label className='font'>Validity End Date</label>
                                                        <input type="date" value={permanentEstablishmentDeclaration.ValidityEndDate || ""} className="form-control readonly" />
                                                    </div>
                                                    <div className='col-md-4'>
                                                        <label className='font fontblock'>View Document</label>
                                                        <span><a href={permanentEstablishmentDeclaration.Attachmenturl || "#"} target="_blank">{permanentEstablishmentDeclaration.Attachmentfilename || "No Document Available"}</a></span>
                                                    </div>

                                                </div>
                                            </div>
                                            <div className="heading1" style={{ marginTop: "10px" }}>
                                                <label>Tax Residency Certificate</label>
                                            </div>
                                            <div className='main-formcontainer'>
                                                <div className='row mb-20'>
                                                    <div className="col-md-4">
                                                        <label className="font fontblock">Document Available</label>
                                                        <select
                                                            value={taxResidencyCertificate.DocumentAvailable || ""}
                                                            className="form-controltext readonly">
                                                            <option value="">Select</option>
                                                            <option value="Yes">Yes</option>
                                                            <option value="No">No</option>
                                                        </select>
                                                    </div>
                                                    <div className='col-md-4'>
                                                        <label className='font'>Document Number</label>
                                                        <input type="text" value={taxResidencyCertificate.DocumentNumber || ""} className="form-control readonly" />
                                                    </div>
                                                    <div className='col-md-4'>
                                                        <label className='font'>Country of Tax Residence</label>
                                                        <input type="text" value={taxResidencyCertificate.CountryOfTaxResidence || ""} className="form-control readonly" />
                                                    </div>
                                                </div>
                                                <div className='row mb-20'>
                                                    <div className="col-md-4">
                                                        <label className="font fontblock">Tax Identification Number</label>
                                                        <input type="text" value={taxResidencyCertificate.TaxIdentificationNumber || "" || ""} className="form-control readonly" />
                                                    </div>
                                                    <div className='col-md-4'>
                                                        <label className='font'>Validity Start Date</label>
                                                        <input type="date" value={taxResidencyCertificate.ValidityStartDate || "" || ""} className="form-control readonly" />
                                                    </div>
                                                    <div className='col-md-4'>
                                                        <label className='font'>Validity End Date</label>
                                                        <input type="date" value={taxResidencyCertificate.ValidityEndDate || ""} className="form-control readonly" />
                                                    </div>
                                                </div>
                                                <div className='row mb-20'>
                                                    <div className="col-md-4">
                                                        <label className="font fontblock">View Document</label>
                                                        <span><a href={taxResidencyCertificate.Attachmenturl || "#"} target="_blank">{taxResidencyCertificate.Attachmentfilename || "No Document Available"}</a></span>
                                                    </div>
                                                </div>
                                            </div>
                                            <div className="heading1" style={{ marginTop: "10px" }}>
                                                <label>Form 10F</label>
                                            </div>
                                            <div className='main-formcontainer'>
                                                <div className='row mb-20'>
                                                    <div className="col-md-4">
                                                        <label className="font fontblock">Document Available</label>
                                                        <select
                                                            value={form10F.DocumentAvailable || ""}
                                                            className="form-controltext readonly">
                                                            <option value="">Select</option>
                                                            <option value="Yes">Yes</option>
                                                            <option value="No">No</option>
                                                        </select>
                                                    </div>
                                                    <div className='col-md-4'>
                                                        <label className='font'>Document Number</label>
                                                        <input type="text" value={form10F.DocumentNumber || ""} className="form-control readonly" />
                                                    </div>
                                                    <div className='col-md-4'>
                                                        <label className='font'>Acknowledgment Number</label>
                                                        <input type="text" value={form10F.AcknowledgmentNumber || ""} className="form-control readonly" />
                                                    </div>
                                                </div>
                                                <div className='row mb-20'>
                                                    <div className="col-md-4">
                                                        <label className="font fontblock">Document Date</label>
                                                        <input type="date" value={form10F.DocumentDate || ""} className="form-control readonly" />
                                                    </div>
                                                    <div className='col-md-4'>
                                                        <label className='font'>Validity Start Date</label>
                                                        <input type="date" value={form10F.ValidityStartDate || ""} className="form-control readonly" />
                                                    </div>
                                                    <div className='col-md-4'>
                                                        <label className='font'>Validity End Date</label>
                                                        <input type="date" value={form10F.ValidityEndDate || ""} className="form-control readonly" />
                                                    </div>
                                                </div>
                                                <div className='row mb-20'>
                                                    <div className="col-md-4">
                                                        <label className="font fontblock">View Document</label>
                                                        <span><a href={form10F.Attachmenturl || "#"} target="_blank">{form10F.Attachmentfilename || "No Document Available"}</a></span>
                                                    </div>
                                                </div>
                                            </div>
                                        </>
                                    )}
                                </CollapsibleSection>

                                <CollapsibleSection title="Summary of WHT Applicability" style={{ marginTop: "10px" }}>
                                    <div className='main-formcontainer'>
                                        <div className='row mb-20'>
                                            <div className="col-md-4">
                                                <div className="date-summary">
                                                    <span className="label">From</span>
                                                    <span className="value">{fromdate}</span>
                                                    <span className="label">To</span>
                                                    <span className="value">{todate}</span>
                                                </div>
                                            </div>
                                        </div>
                                        <div className='row mb-20'>
                                            {paymentType !== "Service-Bill Payment" &&
                                                paymentType !== "Service-Advance Payment" && (
                                                    <div className="col-md-4">
                                                        <label className="font">Eligible amount that can be transmitted without WHT</label>
                                                        <input type="number" value={eligibleAmountWithWHT} onChange={(e) => setEligibleAmountWithWHT(e.target.value)}
                                                            className="form-control readonly"
                                                        />
                                                    </div>
                                                )}
                                            <div className="col-md-4">
                                                <label className="font">Paid Amount</label>
                                                <input type="number" value={paidAmount} onChange={(e) => setPaidAmount(e.target.value)}
                                                    className="form-control readonly"
                                                />
                                            </div>
                                            <div className="col-md-4">
                                                <label className="font">Balance eligible amount (Without withholding Tax)</label>
                                                <input type="number" value={ballenceEligibleAmount} onChange={(e) => setBallenceEligibleAmount(e.target.value)}
                                                    className="form-control readonly"
                                                />
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

                                {paymentType === "Goods-Bill Payment" && (
                                    <>
                                        <div className="heading1" style={{ marginTop: "10px" }}>
                                            <label>Forex Payment Request Details</label>
                                        </div>
                                        <div className='main-formcontainer'>
                                            <div className='row mb-20'>
                                                <div className='col-md-4'>
                                                    <Field label="Type" required>
                                                        <select value={paymentType} onChange={(e) => setPaymentType(e.target.value)} className="form-controltext" disabled>
                                                            <option value="Goods-Bill Payment">Goods-Bill Payment</option>
                                                            <option value="Service-Bill Payment">Service-Bill Payment</option>
                                                            <option value="Goods-Advance Payment">Goods-Advance Payment</option>
                                                            <option value="Service-Advance Payment">Service-Advance Payment</option>
                                                        </select>
                                                    </Field>
                                                </div>
                                            </div>
                                            <div className='row mb-20'>
                                                <div className='col-md-4'>
                                                    <label className='font'>Request Number </label>
                                                    <input type="text" value={requestNumber} className="form-control readonly" />
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className="font">Requested On </label>
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
                                            </div>
                                        </div>
                                        <div style={{ overflowX: "auto" }}>
                                            <table className="custom-table" style={{ marginTop: "10px" }}>
                                                <thead>
                                                    <tr>
                                                        <th>Sr.No</th>
                                                        <th>Invoice No</th>
                                                        <th>Invoice Date</th>
                                                        <th>BOE No</th>
                                                        <th>BOE Date</th>
                                                        <th>MRN No</th>
                                                        <th>Bill of Lading No</th>
                                                        <th>Bill of Lading Date</th>
                                                        <th>Invoice Amount</th>
                                                        <th>Attach Invoice</th>
                                                        <th>Attach Other Docs</th>
                                                        {/* <th>Add/Delete Entry</th> */}
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    {rows.map((row, index) => (
                                                        <tr key={index}>
                                                            <td>{index + 1}</td>

                                                            <td>
                                                                <input
                                                                    value={row.invoiceNo}
                                                                    onChange={(e) =>
                                                                        handleChange(index, "invoiceNo", e.target.value)
                                                                    }
                                                                    readOnly />
                                                            </td>

                                                            <td>
                                                                <input
                                                                    type="date"
                                                                    value={row.invoiceDate}
                                                                    onChange={(e) =>
                                                                        handleChange(index, "invoiceDate", e.target.value)
                                                                    }
                                                                    readOnly />
                                                            </td>

                                                            <td>
                                                                <input
                                                                    value={row.boeNo}
                                                                    onChange={(e) =>
                                                                        handleChange(index, "boeNo", e.target.value)
                                                                    }
                                                                    readOnly />
                                                            </td>

                                                            <td>
                                                                <input
                                                                    type="date"
                                                                    value={row.boeDate}
                                                                    onChange={(e) =>
                                                                        handleChange(index, "boeDate", e.target.value)
                                                                    }
                                                                    readOnly />
                                                            </td>

                                                            <td>
                                                                <input
                                                                    value={row.mrnNo}
                                                                    onChange={(e) =>
                                                                        handleChange(index, "mrnNo", e.target.value)
                                                                    }
                                                                    readOnly />
                                                            </td>

                                                            <td>
                                                                <input
                                                                    value={row.blNo}
                                                                    onChange={(e) =>
                                                                        handleChange(index, "blNo", e.target.value)
                                                                    }
                                                                    readOnly />
                                                            </td>

                                                            <td>
                                                                <input
                                                                    type="date"
                                                                    value={row.blDate}
                                                                    onChange={(e) =>
                                                                        handleChange(index, "blDate", e.target.value)
                                                                    }
                                                                    readOnly />
                                                            </td>

                                                            <td>
                                                                <input
                                                                    type="number"
                                                                    value={row.invoiceAmount}
                                                                    onChange={(e) =>
                                                                        handleChange(index, "invoiceAmount", e.target.value)
                                                                    }
                                                                    readOnly />
                                                            </td>

                                                            <td>
                                                                <div>
                                                                    {invoiceAttachments[index]?.map((file: any, i: number) => (
                                                                        <div key={i}>
                                                                            <a href={file.ServerRelativeUrl} target="_blank">
                                                                                {file.FileName}
                                                                            </a>
                                                                        </div>
                                                                    ))}

                                                                </div>
                                                            </td>

                                                            <td>
                                                                <div>
                                                                    {otherAttachments[index]?.map((file: any, i: number) => (
                                                                        <div key={i}>
                                                                            <a href={file.ServerRelativeUrl} target="_blank">
                                                                                {file.FileName}
                                                                            </a>
                                                                        </div>
                                                                    ))}
                                                                </div>
                                                            </td>

                                                            {/* <td style={{ textAlign: "center" }}>
                                            {index === rows.length - 1 && (
                                                <button
                                                    type="button"
                                                    onClick={addRow}
                                                    style={{
                                                        background: "#28a745",
                                                        color: "white",
                                                        marginRight: "5px",
                                                        border: "none",
                                                        padding: "5px 10px",
                                                        cursor: "pointer",
                                                        borderRadius: "4px"
                                                    }}
                                                >
                                                    +
                                                </button>
                                            )}

                                            {rows.length > 1 && (
                                                <button
                                                    type="button"
                                                    onClick={() => deleteRow(index)}
                                                    style={{
                                                        background: "#dc3545",
                                                        color: "white",
                                                        border: "none",
                                                        padding: "5px 10px",
                                                        cursor: "pointer",
                                                        borderRadius: "4px"
                                                    }}
                                                >
                                                    ✖
                                                </button>
                                            )}
                                        </td> */}

                                                        </tr>
                                                    ))}
                                                </tbody>
                                                {/* <tfoot>
                                <tr>
                                    <td colSpan={8} style={{ textAlign: "right", fontWeight: "bold" }}>
                                        Total Invoice Amount:
                                    </td>
                                    <td style={{ fontWeight: "bold" }}>
                                        {totalInvoiceAmount.toFixed(2)}
                                    </td>
                                    <td colSpan={3}></td>
                                </tr>
                            </tfoot> */}
                                            </table>
                                        </div>
                                        <div style={{ display: "flex", gap: "40px", marginTop: "30px" }}>

                                            {/* BOE TABLE */}
                                            <div style={{ flex: 1 }}>
                                                <p style={{ color: "red", fontSize: "13px" }}>
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
                                                                    <div>
                                                                        {boeLibraryFiles
                                                                            .filter(file => file.BOENo === boe)
                                                                            .map((file, i) => (
                                                                                <div key={i}>
                                                                                    <a href={file.FileRef} target="_blank">
                                                                                        {file.FileLeafRef}
                                                                                    </a>
                                                                                </div>
                                                                            ))}
                                                                    </div>                                                </td>
                                                            </tr>
                                                        ))}
                                                    </tbody>
                                                </table>
                                            </div>
                                            {/* BILL OF LADING TABLE */}
                                            <div style={{ flex: 1 }}>
                                                <p style={{ color: "red", fontSize: "13px" }}>
                                                    Unique Bill of lading no will be listed below
                                                </p>

                                                <table className="custom-table">
                                                    <thead>
                                                        <tr>
                                                            <th>Bill of Lading Number</th>
                                                            <th>Attach Documents</th>
                                                        </tr>
                                                    </thead>
                                                    <tbody>
                                                        {uniqueBlNumbers.map((bl, index) => (
                                                            <tr key={index}>
                                                                <td>{bl}</td>
                                                                <td>
                                                                    <div>
                                                                        {bolLibraryFiles
                                                                            .filter(file => file.BOLNo === bl)
                                                                            .map((file, i) => (
                                                                                <div key={i}>
                                                                                    <a href={file.FileRef} target="_blank">
                                                                                        {file.FileLeafRef}
                                                                                    </a>
                                                                                </div>
                                                                            ))}
                                                                    </div>                                                </td>
                                                            </tr>
                                                        ))}
                                                    </tbody>
                                                </table>
                                            </div>

                                        </div>
                                    </>
                                )}

                                {paymentType === "Service-Bill Payment" && (
                                    <>
                                        <div className="heading1" style={{ marginTop: "10px" }}>
                                            <label>Forex Payment Request Details</label>
                                        </div>
                                        <div className='main-formcontainer'>
                                            <div className='row mb-20'>
                                                <div className='col-md-4'>
                                                    <Field label="Type" required>
                                                        <select value={paymentType} onChange={(e) => setPaymentType(e.target.value)} className="form-controltext" disabled>
                                                            <option value="Goods-Bill Payment">Goods-Bill Payment</option>
                                                            <option value="Service-Bill Payment">Service-Bill Payment</option>
                                                            <option value="Goods-Advance Payment">Goods-Advance Payment</option>
                                                            <option value="Service-Advance Payment">Service-Advance Payment</option>
                                                        </select>
                                                    </Field>
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
                                            </div>
                                        </div>
                                        <div style={{ overflowX: "auto" }}>
                                            <table className="custom-table" style={{ marginTop: "10px" }}>
                                                <thead>
                                                    <tr>
                                                        <th>Sr No</th>
                                                        <th>Invoice No</th>
                                                        <th>Invoice Date</th>
                                                        <th>Invoice Amount</th>
                                                        <th>MRN No</th>
                                                        <th>MRN Date</th>
                                                        <th>Attach Invoice</th>
                                                        <th>Attach Other Document</th>
                                                        {/* <th>Add/Delete Entry</th> */}
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    {rows.map((row, index) => (
                                                        <tr key={index}>
                                                            <td>{index + 1}</td>

                                                            <td>
                                                                <input
                                                                    value={row.invoiceNo}
                                                                    onChange={(e) =>
                                                                        handleChange(index, "invoiceNo", e.target.value)
                                                                    }
                                                                    readOnly />
                                                            </td>

                                                            <td>
                                                                <input
                                                                    type="date"
                                                                    value={row.invoiceDate}
                                                                    onChange={(e) =>
                                                                        handleChange(index, "invoiceDate", e.target.value)
                                                                    }
                                                                    readOnly />
                                                            </td>

                                                            <td>
                                                                <input
                                                                    value={row.invoiceAmount}
                                                                    onChange={(e) =>
                                                                        handleChange(index, "invoiceAmount", e.target.value)
                                                                    }
                                                                    readOnly />
                                                            </td>

                                                            <td>
                                                                <input
                                                                    value={row.mrnNo}
                                                                    onChange={(e) =>
                                                                        handleChange(index, "mrnNo", e.target.value)
                                                                    }
                                                                    readOnly />
                                                            </td>
                                                            <td>
                                                                <input
                                                                    type="date"
                                                                    value={row.mrnDate}
                                                                    onChange={(e) =>
                                                                        handleChange(index, "mrnDate", e.target.value)
                                                                    }
                                                                    readOnly />
                                                            </td>
                                                            <td>
                                                                {invoiceAttachments[index]?.map((file: any, i: number) => (
                                                                    <div key={i}>
                                                                        <a href={file.ServerRelativeUrl} target="_blank">
                                                                            {file.FileName}
                                                                        </a>
                                                                    </div>
                                                                ))}
                                                            </td>

                                                            <td>
                                                                {otherAttachments[index]?.map((file: any, i: number) => (
                                                                    <div key={i}>
                                                                        <a href={file.ServerRelativeUrl} target="_blank">
                                                                            {file.FileName}
                                                                        </a>
                                                                    </div>
                                                                ))}
                                                            </td>




                                                            {/* <td style={{ textAlign: "center" }}>
                                            {index === rows.length - 1 && (
                                                <button
                                                    type="button"
                                                    onClick={addRow}
                                                    style={{
                                                        background: "#28a745",
                                                        color: "white",
                                                        marginRight: "5px",
                                                        border: "none",
                                                        padding: "5px 10px",
                                                        cursor: "pointer",
                                                        borderRadius: "4px"
                                                    }}
                                                >
                                                    +
                                                </button>
                                            )}

                                            {rows.length > 1 && (
                                                <button
                                                    type="button"
                                                    onClick={() => deleteRow(index)}
                                                    style={{
                                                        background: "#dc3545",
                                                        color: "white",
                                                        border: "none",
                                                        padding: "5px 10px",
                                                        cursor: "pointer",
                                                        borderRadius: "4px"
                                                    }}
                                                >
                                                    ✖
                                                </button>
                                            )}
                                        </td> */}

                                                        </tr>
                                                    ))}
                                                </tbody>
                                                {/* <tfoot>
                                <tr>
                                    <td colSpan={3} style={{ textAlign: "right", fontWeight: "bold" }}>
                                        Total Invoice Amount:
                                    </td>
                                    <td style={{ fontWeight: "bold" }}>
                                        {totalInvoiceAmount.toFixed(2)}
                                    </td>
                                    <td colSpan={5}></td>
                                </tr>
                            </tfoot> */}
                                            </table>
                                        </div>
                                    </>
                                )}

                                {paymentType === "Service-Advance Payment" && (
                                    <>
                                        <div className="heading1" style={{ marginTop: "10px" }}>
                                            <label>Forex Payment Request Details</label>
                                        </div>
                                        <div className='main-formcontainer'>
                                            <div className='row mb-20'>
                                                <div className='col-md-4'>
                                                    <Field label="Type" required>
                                                        <select value={paymentType} onChange={(e) => setPaymentType(e.target.value)} className="form-controltext" disabled>
                                                            <option value="Goods-Bill Payment">Goods-Bill Payment</option>
                                                            <option value="Service-Bill Payment">Service-Bill Payment</option>
                                                            <option value="Goods-Advance Payment">Goods-Advance Payment</option>
                                                            <option value="Service-Advance Payment">Service-Advance Payment</option>
                                                        </select>
                                                    </Field>
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
                                                        <th>Sr No</th>
                                                        <th>Performa Invoice No</th>
                                                        <th>Performa Invoice Date</th>
                                                        <th>Performa Invoice Amount</th>
                                                        <th>Attach PO</th>
                                                        <th>Attach PI</th>
                                                        <th>Attach Other Document</th>
                                                        {/* <th>Add/Delete Entry</th> */}
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    {rows.map((row, index) => (
                                                        <tr key={index}>
                                                            <td>{index + 1}</td>

                                                            <td>
                                                                <input
                                                                    value={row.invoiceNo}
                                                                    onChange={(e) =>
                                                                        handleChange(index, "invoiceNo", e.target.value)
                                                                    }
                                                                    readOnly />
                                                            </td>

                                                            <td>
                                                                <input
                                                                    type="date"
                                                                    value={row.invoiceDate}
                                                                    onChange={(e) =>
                                                                        handleChange(index, "invoiceDate", e.target.value)
                                                                    }
                                                                    readOnly />
                                                            </td>

                                                            <td>
                                                                <input
                                                                    value={row.invoiceAmount}
                                                                    onChange={(e) =>
                                                                        handleChange(index, "invoiceAmount", e.target.value)
                                                                    }
                                                                    readOnly />
                                                            </td>


                                                            <td>
                                                                <div>
                                                                    {poAttachments[index]?.map((file: any, i: number) => (
                                                                        <div key={i}>
                                                                            <a href={file.ServerRelativeUrl} target="_blank">
                                                                                {file.FileName}
                                                                            </a>
                                                                        </div>
                                                                    ))}

                                                                </div>
                                                            </td>
                                                            <td>
                                                                <div>
                                                                    {piAttachments[index]?.map((file: any, i: number) => (
                                                                        <div key={i}>
                                                                            <a href={file.ServerRelativeUrl} target="_blank">
                                                                                {file.FileName}
                                                                            </a>
                                                                        </div>
                                                                    ))}

                                                                </div>
                                                            </td>

                                                            <td>
                                                                {otherAttachments[index]?.map((file: any, i: number) => (
                                                                    <div key={i}>
                                                                        <a href={file.ServerRelativeUrl} target="_blank">
                                                                            {file.FileName}
                                                                        </a>
                                                                    </div>
                                                                ))}
                                                            </td>




                                                            {/* <td style={{ textAlign: "center" }}>
                                            {index === rows.length - 1 && (
                                                <button
                                                    type="button"
                                                    onClick={addRow}
                                                    style={{
                                                        background: "#28a745",
                                                        color: "white",
                                                        marginRight: "5px",
                                                        border: "none",
                                                        padding: "5px 10px",
                                                        cursor: "pointer",
                                                        borderRadius: "4px"
                                                    }}
                                                >
                                                    +
                                                </button>
                                            )}

                                            {rows.length > 1 && (
                                                <button
                                                    type="button"
                                                    onClick={() => deleteRow(index)}
                                                    style={{
                                                        background: "#dc3545",
                                                        color: "white",
                                                        border: "none",
                                                        padding: "5px 10px",
                                                        cursor: "pointer",
                                                        borderRadius: "4px"
                                                    }}
                                                >
                                                    ✖
                                                </button>
                                            )}
                                        </td> */}

                                                        </tr>
                                                    ))}
                                                </tbody>
                                                {/* <tfoot>
                                <tr>
                                    <td colSpan={3} style={{ textAlign: "right", fontWeight: "bold" }}>
                                        Total Invoice Amount:
                                    </td>
                                    <td style={{ fontWeight: "bold" }}>
                                        {totalInvoiceAmount.toFixed(2)}
                                    </td>
                                    <td colSpan={4}></td>
                                </tr>
                            </tfoot> */}
                                            </table>
                                        </div>
                                    </>
                                )}

                                {paymentType === "Goods-Advance Payment" && (
                                    <>
                                        <div className="heading1" style={{ marginTop: "10px" }}>
                                            <label>Forex Payment Request Details</label>
                                        </div>
                                        <div className='main-formcontainer'>
                                            <div className='row mb-20'>
                                                <div className='col-md-4'>
                                                    <Field label="Type" required>
                                                        <select value={paymentType} onChange={(e) => setPaymentType(e.target.value)} className="form-controltext" disabled>
                                                            <option value="Goods-Bill Payment">Goods-Bill Payment</option>
                                                            <option value="Service-Bill Payment">Service-Bill Payment</option>
                                                            <option value="Goods-Advance Payment">Goods-Advance Payment</option>
                                                            <option value="Service-Advance Payment">Service-Advance Payment</option>
                                                        </select>
                                                    </Field>
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
                                                        <th>Sr No</th>
                                                        <th>Performa Invoice No</th>
                                                        <th>Performa Invoice Date</th>
                                                        <th>Performa Invoice Amount</th>
                                                        <th>Attach PO</th>
                                                        <th>Attach PI</th>
                                                        <th>Attach Other Document</th>
                                                        {/* <th>Add/Delete Entry</th> */}
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    {rows.map((row, index) => (
                                                        <tr key={index}>
                                                            <td>{index + 1}</td>

                                                            <td>
                                                                <input
                                                                    value={row.invoiceNo}
                                                                    onChange={(e) =>
                                                                        handleChange(index, "invoiceNo", e.target.value)
                                                                    }
                                                                />
                                                            </td>

                                                            <td>
                                                                <input
                                                                    type="date"
                                                                    value={row.invoiceDate}
                                                                    onChange={(e) =>
                                                                        handleChange(index, "invoiceDate", e.target.value)
                                                                    }
                                                                />
                                                            </td>

                                                            <td>
                                                                <input
                                                                    value={row.invoiceAmount}
                                                                    onChange={(e) =>
                                                                        handleChange(index, "invoiceAmount", e.target.value)
                                                                    }
                                                                />
                                                            </td>


                                                            <td>
                                                                <div>
                                                                    {poAttachments[index]?.map((file: any, i: number) => (
                                                                        <div key={i}>
                                                                            <a href={file.ServerRelativeUrl} target="_blank">
                                                                                {file.FileName}
                                                                            </a>
                                                                        </div>
                                                                    ))}
                                                                </div>
                                                            </td>
                                                            <td>
                                                                <div>
                                                                    {piAttachments[index]?.map((file: any, i: number) => (
                                                                        <div key={i}>
                                                                            <a href={file.ServerRelativeUrl} target="_blank">
                                                                                {file.FileName}
                                                                            </a>
                                                                        </div>
                                                                    ))}

                                                                </div>
                                                            </td>

                                                            <td>
                                                                {otherAttachments[index]?.map((file: any, i: number) => (
                                                                    <div key={i}>
                                                                        <a href={file.ServerRelativeUrl} target="_blank">
                                                                            {file.FileName}
                                                                        </a>
                                                                    </div>
                                                                ))}
                                                            </td>




                                                            {/* <td style={{ textAlign: "center" }}>
                                            {index === rows.length - 1 && (
                                                <button
                                                    type="button"
                                                    onClick={addRow}
                                                    style={{
                                                        background: "#28a745",
                                                        color: "white",
                                                        marginRight: "5px",
                                                        border: "none",
                                                        padding: "5px 10px",
                                                        cursor: "pointer",
                                                        borderRadius: "4px"
                                                    }}
                                                >
                                                    +
                                                </button>
                                            )}

                                            {rows.length > 1 && (
                                                <button
                                                    type="button"
                                                    onClick={() => deleteRow(index)}
                                                    style={{
                                                        background: "#dc3545",
                                                        color: "white",
                                                        border: "none",
                                                        padding: "5px 10px",
                                                        cursor: "pointer",
                                                        borderRadius: "4px"
                                                    }}
                                                >
                                                    ✖
                                                </button>
                                            )}
                                        </td> */}

                                                        </tr>
                                                    ))}
                                                </tbody>
                                                {/* <tfoot>
                                <tr>
                                    <td colSpan={3} style={{ textAlign: "right", fontWeight: "bold" }}>
                                        Total Invoice Amount:
                                    </td>
                                    <td style={{ fontWeight: "bold" }}>
                                        {totalInvoiceAmount.toFixed(2)}
                                    </td>
                                    <td colSpan={4}></td>
                                </tr>
                            </tfoot> */}
                                            </table>
                                        </div>
                                    </>
                                )}

                                {(paymentType === "Goods-Advance Payment" && hasAdvanceData) && (
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

                                                    {rowsAdvance.map((row, index) => (
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

                                                                {poAttachmentsTrackeradvance[index]?.length > 0 ? (

                                                                    poAttachmentsTrackeradvance[index].map((file: any, i: number) => (

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

                                                                {piAttachmentsTrackeradvance[index]?.length > 0 ? (

                                                                    piAttachmentsTrackeradvance[index].map((file: any, i: number) => (

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

                                                {/* <tfoot>
                                    <tr>
                                        <td colSpan={7}></td>
                                        <td style={{ fontWeight: "bold" }}>Total Amount</td>
                                        <td>{totalInvoiceAmount.toFixed(2)}</td>
                                        <td colSpan={2}></td>
                                    </tr>
                                </tfoot> */}

                                            </table>
                                        </div>
                                        <div style={{ display: "flex", gap: "40px", marginTop: "20px" }}>

                                            {/* BOE TABLE */}

                                            <div>

                                                {/* <p style={{ color: "red", fontSize: "12px" }}>
                                    Unique BOE no listed below
                                </p> */}

                                                <table className="custom-table">

                                                    <thead>
                                                        <tr>
                                                            <th>BOE No</th>
                                                            <th>Attach Documents</th>
                                                        </tr>
                                                    </thead>

                                                    <tbody>

                                                        {uniqueBoeNumbersadvancetracker.map((boe, index) => (
                                                            <tr key={index}>

                                                                <td>{boe}</td>

                                                                <td>

                                                                    {boeAttachmentstrackeradvance[index]?.length > 0 ? (

                                                                        boeAttachmentstrackeradvance[index].map((file: any, i: number) => (
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
                                                {/* 
                                <p style={{ color: "red", fontSize: "12px" }}>
                                    Unique Bill of Lading listed below
                                </p> */}

                                                <table className="custom-table">

                                                    <thead>
                                                        <tr>
                                                            <th>Bill of Lading</th>
                                                            <th>Attach Documents</th>
                                                        </tr>
                                                    </thead>

                                                    <tbody>

                                                        {uniqueBlNumbersadvancetracker.map((bl, index) => (
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
                                                                    {blAttachmentstrackeradvance[index]?.length > 0 ? (

                                                                        blAttachmentstrackeradvance[index].map((file: any, i: number) => (
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

                                {(paymentType === "Service-Advance Payment" && hasAdvanceData) && (
                                    <>
                                        <p>
                                            <b>Bill Payment Details (for Service Bill Payment)</b>
                                        </p>
                                        <div style={{ overflowX: "auto" }}>
                                            <table className="data-table">

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

                                                    {rowsAdvance.map((row, index) => (
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

                                                                {poAttachmentsTrackeradvance[index]?.length > 0 ? (

                                                                    poAttachmentsTrackeradvance[index].map((file: any, i: number) => (

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

                                                                {piAttachmentsTrackeradvance[index]?.length > 0 ? (

                                                                    piAttachmentsTrackeradvance[index].map((file: any, i: number) => (

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
                                                {/* <tfoot>
                                <tr>
                                    <td colSpan={2}></td>
                                    <td style={{ fontWeight: "bold" }}>Total Amount</td>
                                    <td>{totalInvoiceAmount.toFixed(2)}</td>
                                    <td colSpan={5}></td>
                                </tr>
                            </tfoot> */}
                                            </table>
                                        </div>
                                    </>
                                )}
                                {validationDateshow !== "" && (
                                    <>
                                        <div className="heading1" style={{ marginTop: "10px" }}>
                                            <label>Vouching Details</label>
                                        </div>

                                        <div className="main-formcontainer">
                                            <div className="row mb-20">
                                                <div className="col-md-4">
                                                    <label className="font">Validation Date</label>
                                                    <input
                                                        type="date"
                                                        value={validationDateshow}
                                                        className="form-control readonly"
                                                        readOnly
                                                    />
                                                </div>

                                                {paymentType.includes("Advance") && voucherNumbershow && (
                                                    <div className="col-md-4">
                                                        <label className="font">Voucher Number</label>
                                                        <input
                                                            type="text"
                                                            value={voucherNumbershow}
                                                            className="form-control readonly"
                                                            readOnly
                                                        />
                                                    </div>
                                                )}
                                            </div>
                                        </div>
                                    </>
                                )}
                                {hasPaymentData && (
                                    <>
                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Payment Details</label>
                                </div>
                                <div className='main-formcontainer'>
                                    <div className='row mb-20'>
                                        <div className='col-md-4'>
                                            <label className='font'>Foreign Currency</label>
                                            <input type="text" value={foreignCurrency} className="form-control readonly" />
                                        </div>
                                        <div className='col-md-4'>
                                            <label className="font">Foreign Currency Amount</label>
                                            <input type="text" value={foreignAmountpayment} className="form-control readonly" />
                                        </div>
                                        <div className='col-md-4'>
                                            <label className="font">Exchange Rate</label>
                                            <input type="number" value={exchangeRatepayment} className="form-control readonly" />
                                        </div>
                                    </div>
                                    <div className='row mb-20'>
                                        <div className='col-md-4'>
                                            <label className='font'>INR Amount</label>
                                            <input type="text" value={inrAmountpayment} className="form-control readonly" />
                                        </div>
                                        <div className='col-md-4'>
                                            <label className='font'>Payment Date</label>
                                            <input type="date" value={paymentDatepayment} className="form-control readonly" />
                                        </div>
                                        <div className='col-md-4'>
                                            <label className="font">Payment reference number</label>
                                            <input type="number" value={paymentReferencepayment} className="form-control readonly" />
                                        </div>
                                    </div>
                                    <div className="row mb-20">

                                    </div>
                                    <div className='row mb-20'>

                                        {isServicePayment && (
                                            <div className='col-md-4'>
                                                <label className="font">Form145</label>

                                                {form15CA.length > 0 && (
                                                    <div style={{ marginTop: "10px" }}>
                                                        {form15CA.map((file: any, index: number) => (
                                                            <div
                                                                key={index}
                                                                style={{
                                                                    display: "flex",
                                                                    justifyContent: "space-between",
                                                                    alignItems: "center",
                                                                    marginBottom: "5px",
                                                                    padding: "6px 10px",
                                                                    border: "1px solid #ddd",
                                                                    borderRadius: "4px",
                                                                }}
                                                            >
                                                                <a
                                                                    href={file.ServerRelativeUrl}
                                                                    target="_blank"
                                                                    rel="noreferrer"
                                                                >
                                                                    {file.FileName}
                                                                </a>
                                                            </div>
                                                        ))}
                                                    </div>
                                                )}
                                            </div>
                                        )}
                                        {isServicePayment && (
                                            <div className='col-md-4'>
                                                <label className="font">Form146</label>

                                                {form15CB.length > 0 && (
                                                    <div style={{ marginTop: "10px" }}>
                                                        {form15CB.map((file: any, index: number) => (
                                                            <div
                                                                key={index}
                                                                style={{
                                                                    display: "flex",
                                                                    justifyContent: "space-between",
                                                                    alignItems: "center",
                                                                    marginBottom: "5px",
                                                                    padding: "6px 10px",
                                                                    border: "1px solid #ddd",
                                                                    borderRadius: "4px",
                                                                }}
                                                            >
                                                                <a
                                                                    href={file.ServerRelativeUrl}
                                                                    target="_blank"
                                                                    rel="noreferrer"
                                                                >
                                                                    {file.FileName}
                                                                </a>
                                                            </div>
                                                        ))}
                                                    </div>
                                                )}
                                            </div>
                                        )}
                                        {isGoodsPayment && (
                                            <div className="col-md-4">
                                                <label className="font">
                                                    Attach Swift Copy (Applicable for Goods)
                                                </label>

                                                {swiftCopy.length > 0 && (
                                                    <div style={{ marginTop: "10px" }}>
                                                        {swiftCopy.map((file: any, index: number) => (
                                                            <div
                                                                key={index}
                                                                style={{
                                                                    display: "flex",
                                                                    justifyContent: "space-between",
                                                                    alignItems: "center",
                                                                    marginBottom: "5px",
                                                                    padding: "6px 10px",
                                                                    border: "1px solid #ddd",
                                                                    borderRadius: "4px",
                                                                }}
                                                            >
                                                                <a
                                                                    href={file.ServerRelativeUrl}
                                                                    target="_blank"
                                                                    rel="noreferrer"
                                                                >
                                                                    {file.FileName}
                                                                </a>
                                                            </div>
                                                        ))}
                                                    </div>
                                                )}
                                            </div>
                                        )}
                                    </div>
                                    <div className='row mb-20'>

                                    </div>
                                </div>
                                </>
                                )}



                                <div className='row my-3'>
                                    <div className='col-md-12'>
                                        <div style={{ display: "flex", justifyContent: "center", gap: "5px" }}>
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
            {/* {showVendorPopup && (
                <div className="popup-overlay">
                    <div className="popup-box">
                        <h3>Vendor Not Found</h3>
                        <p>Please add vendor first.</p>

                        <div style={{ marginTop: "20px", display: "flex", gap: "10px", justifyContent: "center" }}>
                            <button
                                className="btn-submit"
                                onClick={() => {
                                    setShowVendorPopup(false);
                                    history.push("/vendor-creation");
                                }}
                            >
                                Go to Vendor Creation
                            </button>

                            {/* <button
                                className="btn-exit"
                                onClick={() => setShowVendorPopup(false)}
                            >
                                Cancel
                            </button>
                        </div>
                    </div>
                </div>
            )} */}


        </>
    );
};

export default ViewRequestForm;


