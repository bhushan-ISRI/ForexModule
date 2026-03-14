import React, { useEffect, useState } from "react";
import "../Pages/Css/NewRequest.scss";
import { IForexModuleProps } from "../IForexModuleProps";
import { useHistory, useParams } from 'react-router-dom';
import { Dropdown, IDropdownOption } from '@fluentui/react';
import SPCRUDOPS from "../../service/BAL/spcrud";
import { Attachment } from "@pnp/sp/attachments";
import { View } from "@pnp/sp/views";
//import USESPCRUD from "../../service/BAL/spcrud";

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
        VendorAddress: ""
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
        }
    }, [Id])
    //----------------------- load data for edit------------------------//
    const loadForexData = async (forexId: string) => {
        const sp = await spCrudOps;

        try {

            // 🔹 Load Parent
            const parent = await sp.getData(
                "ForexRequest",
                "*,AllApprovers,RM/Title,HOD/Title,Author/Id,Currency/Title,Currency/Id,CurrentApprover/Id,CurrentApprover/Title,NextApprovers/Id,NextApprovers/Title,PrevApprovers/Id,PrevApprovers/Title",
                "RM,HOD,Author,Currency,CurrentApprover,NextApprovers,PrevApprovers",
                `ID eq ${forexId}`,
                { column: "ID", isAscending: true },
                1,
                props
            );

            if (parent.length > 0) {
                const data = parent[0];

                setPaymentType(data.ForexType || "");
                setRequestNumber(data.ForexNumber || "");
                setRequestedOn(data.RequestedOn?.split("T")[0] || "");
                setCurrency(data.Currency.Title || "");
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
                setPaymentDate(data.PaymentDate?.split("T")[0] || "");
                setPaymentReferenceNumber(data.PaymentReferenceNumber || "");

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

                        if (a.Status === "Approved") {
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


    //----------------------VendorData-------------------------//
    const getVendorData = async (vendorCode: string) => {

        if (!vendorCode) return;

        (await spCrudOps).getData(
            "VendorMaster",
            "VendorCode,VendorName,VendorNameLegal,VendorShortName,VendorType,City/Title,State/Title,Country/Title,Currency/Title,PostalCode,ContactPersonName,EmailId,PhoneNumber,AlternateContact,BeneficiaryName,BankName,AccountNumberIBAN,SWIFTBICCode,RoutingNumberABA,IFSCCode,IntermediaryBank,IntermediarySWIFTCode,NatureOfPayment/Title,PurposeCodeRBI,BankCountry,BankAddress,VendorAddress",
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
                        City: v.City?.Title || "",
                        State: v.State?.Title || "",
                        Country: v.Country?.Title || "",
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
                        VendorAddress: v.VendorAddress || ""
                    });
                    getTaxDeclarationdata(v.VendorCode);

                } else {
                    //alert("Vendor not found");
                    setShowVendorPopup(true);
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


    return (
        <div className="forex-wrapper">

            {/* ================= HEADER ================= */}
            <div className="forex-header">
                <h2>Forex Payment View Form</h2>
            </div>

            <div className="forex-card">
                {/* =============================Approvers================== */}
                <Section title="Approval Flow">
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
                </Section>

                {/* ================= REQUESTOR ================= */}
                <Section title="Requestor Information">
                    <Grid>
                        <Field label="Type">
                            <select value={paymentType} onChange={(e) => setPaymentType(e.target.value)}>
                                <option value="Goods-Bill Payment">Goods-Bill Payment</option>
                                <option value="Service-Bill Payment">Service-Bill Payment</option>
                                <option value="Goods-Advance Payment">Goods-Advance Payment</option>
                                <option value="Service-Advance Payment">Service-Advance Payment</option>
                            </select>
                        </Field>
                    </Grid>
                    <Grid style={{ marginTop: "20px" }}>
                        <Field label="Employee Code">
                            <input type="text" value={employee.EmployeeCode} readOnly />
                        </Field>

                        <Field label="Employee Name">
                            <input type="text" value={employee.EmployeeName} readOnly />
                        </Field>

                        <Field label="Division">
                            <input type="text" value={employee.Division} readOnly />
                        </Field>

                        <Field label="Location">
                            <input type="text" value={employee.Location} readOnly />
                        </Field>

                        <Field label="RM">
                            <input type="text" value={employee.RM} readOnly />
                        </Field>

                        <Field label="HOD">
                            <input type="text" value={employee.HOD} readOnly />
                        </Field>

                        <Field label="Contact No">
                            <input type="text" value={employee.ContactNo} readOnly />
                        </Field>

                        <Field label="Employee Status">
                            <input type="text" value={employee.EmployeeStatus} readOnly />
                        </Field>

                        <Field label="Email" full>
                            <input type="email" value={employee.Email} readOnly />
                        </Field>
                    </Grid>

                </Section>

                {/* ================= VENDOR ================= */}
                <Section title="Vendor / Beneficiary Details">
                    <Grid>
                        <Field label="Vendor Code">
                            <input
                                value={vendor.VendorCode}
                                onChange={(e) => {
                                    const code = e.target.value;
                                    setVendor({ ...vendor, VendorCode: code });
                                }}
                                onBlur={(e) => getVendorData(e.target.value)}   // fetch when user leaves field
                            />
                        </Field>
                        <Field label="Vendor Name">
                            <input value={vendor.VendorName} readOnly />
                        </Field>
                        <Field label="Address" full><input value={vendor.VendorAddress} readOnly /></Field>
                        <Field label="City">
                            <input value={vendor.City} readOnly />
                        </Field>

                        <Field label="Country">
                            <input value={vendor.Country} readOnly />
                        </Field>
                        <Field label="Pincode"><input value={vendor.PostalCode} readOnly /></Field>
                        <Field label="Bank Name">
                            <input value={vendor.BankName} readOnly />
                        </Field>
                        <Field label="Bank Country"><input value={vendor.BankCountry} readOnly /></Field>
                        <Field label="Bank Swift Code"><input value={vendor.SWIFTBICCode} readOnly /></Field>
                        <Field label="Bank Branch Address"><input value={vendor.BankAddress} readOnly /></Field>
                        <Field label="Bank IBAN / Account No" full><input value={vendor.AccountNumberIBAN} readOnly /></Field>
                    </Grid>
                </Section>

                {/* ================= TAX INFO ================= */}
                <Section title="Tax & Regulatory Information">
                    <Grid>
                        <Field label="Nature of Payment"><input value={paymentType} readOnly /></Field>
                        <Field label="Tax Document Available?">
                            <select onChange={(e) => { setTaxDocumentView(e.target.value) }}>
                                <option>Yes</option>
                                <option>No</option>
                            </select>
                        </Field>
                        {taxDocumentView === "No" && (
                            <Field >
                                <span style={{ color: "red" }}>
                                    (if No, withholding tax will be applicable)
                                </span>
                            </Field>
                        )}

                        {taxDocumentView === "Yes" && (
                            <Field label="DTAA Applicable?">
                                <select value={dTAAApplicable} onChange={(e) => setDTAAApplicable(e.target.value)}>
                                    <option value="">Select</option>
                                    <option value="Yes">Yes</option>
                                    <option value="No">No</option>
                                </select>
                            </Field>
                        )}
                    </Grid>
                </Section>
                {taxDocumentView === "Yes" && (
                    <>
                        {/* 🔹 Permanent Establishment Declaration */}
                        <Section title="Permanent Establishment Declaration">
                            <Grid>

                                <Field label="Document Available">
                                    <select
                                        value={permanentEstablishmentDeclaration.DocumentAvailable || ""}
                                        disabled
                                    >
                                        <option value="">Select</option>
                                        <option value="Yes">Yes</option>
                                        <option value="No">No</option>
                                    </select>
                                </Field>

                                <Field label="Document Number">
                                    <input
                                        value={permanentEstablishmentDeclaration.DocumentNumber || ""}
                                        readOnly
                                    />
                                </Field>

                                <Field label="Document Date">
                                    <input
                                        type="date"
                                        value={permanentEstablishmentDeclaration.DocumentDate || ""}
                                        readOnly
                                    />
                                </Field>

                                <Field label="Validity Start Date">
                                    <input
                                        type="date"
                                        value={permanentEstablishmentDeclaration.ValidityStartDate || ""}
                                        readOnly
                                    />
                                </Field>

                                <Field label="Validity End Date">
                                    <input
                                        type="date"
                                        value={permanentEstablishmentDeclaration.ValidityEndDate || ""}
                                        readOnly
                                    />
                                </Field>

                                <Field label="View Document">
                                    <span><a href={permanentEstablishmentDeclaration.Attachmenturl || "#"} target="_blank">{permanentEstablishmentDeclaration.Attachmentfilename || "No Document Available"}</a></span>
                                </Field>

                            </Grid>
                        </Section>

                        {/* 🔹 Tax Residency Certificate */}
                        <Section title="Tax Residency Certificate">
                            <Grid>

                                <Field label="Document Available">
                                    <select
                                        value={taxResidencyCertificate.DocumentAvailable || ""}
                                        disabled
                                    >
                                        <option value="">Select</option>
                                        <option value="Yes">Yes</option>
                                        <option value="No">No</option>
                                    </select>
                                </Field>

                                <Field label="Document Number">
                                    <input
                                        value={taxResidencyCertificate.DocumentNumber || ""}
                                        readOnly
                                    />
                                </Field>

                                <Field label="Country of Tax Residence">
                                    <input
                                        value={taxResidencyCertificate.CountryOfTaxResidence || ""}
                                        readOnly
                                    />
                                </Field>

                                <Field label="Tax Identification Number">
                                    <input
                                        value={taxResidencyCertificate.TaxIdentificationNumber || ""}
                                        readOnly
                                    />
                                </Field>

                                <Field label="Validity Start Date">
                                    <input
                                        type="date"
                                        value={taxResidencyCertificate.ValidityStartDate || ""}
                                        readOnly
                                    />
                                </Field>

                                <Field label="Validity End Date">
                                    <input
                                        type="date"
                                        value={taxResidencyCertificate.ValidityEndDate || ""}
                                        readOnly
                                    />
                                </Field>

                                <Field label="View Document">
                                    <span><a href={taxResidencyCertificate.Attachmenturl || "#"} target="_blank">{taxResidencyCertificate.Attachmentfilename || "No Document Available"}</a></span>
                                </Field>

                            </Grid>
                        </Section>

                        {/* 🔹 Form 10F */}
                        <Section title="Form 10F">
                            <Grid>

                                <Field label="Document Available">
                                    <select
                                        value={form10F.DocumentAvailable || ""}
                                        disabled
                                    >
                                        <option value="">Select</option>
                                        <option value="Yes">Yes</option>
                                        <option value="No">No</option>
                                    </select>
                                </Field>

                                <Field label="Document Number">
                                    <input
                                        value={form10F.DocumentNumber || ""}
                                        readOnly
                                    />
                                </Field>

                                <Field label="Acknowledgment Number">
                                    <input
                                        value={form10F.AcknowledgmentNumber || ""}
                                        readOnly
                                    />
                                </Field>

                                <Field label="Document Date">
                                    <input
                                        type="date"
                                        value={form10F.DocumentDate || ""}
                                        readOnly
                                    />
                                </Field>

                                <Field label="Validity Start Date">
                                    <input
                                        type="date"
                                        value={form10F.ValidityStartDate || ""}
                                        readOnly
                                    />
                                </Field>

                                <Field label="Validity End Date">
                                    <input
                                        type="date"
                                        value={form10F.ValidityEndDate || ""}
                                        readOnly
                                    />
                                </Field>

                                <Field label="View Document">
                                    <span><a href={form10F.Attachmenturl || "#"} target="_blank">{form10F.Attachmentfilename || "No Document Available"}</a></span>
                                </Field>

                            </Grid>
                        </Section>
                    </>
                )}

                <Section >
                    <Grid>
                        <Field label="From"><input type="date" value={fromdate} /></Field>

                        <Field label="To"><input type="date" value={todate} /></Field>
                    </Grid>
                    <Grid>

                        <Field label="Eligible amount that can be transmitted without WHT">
                            <input type="number" value={eligibleAmountWithWHT || ""} readOnly />
                        </Field>


                        <Field label="Paid Amount">
                            <input type="number" value={paidAmount || ""} readOnly />
                        </Field>

                        <Field label="Balance eligible amount(Without with holding Tax)">
                            <input type="number" value={ballenceEligibleAmount || ""} readOnly />
                        </Field>
                    </Grid>

                </Section>

                {/* ================= FOREX DETAILS ================= */}
                {/* ==============================Goods-Bill Payment============================== */}
                {paymentType === "Goods-Bill Payment" && (
                    <Section title="Forex Payment Request Details">
                        <Grid>
                            <Field label="Request Number"><input value={requestNumber} onChange={(e) => { setRequestNumber(e.target.value) }} /></Field>
                            <Field label="Requested On"><input type="date" value={requestedOn} onChange={(e) => { setRequestedOn(e.target.value) }} /></Field>
                            <Field label="Currency"><input value={currency} onChange={(e) => { setCurrency(e.target.value) }} /></Field>
                            <Field label="Total Amount"><input type="number" value={totalAmount} onChange={(e) => { setTotalAmount(e.target.value) }} /></Field>
                            <Field label="Foreign Bank Charges"><input type="number" value={foreignBankCharges} onChange={(e) => { setForeignBankCharges(e.target.value) }} /></Field>
                            {/* <Field label="PO/Contract No"><input /></Field>
                            <Field label="PO Date"><input type="date" /></Field>
                            <Field label="Expected Settlement Date"><input type="date" /></Field> */}
                        </Grid>

                        <table className="data-table" style={{ marginTop: "10px" }}>
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
                                    <th>Add/Delete Entry</th>
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
                                                value={row.boeNo}
                                                onChange={(e) =>
                                                    handleChange(index, "boeNo", e.target.value)
                                                }
                                            />
                                        </td>

                                        <td>
                                            <input
                                                type="date"
                                                value={row.boeDate}
                                                onChange={(e) =>
                                                    handleChange(index, "boeDate", e.target.value)
                                                }
                                            />
                                        </td>

                                        <td>
                                            <input
                                                value={row.mrnNo}
                                                onChange={(e) =>
                                                    handleChange(index, "mrnNo", e.target.value)
                                                }
                                            />
                                        </td>

                                        <td>
                                            <input
                                                value={row.blNo}
                                                onChange={(e) =>
                                                    handleChange(index, "blNo", e.target.value)
                                                }
                                            />
                                        </td>

                                        <td>
                                            <input
                                                type="date"
                                                value={row.blDate}
                                                onChange={(e) =>
                                                    handleChange(index, "blDate", e.target.value)
                                                }
                                            />
                                        </td>

                                        <td>
                                            <input
                                                type="number"
                                                value={row.invoiceAmount}
                                                onChange={(e) =>
                                                    handleChange(index, "invoiceAmount", e.target.value)
                                                }
                                            />
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

                                        <td style={{ textAlign: "center" }}>
                                            {/* Show PLUS only on last row */}
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

                                            {/* Show DELETE if more than 1 row */}
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
                                        </td>

                                    </tr>
                                ))}
                            </tbody>
                            <tfoot>
                                <tr>
                                    <td colSpan={8} style={{ textAlign: "right", fontWeight: "bold" }}>
                                        Total Invoice Amount:
                                    </td>
                                    <td style={{ fontWeight: "bold" }}>
                                        {totalInvoiceAmount.toFixed(2)}
                                    </td>
                                    <td colSpan={3}></td>
                                </tr>
                            </tfoot>
                        </table>
                        <div style={{ display: "flex", gap: "40px", marginTop: "30px" }}>

                            {/* BOE TABLE */}
                            <div style={{ flex: 1 }}>
                                <p style={{ color: "red", fontSize: "13px" }}>
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

                                <table className="data-table">
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

                    </Section>
                )}
                {/* ==============================service-Bill Payment============================== */}
                {paymentType === "Service-Bill Payment" && (
                    <Section title="Forex Payment Request Details">
                        <Grid>
                            <Field label="Request Number"><input value={requestNumber} onChange={(e) => { setRequestNumber(e.target.value) }} /></Field>
                            <Field label="Requested On"><input type="date" value={requestedOn} onChange={(e) => { setRequestedOn(e.target.value) }} /></Field>
                            <Field label="Currency"><input value={currency} onChange={(e) => { setCurrency(e.target.value) }} /></Field>
                            <Field label="Total Amount"><input type="number" value={totalAmount} onChange={(e) => { setTotalAmount(e.target.value) }} /></Field>
                            <Field label="Foreign Bank Charges"><input type="number" value={foreignBankCharges} onChange={(e) => { setForeignBankCharges(e.target.value) }} /></Field>
                            {/* <Field label="PO/Contract No"><input /></Field>
                            <Field label="PO Date"><input type="date" /></Field>
                            <Field label="Expected Settlement Date"><input type="date" /></Field> */}
                        </Grid>

                        <table className="data-table" style={{ marginTop: "10px" }}>
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
                                    <th>Add/Delete Entry</th>
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
                                            <input
                                                value={row.mrnNo}
                                                onChange={(e) =>
                                                    handleChange(index, "mrnNo", e.target.value)
                                                }
                                            />
                                        </td>
                                        <td>
                                            <input
                                                type="date"
                                                value={row.mrnDate}
                                                onChange={(e) =>
                                                    handleChange(index, "mrnDate", e.target.value)
                                                }
                                            />
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




                                        <td style={{ textAlign: "center" }}>
                                            {/* Show PLUS only on last row */}
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

                                            {/* Show DELETE if more than 1 row */}
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
                                        </td>

                                    </tr>
                                ))}
                            </tbody>
                            <tfoot>
                                <tr>
                                    <td colSpan={3} style={{ textAlign: "right", fontWeight: "bold" }}>
                                        Total Invoice Amount:
                                    </td>
                                    <td style={{ fontWeight: "bold" }}>
                                        {totalInvoiceAmount.toFixed(2)}
                                    </td>
                                    <td colSpan={5}></td>
                                </tr>
                            </tfoot>
                        </table>
                    </Section>
                )}

                {/* ==============================Goods - Advance Payment || Service - Advance Payment============================== */}
                {paymentType === "Service-Advance Payment" && (

                    <Section title="Forex Payment Request Details">
                        <Grid>
                            <Field label="Request Number"><input value={requestNumber} onChange={(e) => { setRequestNumber(e.target.value) }} /></Field>
                            <Field label="Requested On"><input type="date" value={requestedOn} onChange={(e) => { setRequestedOn(e.target.value) }} /></Field>
                            <Field label="Currency"><input value={currency} onChange={(e) => { setCurrency(e.target.value) }} /></Field>
                            <Field label="Total Amount"><input type="number" value={totalAmount} onChange={(e) => { setTotalAmount(e.target.value) }} /></Field>
                            <Field label="Foreign Bank Charges"><input type="number" value={foreignBankCharges} onChange={(e) => { setForeignBankCharges(e.target.value) }} /></Field>
                            <Field label="PO/Contract No"><input value={poContractNo} onChange={(e) => { setPoContractNo(e.target.value) }} /></Field>
                            <Field label="PO Date"><input type="date" value={poDate} onChange={(e) => { setPoDate(e.target.value) }} /></Field>
                            <Field label="Expected Settlement Date"><input type="date" value={expectedSettlementDate} onChange={(e) => { setExpectedSettlementDate(e.target.value) }} /></Field>
                        </Grid>

                        <table className="data-table" style={{ marginTop: "10px" }}>
                            <thead>
                                <tr>
                                    <th>Sr No</th>
                                    <th>Performa Invoice No</th>
                                    <th>Performa Invoice Date</th>
                                    <th>Performa Invoice Amount</th>
                                    <th>Attach PO</th>
                                    <th>Attach PI</th>
                                    <th>Attach Other Document</th>
                                    <th>Add/Delete Entry</th>
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
                                            <input
                                                type="file"

                                            />
                                        </td>
                                        <td>
                                            <input
                                                type="file"


                                            />
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




                                        <td style={{ textAlign: "center" }}>
                                            {/* Show PLUS only on last row */}
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

                                            {/* Show DELETE if more than 1 row */}
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
                                        </td>

                                    </tr>
                                ))}
                            </tbody>
                            <tfoot>
                                <tr>
                                    <td colSpan={3} style={{ textAlign: "right", fontWeight: "bold" }}>
                                        Total Invoice Amount:
                                    </td>
                                    <td style={{ fontWeight: "bold" }}>
                                        {totalInvoiceAmount.toFixed(2)}
                                    </td>
                                    <td colSpan={4}></td>
                                </tr>
                            </tfoot>
                        </table>
                    </Section>
                )}
                {paymentType === "Goods-Advance Payment" && (

                    <Section title="Forex Payment Request Details">
                        <Grid>
                            <Field label="Request Number"><input value={requestNumber} onChange={(e) => { setRequestNumber(e.target.value) }} /></Field>
                            <Field label="Requested On"><input type="date" value={requestedOn} onChange={(e) => { setRequestedOn(e.target.value) }} /></Field>
                            <Field label="Currency"><input value={currency} onChange={(e) => { setCurrency(e.target.value) }} /></Field>
                            <Field label="Total Amount"><input type="number" value={totalAmount} onChange={(e) => { setTotalAmount(e.target.value) }} /></Field>
                            <Field label="Foreign Bank Charges"><input type="number" value={foreignBankCharges} onChange={(e) => { setForeignBankCharges(e.target.value) }} /></Field>
                            <Field label="PO/Contract No"><input value={poContractNo} onChange={(e) => { setPoContractNo(e.target.value) }} /></Field>
                            <Field label="PO Date"><input type="date" value={poDate} onChange={(e) => { setPoDate(e.target.value) }} /></Field>
                            <Field label="Expected Settlement Date"><input type="date" value={expectedSettlementDate} onChange={(e) => { setExpectedSettlementDate(e.target.value) }} /></Field>
                        </Grid>

                        <table className="data-table" style={{ marginTop: "10px" }}>
                            <thead>
                                <tr>
                                    <th>Sr No</th>
                                    <th>Performa Invoice No</th>
                                    <th>Performa Invoice Date</th>
                                    <th>Performa Invoice Amount</th>
                                    <th>Attach PO</th>
                                    <th>Attach PI</th>
                                    <th>Attach Other Document</th>
                                    <th>Add/Delete Entry</th>
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
                                            <input
                                                type="file"

                                            />
                                        </td>
                                        <td>
                                            <input
                                                type="file"


                                            />
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




                                        <td style={{ textAlign: "center" }}>
                                            {/* Show PLUS only on last row */}
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

                                            {/* Show DELETE if more than 1 row */}
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
                                        </td>

                                    </tr>
                                ))}
                            </tbody>
                            <tfoot>
                                <tr>
                                    <td colSpan={3} style={{ textAlign: "right", fontWeight: "bold" }}>
                                        Total Invoice Amount:
                                    </td>
                                    <td style={{ fontWeight: "bold" }}>
                                        {totalInvoiceAmount.toFixed(2)}
                                    </td>
                                    <td colSpan={4}></td>
                                </tr>
                            </tfoot>
                        </table>
                    </Section>
                )}



                {/* ================= CORRESPONDENT ================= */}
                <Section title="Correspondent Bank Details">
                    <Grid>
                        <Field label="Bank Name"><input value={bankname} onChange={(e) => { setBankName(e.target.value) }} /></Field>
                        <Field label="Swift Code"><input value={bankswiftcode} onChange={(e) => { setBankSwiftCode(e.target.value) }} /></Field>
                        <Field label="Bank Account No"><input value={bankaccountno} onChange={(e) => { setBankAccountNo(e.target.value) }} /></Field>
                        <Field label="Remarks" full><textarea rows={3} value={remarks} onChange={(e) => { setRemarks(e.target.value) }}></textarea></Field>
                    </Grid>
                </Section>
                <Section title="Approval Remarks">

                    <Grid>

                        {approverDetails.map((a, i) => (

                            <Field key={i} label={`${a.Role} Remarks`} full>

                                <textarea value={a.Comment || ""} readOnly />

                            </Field>

                        ))}

                    </Grid>

                </Section>
                <Section title="Payment Details">

                    <Grid>

                        <Field label="Foreign Currency">
                            <input value={foreignCurrency} readOnly />
                        </Field>

                        <Field label="Foreign Currency Amount">
                            <input value={foreignCurrencyAmount} readOnly />
                        </Field>

                        <Field label="Exchange Rate">
                            <input value={exchangeRate} readOnly />
                        </Field>

                        <Field label="INR Amount">
                            <input value={inrAmount} readOnly />
                        </Field>

                        <Field label="Payment Date">
                            <input type="date" value={paymentDate} readOnly />
                        </Field>

                        <Field label="Payment Reference Number">
                            <input value={paymentReferenceNumber} readOnly />
                        </Field>

                    </Grid>

                </Section>
                <div className="button-row">
                    {/* <button className="btn-submit" onClick={onsubmit}>Submit</button> */}
                    <button className="btn-exit" onClick={() => history.push("/")}>Exit</button>
                </div>

            </div>
            {showVendorPopup && (
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
                            </button> */}
                        </div>
                    </div>
                </div>
            )}

        </div>
    );
};

export default ViewRequestForm;


