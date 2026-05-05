import React, { useEffect, useState } from "react";
import "../Pages/Css/NewRequest.scss";
import { IForexModuleProps } from "../IForexModuleProps";
import { useHistory, useParams } from 'react-router-dom';
import { Dropdown, IDropdownOption } from '@fluentui/react';
import SPCRUDOPS from "../../service/BAL/spcrud";
import { Attachment } from "@pnp/sp/attachments";
import { View } from "@pnp/sp/views";
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/attachments";
import { SPHttpClient } from "@microsoft/sp-http";
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

const ApprovalRequestForm = (props: IForexModuleProps) => {
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
    const [nextApprovers, setNextApprovers] = useState<number[]>([]);
    const [currentApprover, setCurrentApprover] = useState<number[]>([]);
    const currentUserId = props.context.pageContext.legacyPageContext.userId;
    const [invoiceAttachments, setInvoiceAttachments] = useState<any>({});
    const [otherAttachments, setOtherAttachments] = useState<any>({});
    const [poAttachments, setPoAttachments] = useState<any>({});
    const [piAttachments, setPiAttachments] = useState<any>({});
    const [boeLibraryFiles, setBoeLibraryFiles] = useState<any[]>([]);
    const [bolLibraryFiles, setBolLibraryFiles] = useState<any[]>([]);
    const [approvalSteps, setApprovalSteps] = useState<any[]>([]);

    const [prevApprovers, setPrevApprovers] = useState<number[]>([]);
    const [allApprovers, setAllApprovers] = useState<number[]>([]);

    const [approverRemarks, setApproverRemarks] = useState("");
    const [validationDate, setValidationDate] = useState("");
    const [voucherNumber, setVoucherNumber] = useState("");
    const [vouchingRemarks, setVouchingRemarks] = useState("");
    const [treasuryRemarks, setTreasuryRemarks] = useState("");
    const [currentRole, setCurrentRole] = useState("");

    const [foreignCurrency, setForeignCurrency] = useState("");
    const [foreignAmount, setForeignAmount] = useState("");
    const [exchangeRate, setExchangeRate] = useState("");
    const [inrAmount, setInrAmount] = useState("");
    const [paymentDate, setPaymentDate] = useState("");
    const [paymentReference, setPaymentReference] = useState("");

    const [swiftCopy, setSwiftCopy] = useState<File | null>(null);
    const [form15CA, setForm15CA] = useState<File | null>(null);
    const [form15CB, setForm15CB] = useState<File | null>(null);
    const [allApproversJson, setAllApproversJson] = useState<string>("");
    const [workflowHistory, setWorkflowHistory] = useState<any[]>([]);
    const [eligibleAmountWithWHT, setEligibleAmountWithWHT] = useState("");
    const [paidAmount, setPaidAmount] = useState("");
    const [ballenceEligibleAmount, setBallenceEligibleAmount] = useState("");
    const [foreignCurrencyOptions, setforeignCurrencyOptions] = useState<IDropdownOption[]>([]);

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

        sp.setup({
            spfxContext: props.context
        });

    }, []);

    useEffect(() => {
        getFinancialYearStart();
        getCurrencyData();
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
                "*,AllApprovers,RM/Title,HOD/Title,RM/Id,HOD/Id,Author/Id,Currency/Title,Currency/Id,NextApprovers/Id,NextApprovers/Title,CurrentApprover/Id,CurrentApprover/Title,PrevApprovers/Id,PrevApprovers/Title",
                "RM,HOD,Author,Currency,NextApprovers,CurrentApprover,PrevApprovers",
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
                setNextApprovers(data.NextApprovers?.map((approver: any) => approver.Id) || []);
                setEligibleAmountWithWHT(data.EligibleAmountWithWHT || "");
                setPaidAmount(data.PaidAmount || "");
                setBallenceEligibleAmount(data.BallenceEligibleAmount || "");

                if (data.AllApprovers) {

                    setAllApproversJson(data.AllApprovers);

                    const parsed = JSON.parse(data.AllApprovers);

                    const ids = parsed.map((a: any) => a.Id);

                    setAllApprovers(ids);

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
                setCurrentApprover(data.CurrentApprover ? [data.CurrentApprover.Id] : []);
                if (data.AllApprovers) {

                    const approvers = JSON.parse(data.AllApprovers);

                    const current = approvers.find((a: any) => a.Id === data.CurrentApprover?.Id);

                    if (current) {
                        setCurrentRole(current.Role);
                    }

                }

                setPrevApprovers(
                    data.PrevApprovers?.map((p: any) => p.Id) || []
                );
                const steps: any[] = [];

                steps.push({
                    name: data.EmployeeName,
                    status: "initiator"
                });

                let approvers: any[] = [];

                if (data.AllApprovers) {
                    approvers = JSON.parse(data.AllApprovers);
                }

                approvers.forEach((a: any) => {

                    let status = "pending";

                    if (a.Status === "Approved") {
                        status = "approved";
                    }
                    else if (data.CurrentApprover?.Id === a.Id) {
                        status = "current";
                    }

                    steps.push({
                        id: a.Id,
                        name: `${a.Role} - ${a.Name}`,
                        status
                    });

                });

                setApprovalSteps(steps);
                if (data.WorkFlowHistory) {
                    try {
                        const parsed = JSON.parse(data.WorkFlowHistory);
                        setWorkflowHistory(parsed);
                    } catch {
                        setWorkflowHistory([]);
                    }
                }

            }

            // 🔹 Load Child Rows
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
                const invoiceMap: any = {};
                const otherMap: any = {};
                const poMap: any = {};
                const piMap: any = {};

                child.forEach((item: any, index: number) => {

                    const allFiles = item.AttachmentFiles || [];

                    invoiceMap[index] = allFiles.filter((f: any) =>
                        f.FileName?.startsWith("INV_")
                    );

                    otherMap[index] = allFiles.filter((f: any) =>
                        f.FileName?.startsWith("DOC_")
                    );

                    poMap[index] = allFiles.filter((f: any) =>
                        f.FileName?.startsWith("PO_")
                    );

                    piMap[index] = allFiles.filter((f: any) =>
                        f.FileName?.startsWith("PI_")
                    );
                });

                setInvoiceAttachments(invoiceMap);
                setOtherAttachments(otherMap);
                setPoAttachments(poMap);
                setPiAttachments(piMap);
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

    const getApproversdata = async (): Promise<number[]> => {
        const Sp = await spCrudOps;

        try {

            // 🔹 Change this to your actual request title value
            //  const requestTitle = requestType; // example: "1"

            const res = await Sp.getData(
                "ForexApprovalMAtrix",
                "ID,Approver/Id,Approver/Title,Level,Title",
                "Approver",
                ``,

                { column: "Level", isAscending: true },
                5000,
                props,
                // 🔹 Expand lookup
            );

            if (!res || res.length === 0) {
                console.error("No approvers found in matrix");
                return [];
            }

            // 🔹 Define category execution order
            const categoryOrder = ["Approver", "VouchingApprover", "TreasuryTeam"];

            // 🔹 Sort category first, then level
            const sorted = res.sort((a: any, b: any) => {

                const catDiff =
                    categoryOrder.indexOf(a.Level) -
                    categoryOrder.indexOf(b.Level);

                if (catDiff !== 0) return catDiff;

                return a.Level - b.Level;
            });

            // 🔹 Return only user IDs
            return sorted.map((item: any) => item.Approver?.Id);

        } catch (error) {
            console.error("Error fetching approval matrix:", error);
            return [];
        }
    };

    const validateAndBuildApprovers = async (): Promise<any[]> => {

        const sp = await spCrudOps;

        const approvers: any[] = [];

        // RM
        if (employee.RMId) {

            approvers.push({
                Id: employee.RMId,
                Name: employee.RM,
                Role: "RM",
                Level: 1
            });

        }

        // HOD
        if (employee.HODId) {

            approvers.push({
                Id: employee.HODId,
                Name: employee.HOD,
                Role: "HOD",
                Level: 2
            });


        }
        let requestTypeFilter = "";
        if (
            paymentType === "Goods-Advance Payment" ||
            paymentType === "Service-Advance Payment"
        ) {
            requestTypeFilter = "Advance Payment";
        } else {
            requestTypeFilter = paymentType; // direct match for Bill types
        }

        const matrix = await sp.getData(
            "ForexApprovalMatrix",
            "Title,Role,Approver/Id,Approver/Title,Level,RequestType",
            "Approver",
            `RequestType eq '${requestTypeFilter}' and Status eq 'Active'`,
            { column: "Level", isAscending: true },
            5000,
            props
        );

        matrix.forEach((item: any) => {

            const id = item.Approver?.Id;

            if (!approvers.find(a => a.Id === id)) {

                approvers.push({

                    Id: id,
                    Name: item.Approver?.Title,
                    Role: item.Role,
                    Level: item.Level

                });

            }

        });

        return approvers;

    };

    const getCurrencyData = async () => {
        try {
            const sp = await spCrudOps;
            await sp.getData(
                "Currency",
                "Title,Id",
                "",
                "",
                { column: "Title", isAscending: true },
                5000,
                props
            ).then((res: any[]) => {
                const options = res.map((c: any) => ({
                    key: c.Title,
                    text: c.Title
                }));
                setforeignCurrencyOptions(options);
            });
        } catch (error) { console.error("Error fetching currency data:", error); }
    }


    const onsubmit = async () => {

        const sp = await spCrudOps;

        try {

            const correctApproversObjects = await validateAndBuildApprovers();

            if (!correctApproversObjects.length) {

                alert("Approval matrix not configured");
                return;

            }

            // 🔹 Load existing JSON stored in SharePoint
            let existingApprovers: any[] = [];

            if (allApproversJson) {
                existingApprovers = JSON.parse(allApproversJson);
            }

            // 🔹 Build new approver list but preserve old status
            const approvers = correctApproversObjects.map((newAppr: any) => {

                const existing = existingApprovers.find(
                    (a: any) => a.Id === newAppr.Id
                );

                return existing
                    ? existing
                    : {
                        ...newAppr,
                        Status: "",
                        ActionDate: "",
                        Remarks: ""
                    };

            });

            const currentIndex = approvers.findIndex(
                (a: any) => a.Id === currentUserId
            );

            if (currentIndex === -1) {

                alert("You are not authorized to approve");
                return;

            }

            const currentApproverObj = approvers[currentIndex];

            // ⭐ Update current approver
            approvers[currentIndex].Status = "Approved";
            approvers[currentIndex].ActionDate = new Date().toISOString();
            approvers[currentIndex].Remarks = approverRemarks;


            // 🔹 Find next approver
            const nextApprover =
                currentIndex + 1 < approvers.length
                    ? approvers[currentIndex + 1]
                    : null;

            if (nextApprover && !nextApprover.Status) {
                nextApprover.Status = "Pending";
            }

            let status = "Approved";

            if (nextApprover) {

                if (nextApprover.Role === "HOD")
                    status = "Pending for HOD";

                else if (nextApprover.Role === "Vouching")
                    status = "Pending for Vouching";

                else if (nextApprover.Role === "TreasuryVerification")
                    status = "Pending for Treasury Verification";

                else if (nextApprover.Role === "TreasuryPayment")
                    status = "Pending for Payment";

            }
            else {

                // Final approver logic
                if (currentApproverObj.Role === "TreasuryPayment" && paymentType.includes("Advance")) {
                    status = "Paid";
                } else {
                    status = "Paid & Closed";
                }

            }
            // 🔥 ALWAYS GET LATEST FROM SP
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
            // ⭐ Workflow history update


            let remarksPayload: any = {};
            let finalRemark = "";
            if (currentApproverObj.Role === "RM" || currentApproverObj.Role === "HOD") {

                if (!approverRemarks || approverRemarks.trim() === "") {
                    alert("Please enter remark before approving");
                    return;
                }

                finalRemark = approverRemarks;
            }

            else if (currentApproverObj.Role === "Vouching") {

                if (!vouchingRemarks || vouchingRemarks.trim() === "") {
                    alert("Please enter vouching remark");
                    return;
                }


                finalRemark = vouchingRemarks;
            }

            else if (currentApproverObj.Role === "TreasuryVerification") {

                if (!treasuryRemarks || treasuryRemarks.trim() === "") {
                    alert("Please enter treasury remark");
                    return;
                }

                finalRemark = treasuryRemarks;
            }

            else if (currentApproverObj.Role === "TreasuryPayment") {

                if (!paymentReference || paymentReference.trim() === "") {
                    alert("Please enter payment reference");
                    return;
                }

                finalRemark = paymentReference; // or treasuryRemarks if you want
            }


            let updatedHistory = [...existingHistory];
            updatedHistory.push({
                CurrentApprover: currentApproverObj.Name,
                ActionTaken: "Approved",
                Comment: finalRemark,
                Date: new Date().toISOString(),
                CurrentStatus: status
            });


            if (currentApproverObj.Role === "RM") {
                remarksPayload.RMRemark = approverRemarks;
            }

            if (currentApproverObj.Role === "HOD") {
                remarksPayload.HODRemarks = approverRemarks;
            }

            if (currentApproverObj.Role === "Vouching") {

                remarksPayload.ValidationDate = validationDate;
                remarksPayload.VoucherNumber = voucherNumber;
                remarksPayload.VouchingRemarks = vouchingRemarks;

                if (!validationDate) {
                    alert("Please enter validation date");
                    return;
                }
                if (paymentType.includes("Advance")) {
                    if (!voucherNumber || voucherNumber.trim() === "") {
                        alert("Please enter voucher number");
                        return;
                    }
                }


            }

            if (currentApproverObj.Role === "TreasuryVerification") {
                remarksPayload.TreasuryRemark = treasuryRemarks;
            }

            if (currentApproverObj.Role === "TreasuryPayment") {
                if (!foreignCurrency) {
                    alert("Please select foreign currency");
                    return;
                }
                if (!foreignAmount || isNaN(Number(foreignAmount))) {
                    alert("Please enter a valid foreign amount");
                    return;
                }
                if (!exchangeRate || isNaN(Number(exchangeRate))) {
                    alert("Please enter a valid exchange rate");
                    return;
                }
                if (!inrAmount || isNaN(Number(inrAmount))) {
                    alert("Please enter a valid INR amount");
                    return;
                }
                if (!paymentDate) {
                    alert("Please enter payment date");
                    return;
                }
                if (!paymentReference || paymentReference.trim() === "") {
                    alert("Please enter payment reference");
                    return;
                }

                remarksPayload.ForeignCurrency = foreignCurrency;
                remarksPayload.ForeignCurrencyAmount = foreignAmount;
                remarksPayload.ExchangeRate = exchangeRate;
                remarksPayload.INRAmount = inrAmount;
                remarksPayload.PaymentDate = paymentDate;
                remarksPayload.PaymentReferenceNumber = paymentReference;

            }

            // 🔹 Upload payment documents
            if (currentApproverObj.Role === "TreasuryPayment") {

                const webUrl = props.context.pageContext.web.absoluteUrl;

                if (swiftCopy) {

                    const fileName = `SWIFT_${swiftCopy.name}`;

                    await props.context.spHttpClient.post(
                        `${webUrl}/_api/web/lists/getbytitle('ForexRequest')/items(${Id})/AttachmentFiles/add(FileName='${fileName}')`,
                        SPHttpClient.configurations.v1,
                        { body: swiftCopy }
                    );

                }

                if (form15CA) {

                    const fileName = `15CA_${form15CA.name}`;

                    await props.context.spHttpClient.post(
                        `${webUrl}/_api/web/lists/getbytitle('ForexRequest')/items(${Id})/AttachmentFiles/add(FileName='${fileName}')`,
                        SPHttpClient.configurations.v1,
                        { body: form15CA }
                    );

                }

                if (form15CB) {

                    const fileName = `15CB_${form15CB.name}`;

                    await props.context.spHttpClient.post(
                        `${webUrl}/_api/web/lists/getbytitle('ForexRequest')/items(${Id})/AttachmentFiles/add(FileName='${fileName}')`,
                        SPHttpClient.configurations.v1,
                        { body: form15CB }
                    );

                }

            }

            // ⭐ Update SharePoint
            await sp.updateData(
                "ForexRequest",
                Number(Id),
                {
                    AllApprovers: JSON.stringify(approvers),
                    WorkFlowHistory: JSON.stringify(updatedHistory),
                    CurrentApproverId: nextApprover ? nextApprover.Id : null,
                    Status: status,
                    ...remarksPayload
                },
                props
            );
            if (currentApproverObj.Role === "Vouching") {
                alert("Request has been vouched");
            } else if (currentApproverObj.Role === "TreasuryVerification") {
                alert("Request has been verified");
            } else if (currentApproverObj.Role === "TreasuryPayment") {
                alert("Request has been processed successfully”");
            } else {
                alert("Request has been approved successfully");
            }
            history.push("/ApprovalDashboard");

        } catch (error) {

            console.error(error);
            alert("Approval failed");

        }

    };
    const onReject = async () => {

        const sp = await spCrudOps;

        try {

            // 🔴 Mandatory Remark Check
            if (!approverRemarks || approverRemarks.trim() === "") {
                alert("Please enter remarks before Rejecting the request.");
                return;
            }

            const correctApproversObjects = await validateAndBuildApprovers();

            const approvers = correctApproversObjects;

            const currentIndex = approvers.findIndex(
                a => a.Id === currentUserId
            );

            if (currentIndex === -1) {

                alert("You are not authorized");
                return;

            }
            const currentApproverObj = approvers[currentIndex];

            // ⭐ Update JSON history



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

            let finalRemark = "";
            if (currentApproverObj.Role === "RM" || currentApproverObj.Role === "HOD") {

                if (!approverRemarks || approverRemarks.trim() === "") {
                    alert("Please enter remark before approving");
                    return;
                }

                finalRemark = approverRemarks;
            }

            else if (currentApproverObj.Role === "Vouching") {

                if (!vouchingRemarks || vouchingRemarks.trim() === "") {
                    alert("Please enter vouching remark");
                    return;
                }

                finalRemark = vouchingRemarks;
            }

            else if (currentApproverObj.Role === "TreasuryVerification") {

                if (!treasuryRemarks || treasuryRemarks.trim() === "") {
                    alert("Please enter treasury remark");
                    return;
                }

                finalRemark = treasuryRemarks;
            }

            else if (currentApproverObj.Role === "TreasuryPayment") {

                if (!paymentReference || paymentReference.trim() === "") {
                    alert("Please enter payment reference");
                    return;
                }

                finalRemark = paymentReference; // or treasuryRemarks if you want
            }
            approvers[currentIndex].Status = "Rejected";
            approvers[currentIndex].ActionDate = new Date().toISOString();
            approvers[currentIndex].Remarks = finalRemark;

            let updatedHistory = [...existingHistory];
            updatedHistory.push({
                CurrentApprover: approvers[currentIndex].Name,
                ActionTaken: "Rejected",
                Comment: finalRemark,
                Date: new Date().toISOString(),
                CurrentStatus: "Rejected"
            });

            // let updatedHistory = [...existingHistory];

            // updatedHistory.push({
            //     CurrentApprover: approvers[currentIndex].Name,
            //     ActionTaken: "Rejected",
            //     Comment: approverRemarks,
            //     Date: new Date().toISOString(),
            //     CurrentStatus: "Rejected"
            // });

            await sp.updateData(
                "ForexRequest",
                Number(Id),
                {
                    AllApprovers: JSON.stringify(approvers),
                    WorkFlowHistory: JSON.stringify(updatedHistory),
                    CurrentApproverId: null,
                    Status: "Rejected"
                },
                props
            );

            alert("Request has been rejected by Approver");

            history.push("/ApprovalDashboard");

        } catch (error) {

            console.error(error);
            alert("Sent back failed");

        }

    };

    const onSentBack = async () => {

        const sp = await spCrudOps;

        try {

            // const correctApproversObjects = await validateAndBuildApprovers();


            // 🔴 Mandatory Remark
            // if (!approverRemarks || approverRemarks.trim() === "") {
            //     alert("Please enter remarks before sending back.");
            //     return;
            // }

            // 🔹 1. Get latest item from SP
            const existingItem = await sp.getData(
                "ForexRequest",
                "AllApprovers,WorkFlowHistory",
                "",
                `ID eq ${Id}`,
                { column: "ID", isAscending: true },
                1,
                props
            );

            if (!existingItem.length) {
                alert("Item not found");
                return;
            }

            // 🔹 2. Parse approvers JSON
            let approvers: any[] = [];
            if (existingItem[0].AllApprovers) {
                approvers = JSON.parse(existingItem[0].AllApprovers);
            }

            // 🔹 3. Find current approver index
            const currentIndex = approvers.findIndex(
                (a: any) => a.Id === currentUserId
            );

            if (currentIndex === -1) {
                alert("You are not authorized");
                return;
            }

            const currentApproverObj = approvers[currentIndex];

            // 🔹 4. Mark current as Sent Back


            // 🔹 5. Move to previous approver
            const previousApprover =
                currentIndex - 1 >= 0 ? approvers[currentIndex - 1] : null;

            if (previousApprover) {
                previousApprover.Status = "Pending";
            }
            let finalRemark = "";
            if (currentApproverObj.Role === "RM" || currentApproverObj.Role === "HOD") {

                if (!approverRemarks || approverRemarks.trim() === "") {
                    alert("Please enter remark before approving");
                    return;
                }

                finalRemark = approverRemarks;
            }

            else if (currentApproverObj.Role === "Vouching") {

                if (!vouchingRemarks || vouchingRemarks.trim() === "") {
                    alert("Please enter vouching remark");
                    return;
                }


                finalRemark = vouchingRemarks;
            }

            else if (currentApproverObj.Role === "TreasuryVerification") {

                if (!treasuryRemarks || treasuryRemarks.trim() === "") {
                    alert("Please enter treasury remark");
                    return;
                }

                finalRemark = treasuryRemarks;
            }
            // 🔹 6. Workflow History
            approvers[currentIndex].Status = "Sent Back";
            approvers[currentIndex].ActionDate = new Date().toISOString();
            approvers[currentIndex].Remarks = finalRemark;
            let existingHistory: any[] = [];

            if (existingItem[0].WorkFlowHistory) {
                try {
                    existingHistory = JSON.parse(existingItem[0].WorkFlowHistory);
                } catch {
                    existingHistory = [];
                }
            }

            existingHistory.push({
                CurrentApprover: currentApproverObj.Name,
                ActionTaken: "Sent Back",
                Comment: finalRemark,
                Date: new Date().toISOString(),
                CurrentStatus: "Sent Back"
            });

            // 🔹 7. Update SharePoint
            await sp.updateData(
                "ForexRequest",
                Number(Id),
                {
                    AllApprovers: JSON.stringify(approvers),
                    WorkFlowHistory: JSON.stringify(existingHistory),
                    CurrentApproverId: previousApprover ? previousApprover.Id : null,
                    Status: "Sent Back"
                },
                props
            );

            alert("Request sent back successfully");

            history.push("/ApprovalDashboard");

        } catch (error) {

            console.error(error);
            alert("Send back failed");

        }
    };
    const isServicePayment = paymentType.includes("Service");
    const isGoodsPayment = paymentType.includes("Goods");

    const getApproveButtonText = () => {
        switch (currentRole) {
            case "Vouching":
                return "Vouch";
            case "TreasuryVerification":
                return "Verify";
            case "TreasuryPayment":
                return "Pay & Close";
            default:
                return "Approve";
        }
    };

    return (
        <div className="forex-wrapper">

            {/* ================= HEADER ================= */}
            <div className="forex-header">
                <h2>Forex Payment Approval Form</h2>
            </div>

            <div className="forex-card">
                {/* ================= APPROVAL FLOW ================= */}

                <Section title="Approval Flow">
                    <div className="approval-ribbon">

                        {approvalSteps.map((step, index) => {

                            let className = "pending";

                            if (step.status === "initiator") className = "initiator";
                            if (step.status === "approved") className = "approved";
                            if (step.status === "current") className = "current";

                            return (
                                <div key={index} className={`ribbon-step ${className}`}>
                                    {step.name}
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
                <CollapsibleSection title="Vendor / Beneficiary Details">
                    <Grid>
                        <Field label="Vendor Code">
                            <input
                                value={vendor.VendorCode}

                                onChange={(e) => {
                                    const code = e.target.value;
                                    setVendor({ ...vendor, VendorCode: code });
                                }}
                                onBlur={(e) => getVendorData(e.target.value)}
                                readOnly
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
                </CollapsibleSection>

                {/* ================= TAX INFO ================= */}
                <CollapsibleSection title="Tax & Regulatory Information">
                    <Grid>
                        <Field label="Nature of Payment"><input value={paymentType} readOnly /></Field>
                        <Field label="Tax Document Available?">
                            <select onChange={(e) => { setTaxDocumentView(e.target.value) }} disabled>
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
                                <select value={dTAAApplicable} onChange={(e) => setDTAAApplicable(e.target.value)} disabled>
                                    <option value="">Select</option>
                                    <option value="Yes">Yes</option>
                                    <option value="No">No</option>
                                </select>
                            </Field>
                        )}
                    </Grid>
                </CollapsibleSection>
                {taxDocumentView === "Yes" && (
                    <>
                        {/* 🔹 Permanent Establishment Declaration */}
                        <CollapsibleSection title="Permanent Establishment Declaration">
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
                        </CollapsibleSection>

                        {/* 🔹 Tax Residency Certificate */}
                        <CollapsibleSection title="Tax Residency Certificate">
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
                        </CollapsibleSection>

                        {/* 🔹 Form 10F */}
                        <CollapsibleSection title="Form 10F">
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
                        </CollapsibleSection>
                    </>
                )}

                <CollapsibleSection title="Other Details" >
                    <div className="date-summary">
                        <span className="label">From</span>
                        <span className="value">{fromdate}</span>
                        <span className="label">To</span>
                        <span className="value">{todate}</span>
                    </div>
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

                </CollapsibleSection>

                {/* ================= FOREX DETAILS ================= */}
                {/* ==============================Goods-Bill Payment============================== */}
                {paymentType === "Goods-Bill Payment" && (
                    <Section title="Forex Payment Request Details">
                        <Grid>
                            <Field label="Request Number"><input value={requestNumber} onChange={(e) => { setRequestNumber(e.target.value) }} readOnly /></Field>
                            <Field label="Requested On"><input type="date" value={requestedOn} onChange={(e) => { setRequestedOn(e.target.value) }} readOnly /></Field>
                            <Field label="Currency"><input value={currency} onChange={(e) => { setCurrency(e.target.value) }} readOnly /></Field>
                            <Field label="Total Amount"><input type="number" value={totalAmount} onChange={(e) => { setTotalAmount(e.target.value) }} readOnly /></Field>
                            <Field label="Foreign Bank Charges"><input type="text" value={foreignBankCharges} onChange={(e) => { setForeignBankCharges(e.target.value) }} readOnly /></Field>
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
                                                readOnly
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
                                                {invoiceAttachments[index]?.length > 0 ? (
                                                    invoiceAttachments[index].map((file: any, i: number) => (
                                                        <div key={i}>
                                                            <a href={file.ServerRelativeUrl} target="_blank">
                                                                {file.FileName.replace("INV_", "")}
                                                            </a>
                                                        </div>
                                                    ))
                                                ) : (
                                                    <span>-</span>
                                                )}
                                            </div>                                        </td>

                                        <td>
                                            <div>
                                                {otherAttachments[index]?.map((file: any, i: number) => (
                                                    <div key={i}>
                                                        <a href={file.ServerRelativeUrl} target="_blank">
                                                            {file.FileName.replace("DOC_", "")}
                                                        </a>
                                                    </div>
                                                ))}
                                            </div>                                        </td>

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
                            <Field label="Request Number"><input value={requestNumber} onChange={(e) => { setRequestNumber(e.target.value) }} readOnly /></Field>
                            <Field label="Requested On"><input type="date" value={requestedOn} onChange={(e) => { setRequestedOn(e.target.value) }} readOnly /></Field>
                            <Field label="Currency"><input value={currency} onChange={(e) => { setCurrency(e.target.value) }} /></Field>
                            <Field label="Total Amount"><input type="number" value={totalAmount} onChange={(e) => { setTotalAmount(e.target.value) }} readOnly /></Field>
                            <Field label="Foreign Bank Charges"><input type="text" value={foreignBankCharges} onChange={(e) => { setForeignBankCharges(e.target.value) }} readOnly /></Field>
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
                                            <div>
                                                {invoiceAttachments[index]?.length > 0 ? (
                                                    invoiceAttachments[index].map((file: any, i: number) => (
                                                        <div key={i}>
                                                            <a href={file.ServerRelativeUrl} target="_blank">
                                                                {file.FileName.replace("INV_", "")}
                                                            </a>
                                                        </div>
                                                    ))
                                                ) : (
                                                    <span>-</span>
                                                )}
                                            </div>
                                        </td>

                                        <td>
                                            <div>
                                                {otherAttachments[index]?.map((file: any, i: number) => (
                                                    <div key={i}>
                                                        <a href={file.ServerRelativeUrl} target="_blank">
                                                            {file.FileName.replace("DOC_", "")}
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
                    </Section>
                )}

                {/* ==============================Goods - Advance Payment || Service - Advance Payment============================== */}
                {paymentType === "Service-Advance Payment" && (

                    <Section title="Forex Payment Request Details">
                        <Grid>
                            <Field label="Request Number"><input value={requestNumber} onChange={(e) => { setRequestNumber(e.target.value) }} readOnly /></Field>
                            <Field label="Requested On"><input type="date" value={requestedOn} onChange={(e) => { setRequestedOn(e.target.value) }} readOnly /></Field>
                            <Field label="Currency"><input value={currency} onChange={(e) => { setCurrency(e.target.value) }} readOnly /></Field>
                            <Field label="Total Amount"><input type="number" value={totalAmount} onChange={(e) => { setTotalAmount(e.target.value) }} readOnly /></Field>
                            <Field label="Foreign Bank Charges"><input type="text" value={foreignBankCharges} onChange={(e) => { setForeignBankCharges(e.target.value) }} readOnly /></Field>
                            <Field label="PO/Contract No"><input value={poContractNo} onChange={(e) => { setPoContractNo(e.target.value) }} readOnly /></Field>
                            <Field label="PO Date"><input type="date" value={poDate} onChange={(e) => { setPoDate(e.target.value) }} readOnly /></Field>
                            <Field label="Expected Settlement Date"><input type="date" value={expectedSettlementDate} onChange={(e) => { setExpectedSettlementDate(e.target.value) }} readOnly /></Field>
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
                                                readOnly
                                            />
                                        </td>

                                        <td>
                                            <input
                                                type="date"
                                                value={row.invoiceDate}
                                                onChange={(e) =>
                                                    handleChange(index, "invoiceDate", e.target.value)
                                                }
                                                readOnly
                                            />
                                        </td>

                                        <td>
                                            <input
                                                value={row.invoiceAmount}
                                                onChange={(e) =>
                                                    handleChange(index, "invoiceAmount", e.target.value)
                                                }
                                                readOnly
                                            />
                                        </td>


                                        <td>
                                            <div>
                                                {poAttachments[index]?.map((file: any, i: number) => (
                                                    <div key={i}>
                                                        <a href={file.ServerRelativeUrl} target="_blank">
                                                            {file.FileName.replace("PO_", "")}
                                                        </a>
                                                    </div>
                                                ))}
                                            </div>                                        </td>
                                        <td>
                                            <div>
                                                {piAttachments[index]?.map((file: any, i: number) => (
                                                    <div key={i}>
                                                        <a href={file.ServerRelativeUrl} target="_blank">
                                                            {file.FileName.replace("PI_", "")}
                                                        </a>
                                                    </div>
                                                ))}
                                            </div>                                        </td>

                                        <td>
                                            <div>
                                                {otherAttachments[index]?.map((file: any, i: number) => (
                                                    <div key={i}>
                                                        <a href={file.ServerRelativeUrl} target="_blank">
                                                            {file.FileName.replace("DOC_", "")}
                                                        </a>
                                                    </div>
                                                ))}
                                            </div>                                        </td>




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
                    </Section>
                )}
                {paymentType === "Goods-Advance Payment" && (

                    <Section title="Forex Payment Request Details">
                        <Grid>
                            <Field label="Request Number"><input value={requestNumber} onChange={(e) => { setRequestNumber(e.target.value) }} readOnly /></Field>
                            <Field label="Requested On"><input type="date" value={requestedOn} onChange={(e) => { setRequestedOn(e.target.value) }} readOnly /></Field>
                            <Field label="Currency"><input value={currency} onChange={(e) => { setCurrency(e.target.value) }} readOnly /></Field>
                            <Field label="Total Amount"><input type="number" value={totalAmount} onChange={(e) => { setTotalAmount(e.target.value) }} readOnly /></Field>
                            <Field label="Foreign Bank Charges"><input type="text" value={foreignBankCharges} onChange={(e) => { setForeignBankCharges(e.target.value) }} readOnly /></Field>
                            <Field label="PO/Contract No"><input value={poContractNo} onChange={(e) => { setPoContractNo(e.target.value) }} readOnly /></Field>
                            <Field label="PO Date"><input type="date" value={poDate} onChange={(e) => { setPoDate(e.target.value) }} readOnly /></Field>
                            <Field label="Expected Settlement Date"><input type="date" value={expectedSettlementDate} onChange={(e) => { setExpectedSettlementDate(e.target.value) }} readOnly /></Field>
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
                                                readOnly
                                            />
                                        </td>

                                        <td>
                                            <input
                                                type="date"
                                                value={row.invoiceDate}
                                                onChange={(e) =>
                                                    handleChange(index, "invoiceDate", e.target.value)
                                                }
                                                readOnly
                                            />
                                        </td>

                                        <td>
                                            <input
                                                value={row.invoiceAmount}
                                                onChange={(e) =>
                                                    handleChange(index, "invoiceAmount", e.target.value)
                                                }
                                                readOnly
                                            />
                                        </td>


                                        <td>
                                            <div>
                                                {poAttachments[index]?.map((file: any, i: number) => (
                                                    <div key={i}>
                                                        <a href={file.ServerRelativeUrl} target="_blank">
                                                            {file.FileName.replace("PO_", "")}
                                                        </a>
                                                    </div>
                                                ))}
                                            </div>                                        </td>
                                        <td>
                                            <div>
                                                {piAttachments[index]?.map((file: any, i: number) => (
                                                    <div key={i}>
                                                        <a href={file.ServerRelativeUrl} target="_blank">
                                                            {file.FileName.replace("PI_", "")}
                                                        </a>
                                                    </div>
                                                ))}
                                            </div>                                        </td>

                                        <td>
                                            <div>
                                                {otherAttachments[index]?.map((file: any, i: number) => (
                                                    <div key={i}>
                                                        <a href={file.ServerRelativeUrl} target="_blank">
                                                            {file.FileName.replace("DOC_", "")}
                                                        </a>
                                                    </div>
                                                ))}
                                            </div>                                        </td>




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
                    </Section>
                )}



                {/* ================= CORRESPONDENT ================= */}
                {/* <Section title="Correspondent Bank Details">
                    <Grid>
                        <Field label="Bank Name"><input value={bankname} onChange={(e) => { setBankName(e.target.value) }} readOnly /></Field>
                        <Field label="Swift Code"><input value={bankswiftcode} onChange={(e) => { setBankSwiftCode(e.target.value) }} readOnly /></Field>
                        <Field label="Bank Account No"><input value={bankaccountno} onChange={(e) => { setBankAccountNo(e.target.value) }} readOnly /></Field>
                        <Field label="Remarks" full><textarea rows={3} value={remarks} onChange={(e) => { setRemarks(e.target.value) }} readOnly></textarea></Field>
                        {currentRole === "RM" && (
                            <Field label="RM Remarks" full>
                                <textarea
                                    value={approverRemarks}
                                    onChange={(e) => setApproverRemarks(e.target.value)}
                                />
                            </Field>
                        )}

                        {currentRole === "HOD" && (
                            <Field label="HOD Remarks" full>
                                <textarea
                                    value={approverRemarks}
                                    onChange={(e) => setApproverRemarks(e.target.value)}
                                />
                            </Field>
                        )}


                    </Grid>

                </Section> */}
                <CollapsibleSection title="Workflow History">
                    {workflowHistory.length > 0 ? (
                        <table className="data-table">
                            <thead>
                                <tr>
                                    <th>Action By</th>
                                    {/* <th>Role</th> */}
                                    <th>Action</th>
                                    <th>Remark</th> {/* ✅ NEW COLUMN */}
                                    <th>Date</th>
                                    <th>Status</th>
                                </tr>
                            </thead>
                            <tbody>
                                {workflowHistory.map((item, index) => (
                                    <tr key={index}>
                                        <td>{item.CurrentApprover}</td>   {/* ✅ FIX */}
                                        {/* <td>{item.Role || "-"}</td>       optional */}
                                        <td>{item.ActionTaken}</td>       {/* ✅ FIX */}
                                        <td>{item.Comment}</td>           {/* ✅ FIX */}
                                        <td>
                                            {item.Date
                                                ? new Date(item.Date).toLocaleString("en-GB")
                                                : ""}
                                        </td>
                                        <td>{item.CurrentStatus}</td>     {/* ✅ FIX */}
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    ) : (
                        <p>No workflow history available</p>
                    )}
                </CollapsibleSection>
                {/* ================= VOUCHING DETAILS ================= */}

                {currentRole === "Vouching" && (

                    <Section title="Vouching Details">

                        <Grid>

                            <Field label="Validation Date *">
                                <input
                                    type="date"
                                    value={validationDate}
                                    onChange={(e) => setValidationDate(e.target.value)}
                                />
                            </Field>

                            {/* Hide Voucher Number for Non Advance */}
                            {(paymentType.includes("Advance")) && (
                                <Field label="Voucher Number *">
                                    <input
                                        value={voucherNumber}
                                        onChange={(e) => setVoucherNumber(e.target.value)}
                                    />
                                </Field>
                            )}

                            <Field label="Remarks" full>
                                <textarea
                                    rows={2}
                                    value={vouchingRemarks}
                                    onChange={(e) => setVouchingRemarks(e.target.value)}
                                />
                            </Field>

                        </Grid>

                    </Section>

                )}
                {/* ================= TREASURY VERIFICATION ================= */}

                {currentRole === "TreasuryVerification" && (
                    <Section title="Treasury Verification">
                        <Grid>

                            <Field label="Treasury Remarks" full>
                                <textarea
                                    rows={2}
                                    value={treasuryRemarks}
                                    onChange={(e) => setTreasuryRemarks(e.target.value)}
                                />
                            </Field>

                        </Grid>
                    </Section>
                )}

                {currentRole === "TreasuryPayment" && (

                    <Section title="Payment Details">

                        <Grid>

                            <Field label="Foreign Currency *">
                                <Dropdown
                                    options={foreignCurrencyOptions}
                                    selectedKey={foreignCurrency}
                                    onChange={(e, option) => {
                                        if (option) {
                                            setForeignCurrency(option.key as string);
                                        }
                                    }}
                                />
                            </Field>

                            <Field label="Foreign Currency Amount *">
                                <input
                                    type="number"
                                    value={foreignAmount}
                                    onChange={(e) => setForeignAmount(e.target.value)}
                                />
                            </Field>

                            <Field label="Exchange Rate *">
                                <input
                                    type="number"
                                    value={exchangeRate}
                                    onChange={(e) => setExchangeRate(e.target.value)}
                                />
                            </Field>

                            <Field label="INR Amount *">
                                <input
                                    type="number"
                                    value={inrAmount}
                                    onChange={(e) => setInrAmount(e.target.value)}
                                />
                            </Field>

                        </Grid>

                        <p style={{ color: "red", fontSize: "12px" }}>
                            (if difference in foreign currency & Amount system to display alert message only)
                        </p>

                        <Grid>

                            <Field label="Payment Date *">
                                <input
                                    type="date"
                                    value={paymentDate}
                                    onChange={(e) => setPaymentDate(e.target.value)}
                                />
                            </Field>

                            <Field label="Payment reference number *">
                                <input
                                    value={paymentReference}
                                    onChange={(e) => setPaymentReference(e.target.value)}
                                />
                            </Field>

                            {/* 15CA only for Service */}
                            {isServicePayment && (
                                <Field label="Attach 15CA (Applicable only for Service)">
                                    <input
                                        type="file"
                                        onChange={(e) => setForm15CA(e.target.files?.[0] || null)}
                                    />
                                </Field>
                            )}

                            {/* 15CB only for Service */}
                            {isServicePayment && (
                                <Field label="Attach 15CB (Applicable only for Service)">
                                    <input
                                        type="file"
                                        onChange={(e) => setForm15CB(e.target.files?.[0] || null)}
                                    />
                                </Field>
                            )}

                        </Grid>

                        <Grid>

                            {/* Swift Copy only for Goods */}
                            {isGoodsPayment && (
                                <Field label="Attach Swift Copy (Applicable for Goods)">
                                    <input
                                        type="file"
                                        onChange={(e) => setSwiftCopy(e.target.files?.[0] || null)}
                                    />
                                </Field>
                            )}

                        </Grid>

                    </Section>

                )}
                {(currentRole === "RM" || currentRole === "HOD") && (
                    <Section title=" Remarks Section ">
                        <Grid>

                            <Field label="Approver Remarks" full>
                                <textarea
                                    rows={2}
                                    value={approverRemarks}
                                    onChange={(e) => setApproverRemarks(e.target.value)}
                                />
                            </Field>

                        </Grid>
                    </Section>
                )}
                <div className="button-row">
                    <button className="btn-submit" onClick={onsubmit}>
                        {getApproveButtonText()}
                    </button>
                    <button className="btn-exit" onClick={onSentBack}>Send Back</button>

                    <button className="btn-Reject" onClick={onReject}>Reject</button>
                    <button className="btn-exit" onClick={() => history.goBack()}>Exit</button>
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

export default ApprovalRequestForm;




