import React, { useEffect, useState } from "react";
import "../Pages/Css/NewRequest.scss";
import { IForexModuleProps } from "../IForexModuleProps";
import { useHistory } from 'react-router-dom';
import { ComboBox, Dropdown, IComboBox, IComboBoxOption, IDropdownOption } from '@fluentui/react';
import SPCRUDOPS from "../../service/BAL/spcrud";
import '../../assets/bootstrap/css/bootstrap.css';
import { Attachment } from "@pnp/sp/attachments";
import { set } from "@microsoft/sp-lodash-subset/lib/index";
import { SPHttpClient, ISPHttpClientOptions } from '@microsoft/sp-http';
import "@pnp/sp/webs";
import "@pnp/sp/files";
import "@pnp/sp/folders";

import logo from "../../assets/sona-comstarlogo.png";

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

interface IApproverDetails {
    Id: number;
    Name: string;
    Role: string;
    Level: number;
}

const Section = ({ title, children }: any) => (
    <div className="form-section">
        <h3>{title}</h3>
        {children}
    </div>
);

const Grid = ({ children }: any) => (
    <div className="form-grid">{children}</div>
);

const Field = ({ label, children, full, required }: any) => (
    <div className={full ? "form-field full" : "form-field"}>
        <label>
            {label}
            {required && <span className="required-star">*</span>}
        </label>
        {children}
    </div>
);
const NewRequest = (props: IForexModuleProps) => {
    const history = useHistory();
    const spCrudOps = SPCRUDOPS();
    const [paymentType, setPaymentType] = useState("Goods-Bill Payment");
    const [taxDocumentView, setTaxDocumentView] = useState("Yes");
    const [paymenttypeDropdownValue, setPaymentTypeDropdownValue] = useState<IDropdownOption>();
    const [currencyOptions, setCurrencyOptions] = useState<IDropdownOption[]>([]);
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
    const [totalAmount, setTotalAmount] = useState(0);
    const [foreignBankCharges, setForeignBankCharges] = useState("");
    const [poContractNo, setPoContractNo] = useState("");
    const [poDate, setPoDate] = useState("");
    const [expectedSettlementDate, setExpectedSettlementDate] = useState("");
    const [incrimentalId, setIncrimentalId] = useState(0);
    const [nextNo, setNextNo] = useState(0);
    const [vendorOptions, setVendorOptions] = useState<IComboBoxOption[]>([]);
    const [approvers, setApprovers] = useState<number[]>([]);
    const [eligibleAmountWithWHT, setEligibleAmountWithWHT] = useState("");
    const [paidAmount, setPaidAmount] = useState("");
    const [ballenceEligibleAmount, setBallenceEligibleAmount] = useState("");
    const [approverDetails, setApproverDetails] = useState<IApproverDetails[]>([]);

    const [boeFiles, setBoeFiles] = useState<{ [key: string]: File[] }>({});
    const [blFiles, setBlFiles] = useState<{ [key: string]: File[] }>({});

    const [invoiceFiles, setInvoiceFiles] = useState<{ [key: number]: File[] }>({});
    const [otherFiles, setOtherFiles] = useState<{ [key: number]: File[] }>({});

    const [poFiles, setPoFiles] = useState<{ [key: number]: File[] }>({});
    const [piFiles, setPiFiles] = useState<{ [key: number]: File[] }>({});
    const [advanceOtherFiles, setAdvanceOtherFiles] = useState<{ [key: number]: File[] }>({});

    const [loading, setLoading] = useState(false);
    const [workflowHistoryData, setWorkflowHistoryData] = useState<any[]>([]);
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

    useEffect(() => {
        setTotalAmount(totalInvoiceAmount);
    }, [totalInvoiceAmount]);

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
        getuserData();
        getFinancialYearStart();
        generateRequestNumber();
        loadVendorOptions();
        getCurrencyData();


    }, [])
    //---------------------------GetCurrencyData----------------------------//
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
                    key: c.Id,
                    text: c.Title
                }));
                setCurrencyOptions(options);
            });
        } catch (error) { console.error("Error fetching currency data:", error); }
    }
    //---------------------------COUNTER FOR REQUEST NUMBER-------------------------//
    const getFinancialYear = () => {
        const today = new Date();
        const year = today.getFullYear();
        const month = today.getMonth(); // 0 = Jan

        const startYear = month < 3 ? year - 1 : year;
        const endYear = startYear + 1;

        const shortStart = startYear.toString().slice(-2);
        const shortEnd = endYear.toString().slice(-2);

        return `${shortStart}-${shortEnd}`;
    };

    const generateRequestNumber = async () => {

        try {
            const sp = await spCrudOps;
            const fy = getFinancialYear();
            const counterItem = await sp.getData("ApplicationNumber", "Id,IDNo", "", "", { column: "ID", isAscending: true }, 1, props);
            if (!counterItem || counterItem.length === 0) {
                console.error("Counter row not found in ApplicationNumber list.");
                return;
            }

            const itemId = counterItem[0].ID;
            const currentNo = Number(counterItem[0].IDNo) || 0;

            const nextNo = currentNo + 1;
            setNextNo(nextNo);
            setIncrimentalId(itemId);
            const paddedNumber = nextNo.toString().padStart(4, "0");

            const formattedNumber = `Forex/${fy}/${paddedNumber}`;

            // await sp.updateData(
            //     "ApplicationNumber",
            //     itemId,
            //     { IDNo: nextNo },
            //     props
            // );
            setRequestNumber(formattedNumber);
        } catch (error) { console.error("Error generating request number:", error); }
    }

    const buildApprovalFlow = async (employeeData: any, selectedType: string) => {
        try {
            const sp = await spCrudOps;

            const baseApprovers: IApproverDetails[] = [];

            if (employeeData.RMId) {
                baseApprovers.push({
                    Id: employeeData.RMId,
                    Name: employeeData.RM,
                    Role: "RM",
                    Level: 1
                });
            }

            if (employeeData.HODId) {
                baseApprovers.push({
                    Id: employeeData.HODId,
                    Name: employeeData.HOD,
                    Role: "HOD",
                    Level: 2
                });
            }

            // 🔥 Map UI type to Backend RequestType
            let requestTypeFilter = "";

            if (
                selectedType === "Goods-Advance Payment" ||
                selectedType === "Service-Advance Payment"
            ) {
                requestTypeFilter = "Advance Payment";
            } else {
                requestTypeFilter = selectedType; // direct match for Bill types
            }

            const matrixData = await sp.getData(
                "ForexApprovalMatrix",
                "Title,Role,Approver/Id,Approver/Title,Level,RequestType",
                "Approver",
                `RequestType eq '${requestTypeFilter}' and Status eq 'Active'`,
                { column: "Level", isAscending: true },
                5000,
                props
            );

            const matrixApprovers = matrixData.map((item: any, index: number) => ({
                Id: item.Approver?.Id,
                Name: item.Approver?.Title,
                Role: item.Role,
                Level: baseApprovers.length + index + 1
            }));

            const fullFlow = [...baseApprovers, ...matrixApprovers];

            const uniqueFlow = fullFlow.filter(
                (value, index, self) =>
                    self.findIndex(v => v.Id === value.Id) === index
            );

            setApproverDetails(uniqueFlow);
            setApprovers(uniqueFlow.map(a => a.Id));

        } catch (error) {
            console.error("Error building approval flow:", error);
        }
    };
    //=----------------------- userdata------------------------//
    const getuserData = async () => {
        (await spCrudOps).getData(
            "EmployeeMaster",
            "EmployeeCode,EmployeeName,Division,Location,RM/EMail,RM/Id,HOD/EMail,HOD/Id,RM/Title,HOD/Title,ContactNo,EmployeeStatus,Email,Employee/Id,Employee/Title",
            "RM,HOD,Employee",
            `EmployeeId eq '${props.context.pageContext.legacyPageContext.userId}'`,
            { column: "ID", isAscending: true },
            1,
            props
        )
            .then((res: any[]) => {
                if (res.length > 0) {
                    const userData = res[0];

                    setEmployee({
                        EmployeeCode: userData.EmployeeCode || "",
                        EmployeeName: userData.EmployeeName || "",
                        Division: userData.Division || "",
                        Location: userData.Location || "",
                        RM: userData.RM?.Title || "",
                        HOD: userData.HOD?.Title || "",
                        ContactNo: userData.ContactNo || "",
                        EmployeeStatus: userData.EmployeeStatus || "",
                        Email: userData.Email || "",
                        RMId: userData.RM?.Id || 0,
                        HODId: userData.HOD?.Id || 0
                    });

                    const rmId = userData.RM?.Id;
                    const hodId = userData.HOD?.Id;

                    const userApprovers = [rmId, hodId]
                        .filter((id): id is number => !!id);

                    const uniqueApprovers = userApprovers.filter(
                        (value, index, self) => self.indexOf(value) === index
                    );

                    setApprovers(uniqueApprovers);

                    // 🔥 NEW: Fetch Roles From Matrix
                    buildApprovalFlow({
                        RMId: userData.RM?.Id,
                        HODId: userData.HOD?.Id,
                        RM: userData.RM?.Title,
                        HOD: userData.HOD?.Title
                    }, paymentType);

                } else {
                    console.log("No user data found for the current email.");
                }
            })
            .catch((error: any) => {
                console.error("Error fetching user data:", error);
            });
    };
    //---------------------------------lOADVENDOR DATA-------------------------//
    const loadVendorOptions = async () => {
        const sp = await spCrudOps;

        try {
            const vendors = await sp.getData(
                "VendorMaster",
                "ID,VendorCode,VendorName",
                "",
                "",
                { column: "VendorCode", isAscending: true },
                5000,
                props
            );

            const options = vendors.map((v: any) => ({
                key: v.VendorCode,
                text: `${v.VendorCode}`
            }));

            setVendorOptions(options);

        } catch (error) {
            console.error("Error loading vendors:", error);
        }
    };

    //----------------------VendorData-------------------------//
    const getVendorData = async (vendorCode: string) => {

        if (!vendorCode) return;

        (await spCrudOps).getData(
            "VendorMaster",
            "VendorCode,VendorName,VendorNameLegal,VendorShortName,VendorType,City/Title,State/Title,Country/Title,Currency/Title,PostalCode,ContactPersonName,EmailId,PhoneNumber,AlternateContact,BeneficiaryName,BankName,AccountNumberIBAN,SWIFTBICCode,RoutingNumberABA,IFSCCode,IntermediaryBank,IntermediarySWIFTCode,NatureOfPayment/Title,PurposeCodeRBI,BankCountry,BankAddress,VendorAddress,BalanceEligibleAmount,ApprovedAmountPaidAmount,EligibleAmountWithoutWHT,TaxDocumentAvailable,DTAAApplicable",
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

    //     const getApprovers = async () => {
    //         try {
    //             (await spCrudOps).getData(
    //                 "ForexApprovalMAtrix",
    //                 "Title,Role,Approver/Id,Approver/Title,Level",
    //                 "Approver",
    //                 `Title eq 'Approver'`,
    //                 { column: "ID", isAscending: true },
    //                 5000,
    //                 props
    //             ).then((res: any[]) => {
    //                 const approverIds = res
    //                     .map(item => item.Approver?.Id)
    //                     .filter((id): id is number => !!id);

    //                setApprovers(prev => {
    //     const newApprovers = approverIds.filter(id => !prev.includes(id));
    //     return [...prev, ...newApprovers];
    // });

    //               //  setApprovers(approverIds);

    //             })

    //         } catch (error) { console.error("Error fetching approvers:", error); }
    //     }




    const uploadToLibrary = async (
        libraryName: string,
        fileName: string,
        file: File,
        metadata: any
    ) => {

        const webUrl = props.context.pageContext.web.absoluteUrl;
        const serverRelativeUrl = props.context.pageContext.web.serverRelativeUrl;
        const folderPath = `${serverRelativeUrl}/${libraryName}`;

        // ================= STEP 1: UPLOAD FILE =================

        const uploadUrl =
            `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${folderPath}')/Files/add(url='${fileName}',overwrite=true)`;

        const uploadOptions: ISPHttpClientOptions = {
            body: file
        };

        const uploadResponse = await props.context.spHttpClient.post(
            uploadUrl,
            SPHttpClient.configurations.v1,
            uploadOptions
        );

        if (!uploadResponse.ok) {
            const errorText = await uploadResponse.text();
            console.error("Upload error:", errorText);
            throw new Error("File upload failed");
        }

        const uploadResult = await uploadResponse.json();

        // ================= STEP 2: GET ENTITY TYPE =================

        const entityTypeResponse = await props.context.spHttpClient.get(
            `${webUrl}/_api/web/lists/getbytitle('${libraryName}')?$select=ListItemEntityTypeFullName`,
            SPHttpClient.configurations.v1
        );

        const entityTypeData = await entityTypeResponse.json();
        const entityTypeName = entityTypeData.ListItemEntityTypeFullName;

        // ================= STEP 3: UPDATE METADATA =================
        // ================= STEP 3: UPDATE METADATA =================

        const itemUrl =
            `${webUrl}/_api/web/GetFileByServerRelativeUrl('${uploadResult.ServerRelativeUrl}')/ListItemAllFields`;

        const updateOptions: ISPHttpClientOptions = {
            headers: {
                "IF-MATCH": "*",
                "X-HTTP-Method": "MERGE",
                "Content-Type": "application/json"
            },
            body: JSON.stringify(metadata)
        };

        const updateResponse = await props.context.spHttpClient.post(
            itemUrl,
            SPHttpClient.configurations.v1,
            updateOptions
        );

        if (!updateResponse.ok) {
            const errorText = await updateResponse.text();
            console.error("Metadata update error:", errorText);
            throw new Error("Metadata update failed");
        }
    };


    const validateForm = () => {

        // ================= HEADER VALIDATION =================
        if (!vendor.VendorCode) {
            alert("Vendor is required.");
            return;
        }

        if (!requestedOn) {
            alert("Requested On is required");
            return false;
        }

        if (!currency) {
            alert("Currency is required");
            return false;
        }

        if (!foreignBankCharges) {
            alert("Foreign Bank Charges is required");
            return false;
        }
        if (paymentType.includes("Advance Payment")) {

            if (!poContractNo) {
                alert("PO/Contract No is required");
                return false;
            }

            if (!poDate) {
                alert("PO Date is required");
                return false;
            }

            if (!expectedSettlementDate) {
                alert("Expected Settlement Date is required");
                return false;
            }

            if (!requestedOn) {
                alert("Requested On is required");
                return false;
            }

            if (!currency) {
                alert("Currency is required");
                return false;
            }

            if (!foreignBankCharges) {
                alert("Foreign Bank Charges is required");
                return false;
            }
        }

        // ================= ROW VALIDATION =================
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

            // ================= GOODS BILL =================
            if (paymentType === "Goods-Bill Payment") {

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
                    alert(`BL No required in row ${i + 1}`);
                    return false;
                }

                if (!row.blDate) {
                    alert(`BL Date required in row ${i + 1}`);
                    return false;
                }

                if (!invoiceFiles[i] || invoiceFiles[i].length === 0) {
                    alert(`Invoice attachment required in row ${i + 1}`);
                    return false;
                }
            }

            // ================= SERVICE BILL =================
            if (paymentType === "Service-Bill Payment") {

                if (!row.mrnNo) {
                    alert(`MRN No required in row ${i + 1}`);
                    return false;
                }

                if (!row.mrnDate) {
                    alert(`MRN Date required in row ${i + 1}`);
                    return false;
                }

                if (!invoiceFiles[i] || invoiceFiles[i].length === 0) {
                    alert(`Invoice attachment required in row ${i + 1}`);
                    return false;
                }
            }

            // ================= ADVANCE =================
            if (
                paymentType === "Service-Advance Payment" ||
                paymentType === "Goods-Advance Payment"
            ) {
                if (!poFiles[i] || poFiles[i].length === 0) {
                    alert(`PO attachment required in row ${i + 1}`);
                    return false;
                }

                if (!piFiles[i] || piFiles[i].length === 0) {
                    alert(`PI attachment required in row ${i + 1}`);
                    return false;
                }
            }
        }

        // ================= UNIQUE BOE FILES =================
        if (paymentType === "Goods-Bill Payment") {

            for (let boe of uniqueBoeNumbers) {
                if (!boeFiles[boe] || boeFiles[boe].length === 0) {
                    alert(`BOE document required for BOE No: ${boe}`);
                    return false;
                }
            }

            for (let bl of uniqueBlNumbers) {
                if (!blFiles[bl] || blFiles[bl].length === 0) {
                    alert(`BL document required for BL No: ${bl}`);
                    return false;
                }
            }
        }

        return true;
    };

    const onsubmit = async () => {

        setLoading(true);

        if (!validateForm()) {
            setLoading(false);
            return;
        }

        try {

            const sp = await spCrudOps;

            // 🔹 1️⃣ Validate Before Saving
            // if (!vendor.VendorCode) {
            //     alert("Vendor is required.");
            //     return;
            // }

            if (rows.length === 0) {
                alert("Please add at least one invoice row.");
                return;
            }
            if (isNaN(Number(totalAmount)) || Number(totalAmount) <= 0) {
                alert("Please enter a valid Total Amount.");
                return;
            }

            const approverslist = approvers || [];

            // Current approver
            const currentApprover =
                approverslist.length > 0 ? approverslist[0] : null;

            // ONLY ONE next approver
            const nextApprover =
                approverslist.length > 1 ? approverslist[1] : null;

            // Save ALL approvers as JSON
            const allApproversJson = JSON.stringify(approverDetails);

            const now = new Date().toLocaleString();

            const workflowHistory = [
                {
                    CurrentApprover: employee.EmployeeName,
                    ActionTaken: "Request Created",
                    Comment: remarks || "",
                    Date: new Date().toISOString(),
                    CurrentStatus: "Submitted"
                }
            ];

            const initialComment =
                `${employee.EmployeeName}: ${remarks || "Request initiated"} [${now}]`;
            // 🔹 2️⃣ Insert Parent
            const parentResponse = await sp.insertData(
                "ForexRequest",
                {
                    ForexType: paymentType,
                    EmployeeCode: employee.EmployeeCode,
                    EmployeeName: employee.EmployeeName,
                    Division: employee.Division,
                    Location: employee.Location,
                    RMId: employee.RMId,
                    HODId: employee.HODId,
                    ContactNo: employee.ContactNo?.toString() || "",
                    Email: employee.Email,
                    BankName: bankname || "",
                    BankAccNo: bankaccountno || "",
                    Remarks: remarks || "",
                    Status: "Pending for RM Approval",
                    NatureOfPayment: paymentType,
                    DocumentIsAvailable: taxDocumentView,
                    DTAAApplicable: dTAAApplicable,
                    ForexNumber: requestNumber,
                    TotalAmount: "" + totalAmount,
                    ForeignBankCharges: (foreignBankCharges) || "",
                    RequestedOn: requestedOn || null,
                    VendorCode: vendor.VendorCode,
                    VendorName: vendor.VendorName,
                    poContractNo: poContractNo || "",
                    poDate: poDate || null,
                    expectedSettlementDate: expectedSettlementDate || null,
                    BankSwiftCode: bankswiftcode || "",
                    CurrentApproverId: currentApprover,
                    // ONLY ONE next approver
                    NextApproversId: { results: nextApprover ? [nextApprover] : [] },                    // JSON column
                    AllApprovers: "" + allApproversJson,
                    EmployeeStatus: employee.EmployeeStatus || "",
                    BallenceEligibleAmount: "" + ballenceEligibleAmount,
                    PaidAmount: "" + paidAmount,
                    EligibleAmountWithWHT: "" + eligibleAmountWithWHT,
                    CurrencyId: currency || 0,
                    CommentHistory: initialComment,
                    WorkFlowHistory: JSON.stringify(workflowHistory),

                },
                props
            );

            const requestId = parentResponse.data.ID;

            console.log("✅ Parent Saved ID:", requestId);

            for (let index = 0; index < rows.length; index++) {

                const row = rows[index];
                if (!row.invoiceNo) continue;

                const childResponse = await sp.insertData(
                    "ForexServicesBillPayment",
                    {
                        ForexIDId: requestId,
                        SrNo: "" + (index + 1),
                        InvoiceNumber: row.invoiceNo || "",
                        InvoiceDate: row.invoiceDate || null,
                        InvoiceAmount: row.invoiceAmount || "",
                        MRNNumber: row.mrnNo || "",
                        MRNDate: row.mrnDate || null,
                        BillofLandingNo: row.blNo || "",
                        BillOfLandingdate: row.blDate || null,
                        BOENo: row.boeNo || "",
                        BOEDate: row.boeDate || null
                    },
                    props
                );

                const childItemId = childResponse.data.ID;
                const webUrl = props.context.pageContext.web.absoluteUrl;

                // ================= BILL MODE =================
                if (paymentType === "Goods-Bill Payment" || paymentType === "Service-Bill Payment") {

                    for (const file of (invoiceFiles[index] || [])) {

                        const fileName = `INV_${row.invoiceNo}_${file.name}`;

                        await props.context.spHttpClient.post(
                            `${webUrl}/_api/web/lists/getbytitle('ForexServicesBillPayment')/items(${childItemId})/AttachmentFiles/add(FileName='${fileName}')`,
                            SPHttpClient.configurations.v1,
                            { body: file }
                        );
                    }

                    for (const file of (otherFiles[index] || [])) {

                        const fileName = `DOC_${row.invoiceNo}_${file.name}`;

                        await props.context.spHttpClient.post(
                            `${webUrl}/_api/web/lists/getbytitle('ForexServicesBillPayment')/items(${childItemId})/AttachmentFiles/add(FileName='${fileName}')`,
                            SPHttpClient.configurations.v1,
                            { body: file }
                        );
                    }
                }

                // ================= ADVANCE MODE =================
                if (paymentType === "Goods-Advance Payment" || paymentType === "Service-Advance Payment") {

                    for (const file of (poFiles[index] || [])) {

                        const fileName = `PO_${row.invoiceNo}_${file.name}`;

                        await props.context.spHttpClient.post(
                            `${webUrl}/_api/web/lists/getbytitle('ForexServicesBillPayment')/items(${childItemId})/AttachmentFiles/add(FileName='${fileName}')`,
                            SPHttpClient.configurations.v1,
                            { body: file }
                        );
                    }

                    for (const file of (piFiles[index] || [])) {

                        const fileName = `PI_${row.invoiceNo}_${file.name}`;

                        await props.context.spHttpClient.post(
                            `${webUrl}/_api/web/lists/getbytitle('ForexServicesBillPayment')/items(${childItemId})/AttachmentFiles/add(FileName='${fileName}')`,
                            SPHttpClient.configurations.v1,
                            { body: file }
                        );
                    }

                    for (const file of (advanceOtherFiles[index] || [])) {

                        const fileName = `DOC_${row.invoiceNo}_${file.name}`;

                        await props.context.spHttpClient.post(
                            `${webUrl}/_api/web/lists/getbytitle('ForexServicesBillPayment')/items(${childItemId})/AttachmentFiles/add(FileName='${fileName}')`,
                            SPHttpClient.configurations.v1,
                            { body: file }
                        );
                    }
                }
            }
            await sp.updateData(
                "ApplicationNumber",
                incrimentalId,
                { IDNo: nextNo },
                props
            );

            // ================= BOE LIBRARY UPLOAD =================

            // ================= BOE MULTIPLE FILE UPLOAD =================
            for (const boeNo of uniqueBoeNumbers) {

                const files = boeFiles[boeNo] || [];

                for (const file of files) {

                    await uploadToLibrary(
                        "BOEAttachments",
                        `${requestId}_${boeNo}_${Date.now()}_${file.name}`,
                        file,
                        {
                            Title: file.name,
                            BOENo: boeNo,
                            ReqeuestId: requestId.toString()
                        }
                    );
                }
            }

            // ================= BILL OF LADING LIBRARY UPLOAD =================
            // ================= BOL MULTIPLE FILE UPLOAD =================

            for (const blNo of uniqueBlNumbers) {

                const files = blFiles[blNo] || [];

                for (const file of files) {

                    await uploadToLibrary(
                        "BillOfLandingAttachment",
                        `${requestId}_${blNo}_${Date.now()}_${file.name}`,
                        file,
                        {
                            Title: file.name,
                            BOLNo: blNo,
                            ReqeuestId: requestId.toString()
                        }
                    );
                }
            }

            alert("✅ Data submitted successfully!");
            history.push("/");


        } catch (error) {

            console.error("❌ Error submitting data:", error);
            alert("Something went wrong. Please check console.");

        }
        finally {

            setLoading(false); // stop loader

        }
    };
    const handledraft = async () => {

        try {

            setLoading(true);

            const sp = await spCrudOps;

            const workflowHistory = [
                {
                    CurrentApprover: employee.EmployeeName,
                    ActionTaken: "Draft Saved",
                    Comment: remarks || "",
                    Date: new Date().toISOString(),
                    CurrentStatus: "Draft"
                }
            ];
            const approverslist = approvers || [];

            // Current approver
            const currentApprover =
                approverslist.length > 0 ? approverslist[0] : null;

            // ONLY ONE next approver
            const nextApprover =
                approverslist.length > 1 ? approverslist[1] : null;

            // Save ALL approvers as JSON
            const allApproversJson = JSON.stringify(approverDetails);

            const parentResponse = await sp.insertData(
                "ForexRequest",
                {
                    ForexType: paymentType,
                    EmployeeCode: employee.EmployeeCode,
                    EmployeeName: employee.EmployeeName,
                    Division: employee.Division,
                    Location: employee.Location,
                    RMId: employee.RMId || null,
                    HODId: employee.HODId || null,
                    ContactNo: employee.ContactNo?.toString() || "",
                    Email: employee.Email,

                    BankName: bankname || "",
                    BankAccNo: bankaccountno || "",
                    BankSwiftCode: bankswiftcode || "",

                    Remarks: remarks || "",

                    Status: "Draft",

                    NatureOfPayment: paymentType,
                    DocumentIsAvailable: taxDocumentView,
                    DTAAApplicable: dTAAApplicable,

                    ForexNumber: requestNumber,

                    TotalAmount: "" + totalAmount,

                    ForeignBankCharges: foreignBankCharges || "",

                    RequestedOn: requestedOn || null,

                    VendorCode: vendor.VendorCode || "",
                    VendorName: vendor.VendorName || "",

                    poContractNo: poContractNo || "",
                    poDate: poDate || null,
                    expectedSettlementDate: expectedSettlementDate || null,

                    EmployeeStatus: employee.EmployeeStatus || "",

                    BallenceEligibleAmount: "" + ballenceEligibleAmount,
                    PaidAmount: "" + paidAmount,
                    EligibleAmountWithWHT: "" + eligibleAmountWithWHT,

                    CurrencyId: currency || null,

                    WorkFlowHistory: JSON.stringify(workflowHistory),

                    // Draft -> no approval flow
                    CurrentApproverId: null,
                    NextApproversId: { results: [] },
                    AllApprovers: "" + allApproversJson

                },
                props
            );

            const requestId = parentResponse.data.ID;

            console.log("Draft Saved ID:", requestId);

            // ================= CHILD ROWS =================

            for (let index = 0; index < rows.length; index++) {

                const row = rows[index];

                // Skip empty rows
                if (
                    !row.invoiceNo &&
                    !row.invoiceDate &&
                    !row.invoiceAmount
                ) continue;

                const childResponse = await sp.insertData(
                    "ForexServicesBillPayment",
                    {
                        ForexIDId: requestId,

                        SrNo: "" + (index + 1),

                        InvoiceNumber: row.invoiceNo || "",
                        InvoiceDate: row.invoiceDate || null,

                        InvoiceAmount: row.invoiceAmount || "",

                        MRNNumber: row.mrnNo || "",
                        MRNDate: row.mrnDate || null,

                        BillofLandingNo: row.blNo || "",
                        BillOfLandingdate: row.blDate || null,

                        BOENo: row.boeNo || "",
                        BOEDate: row.boeDate || null
                    },
                    props
                );

                const childItemId = childResponse.data.ID;

                const webUrl = props.context.pageContext.web.absoluteUrl;

                // ================= INVOICE FILES =================

                // ================= BILL MODE =================
                if (paymentType === "Goods-Bill Payment" || paymentType === "Service-Bill Payment") {
                    if (invoiceFiles[index] && invoiceFiles[index].length > 0) {
                        for (const file of (invoiceFiles[index] || [])) {

                            const fileName = `INV_${row.invoiceNo}_${file.name}`;

                            await props.context.spHttpClient.post(
                                `${webUrl}/_api/web/lists/getbytitle('ForexServicesBillPayment')/items(${childItemId})/AttachmentFiles/add(FileName='${fileName}')`,
                                SPHttpClient.configurations.v1,
                                { body: file }
                            );
                        }
                    }
                    if (otherFiles[index] && otherFiles[index].length > 0) {
                        for (const file of (otherFiles[index] || [])) {

                            const fileName = `DOC_${row.invoiceNo}_${file.name}`;

                            await props.context.spHttpClient.post(
                                `${webUrl}/_api/web/lists/getbytitle('ForexServicesBillPayment')/items(${childItemId})/AttachmentFiles/add(FileName='${fileName}')`,
                                SPHttpClient.configurations.v1,
                                { body: file }
                            );
                        }
                    }
                }
                // ================= ADVANCE MODE =================
                if (paymentType === "Goods-Advance Payment" || paymentType === "Service-Advance Payment") {
                    if (poFiles[index] && poFiles[index].length > 0) {
                        for (const file of (poFiles[index] || [])) {

                            const fileName = `PO_${row.invoiceNo}_${file.name}`;

                            await props.context.spHttpClient.post(
                                `${webUrl}/_api/web/lists/getbytitle('ForexServicesBillPayment')/items(${childItemId})/AttachmentFiles/add(FileName='${fileName}')`,
                                SPHttpClient.configurations.v1,
                                { body: file }
                            );
                        }
                    }
                    if (piFiles[index] && piFiles[index].length > 0) {
                        for (const file of (piFiles[index] || [])) {

                            const fileName = `PI_${row.invoiceNo}_${file.name}`;

                            await props.context.spHttpClient.post(
                                `${webUrl}/_api/web/lists/getbytitle('ForexServicesBillPayment')/items(${childItemId})/AttachmentFiles/add(FileName='${fileName}')`,
                                SPHttpClient.configurations.v1,
                                { body: file }
                            );
                        }
                    }
                    if (advanceOtherFiles[index] && advanceOtherFiles[index].length > 0) {
                        for (const file of (advanceOtherFiles[index] || [])) {

                            const fileName = `DOC_${row.invoiceNo}_${file.name}`;

                            await props.context.spHttpClient.post(
                                `${webUrl}/_api/web/lists/getbytitle('ForexServicesBillPayment')/items(${childItemId})/AttachmentFiles/add(FileName='${fileName}')`,
                                SPHttpClient.configurations.v1,
                                { body: file }
                            );
                        }
                    }
                }

            }
            for (const boeNo of uniqueBoeNumbers) {

                if (!boeFiles[boeNo] || boeFiles[boeNo].length > 0) {
                    const files = boeFiles[boeNo] || [];

                    for (const file of files) {

                        await uploadToLibrary(
                            "BOEAttachments",
                            `${requestId}_${boeNo}_${Date.now()}_${file.name}`,
                            file,
                            {
                                Title: file.name,
                                BOENo: boeNo,
                                ReqeuestId: requestId.toString()
                            }
                        );
                    }
                }
            }

            for (const blNo of uniqueBlNumbers) {
                if (!blFiles[blNo] || blFiles[blNo].length > 0) {

                    const files = blFiles[blNo] || [];

                    for (const file of files) {

                        await uploadToLibrary(
                            "BillOfLandingAttachment",
                            `${requestId}_${blNo}_${Date.now()}_${file.name}`,
                            file,
                            {
                                Title: file.name,
                                BOLNo: blNo,
                                ReqeuestId: requestId.toString()
                            }
                        );
                    }
                }
            }

            alert("Draft saved successfully!");

            history.push("/");

        } catch (error) {

            console.error("Error saving draft:", error);

            alert("Something went wrong while saving draft.");

        } finally {

            setLoading(false);

        }
    };
    const formatDate = (date: any) => {
        if (!date) return "";

        const d = new Date(date);
        const day = String(d.getDate()).padStart(2, "0");
        const month = String(d.getMonth() + 1).padStart(2, "0");
        const year = d.getFullYear();

        return `${day}-${month}-${year}`;
    };

    return (
        <>

            <div className='MainUplodForm' style={{ margin: "5px 0px" }}>
                <div className='row'>
                    <div className='col-md-12'>
                        <div className='Main-Boxpoup'>
                            <div className="bordered">
                                <a><img src={logo} /></a>
                                <h1>Forex Payment Request Form</h1>
                            </div>
                            <div className="approval-ribbon">

                                <div className="ribbon-step initiator">
                                    {employee.EmployeeName}
                                </div>

                                {approverDetails.map((approver, index) => (
                                    <div key={index} className="ribbon-step approver">
                                        {approver.Name}
                                    </div>
                                ))}

                            </div>
                            <div className='borderedbox'>
                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Requestor Information</label>
                                </div>
                                <div className='main-formcontainer'>
                                    <div className='row mb-20'>
                                        <div className='col-md-4'>
                                            <Field label="Type" required>
                                                <select
                                                    value={paymentType} className="form-controltext"
                                                    onChange={(e) => {
                                                        const selected = e.target.value;
                                                        setPaymentType(selected);

                                                        buildApprovalFlow({
                                                            RMId: employee.RMId,
                                                            HODId: employee.HODId,
                                                            RM: employee.RM,
                                                            HOD: employee.HOD
                                                        }, selected);
                                                    }}
                                                >
                                                    <option value="Goods-Bill Payment">Goods-Bill Payment</option>
                                                    <option value="Service-Bill Payment">Service-Bill Payment</option>
                                                    <option value="Goods-Advance Payment">Goods-Advance Payment</option>
                                                    <option value="Service-Advance Payment">Service-Advance Payment</option>
                                                </select>
                                            </Field>
                                            {/* <label className='font'>Type <span className='Mantorystar'>*</span></label>
                                            <select
                                                value={paymentType} className="form-control readonlytext"
                                                onChange={(e) => {
                                                    const selected = e.target.value;
                                                    setPaymentType(selected);
                                                    buildApprovalFlow({
                                                        RMId: employee.RMId,
                                                        HODId: employee.HODId,
                                                        RM: employee.RM,
                                                        HOD: employee.HOD
                                                    }, selected);
                                                }}
                                            >
                                                <option value="Goods-Bill Payment">Goods-Bill Payment</option>
                                                <option value="Service-Bill Payment">Service-Bill Payment</option>
                                                <option value="Goods-Advance Payment">Goods-Advance Payment</option>
                                                <option value="Service-Advance Payment">Service-Advance Payment</option>
                                            </select> */}
                                        </div>
                                    </div>
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
                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Vendor / Beneficiary Details</label>
                                </div>
                                <div className='main-formcontainer'>
                                    <div className='row mb-20'>
                                        <div className='col-md-4'>
                                            <label className='font'>Vendor Code</label>
                                            <Dropdown
                                                placeholder="Select Vendor Code"
                                                className="form-controltext"
                                                options={vendorOptions as IDropdownOption[]}
                                                selectedKey={vendor.VendorCode}
                                                onChange={(event, option) => {
                                                    if (option) {
                                                        const code = option.key as string;
                                                        setVendor(prev => ({ ...prev, VendorCode: code }));
                                                        getVendorData(code);
                                                    }
                                                }}
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
                                            <input type="text" value={vendor.PostalCode} className="form-control readonly" />
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
                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Tax & Regulatory Information</label>
                                </div>
                                <div className='main-formcontainer'>
                                    <div className='row mb-20'>
                                        <div className='col-md-4'>
                                            <label className='font'>Nature of Payment</label>
                                            <input type="text" value={paymentType} className="form-control readonly" />
                                        </div>
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
                                        <div className="heading1" style={{ marginTop: "10px" }}>
                                            <label>Summary of WHT Applicability</label>
                                        </div>
                                        <div className='main-formcontainer'>
                                            <div className='row mb-20'>
                                                <div className="col-md-4">
                                                    <div className="date-summary">
                                                        <span className="label">From</span>
                                                        <span className="value">{formatDate(fromdate)}</span>
                                                        <span className="label">To</span>
                                                        <span className="value">{formatDate(todate)},</span>
                                                        <span className="label"></span>
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

                                        {paymentType === "Goods-Bill Payment" && (
                                            <>
                                                <div className="heading1" style={{ marginTop: "10px" }}>
                                                    <label>Forex Payment Request Details</label>
                                                </div>
                                                <div className='main-formcontainer'>
                                                    <div className='row mb-20'>
                                                        <div className='col-md-4'>
                                                            <label className='font'>Request Number <span className="Mantorystar">*</span></label>
                                                            <input type="text" value={requestNumber} className="form-control readonly" />
                                                        </div>
                                                        <div className='col-md-4'>
                                                            <label className="font">Requested On <span className="Mantorystar">*</span></label>
                                                            <input type="text" value={employee.EmployeeName} className="form-control" />
                                                        </div>
                                                        <div className="col-md-4">
                                                            <label className="font">Currency <span className="Mantorystar">*</span></label>
                                                            <Dropdown
                                                                options={currencyOptions}
                                                                selectedKey={currency}
                                                                onChange={(e, option) => {
                                                                    if (option) {
                                                                        setCurrency(option.key as string);
                                                                    }
                                                                }} className="form-controltext"
                                                            />
                                                        </div>
                                                    </div>
                                                    <div className='row mb-20'>
                                                        <div className='col-md-4'>
                                                            <label className='font'>Total Amount</label>
                                                            <input type="text" value={totalInvoiceAmount} className="form-control readonly" />
                                                        </div>
                                                        <div className="col-md-4">
                                                            <label className="font">Foreign Bank Charges <span className="Mantorystar">*</span></label>
                                                            <select
                                                                className="form-controltext"
                                                                value={foreignBankCharges}
                                                                onChange={(e) => setForeignBankCharges(e.target.value)}
                                                            >
                                                                <option value="">Select</option>
                                                                <option value="Beneficiary">Beneficiary</option>
                                                                <option value="Our">Our</option>
                                                                <option value="Shared">Shared</option>
                                                            </select>
                                                        </div>
                                                    </div>
                                                </div>

                                                <div style={{ overflowX: "auto" }}>
                                                    <table className="custom-table" style={{ marginTop: "10px" }}>
                                                        <thead>
                                                            <tr>
                                                                <th>Sr.No</th>
                                                                <th>Invoice No <span className="required" style={{ color: "red" }}>*</span></th>
                                                                <th>Invoice Date <span className="required" style={{ color: "red" }}>*</span></th>
                                                                <th>BOE No <span className="required" style={{ color: "red" }}>*</span></th>
                                                                <th>BOE Date <span className="required" style={{ color: "red" }}>*</span></th>
                                                                <th>MRN No <span className="required" style={{ color: "red" }}>*</span></th>
                                                                <th>Bill of Lading No <span className="required" style={{ color: "red" }}>*</span></th>
                                                                <th>Bill of Lading Date <span className="required" style={{ color: "red" }}>*</span></th>
                                                                <th>Invoice Amount <span className="required" style={{ color: "red" }}>*</span></th>
                                                                <th>Attach Invoice <span className="required" style={{ color: "red" }}>*</span></th>
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
                                                                        <input
                                                                            type="file"
                                                                            multiple
                                                                            onChange={(e) => {
                                                                                if (e.target.files) {
                                                                                    const filesArray = Array.from(e.target.files);

                                                                                    setInvoiceFiles(prev => ({
                                                                                        ...prev,
                                                                                        [index]: [...(prev[index] || []), ...filesArray]
                                                                                    }));
                                                                                }
                                                                            }}
                                                                        />
                                                                    </td>

                                                                    <td>
                                                                        <input
                                                                            type="file"
                                                                            multiple
                                                                            onChange={(e) => {
                                                                                if (e.target.files) {
                                                                                    const filesArray = Array.from(e.target.files);

                                                                                    setOtherFiles(prev => ({
                                                                                        ...prev,
                                                                                        [index]: [...(prev[index] || []), ...filesArray]
                                                                                    }));
                                                                                }
                                                                            }}
                                                                        />
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
                                                        {/*<tfoot>
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
                                                                    <th>BOE No </th>
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
                                                                                multiple
                                                                                onChange={(e) => {
                                                                                    if (e.target.files) {
                                                                                        const filesArray = Array.from(e.target.files);

                                                                                        setBoeFiles(prev => ({
                                                                                            ...prev,
                                                                                            [boe]: [...(prev[boe] || []), ...filesArray]
                                                                                        }));
                                                                                    }
                                                                                }}
                                                                            />
                                                                        </td>
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
                                                                            <input
                                                                                type="file"
                                                                                multiple
                                                                                onChange={(e) => {
                                                                                    if (e.target.files) {
                                                                                        const filesArray = Array.from(e.target.files);

                                                                                        setBlFiles(prev => ({
                                                                                            ...prev,
                                                                                            [bl]: [...(prev[bl] || []), ...filesArray]
                                                                                        }));
                                                                                    }
                                                                                }}
                                                                            />
                                                                        </td>
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
                                                            <label className='font'>Request Number</label>
                                                            <input type="text" value={requestNumber} className="form-control readonly" />
                                                        </div>
                                                        <div className='col-md-4'>
                                                            <label className="font">Requested On</label>
                                                            <input type="text" value={requestedOn} className="form-control" onChange={(e) => { setRequestedOn(e.target.value) }} />
                                                        </div>
                                                        <div className="col-md-4">
                                                            <label className="font">Currency</label>
                                                            <Dropdown
                                                                options={currencyOptions}
                                                                selectedKey={currency}
                                                                onChange={(e, option) => {
                                                                    if (option) {
                                                                        setCurrency(option.key as string);
                                                                    }
                                                                }} className="form-controltext"
                                                            />
                                                        </div>
                                                    </div>
                                                    <div className='row mb-20'>
                                                        <div className='col-md-4'>
                                                            <label className='font'>Total Amount</label>
                                                            <input type="text" value={totalInvoiceAmount} className="form-control readonly" />
                                                        </div>
                                                        <div className="col-md-4">
                                                            <label className="font">Foreign Bank Charges </label>
                                                            <select
                                                                className="form-controltext"
                                                                value={foreignBankCharges}
                                                                onChange={(e) => setForeignBankCharges(e.target.value)}
                                                            >
                                                                <option value="">Select</option>
                                                                <option value="Beneficiary">Beneficiary</option>
                                                                <option value="Our">Our</option>
                                                                <option value="Shared">Shared</option>
                                                            </select>
                                                        </div>
                                                    </div>
                                                </div>
                                                <div style={{ overflowX: "auto" }}>
                                                    <table className="custom-table" style={{ marginTop: "10px" }}>
                                                        <thead>
                                                            <tr>
                                                                <th>Sr No</th>
                                                                <th>Invoice No <span className="required" style={{ color: "red" }}>*</span></th>
                                                                <th>Invoice Date <span className="required" style={{ color: "red" }}>*</span></th>
                                                                <th>Invoice Amount <span className="required" style={{ color: "red" }}>*</span></th>
                                                                <th>MRN No <span className="required" style={{ color: "red" }}>*</span></th>
                                                                <th>MRN Date <span className="required" style={{ color: "red" }}>*</span></th>
                                                                <th>Attach Invoice <span className="required" style={{ color: "red" }}>*</span></th>
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
                                                                        <input
                                                                            type="file"
                                                                            multiple
                                                                            onChange={(e) => {
                                                                                if (e.target.files) {
                                                                                    const filesArray = Array.from(e.target.files);

                                                                                    setInvoiceFiles(prev => ({
                                                                                        ...prev,
                                                                                        [index]: [...(prev[index] || []), ...filesArray]
                                                                                    }));
                                                                                }
                                                                            }}
                                                                        />
                                                                    </td>

                                                                    <td>
                                                                        <input
                                                                            type="file"
                                                                            multiple
                                                                            onChange={(e) => {
                                                                                if (e.target.files) {
                                                                                    const filesArray = Array.from(e.target.files);

                                                                                    setOtherFiles(prev => ({
                                                                                        ...prev,
                                                                                        [index]: [...(prev[index] || []), ...filesArray]
                                                                                    }));
                                                                                }
                                                                            }}
                                                                        />
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
                                                        {/*  <tfoot>
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
                                                            <label className='font'>Request Number</label>
                                                            <input type="text" value={requestNumber} className="form-control readonly" />
                                                        </div>
                                                        <div className='col-md-4'>
                                                            <label className="font">Requested On</label>
                                                            <input type="text" className="form-control" value={requestedOn} onChange={(e) => { setRequestedOn(e.target.value) }} />
                                                        </div>
                                                        <div className="col-md-4">
                                                            <label className="font">Currency</label>
                                                            <Dropdown
                                                                options={currencyOptions}
                                                                selectedKey={currency}
                                                                onChange={(e, option) => {
                                                                    if (option) {
                                                                        setCurrency(option.key as string);
                                                                    }
                                                                }} className="form-controltext"
                                                            />
                                                        </div>
                                                    </div>
                                                    <div className='row mb-20'>
                                                        <div className='col-md-4'>
                                                            <label className='font'>Total Amount</label>
                                                            <input type="text" value={totalInvoiceAmount} className="form-control readonly" />
                                                        </div>
                                                        <div className="col-md-4">
                                                            <label className="font">Foreign Bank Charges </label>
                                                            <select
                                                                className="form-controltext"
                                                                value={foreignBankCharges}
                                                                onChange={(e) => setForeignBankCharges(e.target.value)}
                                                            >
                                                                <option value="">Select</option>
                                                                <option value="Beneficiary">Beneficiary</option>
                                                                <option value="Our">Our</option>
                                                                <option value="Shared">Shared</option>
                                                            </select>
                                                        </div>
                                                        <div className='col-md-4'>
                                                            <label className='font'>PO/Contract No</label>
                                                            <input type="text" className="form-control" value={poContractNo} onChange={(e) => { setPoContractNo(e.target.value) }} />
                                                        </div>
                                                    </div>
                                                    <div className='row mb-20'>
                                                        <div className='col-md-4'>
                                                            <label className='font'>PO Date</label>
                                                            <input type="date" className="form-control" value={poDate} onChange={(e) => { setPoDate(e.target.value) }} />
                                                        </div>
                                                        <div className="col-md-4">
                                                            <label className="font">Expected Settlement Date </label>
                                                            <input type="date" className="form-control" value={expectedSettlementDate} onChange={(e) => { setExpectedSettlementDate(e.target.value) }} />
                                                        </div>
                                                    </div>
                                                </div>
                                                <div style={{ overflowX: "auto" }}>
                                                    <table className="custom-table" style={{ marginTop: "10px" }}>
                                                        <thead>
                                                            <tr>
                                                                <th>Sr No</th>
                                                                <th>Performa Invoice No <span className="required" style={{ color: "red" }}>*</span></th>
                                                                <th>Performa Invoice Date <span className="required" style={{ color: "red" }}>*</span></th>
                                                                <th>Performa Invoice Amount <span className="required" style={{ color: "red" }}>*</span></th>
                                                                <th>Attach PO <span className="required" style={{ color: "red" }}>*</span></th>
                                                                <th>Attach PI <span className="required" style={{ color: "red" }}>*</span></th>
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
                                                                            multiple
                                                                            onChange={(e) => {
                                                                                if (e.target.files) {
                                                                                    const filesArray = Array.from(e.target.files);

                                                                                    setPoFiles(prev => ({
                                                                                        ...prev,
                                                                                        [index]: [...(prev[index] || []), ...filesArray]
                                                                                    }));
                                                                                }
                                                                            }}
                                                                        />
                                                                    </td>
                                                                    <td>
                                                                        <input
                                                                            type="file"
                                                                            multiple
                                                                            onChange={(e) => {
                                                                                if (e.target.files) {
                                                                                    const filesArray = Array.from(e.target.files);

                                                                                    setPiFiles(prev => ({
                                                                                        ...prev,
                                                                                        [index]: [...(prev[index] || []), ...filesArray]
                                                                                    }));
                                                                                }
                                                                            }}
                                                                        />
                                                                    </td>

                                                                    <td>
                                                                        <input
                                                                            type="file"
                                                                            multiple
                                                                            onChange={(e) => {
                                                                                if (e.target.files) {
                                                                                    const filesArray = Array.from(e.target.files);

                                                                                    setAdvanceOtherFiles(prev => ({
                                                                                        ...prev,
                                                                                        [index]: [...(prev[index] || []), ...filesArray]
                                                                                    }));
                                                                                }
                                                                            }}
                                                                        />
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
                                                        {/*  <tfoot>
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
                                                            <label className='font'>Request Number</label>
                                                            <input type="text" value={requestNumber} className="form-control readonly" />
                                                        </div>
                                                        <div className='col-md-4'>
                                                            <label className="font">Requested On</label>
                                                            <input type="text" className="form-control" value={requestedOn} onChange={(e) => { setRequestedOn(e.target.value) }} />
                                                        </div>
                                                        <div className="col-md-4">
                                                            <label className="font">Currency</label>
                                                            <Dropdown
                                                                options={currencyOptions}
                                                                selectedKey={currency}
                                                                onChange={(e, option) => {
                                                                    if (option) {
                                                                        setCurrency(option.key as string);
                                                                    }
                                                                }} className="form-controltext"
                                                            />
                                                        </div>
                                                    </div>
                                                    <div className='row mb-20'>
                                                        <div className='col-md-4'>
                                                            <label className='font'>Total Amount</label>
                                                            <input type="text" value={totalInvoiceAmount} className="form-control readonly" />
                                                        </div>
                                                        <div className="col-md-4">
                                                            <label className="font">Foreign Bank Charges </label>
                                                            <select
                                                                className="form-controltext"
                                                                value={foreignBankCharges}
                                                                onChange={(e) => setForeignBankCharges(e.target.value)}
                                                            >
                                                                <option value="">Select</option>
                                                                <option value="Beneficiary">Beneficiary</option>
                                                                <option value="Our">Our</option>
                                                                <option value="Shared">Shared</option>
                                                            </select>
                                                        </div>
                                                        <div className='col-md-4'>
                                                            <label className='font'>PO/Contract No</label>
                                                            <input type="text" className="form-control" value={poContractNo} onChange={(e) => { setPoContractNo(e.target.value) }} />
                                                        </div>
                                                    </div>
                                                    <div className='row mb-20'>
                                                        <div className='col-md-4'>
                                                            <label className='font'>PO Date</label>
                                                            <input type="date" className="form-control" value={poDate} onChange={(e) => { setPoDate(e.target.value) }} />
                                                        </div>
                                                        <div className="col-md-4">
                                                            <label className="font">Expected Settlement Date </label>
                                                            <input type="date" className="form-control" value={expectedSettlementDate} onChange={(e) => { setExpectedSettlementDate(e.target.value) }} />
                                                        </div>
                                                    </div>
                                                </div>
                                                <div style={{ overflowX: "auto" }}>
                                                    <table className="custom-table" style={{ marginTop: "10px" }}>
                                                        <thead>
                                                            <tr>
                                                                <th>Sr No</th>
                                                                <th>Performa Invoice No <span className="required" style={{ color: "red" }}>*</span></th>
                                                                <th>Performa Invoice Date <span className="required" style={{ color: "red" }}>*</span></th>
                                                                <th>Performa Invoice Amount <span className="required" style={{ color: "red" }}>*</span></th>
                                                                <th>Attach PO <span className="required" style={{ color: "red" }}>*</span></th>
                                                                <th>Attach PI <span className="required" style={{ color: "red" }}>*</span></th>
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
                                                                            multiple
                                                                            onChange={(e) => {
                                                                                if (e.target.files) {
                                                                                    const filesArray = Array.from(e.target.files);

                                                                                    setPoFiles(prev => ({
                                                                                        ...prev,
                                                                                        [index]: [...(prev[index] || []), ...filesArray]
                                                                                    }));
                                                                                }
                                                                            }}
                                                                        />
                                                                    </td>
                                                                    <td>
                                                                        <input
                                                                            type="file"
                                                                            multiple
                                                                            onChange={(e) => {
                                                                                if (e.target.files) {
                                                                                    const filesArray = Array.from(e.target.files);

                                                                                    setPiFiles(prev => ({
                                                                                        ...prev,
                                                                                        [index]: [...(prev[index] || []), ...filesArray]
                                                                                    }));
                                                                                }
                                                                            }}
                                                                        />
                                                                    </td>

                                                                    <td>
                                                                        <input
                                                                            type="file"
                                                                            multiple
                                                                            onChange={(e) => {
                                                                                if (e.target.files) {
                                                                                    const filesArray = Array.from(e.target.files);

                                                                                    setAdvanceOtherFiles(prev => ({
                                                                                        ...prev,
                                                                                        [index]: [...(prev[index] || []), ...filesArray]
                                                                                    }));
                                                                                }
                                                                            }}
                                                                        />
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
                                                        {/*  <tfoot>
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

                                        <div className='row mb-20'>
                                            <div className='col-md-4'>
                                                <label className='font'>Remarks </label>
                                                <textarea rows={4} cols={4} value={remarks} className="form-control" onChange={(e) => { setRemarks(e.target.value) }}></textarea>
                                            </div>
                                        </div>

                                        <div className='row my-3'>
                                            <div className='col-md-12'>
                                                <div style={{ display: "flex", justifyContent: "center", gap: "5px" }}>
                                                    <button onClick={onsubmit} disabled={loading} className="Submit-btn">
                                                        {loading ? "Submitting..." : "Submit"}
                                                    </button>
                                                    <button onClick={!loading ? handledraft : undefined} className="SendBack-btn">
                                                        {loading ? "Submitting..." : "Save as Draft"}
                                                    </button>
                                                    <button onClick={() => history.push("/")} className="Exit-btn">
                                                        Exit
                                                    </button>
                                                </div>
                                            </div>
                                        </div>

                                    </>
                                )}
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            <div>
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
        </>
    );
};

export default NewRequest;

