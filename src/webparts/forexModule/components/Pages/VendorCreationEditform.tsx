import * as React from "react";
// import { useState } from "react";
import "../Pages/Css/VendorReviewForm.scss";
import { IForexModuleProps } from "../IForexModuleProps";
import { sp } from "@pnp/sp/presets/all";
import SPCRUDOPS from "../../service/BAL/spcrud";
import { useEffect, useState } from "react";
import { useParams, useHistory } from "react-router-dom";
import logo from "../../assets/sona-comstarlogo.png";
import { boolean } from "yup";

const VendorCreationEditForm: React.FC<IForexModuleProps> = (props) => {
    const [formData, setFormData] = useState({
        natureOfPayment: "",
        taxDocumentAvailable: "Yes",

        peDocumentAvailable: "Yes",
        peDocumentNumber: "",
        peDocumentDate: "",
        peStartDate: "",
        peEndDate: "",
        sepClause: false,
        trcDocumentAvailable: "Yes",
        trcDocumentNumber: "",
        trcDocumentDate: "",
        taxIdentificationNumber: "",
        trcStartDate: "",
        trcEndDate: "",
        countryOfTaxResidence: "",

        form10FDocumentAvailable: "Yes",
        form10FDocumentNumber: "",
        form10FDocumentDate: "",
        acknowledgmentNumber: "",
        form10FStartDate: "",
        form10FEndDate: "",

        eligibleAmount: "",
        approvedAmount: "",
        balanceAmount: "",
        fromDate: "",
        toDate: "",
    });
    const itemId = useParams<{ Id: string }>();
    const history = useHistory();

    const spCrudOps = SPCRUDOPS();
    const [vendorInvoice, setVendorInvoice] = useState<any>(null);
    const [trcFile, setTrcFile] = useState<any>(null);
    const [bankFile, setBankFile] = useState<any>(null);
    const [kycFile, setKycFile] = useState<any>(null);
    const [otherFile, setOtherFile] = useState<any>(null);
    // const [peFileAttachments, setPeFileAttachments] = useState<any[]>([]);
    // const [trcFileAttachments, setTrcFileAttachments] = useState<any[]>([]);
    // const [form10FAttachments, setForm10FAttachments] = useState<any[]>([]);
    const [peFile, setPeFile] = useState<any[]>([]);
    const [trcDeclarationFile, setTrcDeclarationFile] = useState<any[]>([]);
    const [form10FFile, setForm10FFile] = useState<any[]>([]);
    const [natureOfPaymentOptions, setNatureOfPaymentOptions] = useState<any[]>([]);
    const [vendorInfo, setVendorInfo] = useState<any>(null);
    const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);
    const [countryOfTaxResidence, setCountryOfTaxResidence] = React.useState("");
    const [countries, setCountries] = React.useState<any[]>([]);
    const [dtaaApplicable, setDTAAApplicable] = React.useState("");
    const [peDeclarationId, setPeDeclarationId] = useState<number>(0);
    const [trcDeclarationId, setTrcDeclarationId] = useState<number>(0);
    const [form10FDeclarationId, setForm10FDeclarationId] = useState<number>(0);
    const [sepClause, setSepClause] = useState<boolean>(false);
    //const [approvalMatrix, setApprovalMatrix] = React.useState<any[]>([]);

    // const handleChange = (
    //     e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement>
    // ) => {
    //     setFormData({
    //         ...formData,
    //         [e.target.name]: e.target.value,
    //     });
    // };
    useEffect(() => {
        loadNatureOfPayment();
        loadcontries();
        loadVendorInfo(itemId.Id);
        loadApprovalMatrix();
        loadTaxDeclaration(itemId.Id);

    }, []);
    const loadTaxDeclaration = async (vendorId: string | undefined) => {
        try {
            if (!vendorId) return;

            const data = await sp.web.lists
                .getByTitle("VendorTaxDeclaration")
                .items
                .select(
                    "*",
                    "AttachmentFiles"
                )
                .expand("AttachmentFiles")
                .filter(`VendorMasterIdId eq ${vendorId}`)
                .get();

            const peData = data.find(
                (x: any) => x.DeclarationType === "Permanent Establishment"
            );

            const trcData = data.find(
                (x: any) => x.DeclarationType === "TAX Residency Certificate"
            );

            const form10fData = data.find(
                (x: any) => x.DeclarationType === "Form 10 F"
            );
            setPeDeclarationId(peData?.Id || 0);
            setTrcDeclarationId(trcData?.Id || 0);
            setForm10FDeclarationId(form10fData?.Id || 0);

            setFormData((prev) => ({
                ...prev,

                peDocumentAvailable: peData?.DocumentAvailable || "No",
                peDocumentNumber: peData?.DocumentNumber || "",
                peDocumentDate: peData?.DocumentDate
                    ? new Date(peData.DocumentDate).toISOString().split("T")[0]
                    : "",
                peStartDate: peData?.ValidityStartDate
                    ? new Date(peData.ValidityStartDate).toISOString().split("T")[0]
                    : "",
                peEndDate: peData?.ValidityEndDate
                    ? new Date(peData.ValidityEndDate).toISOString().split("T")[0]
                    : "",

                trcDocumentAvailable: trcData?.DocumentAvailable || "No",
                trcDocumentNumber: trcData?.DocumentNumber || "",
                trcDocumentDate: trcData?.DocumentDate
                    ? new Date(trcData.DocumentDate).toISOString().split("T")[0]
                    : "",
                taxIdentificationNumber:
                    trcData?.TaxIdentificationNumber || "",
                trcStartDate: trcData?.ValidityStartDate
                    ? new Date(trcData.ValidityStartDate).toISOString().split("T")[0]
                    : "",
                trcEndDate: trcData?.ValidityEndDate
                    ? new Date(trcData.ValidityEndDate).toISOString().split("T")[0]
                    : "",

                form10FDocumentAvailable:
                    form10fData?.DocumentAvailable || "No",
                form10FDocumentNumber:
                    form10fData?.DocumentNumber || "",
                form10FDocumentDate:
                    form10fData?.DocumentDate
                        ? new Date(form10fData.DocumentDate)
                            .toISOString()
                            .split("T")[0]
                        : "",
                acknowledgmentNumber:
                    form10fData?.AcknowledgmentNumber || "",
                form10FStartDate:
                    form10fData?.ValidityStartDate
                        ? new Date(form10fData.ValidityStartDate)
                            .toISOString()
                            .split("T")[0]
                        : "",
                form10FEndDate:
                    form10fData?.ValidityEndDate
                        ? new Date(form10fData.ValidityEndDate)
                            .toISOString()
                            .split("T")[0]
                        : "",
                sepClause: peData?.SEPClause ?? false
            }));
            setPeFile(peData?.AttachmentFiles || []);
            setTrcDeclarationFile(trcData?.AttachmentFiles || []);
            setForm10FFile(form10fData?.AttachmentFiles || []);
            setCountryOfTaxResidence(
                trcData?.CountryOfTaxResidence || ""
            );

        } catch (error) {
            console.log("Load Tax Declaration Error", error);
        }
    };
    const loadcontries = async () => {
        const spx = await spCrudOps;
        const countryData = await spx.getData(
            "CountryMaster",
            "Id,Country",
            "",
            "",
            { column: "Country", isAscending: true },
            5000,
            props
        );
        setCountries(countryData);
    }

    const loadApprovalMatrix = async () => {

        try {

            const spx = await spCrudOps;

            const data = await spx.getData(
                "ForexApprovalMAtrix",
                "ID,Approver/Title,Approver/Id,RequestType,Status",
                "Approver",
                "RequestType eq 'IDT Checker' and Status eq 'Active'",
                { column: "ID", isAscending: true },
                5000,
                props
            );

            console.log("Approval Matrix", data);

            setApprovalMatrix(data);

        } catch (error) {

            console.log(error);

        }
    };

    const loadVendorInfo = async (itemId: string | undefined) => {

        try {

            const spx = await spCrudOps;

            //    const itemId = useParams<{ Id: string }>();

            if (!itemId) return;

            const data = await spx.getData(
                "VendorMaster",
                "*,Country/Country,State/Title,City/City,Currency/Currency",
                "Country,State,City,Currency",
                `Id eq ${itemId}`,
                { column: "Id", isAscending: false },
                1,
                props
            );

            if (data && data.length > 0) {

                console.log("Vendor Info", data[0]);
                setDTAAApplicable(data[0].DTAAApplicable);

                setVendorInfo(data[0]);
                if (data && data.length > 0) {

                    setVendorInfo(data[0]);

                    setFormData((prev) => ({
                        ...prev,
                        taxDocumentAvailable:
                            data[0]?.TaxDocumentAvailable || "Yes"
                    }));

                    setDTAAApplicable(
                        data[0]?.DTAAApplicable || ""
                    );
                }
            }

        } catch (error) {

            console.log("Vendor Load Error", error);

        }
    };
    const loadNatureOfPayment = async () => {

        try {

            sp.setup({
                spfxContext: props.context
            });

            const data = await sp.web.lists
                .getByTitle("NatureOfPaymentMaster")
                .items
                .select("Id", "Title")
                .get();

            setNatureOfPaymentOptions(data);

        } catch (error) {

            console.log("Lookup Load Error", error);

        }
    };
    const getUniqueFileName = (file: any) => {

        const time = new Date().getTime();

        return `${time}_${file.name}`;
    };

    const handleSubmit = async () => {
        try {
            const spsf = await spCrudOps;
            // sp.setup({
            //     spfxContext: props.context
            // });
            // Permanent Establishment Declaration
            if (formData.peDocumentAvailable === "Yes" && formData.taxDocumentAvailable === "Yes") {

                if (!formData.peDocumentNumber?.trim()) {
                    alert("PE Declaration Document Number is mandatory.");
                    return;
                }

                if (!formData.peDocumentDate) {
                    alert("PE Declaration Document Date is mandatory.");
                    return;
                }

                if (!formData.peStartDate) {
                    alert("PE Declaration Validity Start Date is mandatory.");
                    return;
                }

                if (!formData.peEndDate) {
                    alert("PE Declaration Validity End Date is mandatory.");
                    return;
                }

                if (!peFile) {
                    alert("PE Declaration Document Upload is mandatory.");
                    return;
                }
            }

            // Tax Residency Certificate
            if (formData.trcDocumentAvailable === "Yes") {

                if (!formData.trcDocumentNumber?.trim()) {
                    alert("TRC Document Number is mandatory.");
                    return;
                }

                if (!formData.trcDocumentDate) {
                    alert("TRC Document Date is mandatory.");
                    return;
                }

                if (!formData.taxIdentificationNumber?.trim()) {
                    alert("Tax Identification Number is mandatory.");
                    return;
                }

                // if (!countryOfTaxResidence) {
                //     alert("Country Of Tax Residence is mandatory.");
                //     return;
                // }

                if (!formData.trcStartDate) {
                    alert("TRC Validity Start Date is mandatory.");
                    return;
                }

                if (!formData.trcEndDate) {
                    alert("TRC Validity End Date is mandatory.");
                    return;
                }

                if (!trcDeclarationFile) {
                    alert("TRC Upload Document is mandatory.");
                    return;
                }
            }

            // Form 10F
            if (formData.form10FDocumentAvailable === "Yes") {

                if (!formData.form10FDocumentNumber?.trim()) {
                    alert("Form 10F Document Number is mandatory.");
                    return;
                }

                if (!formData.form10FDocumentDate) {
                    alert("Form 10F Document Date is mandatory.");
                    return;
                }

                if (!formData.acknowledgmentNumber?.trim()) {
                    alert("Form 10F Acknowledgment Number is mandatory.");
                    return;
                }

                if (!formData.form10FStartDate) {
                    alert("Form 10F Validity Start Date is mandatory.");
                    return;
                }

                if (!formData.form10FEndDate) {
                    alert("Form 10F Validity End Date is mandatory.");
                    return;
                }

                if (!form10FFile) {
                    alert("Form 10F Upload Document is mandatory.");
                    return;
                }
            }
            if (!dtaaApplicable) {
                alert('DTAA Applicable is mandetory');
                return;
            }
            // ==============================
            // SAVE IN VendorMaster
            // ==============================

            const vendorData: any = {
                Title: "Vendor",

                // NatureOfPaymentId: parseInt(formData.natureOfPayment),
                TaxDocumentAvailable: formData.taxDocumentAvailable,
                PEDeclaration: formData.peDocumentAvailable,
                CountryOfTaxResidence: countryOfTaxResidence,
                EligibleAmountWithoutWHT: formData.eligibleAmount,
                ApprovedAmountPaidAmount: formData.approvedAmount,
                BalanceEligibleAmount: formData.balanceAmount,
                FromDate: new Date(formData.fromDate),
                ToDate: new Date(formData.toDate),
                DTAAApplicable: dtaaApplicable,
                CurrentApproverId: approvalMatrix.length > 0 ? approvalMatrix[0].Approver.Id : null,
                ApprovedByIDTChecker: "Yes",
                RequestStatus: "Pending With IDT"


            };

            const vendorResponse = await spsf.updateData(
                "VendorMaster", Number(itemId.Id), vendorData, props
            )


            // const vendorId = vendorResponse.data.Id;

            // ==============================
            // Upload Attachments
            // ==============================

            const item = sp.web.lists
                .getByTitle("VendorMaster")
                .items.getById(Number(itemId.Id));

            if (vendorInvoice) {
                await item.attachmentFiles.add(
                    `VendorInvoice_${getUniqueFileName(vendorInvoice)}`,
                    vendorInvoice
                );
            }

            if (trcFile) {
                await item.attachmentFiles.add(
                    `TRC_${getUniqueFileName(trcFile)}`,
                    trcFile
                );
            }

            if (bankFile) {
                await item.attachmentFiles.add(
                    `BankConfirmation_${getUniqueFileName(bankFile)}`,
                    bankFile
                );
            }

            if (kycFile) {
                await item.attachmentFiles.add(
                    `KYC_${getUniqueFileName(kycFile)}`,
                    kycFile
                );
            }

            if (otherFile) {
                await item.attachmentFiles.add(
                    `OtherDocument_${getUniqueFileName(otherFile)}`,
                    otherFile
                );
            }

            // ==============================
            // SAVE PE DECLARATION
            // ==============================

            const peResponse = await sp.web.lists
                .getByTitle("VendorTaxDeclaration")
                .items.getById(peDeclarationId)
                .update({
                    Title: "Permanent Establishment",

                    VendorMasterIdId: Number(itemId.Id),

                    DeclarationType: "Permanent Establishment",

                    DocumentAvailable:
                        formData.peDocumentAvailable,

                    DocumentNumber:
                        formData.peDocumentNumber,

                    DocumentDate:
                        new Date(formData.peDocumentDate) || null,

                    SEPClause: sepClause,

                    ValidityStartDate:
                        new Date(formData.peStartDate) || null,

                    ValidityEndDate:
                        new Date(formData.peEndDate) || null,
                    VendorCodeId: Number(vendorInfo.Id)
                });

            // Upload PE File
            if (peFile?.length > 0) {
                for (const file of peFile) {

                    if (!file.name) continue; // skip existing attachments

                    await sp.web.lists
                        .getByTitle("VendorTaxDeclaration")
                        .items.getById(peDeclarationId)
                        .attachmentFiles.add(
                            getUniqueFileName(file),
                            file
                        );
                }
            }

            // ==============================
            // SAVE TRC
            // ==============================

            const trcResponse = await sp.web.lists
                .getByTitle("VendorTaxDeclaration")
                .items.getById(trcDeclarationId)
                .update({
                    Title: "TAX Residency Certificate",

                    VendorMasterIdId: Number(itemId.Id),

                    DeclarationType: "TAX Residency Certificate",

                    DocumentAvailable:
                        formData.trcDocumentAvailable,

                    DocumentNumber:
                        formData.trcDocumentNumber,

                    DocumentDate:
                        formData.trcDocumentDate || null,

                    CountryOfTaxResidence:
                        countryOfTaxResidence,

                    TaxIdentificationNumber:
                        formData.taxIdentificationNumber,

                    ValidityStartDate:
                        formData.trcStartDate || null,

                    ValidityEndDate:
                        formData.trcEndDate || null,
                    VendorCodeId: Number(vendorInfo.Id)
                });
            if (trcDeclarationFile?.length > 0) {
                for (const file of trcDeclarationFile) {

                    if (!file.name) continue; // skip existing attachments

                    await sp.web.lists
                        .getByTitle("VendorTaxDeclaration")
                        .items.getById(trcDeclarationId)
                        .attachmentFiles.add(
                            getUniqueFileName(file),
                            file
                        );
                }
            }

            // ==============================
            // SAVE FORM 10F
            // ==============================

            const form10Response = await sp.web.lists
                .getByTitle("VendorTaxDeclaration")
                .items.getById(form10FDeclarationId)
                .update({
                    Title: "Form 10 F",

                    VendorMasterIdId: Number(itemId.Id),

                    DeclarationType: "Form 10 F",

                    DocumentAvailable:
                        formData.form10FDocumentAvailable,

                    DocumentNumber:
                        formData.form10FDocumentNumber,

                    DocumentDate:
                        formData.form10FDocumentDate || null,

                    AcknowledgmentNumber:
                        formData.acknowledgmentNumber,

                    ValidityStartDate:
                        formData.form10FStartDate || null,

                    ValidityEndDate:
                        formData.form10FEndDate || null,
                    VendorCodeId: Number(vendorInfo.Id)
                });
            if (form10FFile?.length > 0) {
                for (const file of form10FFile) {

                    if (!file.name) continue; // skip existing attachments

                    await sp.web.lists
                        .getByTitle("VendorTaxDeclaration")
                        .items.getById(form10FDeclarationId)
                        .attachmentFiles.add(
                            getUniqueFileName(file),
                            file
                        );
                }
            }



            alert("Vendor Saved Successfully");
            history.push("/VendorCreationDashboard");

        } catch (error) {

            console.log("Save Error", error);
            alert("Error while saving");

        }
    };
    const validateFile = (file: File | null): boolean => {

        if (!file) return true;

        const allowedExtensions = [
            ".pdf",
            ".doc",
            ".docx",
            ".xls",
            ".xlsx",
            ".jpg",
            ".jpeg",
            ".png"
        ];

        const fileName = file.name.toLowerCase();

        const isValid = allowedExtensions.some(
            ext => fileName.endsWith(ext)
        );

        if (!isValid) {

            alert(
                "Only PDF, Word, Excel, JPG, JPEG and PNG files are allowed."
            );

            return false;
        }
        const maxSize = 20 * 1024 * 1024;

        if (file.size > maxSize) {

            alert(
                `File '${file.name}' exceeds the maximum size limit of 20 MB.`
            );

            return false;
        }


        return true;
    };
    const handleChange = (
        e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement>
    ) => {

        const updatedForm = {
            ...formData,
            [e.target.name]: e.target.value
        };
        setFormData(updatedForm);
        // const eligible = Number(updatedForm.eligibleAmount || 0);

        // const approved = Number(updatedForm.approvedAmount || 0);

        // const balanceAmount = eligible - approved;

        // updatedForm.balanceAmount = String(balanceAmount);

        // setFormData(updatedForm);

        // if (balanceAmount < 0) {
        //     alert(
        //         "WHT would be applicable on this transaction as threshold limit has exceeded."
        //     );
        //     return;
        // }
    };
    const handleDeleteFile = async (fileUrl: string, index: number) => {
        try {
            await sp.web
                .getFileByServerRelativePath(fileUrl)
                .recycle();

            setPeFile((prev: any[]) =>
                prev.filter((_, i) => i !== index)
            );
        } catch (error) {
            console.error("Error deleting file:", error);
        }
    };
    const handleDeleteForm10FFile = async (
        fileUrl: string,
        index: number
    ) => {
        try {
            // Delete from SharePoint
            await sp.web
                .getFileByServerRelativePath(fileUrl)
                .recycle();

            // Remove from state
            setForm10FFile((prev: any[]) =>
                prev.filter((_, i) => i !== index)
            );

            alert("File deleted successfully");
        } catch (error) {
            console.error("Error deleting file:", error);
            alert("Failed to delete file");
        }
    };

    const handleDeleteTRCFile = async (
        fileUrl: string,
        index: number
    ) => {
        try {
            // Delete from SharePoint if already uploaded
            await sp.web
                .getFileByServerRelativePath(fileUrl)
                .recycle();

            // Remove from state
            setTrcDeclarationFile((prev: any[]) =>
                prev.filter((_, i) => i !== index)
            );

            alert("TRC file deleted successfully");
        } catch (error) {
            console.error("Error deleting TRC file:", error);
        }
    };
    const isTaxDocAvailable = formData.taxDocumentAvailable === "Yes";
    return (
        <>
            <div className='MainUplodForm' style={{ margin: "5px 0px" }}>
                <div className='row'>
                    <div className='col-md-12'>
                        <div className='Main-Boxpoup'>
                            <div className="bordered">
                                <a><img src={logo} /></a>
                                <h1>Vendor Taxation Document Upload</h1>
                            </div>
                            <div className='borderedbox'>
                                {
                                    vendorInfo && (
                                        <>
                                            <div className="heading1" style={{ marginTop: "10px" }}>
                                                <label>Vendor Information</label>
                                            </div>
                                            <div className='main-formcontainer'>
                                                <div className="row mb-20">

                                                    <div className="col-md-3 form-group">
                                                        <label>Oracle Vendor Code</label>
                                                        <input
                                                            value={vendorInfo.VendorCode || ""}
                                                            disabled
                                                        />
                                                    </div>
                                                    <div className="col-md-3 form-group">
                                                        <label>Oracle Vendor Name</label>
                                                        <input
                                                            value={vendorInfo.VendorName || ""}
                                                            disabled
                                                        />
                                                    </div>

                                                    <div className="col-md-3 form-group">
                                                        <label>Vendor Name (Legal)</label>
                                                        <input
                                                            value={vendorInfo.VendorName || ""}
                                                            disabled
                                                        />
                                                    </div>

                                                    <div className="col-md-3 form-group">
                                                        <label>Vendor Short Name</label>
                                                        <input
                                                            value={vendorInfo.VendorShortName || ""}
                                                            disabled
                                                        />
                                                    </div>

                                                    {/* <div className="col-md-3 form-group">
                                                        <label>Vendor Type</label>
                                                        <input
                                                            value={vendorInfo.VendorType || ""}
                                                            disabled
                                                        />
                                                    </div> */}

                                                </div>
                                                <div className="row mb-20">

                                                    <div className="col-md-3 form-group">
                                                        <label>Country</label>
                                                        <input
                                                            value={vendorInfo.Country?.Country || ""}
                                                            disabled
                                                        />
                                                    </div>

                                                    <div className="col-md-3 form-group">
                                                        <label>State</label>
                                                        <input
                                                            value={vendorInfo.state0 || ""}
                                                            disabled
                                                        />
                                                    </div>

                                                    <div className="col-md-3 form-group">
                                                        <label>City</label>
                                                        <input
                                                            value={vendorInfo.city0 || ""}
                                                            disabled
                                                        />
                                                    </div>

                                                    <div className="col-md-3 form-group">
                                                        <label>Currency</label>
                                                        <input
                                                            value={vendorInfo.Currency?.Currency || ""}
                                                            disabled
                                                        />
                                                    </div>

                                                </div>
                                                <div className="row mb-20">

                                                    <div className="col-md-3 form-group">
                                                        <label>Address Line 1</label>
                                                        <textarea
                                                            value={vendorInfo.VendorAddress || ""}
                                                            disabled
                                                        />
                                                    </div>

                                                    {/* <div className="col-md-3">
                                    <label>Address Line 2</label>
                                    <textarea
                                        value={vendorInfo.AddressLine2 || ""}
                                        disabled
                                    />
                                </div> */}

                                                    <div className="col-md-3 form-group">
                                                        <label>Postal Code</label>
                                                        <input
                                                            value={vendorInfo.Pincode || ""}
                                                            disabled
                                                        />
                                                    </div>

                                                </div>
                                            </div>
                                        </>)
                                }
                                {
                                    vendorInfo && (
                                        <>
                                            <div className="heading1" style={{ marginTop: "10px" }}>
                                                <label>Contact Information</label>
                                            </div>
                                            <div className='main-formcontainer'>
                                                <div className="row mb-20">

                                                    <div className="col-md-3  form-group">
                                                        <label>Contact Person</label>
                                                        <input
                                                            value={vendorInfo.ContactPersonName || ""}
                                                            disabled
                                                        />
                                                    </div>

                                                    <div className="col-md-3  form-group">
                                                        <label>Email</label>
                                                        <input
                                                            value={vendorInfo.EmailId || ""}
                                                            disabled
                                                        />
                                                    </div>

                                                    <div className="col-md-3  form-group">
                                                        <label>Phone Number</label>
                                                        <input
                                                            value={vendorInfo.PhoneNumber || ""}
                                                            disabled
                                                        />
                                                    </div>

                                                    <div className="col-md-3  form-group">
                                                        <label>Alternate Contact</label>
                                                        <input
                                                            value={vendorInfo.AlternateContact || ""}
                                                            disabled
                                                        />
                                                    </div>

                                                </div>
                                            </div>
                                        </>
                                    )
                                }
                                {
                                    vendorInfo && (
                                        <>
                                            <div className="heading1" style={{ marginTop: "10px" }}>
                                                <label>Banking Details</label>
                                            </div>
                                            <div className='main-formcontainer'>
                                                <div className="row mb-20">

                                                    <div className="col-md-3  form-group">
                                                        <label>Beneficiary Name</label>
                                                        <input
                                                            value={vendorInfo.BeneficiaryName || ""}
                                                            disabled
                                                        />
                                                    </div>

                                                    <div className="col-md-3  form-group">
                                                        <label>Bank Name</label>
                                                        <input
                                                            value={vendorInfo.BankName || ""}
                                                            disabled
                                                        />
                                                    </div>
                                                    <div className="col-md-3  form-group">
                                                        <label>Bank Address</label>
                                                        <input
                                                            value={vendorInfo.BankAddress || ""}
                                                            disabled
                                                        />
                                                    </div>

                                                    <div className="col-md-3  form-group">
                                                        <label>Account Number / IBAN</label>
                                                        <input
                                                            value={vendorInfo.AccountNumberIBAN || ""}
                                                            disabled
                                                        />
                                                    </div>

                                                    <div className="col-md-3  form-group">
                                                        <label>SWIFT / BIC</label>
                                                        <input
                                                            value={vendorInfo.SWIFTBICCode || ""}
                                                            disabled
                                                        />
                                                    </div>
                                                    <div className="col-md-3  form-group">
                                                        <label>Routing Number / ABA</label>
                                                        <input
                                                            value={vendorInfo.RoutingNumberABA || ""}
                                                            disabled
                                                        />
                                                    </div>

                                                    {/* <div className="col-md-3  form-group">
                                                        <label>IFSC Code</label>
                                                        <input
                                                            value={vendorInfo.IFSCCode || ""}
                                                            disabled
                                                        />
                                                    </div>


                                                    <div className="col-md-3">
                                                        <label className='font'>Intermediary Bank</label>
                                                        <input
                                                            className="form-control"
                                                            value={vendorInfo.IntermediaryBank}
                                                            // onChange={(e) => setIntermediaryBank(e.target.value)}
                                                            disabled
                                                        />
                                                    </div> */}

                                                </div>
                                            </div>
                                        </>
                                    )
                                }
                                <div className='main-formcontainer' style={{ marginTop: "10px" }}>
                                    <div className="row mb-20">
                                        {/* <div className="col-md-3  form-group">
                                            <label>Nature of Payment</label>

                                            <select
                                                name="natureOfPayment"
                                                value={formData.natureOfPayment}
                                                onChange={handleChange}
                                            >
                                                <option value="">Select</option>

                                                {
                                                    natureOfPaymentOptions.map((item: any) => (
                                                        <option key={item.Id} value={item.Id}>
                                                            {item.Title}
                                                        </option>
                                                    ))
                                                }
                                            </select>
                                        </div> */}

                                        <div className="col-md-3  form-group">
                                            <label>
                                                Whether tax document is available
                                                <span className="red-text">
                                                    {" "}
                                                    (If No, withholding tax applicable)
                                                </span>
                                            </label>

                                            <select
                                                name="taxDocumentAvailable"
                                                value={formData.taxDocumentAvailable}
                                                onChange={(e) => {
                                                    const value = e.target.value;

                                                    setFormData({
                                                        ...formData,
                                                        taxDocumentAvailable: value,

                                                        ...(value === "No" && {
                                                            peDocumentAvailable: "No",
                                                            trcDocumentAvailable: "No",
                                                            form10FDocumentAvailable: "No"
                                                        })
                                                    });
                                                }}
                                            >
                                                <option value="Yes">Yes</option>
                                                <option value="No">No</option>
                                            </select>
                                        </div>
                                        <div className="col-md-3">
                                            <label className='font'>
                                                DTAA Applicable <span style={{ color: "red" }}>*</span>
                                            </label>

                                            <select
                                                className="form-control"
                                                value={dtaaApplicable}
                                                onChange={(e) => setDTAAApplicable(e.target.value)}
                                            >
                                                <option value="">Select</option>
                                                <option value="Yes">Yes</option>
                                                <option value="No">No</option>
                                            </select>
                                        </div>
                                    </div>
                                </div>
                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Permanent Establishment Declaration</label>
                                </div>
                                <div className='main-formcontainer'>
                                    <div className="row mb-20">
                                        <div className="col-md-3  form-group">
                                            <label>Document Available </label>

                                            <select
                                                name="peDocumentAvailable"
                                                value={formData.peDocumentAvailable}
                                                onChange={handleChange}
                                                disabled={!isTaxDocAvailable}
                                            >
                                                <option>Yes</option>
                                                <option>No</option>
                                            </select>
                                        </div>

                                        <div className="col-md-3  form-group">
                                            <label>Document Number </label>

                                            <input
                                                type="text"
                                                name="peDocumentNumber"
                                                value={formData.peDocumentNumber}
                                                onChange={handleChange}
                                                disabled={!isTaxDocAvailable}
                                            />
                                        </div>

                                        <div className="col-md-3  form-group">
                                            <label>Document Date </label>

                                            <input
                                                type="date"
                                                name="peDocumentDate"
                                                value={formData.peDocumentDate}
                                                onChange={handleChange}
                                                disabled={!isTaxDocAvailable}
                                            />
                                        </div>

                                        <div className="col-md-3  form-group">
                                            <label>Upload Document </label>
                                            <input
                                                type="file"
                                                onChange={(e: any) =>
                                                    setPeFile(Array.from(e.target.files || []))
                                                }
                                                disabled={!isTaxDocAvailable}
                                            />
                                            {peFile.length > 0 && (
                                                <div className="mt-2">
                                                    {peFile.map((file: any, index: number) => (
                                                        <div
                                                            key={index}
                                                            className="d-flex align-items-center mb-1"
                                                        >
                                                            <a
                                                                href={file.ServerRelativeUrl || "#"}
                                                                target="_blank"
                                                                rel="noreferrer"
                                                            >
                                                                {file.FileName || file.name}
                                                            </a>

                                                            <button
                                                                type="button"
                                                                className="btn btn-danger btn-sm ms-2"
                                                                onClick={() =>
                                                                    handleDeleteFile(file.ServerRelativeUrl, index)
                                                                }
                                                            >
                                                                Delete
                                                            </button>
                                                        </div>
                                                    ))}
                                                </div>
                                            )}
                                        </div>
                                    </div>
                                    <div className="row mb-20">
                                        <div className="col-md-3  form-group">
                                            <div style={{ display: "flex", gap: "5px", alignItems: "center", justifyContent: "start" }}>
                                                <input type="checkbox" checked={formData.sepClause} onChange={(e) => setSepClause(e.target.checked)} style={{ width: "15px", height: "15px" }} />
                                                <label>SEP Clause</label>
                                            </div>
                                        </div>
                                        <div className="col-md-3  form-group">
                                            <label>Validity Start Date </label>

                                            <input
                                                type="date"
                                                name="peStartDate"
                                                value={formData.peStartDate}
                                                onChange={handleChange}
                                                disabled={!isTaxDocAvailable}

                                            />
                                        </div>
                                        <div className="col-md-3  form-group">
                                            <label>Validity End Date </label>

                                            <input
                                                type="date"
                                                name="peEndDate"
                                                value={formData.peEndDate}
                                                onChange={handleChange}
                                                disabled={!isTaxDocAvailable}
                                            />
                                        </div>
                                    </div>
                                </div>
                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Tax Residency Certificate</label>
                                </div>
                                <div className='main-formcontainer'>
                                    <div className="row mb-20">
                                        <div className="col-md-3  form-group">
                                            <label>Document Available</label>

                                            <select
                                                name="trcDocumentAvailable"
                                                value={formData.trcDocumentAvailable}
                                                onChange={handleChange}
                                                disabled={!isTaxDocAvailable}
                                            >
                                                <option>Yes</option>
                                                <option>No</option>
                                            </select>
                                        </div>

                                        <div className="col-md-3  form-group">
                                            <label>Document Number </label>

                                            <input
                                                type="text"
                                                name="trcDocumentNumber"
                                                value={formData.trcDocumentNumber}
                                                onChange={handleChange}
                                                disabled={!isTaxDocAvailable}
                                            />
                                        </div>

                                        <div className="col-md-3  form-group">
                                            <label>Document Date </label>

                                            <input
                                                type="date"
                                                name="trcDocumentDate"
                                                value={formData.trcDocumentDate}
                                                onChange={handleChange}
                                                disabled={!isTaxDocAvailable}
                                            />
                                        </div>

                                        <div className="col-md-3  form-group">
                                            <label>Tax Identification Number </label>

                                            <input
                                                type="text"
                                                name="taxIdentificationNumber"
                                                value={formData.taxIdentificationNumber}
                                                onChange={handleChange}
                                                disabled={!isTaxDocAvailable}
                                            />
                                        </div>

                                        <div className="col-md-3  form-group">
                                            <label>Upload Document </label>
                                            <input
                                                type="file"
                                                onChange={(e: any) =>
                                                    setTrcDeclarationFile(Array.from(e.target.files || []))
                                                }
                                                //onChange={(e: any) => setTrcDeclarationFile(e.target.files[0])}
                                                disabled={!isTaxDocAvailable}
                                            />
                                            {trcDeclarationFile.map((file: any, index: number) => (
                                                <div key={index}>
                                                    {file.ServerRelativeUrl ? (
                                                        <a
                                                            href={file.ServerRelativeUrl}
                                                            target="_blank"
                                                            rel="noreferrer"
                                                        >
                                                            {file.FileName}
                                                        </a>
                                                    ) : (
                                                        <span>{file.name}</span>
                                                    )}

                                                    <button
                                                        type="button"
                                                        className="btn btn-danger btn-sm ms-2"
                                                        onClick={() => handleDeleteTRCFile(file.ServerRelativeUrl, index)}
                                                    >
                                                        Delete
                                                    </button>
                                                </div>
                                            ))}
                                        </div>


                                        {/* <div className="col-md-3  form-group">
                                            <label>Country of Tax Residence <span style={{ color: "red" }}>*</span></label>

                                            <select
                                                className="form-control"
                                                value={countryOfTaxResidence}
                                                onChange={(e) => setCountryOfTaxResidence(e.target.value)}
                                                disabled={!isTaxDocAvailable}
                                            >
                                                <option value="">Select Country</option>

                                                {countries.map((x: any) => (
                                                    <option
                                                        key={x.Country}
                                                        value={x.Country}
                                                    >
                                                        {x.Country}
                                                    </option>
                                                ))}
                                            </select>
                                        </div> */}

                                        <div className="col-md-3  form-group">
                                            <label>Validity Start Date </label>

                                            <input
                                                type="date"
                                                name="trcStartDate"
                                                value={formData.trcStartDate}
                                                onChange={handleChange}
                                                disabled={!isTaxDocAvailable}
                                            />
                                        </div>

                                        <div className="col-md-3  form-group">
                                            <label>Validity End Date </label>

                                            <input
                                                type="date"
                                                name="trcEndDate"
                                                value={formData.trcEndDate}
                                                onChange={handleChange}
                                                disabled={!isTaxDocAvailable}
                                            />
                                        </div>
                                    </div>
                                </div>
                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Form 10F</label>
                                </div>
                                <div className='main-formcontainer'>
                                    <div className="row mb-20">
                                        <div className="col-md-3  form-group">
                                            <label>Document Available </label>

                                            <select
                                                name="form10FDocumentAvailable"
                                                value={formData.form10FDocumentAvailable}
                                                onChange={handleChange}
                                                disabled={!isTaxDocAvailable}
                                            >
                                                <option>Yes</option>
                                                <option>No</option>
                                            </select>
                                        </div>

                                        <div className="col-md-3  form-group">
                                            <label>Document Number </label>

                                            <input
                                                type="text"
                                                name="form10FDocumentNumber"
                                                value={formData.form10FDocumentNumber}
                                                onChange={handleChange}
                                                disabled={!isTaxDocAvailable}
                                            />
                                        </div>

                                        <div className="col-md-3  form-group">
                                            <label>Document Date </label>

                                            <input
                                                type="date"
                                                name="form10FDocumentDate"
                                                value={formData.form10FDocumentDate}
                                                onChange={handleChange}
                                                disabled={!isTaxDocAvailable}
                                            />
                                        </div>

                                        <div className="col-md-3  form-group">
                                            <label>Acknowledgment Number </label>

                                            <input
                                                type="text"
                                                name="acknowledgmentNumber"
                                                value={formData.acknowledgmentNumber}
                                                onChange={handleChange}
                                                disabled={!isTaxDocAvailable}
                                            />
                                        </div>
                                        <div className="col-md-3  form-group">
                                            <label>Upload Document </label>
                                            <input
                                                type="file"
                                                // onChange={(e: any) => setForm10FFile(e.target.files[0])}
                                                onChange={(e: any) =>
                                                    setForm10FFile(Array.from(e.target.files || []))
                                                }
                                                disabled={!isTaxDocAvailable}
                                            />
                                            {form10FFile.map((file: any, index: number) => (
                                                <div key={index}>
                                                    {file.ServerRelativeUrl ? (
                                                        <a
                                                            href={file.ServerRelativeUrl}
                                                            target="_blank"
                                                            rel="noreferrer"
                                                        >
                                                            {file.FileName}
                                                        </a>
                                                    ) : (
                                                        <span>{file.name}</span>
                                                    )}

                                                    <button
                                                        type="button"
                                                        className="btn btn-danger btn-sm ms-2"
                                                        onClick={() =>
                                                            handleDeleteForm10FFile(file.ServerRelativeUrl, index)
                                                        }
                                                    >
                                                        Delete
                                                    </button>
                                                </div>
                                            ))}
                                        </div>

                                        <div className="col-md-3  form-group">
                                            <label>Validity Start Date </label>

                                            <input
                                                type="date"
                                                name="form10FStartDate"
                                                value={formData.form10FStartDate}
                                                onChange={handleChange}
                                                disabled={!isTaxDocAvailable}
                                            />
                                        </div>

                                        <div className="col-md-3  form-group">
                                            <label>Validity End Date </label>

                                            <input
                                                type="date"
                                                name="form10FEndDate"
                                                value={formData.form10FEndDate}
                                                onChange={handleChange}
                                                disabled={!isTaxDocAvailable}
                                            />
                                        </div>
                                    </div>
                                </div>
                                {/* <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Threshold Details</label>
                                </div>
                                <div className='main-formcontainer'>
                                    <div className="row mb-20">
                                        <div className="col-md-3  form-group">
                                            <label>Eligible Amount</label>

                                            <input
                                                type="number"
                                                name="eligibleAmount"
                                                value={formData.eligibleAmount}
                                                onChange={handleChange}
                                            />
                                        </div>

                                        <div className="col-md-3  form-group">
                                            <label>Approved / Paid Amount</label>

                                            <input
                                                type="number"
                                                name="approvedAmount"
                                                value={formData.approvedAmount}
                                                onChange={handleChange}
                                            />
                                        </div>

                                        <div className="col-md-3  form-group">
                                            <label>Balance Amount</label>

                                            <input
                                                type="number"
                                                name="balanceAmount"
                                                value={formData.balanceAmount}
                                                onChange={handleChange}
                                            />
                                        </div>

                                        <div className="col-md-3  form-group">
                                            <label>From Date</label>

                                            <input
                                                type="date"
                                                name="fromDate"
                                                value={formData.fromDate}
                                                onChange={handleChange}
                                            />
                                        </div>

                                        <div className="col-md-3  form-group">
                                            <label>To Date</label>

                                            <input
                                                type="date"
                                                name="toDate"
                                                value={formData.toDate}
                                                onChange={handleChange}
                                            />
                                        </div>
                                    </div>
                                </div>
                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Compliance & Supporting Documents</label>
                                </div>
                                <div className='main-formcontainer'>
                                    <div className="row mb-20">
                                        <div className="col-md-4  form-group">
                                            <label>Vendor Invoice / Proforma Invoice</label>
                                            <input
                                                type="file"
                                                accept=".pdf,.doc,.docx,.xls,.xlsx,.jpg,.jpeg"
                                                onChange={(e: any) => {
                                                    if (validateFile(e.target.files[0])) {
                                                        setVendorInvoice(e.target.files[0]);
                                                    }
                                                }}
                                            />

                                        </div>

                                        <div className="col-md-4  form-group">
                                            <label>TRC</label>
                                            <input
                                                type="file"
                                                accept=".pdf,.doc,.docx,.xls,.xlsx,.jpg,.jpeg"
                                                onChange={(e: any) => setTrcFile(e.target.files[0])}
                                            />
                                        </div>

                                        <div className="col-md-4  form-group">
                                            <label>Bank Confirmation / Cancelled Cheque</label>
                                            <input
                                                type="file"
                                                accept=".pdf,.doc,.docx,.xls,.xlsx,.jpg,.jpeg"
                                                onChange={(e: any) => setBankFile(e.target.files[0])}
                                            />
                                        </div>

                                        <div className="col-md-4 form-group">
                                            <label>KYC Documents</label>
                                            <input
                                                type="file"
                                                accept=".pdf,.doc,.docx,.xls,.xlsx,.jpg,.jpeg"
                                                onChange={(e: any) => setKycFile(e.target.files[0])}
                                            />
                                        </div>

                                        <div className="col-md-4  form-group">
                                            <label>Other Documents</label>
                                            <input
                                                type="file"
                                                accept=".pdf,.doc,.docx,.xls,.xlsx,.jpg,.jpeg"
                                                onChange={(e: any) => {
                                                    if (validateFile(e.target.files[0])) {
                                                        setOtherFile(e.target.files[0]);
                                                    }
                                                }}
                                            />
                                        </div>
                                    </div>
                                    <p className="mandatory-text">
                                        ** Mandatory taxation document upload
                                    </p>
                                </div> */}
                                <div style={{ margin: "10px", display: "flex", justifyContent: "center", gap: "5px", alignItems: "center" }}>
                                    <a className="Submit-btn" onClick={handleSubmit}> Submit </a>
                                    <a className="Exit-btn" onClick={() => history.push("/VendorCreationDashboard")}>Exit</a>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

        </>
    );
};

export default VendorCreationEditForm;