import * as React from "react";
// import { useState } from "react";
import "../Pages/Css/VendorReviewForm.scss";
import { IForexModuleProps } from "../IForexModuleProps";
import { sp } from "@pnp/sp/presets/all";
import SPCRUDOPS from "../../service/BAL/spcrud";
import { useEffect, useState } from "react";
import { useParams, useHistory } from "react-router-dom";

const VendorCreationForm: React.FC<IForexModuleProps> = (props) => {
    const [formData, setFormData] = useState({
        natureOfPayment: "",
        taxDocumentAvailable: "Yes",

        peDocumentAvailable: "Yes",
        peDocumentNumber: "",
        peDocumentDate: "",
        peStartDate: "",
        peEndDate: "",

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

    const [peFile, setPeFile] = useState<any>(null);
    const [trcDeclarationFile, setTrcDeclarationFile] = useState<any>(null);
    const [form10FFile, setForm10FFile] = useState<any>(null);
    const [natureOfPaymentOptions, setNatureOfPaymentOptions] = useState<any[]>([]);
    const [vendorInfo, setVendorInfo] = useState<any>(null);
const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);

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
        loadVendorInfo(itemId.Id);
            loadApprovalMatrix();

    }, []);

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

                setVendorInfo(data[0]);
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

            // ==============================
            // SAVE IN VendorMaster
            // ==============================

            const vendorData: any = {
                Title: "Vendor",

                NatureOfPaymentId: parseInt(formData.natureOfPayment),
                TaxDocumentAvailable: formData.taxDocumentAvailable,
                PEDeclaration: formData.peDocumentAvailable,
                CountryOfTaxResidence: formData.countryOfTaxResidence,
                EligibleAmountWithoutWHT: formData.eligibleAmount,
                ApprovedAmountPaidAmount: formData.approvedAmount,
                BalanceEligibleAmount: formData.balanceAmount,
                FromDate: new Date(formData.fromDate),
                ToDate: new Date(formData.toDate),
                DTAAApplicable: "Yes",
                CurrentApproverId: approvalMatrix.length > 0 ? approvalMatrix[0].Approver.Id : null
                

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
                    getUniqueFileName(vendorInvoice),
                    vendorInvoice
                );
            }

            if (trcFile) {
                await item.attachmentFiles.add(
                    getUniqueFileName(trcFile),
                    trcFile
                );
            }

            if (bankFile) {
                await item.attachmentFiles.add(
                    getUniqueFileName(bankFile),
                    bankFile
                );
            }

            if (kycFile) {
                await item.attachmentFiles.add(
                    getUniqueFileName(kycFile),
                    kycFile
                );
            }

            if (otherFile) {
                await item.attachmentFiles.add(
                    getUniqueFileName(otherFile),
                    otherFile
                );
            }

            // ==============================
            // SAVE PE DECLARATION
            // ==============================

            const peResponse = await sp.web.lists
                .getByTitle("VendorTaxDeclaration")
                .items.add({
                    Title: "Permanent Establishment",

                    VendorMasterIdId: Number(itemId.Id),

                    DeclarationType: "Permanent Establishment",

                    DocumentAvailable:
                        formData.peDocumentAvailable,

                    DocumentNumber:
                        formData.peDocumentNumber,

                    DocumentDate:
                        new Date(formData.peDocumentDate) || null,

                    SEPClause: true,

                    ValidityStartDate:
                        new Date(formData.peStartDate) || null,

                    ValidityEndDate:
                        new Date(formData.peEndDate) || null,
                        VendorCodeId: Number(vendorInfo.Id)
                });

            // Upload PE File

            if (peFile) {

                await sp.web.lists
                    .getByTitle("VendorTaxDeclaration")
                    .items.getById(peResponse.data.Id)
                    .attachmentFiles.add(
                        getUniqueFileName(peFile),
                        peFile
                    );
            }

            // ==============================
            // SAVE TRC
            // ==============================

            const trcResponse = await sp.web.lists
                .getByTitle("VendorTaxDeclaration")
                .items.add({
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
                        formData.countryOfTaxResidence,

                    TaxIdentificationNumber:
                        formData.taxIdentificationNumber,

                    ValidityStartDate:
                        formData.trcStartDate || null,

                    ValidityEndDate:
                        formData.trcEndDate || null,
                         VendorCodeId: Number(vendorInfo.Id)
                });

            if (trcDeclarationFile) {

                await sp.web.lists
                    .getByTitle("VendorTaxDeclaration")
                    .items.getById(trcResponse.data.Id)
                    .attachmentFiles.add(
                        getUniqueFileName(trcDeclarationFile),
                        trcDeclarationFile
                    );
            }

            // ==============================
            // SAVE FORM 10F
            // ==============================

            const form10Response = await sp.web.lists
                .getByTitle("VendorTaxDeclaration")
                .items.add({
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

            if (form10FFile) {

                await sp.web.lists
                    .getByTitle("VendorTaxDeclaration")
                    .items.getById(form10Response.data.Id)
                    .attachmentFiles.add(
                        getUniqueFileName(form10FFile),
                        form10FFile
                    );
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

    const eligible =
        Number(updatedForm.eligibleAmount || 0);

    const approved =
        Number(updatedForm.approvedAmount || 0);

    updatedForm.balanceAmount =
        String(eligible - approved);

    setFormData(updatedForm);
};
    return (
        <div className="vendor-tax-container">
            <h2 className="title">Vendor Taxation Document Upload</h2>

            {/* TAX SECTION */}
            <div className="section">
                <h3>Tax & Regulatory Information</h3>
                {
                    vendorInfo && (

                        <div className="section">

                            <h3>Vendor Information bhushan</h3>

                            <div className="grid-4">

                                <div className="form-group">
                                    <label>Oracle Vendor Code</label>
                                    <input
                                        value={vendorInfo.VendorCode || ""}
                                        disabled
                                    />
                                </div>

                                <div className="form-group">
                                    <label>Vendor Name (Legal)</label>
                                    <input
                                        value={vendorInfo.VendorName || ""}
                                        disabled
                                    />
                                </div>

                                <div className="form-group">
                                    <label>Vendor Short Name</label>
                                    <input
                                        value={vendorInfo.VendorShortName || ""}
                                        disabled
                                    />
                                </div>

                                <div className="form-group">
                                    <label>Vendor Type</label>
                                    <input
                                        value={vendorInfo.VendorType || ""}
                                        disabled
                                    />
                                </div>

                            </div>

                            <div className="grid-4">

                                <div className="form-group">
                                    <label>Country</label>
                                    <input
                                        value={vendorInfo.Country?.Country || ""}
                                        disabled
                                    />
                                </div>

                                <div className="form-group">
                                    <label>State</label>
                                    <input
                                        value={vendorInfo.State?.Title || ""}
                                        disabled
                                    />
                                </div>

                                <div className="form-group">
                                    <label>City</label>
                                    <input
                                        value={vendorInfo.City?.City || ""}
                                        disabled
                                    />
                                </div>

                                <div className="form-group">
                                    <label>Currency</label>
                                    <input
                                        value={vendorInfo.Currency?.Currency || ""}
                                        disabled
                                    />
                                </div>

                            </div>

                            <div className="grid-3">

                                <div className="form-group">
                                    <label>Address Line 1</label>
                                    <textarea
                                        value={vendorInfo.VendorAddress || ""}
                                        disabled
                                    />
                                </div>

                                {/* <div className="form-group">
                                    <label>Address Line 2</label>
                                    <textarea
                                        value={vendorInfo.AddressLine2 || ""}
                                        disabled
                                    />
                                </div> */}

                                <div className="form-group">
                                    <label>Postal Code</label>
                                    <input
                                        value={vendorInfo.Pincode || ""}
                                        disabled
                                    />
                                </div>

                            </div>

                        </div>

                    )
                }
                {
                    vendorInfo && (

                        <div className="section">

                            <h3>Contact Information</h3>

                            <div className="grid-4">

                                <div className="form-group">
                                    <label>Contact Person</label>
                                    <input
                                        value={vendorInfo.ContactPersonName || ""}
                                        disabled
                                    />
                                </div>

                                <div className="form-group">
                                    <label>Email</label>
                                    <input
                                        value={vendorInfo.EmailId || ""}
                                        disabled
                                    />
                                </div>

                                <div className="form-group">
                                    <label>Phone Number</label>
                                    <input
                                        value={vendorInfo.PhoneNumber || ""}
                                        disabled
                                    />
                                </div>

                                <div className="form-group">
                                    <label>Alternate Contact</label>
                                    <input
                                        value={vendorInfo.AlternateContact || ""}
                                        disabled
                                    />
                                </div>

                            </div>

                        </div>

                    )
                }
                {
                    vendorInfo && (

                        <div className="section">

                            <h3>Banking Details</h3>

                            <div className="grid-4">

                                <div className="form-group">
                                    <label>Beneficiary Name</label>
                                    <input
                                        value={vendorInfo.BeneficiaryName || ""}
                                        disabled
                                    />
                                </div>

                                <div className="form-group">
                                    <label>Bank Name</label>
                                    <input
                                        value={vendorInfo.BankName || ""}
                                        disabled
                                    />
                                </div>

                                <div className="form-group">
                                    <label>Account Number / IBAN</label>
                                    <input
                                        value={vendorInfo.AccountNumberIBAN || ""}
                                        disabled
                                    />
                                </div>

                                <div className="form-group">
                                    <label>SWIFT / BIC</label>
                                    <input
                                        value={vendorInfo.SWIFTBICCode || ""}
                                        disabled
                                    />
                                </div>

                                <div className="form-group">
                                    <label>IFSC Code</label>
                                    <input
                                        value={vendorInfo.IFSCCode || ""}
                                        disabled
                                    />
                                </div>

                                <div className="form-group">
                                    <label>Routing Number / ABA</label>
                                    <input
                                        value={vendorInfo.RoutingNumberABA || ""}
                                        disabled
                                    />
                                </div>

                            </div>

                        </div>

                    )
                }
                <div className="grid-4">
                    <div className="form-group">
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
                    </div>

                    <div className="form-group">
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
                            onChange={handleChange}
                        >
                            <option>Yes</option>
                            <option>No</option>
                        </select>
                    </div>
                </div>
            </div>

            {/* PE DECLARATION */}
            <div className="section">
                <h3>Permanent Establishment Declaration</h3>

                <div className="grid-5">
                    <div className="form-group">
                        <label>Document Available</label>

                        <select
                            name="peDocumentAvailable"
                            value={formData.peDocumentAvailable}
                            onChange={handleChange}
                        >
                            <option>Yes</option>
                            <option>No</option>
                        </select>
                    </div>

                    <div className="form-group">
                        <label>Document Number</label>

                        <input
                            type="text"
                            name="peDocumentNumber"
                            value={formData.peDocumentNumber}
                            onChange={handleChange}
                        />
                    </div>

                    <div className="form-group">
                        <label>Document Date</label>

                        <input
                            type="date"
                            name="peDocumentDate"
                            value={formData.peDocumentDate}
                            onChange={handleChange}
                        />
                    </div>

                    <div className="form-group">
                        <label>Upload Document</label>
                        <input
                            type="file"
                            onChange={(e: any) => setPeFile(e.target.files[0])}
                        />
                    </div>
                </div>

                <div className="grid-4 mt-20">
                    <div className="checkbox-group">
                        <input type="checkbox" />
                        <label>SEP Clause</label>
                    </div>

                    <div className="form-group">
                        <label>Validity Start Date</label>

                        <input
                            type="date"
                            name="peStartDate"
                            value={formData.peStartDate}
                            onChange={handleChange}
                        />
                    </div>

                    <div className="form-group">
                        <label>Validity End Date</label>

                        <input
                            type="date"
                            name="peEndDate"
                            value={formData.peEndDate}
                            onChange={handleChange}
                        />
                    </div>
                </div>
            </div>

            {/* TRC */}
            <div className="section">
                <h3>Tax Residency Certificate</h3>

                <div className="grid-5">
                    <div className="form-group">
                        <label>Document Available</label>

                        <select
                            name="trcDocumentAvailable"
                            value={formData.trcDocumentAvailable}
                            onChange={handleChange}
                        >
                            <option>Yes</option>
                            <option>No</option>
                        </select>
                    </div>

                    <div className="form-group">
                        <label>Document Number</label>

                        <input
                            type="text"
                            name="trcDocumentNumber"
                            value={formData.trcDocumentNumber}
                            onChange={handleChange}
                        />
                    </div>

                    <div className="form-group">
                        <label>Document Date</label>

                        <input
                            type="date"
                            name="trcDocumentDate"
                            value={formData.trcDocumentDate}
                            onChange={handleChange}
                        />
                    </div>

                    <div className="form-group">
                        <label>Tax Identification Number</label>

                        <input
                            type="text"
                            name="taxIdentificationNumber"
                            value={formData.taxIdentificationNumber}
                            onChange={handleChange}
                        />
                    </div>

                    <div className="form-group">
                        <label>Upload Document</label>
                        <input
                            type="file"
                            onChange={(e: any) => setTrcDeclarationFile(e.target.files[0])}
                        />
                    </div>
                </div>

                <div className="grid-4 mt-20">
                    <div className="form-group">
                        <label>Country of Tax Residence</label>

                        <input
                            type="text"
                            name="countryOfTaxResidence"
                            value={formData.countryOfTaxResidence}
                            onChange={handleChange}
                        />
                    </div>

                    <div className="form-group">
                        <label>Validity Start Date</label>

                        <input
                            type="date"
                            name="trcStartDate"
                            value={formData.trcStartDate}
                            onChange={handleChange}
                        />
                    </div>

                    <div className="form-group">
                        <label>Validity End Date</label>

                        <input
                            type="date"
                            name="trcEndDate"
                            value={formData.trcEndDate}
                            onChange={handleChange}
                        />
                    </div>
                </div>
            </div>

            {/* FORM 10F */}
            <div className="section">
                <h3>Form 10F</h3>

                <div className="grid-5">
                    <div className="form-group">
                        <label>Document Available</label>

                        <select
                            name="form10FDocumentAvailable"
                            value={formData.form10FDocumentAvailable}
                            onChange={handleChange}
                        >
                            <option>Yes</option>
                            <option>No</option>
                        </select>
                    </div>

                    <div className="form-group">
                        <label>Document Number</label>

                        <input
                            type="text"
                            name="form10FDocumentNumber"
                            value={formData.form10FDocumentNumber}
                            onChange={handleChange}
                        />
                    </div>

                    <div className="form-group">
                        <label>Document Date</label>

                        <input
                            type="date"
                            name="form10FDocumentDate"
                            value={formData.form10FDocumentDate}
                            onChange={handleChange}
                        />
                    </div>

                    <div className="form-group">
                        <label>Acknowledgment Number</label>

                        <input
                            type="text"
                            name="acknowledgmentNumber"
                            value={formData.acknowledgmentNumber}
                            onChange={handleChange}
                        />
                    </div>

                    <div className="form-group">
                        <label>Upload Document</label>
                        <input
                            type="file"
                            onChange={(e: any) => setForm10FFile(e.target.files[0])}
                        />
                    </div>
                </div>

                <div className="grid-4 mt-20">
                    <div className="form-group">
                        <label>Validity Start Date</label>

                        <input
                            type="date"
                            name="form10FStartDate"
                            value={formData.form10FStartDate}
                            onChange={handleChange}
                        />
                    </div>

                    <div className="form-group">
                        <label>Validity End Date</label>

                        <input
                            type="date"
                            name="form10FEndDate"
                            value={formData.form10FEndDate}
                            onChange={handleChange}
                        />
                    </div>
                </div>
            </div>

            {/* ELIGIBLE AMOUNT */}
            <div className="section">
                <h3>Threshold Details</h3>

                <div className="grid-3">
                    <div className="form-group">
                        <label>Eligible Amount</label>

                        <input
                            type="number"
                            name="eligibleAmount"
                            value={formData.eligibleAmount}
                            onChange={handleChange}
                        />
                    </div>

                    <div className="form-group">
                        <label>Approved / Paid Amount</label>

                        <input
                            type="number"
                            name="approvedAmount"
                            value={formData.approvedAmount}
                            onChange={handleChange}
                        />
                    </div>

                    <div className="form-group">
                        <label>Balance Amount</label>

                        <input
                            type="number"
                            name="balanceAmount"
                            value={formData.balanceAmount}
                            onChange={handleChange}
                        />
                    </div>
                </div>

                <div className="grid-3 mt-20">
                    <div className="form-group">
                        <label>From Date</label>

                        <input
                            type="date"
                            name="fromDate"
                            value={formData.fromDate}
                            onChange={handleChange}
                        />
                    </div>

                    <div className="form-group">
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

            {/* DOCUMENTS */}
            <div className="section">
                <h3>Compliance & Supporting Documents</h3>

                <div className="grid-2">
                    <div className="form-group">
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

                    <div className="form-group">
                        <label>TRC</label>
                        <input
                            type="file"
                              accept=".pdf,.doc,.docx,.xls,.xlsx,.jpg,.jpeg"
                            onChange={(e: any) => setTrcFile(e.target.files[0])}
                        />
                    </div>

                    <div className="form-group">
                        <label>Bank Confirmation / Cancelled Cheque</label>
                        <input
                            type="file"
                              accept=".pdf,.doc,.docx,.xls,.xlsx,.jpg,.jpeg"
                            onChange={(e: any) => setBankFile(e.target.files[0])}
                        />
                    </div>

                    <div className="form-group">
                        <label>KYC Documents</label>
                        <input
                            type="file"
                              accept=".pdf,.doc,.docx,.xls,.xlsx,.jpg,.jpeg"
                            onChange={(e: any) => setKycFile(e.target.files[0])}
                        />
                    </div>

                    <div className="form-group">
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
            </div>

            {/* BUTTONS */}
            <div className="button-group">
                <button className="submit-btn" onClick={handleSubmit}>Submit</button>
                <button className="exit-btn">Exit</button>
            </div>
        </div>
    );
};

export default VendorCreationForm;