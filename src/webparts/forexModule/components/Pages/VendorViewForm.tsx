import * as React from "react";
import { useEffect, useState } from "react";
import "../Pages/Css/VendorReviewForm.scss";
import { IForexModuleProps } from "../IForexModuleProps";
import { sp } from "@pnp/sp/presets/all";
import { useHistory, useParams } from "react-router-dom";
import SPCRUDOPS from "../../service/BAL/spcrud";
// import { useHistory } from "react-router-dom";

import logo from "../../assets/sona-comstarlogo.png";
import view from "../../assets/Eye.png";

const VendorViewForm: React.FC<IForexModuleProps> = (props) => {

    const [vendorData, setVendorData] = useState<any>({});
    const [taxDeclarations, setTaxDeclarations] = useState<any[]>([]);
    const [remarks, setRemarks] = useState("");
    const [loading, setLoading] = useState(false);
    const [peData, setPeData] = useState<any>(null);
    const [trcData, setTrcData] = useState<any>(null);
    const [form10FData, setForm10FData] = useState<any>(null);
    const spCrudOps = SPCRUDOPS();
    const history = useHistory();
    const vendorId = useParams<{ Id: string }>();
    useEffect(() => {

        loadVendorData();

    }, []);
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
    const loadVendorData = async () => {

        try {

            setLoading(true);

            const spx = await spCrudOps;

            // =========================
            // Vendor Master
            // =========================

            const vendorData = await spx.getData(
                "VendorMaster",
                "*,NatureOfPayment/Title,Country/Country,State/Title,City/City,Currency/Currency,AttachmentFiles",
                "NatureOfPayment,Country,State,City,Currency,AttachmentFiles",
                `Id eq ${Number(vendorId.Id)}`,
                { column: "Id", isAscending: false },
                1,
                props
            );

            console.log("Vendor Data", vendorData);

            if (vendorData && vendorData.length > 0) {

                setVendorData(vendorData[0]);

            }

            // =========================
            // Tax Declaration
            // =========================

            const declarations = await spx.getData(
                "VendorTaxDeclaration",
                "*,AttachmentFiles",
                "AttachmentFiles",
                `VendorMasterIdId eq ${Number(vendorId.Id)}`,
                { column: "Id", isAscending: false },
                5000,
                props
            );

            console.log("Declarations", declarations);

            console.log("Declarations", declarations);

            const pe = declarations.find(
                (x: any) => x.DeclarationType === "Permanent Establishment"
            );

            const trc = declarations.find(
                (x: any) => x.DeclarationType === "TAX Residency Certificate"
            );

            const form10f = declarations.find(
                (x: any) => x.DeclarationType === "Form 10 F"
            );

            setPeData(pe);
            setTrcData(trc);
            setForm10FData(form10f);

            setTaxDeclarations(declarations || []);

        } catch (error) {

            console.log("Load Error", error);

        } finally {

            setLoading(false);

        }
    };



    if (loading) {

        return <div>Loading...</div>;
    }
    const formatDate = (date: any) => {
        if (!date) return "";
        return new Date(date).toLocaleDateString("en-GB");
    };

    return (

        <>

            <div className='MainUplodForm' style={{ margin: "5px 0px" }}>
                <div className='row'>
                    <div className='col-md-12'>
                        <div className='Main-Boxpoup'>
                            <div className="bordered">
                                <a><img src={logo} /></a>
                                <h1>Vendor Checker Form</h1>
                            </div>
                            <div className='borderedbox'>
                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label> Vendor Basic Details</label>
                                </div>
                                <div className='main-formcontainer'>
                                    <div className="row mb-20">
                                        <div className="col-md-3">
                                            <label className="font fontblock">Oracle Vendor Code</label>
                                            <span>
                                                {vendorData?.VendorCode}
                                            </span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">Oracle Vendor Name</label>
                                            <span>
                                                {vendorData?.VendorName}
                                            </span>
                                        </div>
                                        <div className="col-md-3">
                                            <label className="font fontblock">
                                                Vendor Name (Legal)
                                            </label>

                                            <span>
                                                {vendorData?.VendorNameLegal}
                                            </span>
                                        </div>

                                        {/* <div className="col-md-3">
                                            <label className="font fontblock">
                                                Vendor Short Name
                                            </label>

                                            <span>
                                                {vendorData?.VendorShortName}
                                            </span>
                                        </div> */}

                                        {/* <div className="col-md-3">
                                            <label className="font fontblock">
                                                Vendor Type
                                            </label>

                                            <span>
                                                {vendorData?.VendorType}
                                            </span>
                                        </div> */}
                                        <div className="col-md-3">
                                            <label className="font fontblock">
                                                Currency
                                            </label>

                                            <span>
                                                {vendorData?.Currency?.Currency}
                                            </span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">
                                                Country
                                            </label>

                                            <span>
                                                {vendorData?.Country?.Country}
                                            </span>
                                        </div>
                                        <div className="col-md-3">
                                            <label className="font fontblock">
                                                State
                                            </label>

                                            <span>
                                                {vendorData?.state0}
                                            </span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">
                                                City
                                            </label>

                                            <span>
                                                {vendorData?.city0}
                                            </span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">
                                                Address Line 1
                                            </label>

                                            <span>
                                                {vendorData?.VendorAddress}
                                            </span>
                                        </div>

                                        {/* <div className="col-md-3">
                        <label className= "font fontblock">
                            Address Line 2
                        </label>

                        <span>
                            {vendorData?.VendorAddress}
                        </span>
                    </div> */}




                                        <div className="col-md-3">
                                            <label className="font fontblock">
                                                Postal Code
                                            </label>

                                            <span>
                                                {vendorData?.Pincode}
                                            </span>
                                        </div>

                                    </div>
                                </div>
                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label> Contact Information</label>
                                </div>
                                <div className='main-formcontainer'>
                                    <div className="row mb-20">

                                        <div className="col-md-3">
                                            <label className="font fontblock">
                                                Contact Person Name
                                            </label>

                                            <span>
                                                {vendorData?.ContactPersonName}
                                            </span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">
                                                Email Id
                                            </label>

                                            <span>
                                                {vendorData?.EmailId}
                                            </span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">
                                                Phone Number
                                            </label>

                                            <span>
                                                {vendorData?.PhoneNumber}
                                            </span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">
                                                Alternate Contact
                                            </label>

                                            <span>
                                                {vendorData?.AlternateContact}
                                            </span>
                                        </div>

                                    </div>
                                </div>
                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Banking Details</label>
                                </div>
                                <div className='main-formcontainer'>
                                    <div className="row mb-20">

                                        <div className="col-md-3">
                                            <label className="font fontblock">
                                                Beneficiary Name
                                            </label>

                                            <span>
                                                {vendorData?.BeneficiaryName}
                                            </span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">
                                                Bank Name
                                            </label>

                                            <span>
                                                {vendorData?.BankName}
                                            </span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">
                                                Bank Address
                                            </label>

                                            <span>
                                                {vendorData?.BankAddress}
                                            </span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">
                                                Account Number / IBAN
                                            </label>

                                            <span>
                                                {vendorData?.AccountNumberIBAN}
                                            </span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">
                                                SWIFT / BIC Code
                                            </label>

                                            <span>
                                                {vendorData?.SWIFTBICCode}
                                            </span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">
                                                Routing Number / ABA
                                            </label>

                                            <span>
                                                {vendorData?.RoutingNumberABA}
                                            </span>
                                        </div>

                                        {/* <div className="col-md-3">
                                            <label className="font fontblock">
                                                IFSC Code
                                            </label>

                                            <span>
                                                {vendorData?.IFSCCode}
                                            </span>
                                        </div> */}

                                        {/* <div className="col-md-3">
                                            <label className="font fontblock">
                                                Intermediary Bank
                                            </label>

                                            <span>
                                                {vendorData?.IntermediaryBank}
                                            </span>
                                        </div> */}

                                        {/* <div className="col-md-3">
                        <label className= "font fontblock">
                            Intermediary SWIFT Code
                        </label>

                        <span>
                            {vendorData?.IntermediarySWIFTCode}
                        </span>
                    </div> */}

                                    </div>
                                </div>
                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Tax & Regulatory Information</label>
                                </div>
                                <div className='main-formcontainer'>
                                    <div className="row mb-20">

                                        <div className="col-md-3">
                                            <label className="font fontblock">
                                                Nature Of Payment
                                            </label>

                                            <span>
                                                {
                                                    vendorData?.NatureOfPayment?.Title
                                                }
                                            </span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">
                                                Tax Document Available
                                            </label>

                                            <span>
                                                {
                                                    vendorData?.TaxDocumentAvailable
                                                }
                                            </span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">
                                                DTAA Applicable
                                            </label>

                                            <span>
                                                {
                                                    vendorData?.DTAAApplicable
                                                }
                                            </span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">
                                                Country Of Tax Residence
                                            </label>

                                            <span>
                                                {
                                                    vendorData?.CountryOfTaxResidence
                                                }
                                            </span>
                                        </div>

                                    </div>
                                </div>
                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Permanent Establishment Declaration</label>
                                </div>
                                <div className='main-formcontainer'>
                                    <div className="row mb-20">
                                        <div className="col-md-3">
                                            <label className="font fontblock">Document Available</label>
                                            <span>{peData?.DocumentAvailable}</span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">Document Number</label>
                                            <span>{peData?.DocumentNumber}</span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">Document Date</label>
                                            <span>
                                                {formatDate(peData?.DocumentDate)}
                                            </span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">SEP Clause</label>
                                            <span>{peData?.SEPClause ? "Yes" : "No"}</span>
                                        </div>
                                        <div className="col-md-3">
                                            <label className="font fontblock">Uploaded Document</label>

                                            {peData?.AttachmentFiles?.length > 0 ? (
                                                peData.AttachmentFiles.map((file: any, index: number) => (
                                                    <div key={index}>
                                                        <a
                                                            href={file.ServerRelativeUrl}
                                                            target="_blank"
                                                            rel="noreferrer"
                                                        >
                                                            {file.FileName}
                                                        </a>
                                                    </div>
                                                ))
                                            ) : (
                                                <span>No Document Uploaded</span>
                                            )}
                                        </div>
                                        <div className="col-md-3">
                                            <label className="font fontblock">Validity Start Date</label>
                                            <span>
                                                {peData?.ValidityStartDate
                                                    ? new Date(peData.ValidityStartDate).toLocaleDateString()
                                                    : ""}
                                            </span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">Validity End Date</label>
                                            <span>
                                                {peData?.ValidityEndDate
                                                    ? new Date(peData.ValidityEndDate).toLocaleDateString()
                                                    : ""}
                                            </span>
                                        </div>
                                    </div>
                                </div>
                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Tax Residency Certificate</label>
                                </div>
                                <div className='main-formcontainer'>
                                    <div className="row mb-20">
                                        <div className="col-md-3">
                                            <label className="font fontblock">Document Number</label>
                                            <span>{trcData?.DocumentNumber}</span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">Country Of Tax Residence</label>
                                            <span>{trcData?.CountryOfTaxResidence}</span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">TIN</label>
                                            <span>{trcData?.TaxIdentificationNumber}</span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">Document Available</label>
                                            <span>{trcData?.DocumentAvailable}</span>
                                        </div>
                                        <div className="col-md-3">
                                            <label className="font fontblock">Tax Identification Number</label>
                                            <span>{trcData?.TaxIdentificationNumber}</span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">Validity Start Date</label>
                                            <span>
                                                {trcData?.ValidityStartDate
                                                    ? new Date(trcData.ValidityStartDate).toLocaleDateString()
                                                    : ""}
                                            </span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">Validity End Date</label>
                                            <span>
                                                {trcData?.ValidityEndDate
                                                    ? new Date(trcData.ValidityEndDate).toLocaleDateString()
                                                    : ""}
                                            </span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">Uploaded Document</label>

                                            {trcData?.AttachmentFiles?.length > 0 ? (
                                                trcData.AttachmentFiles.map((file: any, index: number) => (
                                                    <div key={index}>
                                                        <a
                                                            href={file.ServerRelativeUrl}
                                                            target="_blank"
                                                            rel="noreferrer"
                                                        >
                                                            {file.FileName}
                                                        </a>
                                                    </div>
                                                ))
                                            ) : (
                                                <span>No Document Uploaded</span>
                                            )}
                                        </div>
                                    </div>
                                </div>
                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Form 10F</label>
                                </div>
                                <div className='main-formcontainer'>
                                    <div className="row mb-20">

                                        <div className="col-md-3">
                                            <label className="font fontblock">Document Number</label>
                                            <span>{form10FData?.DocumentNumber}</span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">Acknowledgment Number</label>
                                            <span>{form10FData?.AcknowledgmentNumber}</span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">Document Available</label>
                                            <span>{form10FData?.DocumentAvailable}</span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">Document Date</label>
                                            <span>
                                                {form10FData?.DocumentDate
                                                    ? new Date(form10FData.DocumentDate).toLocaleDateString()
                                                    : ""}
                                            </span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">Validity Start Date</label>
                                            <span>
                                                {form10FData?.ValidityStartDate
                                                    ? new Date(form10FData.ValidityStartDate).toLocaleDateString()
                                                    : ""}
                                            </span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">Validity End Date</label>
                                            <span>
                                                {form10FData?.ValidityEndDate
                                                    ? new Date(form10FData.ValidityEndDate).toLocaleDateString()
                                                    : ""}
                                            </span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">Uploaded Document</label>

                                            {form10FData?.AttachmentFiles?.length > 0 ? (
                                                form10FData.AttachmentFiles.map(
                                                    (file: any, index: number) => (
                                                        <div key={index}>
                                                            <a
                                                                href={file.ServerRelativeUrl}
                                                                target="_blank"
                                                                rel="noreferrer"
                                                            >
                                                                {file.FileName}
                                                            </a>
                                                        </div>
                                                    )
                                                )
                                            ) : (
                                                <span>No Document Uploaded</span>
                                            )}
                                        </div>

                                    </div>
                                </div>
                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Compliance & Supporting Documents</label>
                                </div>
                                <div className='main-formcontainer'>
                                    <div className='row mb-20'>
                                        <div className='col-md-12'>
                                            <div className="attachment-table-container">
                                                <table className="attachment-table">
                                                    <thead>
                                                        <tr>
                                                            <th>Document Type</th>
                                                            <th>File Name</th>
                                                            <th>Action</th>
                                                        </tr>
                                                    </thead>

                                                    <tbody>
                                                        {vendorData?.AttachmentFiles?.map(
                                                            (file: any, index: number) => {

                                                                let documentType = "Attachment";

                                                                if (file.FileName.startsWith("VendorInvoice_")) {
                                                                    documentType = "Vendor Invoice / Proforma Invoice";
                                                                }
                                                                else if (file.FileName.startsWith("TRC_")) {
                                                                    documentType = "TRC";
                                                                }
                                                                else if (file.FileName.startsWith("KYC_")) {
                                                                    documentType = "KYC Documents";
                                                                }
                                                                else if (file.FileName.startsWith("BankConfirmation_")) {
                                                                    documentType = "Bank Confirmation / Cancelled Cheque";
                                                                }
                                                                else if (file.FileName.startsWith("OtherDocument_")) {
                                                                    documentType = "Other Documents";
                                                                }

                                                                return (
                                                                    <tr key={index}>
                                                                        <td>{documentType}</td>

                                                                        <td>
                                                                            {file.FileName}
                                                                        </td>

                                                                        <td>
                                                                            <a
                                                                                href={file.ServerRelativeUrl}
                                                                                target="_blank"
                                                                                rel="noreferrer"
                                                                                className="view-icon"
                                                                            >
                                                                                <img
                                                                                    src={view}
                                                                                    alt="View"
                                                                                    width="18"
                                                                                />
                                                                            </a>
                                                                        </td>
                                                                    </tr>
                                                                );
                                                            }
                                                        )}
                                                    </tbody>
                                                </table>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Threshold Details</label>
                                </div>

                                <div className="main-formcontainer">
                                    <div className="row mb-20">

                                        <div className="col-md-3">
                                            <label className="font fontblock">
                                                Eligible Amount
                                            </label>
                                            <span>
                                                {vendorData?.EligibleAmountWithoutWHT}
                                            </span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">
                                                Approved / Paid Amount
                                            </label>
                                            <span>
                                                {vendorData?.ApprovedAmountPaidAmount}
                                            </span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">
                                                Balance Amount
                                            </label>
                                            <span>
                                                {vendorData?.BalanceEligibleAmount}
                                            </span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">
                                                From Date
                                            </label>
                                            <span>
                                                {formatDate(vendorData?.FromDate)}
                                            </span>
                                        </div>

                                        <div className="col-md-3">
                                            <label className="font fontblock">
                                                To Date
                                            </label>
                                            <span>
                                                {formatDate(vendorData?.ToDate)}
                                            </span>
                                        </div>

                                    </div>
                                </div>
                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Remarks</label>
                                </div>

                                <div className="main-formcontainer">
                                    <div className="row mb-20">
                                        <div className="col-md-12">
                                            <span>
                                                {vendorData?.ApproverComments || "No Remarks"}
                                            </span>
                                        </div>
                                    </div>
                                </div>
                                <div style={{ margin: "10px", display: "flex", justifyContent: "center", gap: "5px", alignItems: "center" }}>
                                    {/* <a className="Submit-btn" onClick={handleApprove}>Submit</a>
                                    <a className="Reject-btn" onClick={handleReject}>Reject</a> */}
                                    <a className="Exit-btn" onClick={history.goBack}>Exit</a>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>














            {/* ========================= */}
            {/* TAX DECLARATIONS */}
            {/* ========================= */}

            {/* <div className="section">

                <h3>
                    Tax Declarations
                </h3>

                {
                    taxDeclarations.map(
                        (item: any, index: number) => (

                            <div
                                className="declaration-card"
                                key={index}
                            >

                                <div className="grid-4">

                                    <div className="field">
                                        <label className= "font fontblock">
                                            Declaration Type
                                        </label>

                                        <span>
                                            {
                                                item.DeclarationType
                                            }
                                        </span>
                                    </div>

                                    <div className="field">
                                        <label className= "font fontblock">
                                            Document Available
                                        </label>

                                        <span>
                                            {
                                                item.DocumentAvailable
                                            }
                                        </span>
                                    </div>

                                    <div className="field">
                                        <label className= "font fontblock">
                                            Document Number
                                        </label>

                                        <span>
                                            {
                                                item.DocumentNumber
                                            }
                                        </span>
                                    </div>

                                    <div className="field">
                                        <label className= "font fontblock">
                                            Document Date
                                        </label>

                                        <span>
                                            {
                                                item.DocumentDate
                                                    ? new Date(
                                                        item.DocumentDate
                                                    ).toLocaleDateString()
                                                    : ""
                                            }
                                        </span>
                                    </div>

                                    <div className="field">
                                        <label className= "font fontblock">
                                            Validity Start Date
                                        </label>

                                        <span>
                                            {
                                                item.ValidityStartDate
                                                    ? new Date(
                                                        item.ValidityStartDate
                                                    ).toLocaleDateString()
                                                    : ""
                                            }
                                        </span>
                                    </div>

                                    <div className="field">
                                        <label className= "font fontblock">
                                            Validity End Date
                                        </label>

                                        <span>
                                            {
                                                item.ValidityEndDate
                                                    ? new Date(
                                                        item.ValidityEndDate
                                                    ).toLocaleDateString()
                                                    : ""
                                            }
                                        </span>
                                    </div>

                                    <div className="field">
                                        <label className= "font fontblock">
                                            Tax Identification Number
                                        </label>

                                        <span>
                                            {
                                                item.TaxIdentificationNumber
                                            }
                                        </span>
                                    </div>

                                    <div className="field">
                                        <label className= "font fontblock">
                                            Acknowledgment Number
                                        </label>

                                        <span>
                                            {
                                                item.AcknowledgmentNumber
                                            }
                                        </span>
                                    </div>

                                    <div className="field">
                                        <label className= "font fontblock">
                                            SEP Clause
                                        </label>

                                        <span>
                                            {
                                                item.SEPClause
                                                    ? "Yes"
                                                    : "No"
                                            }
                                        </span>
                                    </div>

                                </div>


                                <div className="attachment-grid mt-15">

                                    {
                                        item.AttachmentFiles?.map(
                                            (
                                                file: any,
                                                fileIndex: number
                                            ) => (

                                                <a
                                                    key={fileIndex}
                                                    href={file.ServerRelativeUrl}
                                                    target="_blank"
                                                    rel="noreferrer"
                                                    className="attachment-link"
                                                >
                                                    {
                                                        file.FileName
                                                    }
                                                </a>
                                            )
                                        )
                                    }

                                </div>

                            </div>
                        )
                    )
                }

            </div> */}





            {/* ========================= */}
            {/* THRESHOLD DETAILS */}
            {/* ========================= */}

            {/* <div className="section">

                <h3>
                    Threshold Details
                </h3>

                <div className="grid-4">

                    <div className="field">
                        <label className= "font fontblock">
                            Eligible Amount Without WHT
                        </label>

                        <span>
                            {
                                vendorData?.EligibleAmountWithoutWHT
                            }
                        </span>
                    </div>

                    <div className="field">
                        <label className= "font fontblock">
                            Approved / Paid Amount
                        </label>

                        <span>
                            {
                                vendorData?.ApprovedAmountPaidAmount
                            }
                        </span>
                    </div>

                    <div className="field">
                        <label className= "font fontblock">
                            Balance Eligible Amount
                        </label>

                        <span>{vendorData?.BalanceEligibleAmount} </span>
                    </div>

                    <div className="field">
                        <label className= "font fontblock">
                            From Date
                        </label>

                        <span>
                            {
                                vendorData?.FromDate
                            }
                        </span>
                    </div>

                    <div className="field">
                        <label className= "font fontblock">
                            To Date
                        </label>

                        <span>
                            {
                                vendorData?.ToDate
                                 
                            }
                        </span>
                    </div>

                </div>

            </div> */}


        </>
    );
};

export default VendorViewForm;