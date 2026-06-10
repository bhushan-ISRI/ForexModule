import * as React from "react";
import { useEffect, useState } from "react";
import "../Pages/Css/VendorReviewForm.scss";
import { IForexModuleProps } from "../IForexModuleProps";
import { sp } from "@pnp/sp/presets/all";
import { useHistory, useParams } from "react-router-dom";
import SPCRUDOPS from "../../service/BAL/spcrud";
// import { useHistory } from "react-router-dom";
const VendorApprovalForm: React.FC<IForexModuleProps> = (props) => {

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
                (x: any) => x.DeclarationType === "PE Declaration"
            );

            const trc = declarations.find(
                (x: any) => x.DeclarationType === "TRC"
            );

            const form10f = declarations.find(
                (x: any) => x.DeclarationType === "Form10F"
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

    // =============================
    // APPROVE
    // =============================

    const handleApprove = async () => {

        try {

            await sp.web.lists
                .getByTitle("VendorMaster")
                .items.getById(Number(vendorId.Id))
                .update({

                    RequestStatus: "Approved",

                    ApproverComments: remarks,

                    // ApprovedDate: new Date()
                });

            alert("Vendor Approved");
            history.push("/VendorApprovalDashboard");

        } catch (error) {

            console.log(error);

        }
    };

    // =============================
    // REJECT
    // =============================

    const handleReject = async () => {

        try {

            await sp.web.lists
                .getByTitle("VendorMaster")
                .items.getById(Number(vendorId.Id))
                .update({

                    RequestStatus: "Rejected",

                    ApproverComments: remarks,

                    // ApprovedDate: new Date()
                });

            alert("Vendor Rejected");
            history.push("/VendorApprovalDashboard");

        } catch (error) {

            console.log(error);

        }
    };

    if (loading) {

        return <div>Loading...</div>;
    }

    return (

        <div className="vendor-approval-container">

            <h2 className="main-title">
                Vendor Checker Form
            </h2>

            {/* ========================= */}
            {/* BASIC DETAILS */}
            {/* ========================= */}

            <div className="section">

                <h3>
                    Vendor Basic Details
                </h3>

                <div className="grid-4">
                    <div className="field">
                        <label>Oracle Vendor Code</label>
                           <span>
                            {vendorData?.VendorCode}
                        </span>
                    </div>

                    <div className="field">
                        <label>Oracle Vendor Name</label>
                           <span>
                            {vendorData?.VendorName}
                        </span>
                    </div>
                    <div className="field">
                        <label>
                            Vendor Name (Legal)
                        </label>

                        <span>
                            {vendorData?.VendorNameLegal}
                        </span>
                    </div>

                    <div className="field">
                        <label>
                            Vendor Short Name
                        </label>

                        <span>
                            {vendorData?.VendorShortName}
                        </span>
                    </div>

                    <div className="field">
                        <label>
                            Vendor Type
                        </label>

                        <span>
                            {vendorData?.VendorType}
                        </span>
                    </div>

                    <div className="field">
                        <label>
                            Country
                        </label>

                        <span>
                            {vendorData?.Country?.Country}
                        </span>
                    </div>

                    <div className="field">
                        <label>
                            Currency
                        </label>

                        <span>
                            {vendorData?.Currency?.Currency}
                        </span>
                    </div>

                    <div className="field">
                        <label>
                            Address Line 1
                        </label>

                        <span>
                            {vendorData?.Address}
                        </span>
                    </div>

                    {/* <div className="field">
                        <label>
                            Address Line 2
                        </label>

                        <span>
                            {vendorData?.VendorAddress}
                        </span>
                    </div> */}

                    <div className="field">
                        <label>
                            City
                        </label>

                        <span>
                            {vendorData?.City?.City}
                        </span>
                    </div>

                    <div className="field">
                        <label>
                            State
                        </label>

                        <span>
                            {vendorData?.State?.Title}
                        </span>
                    </div>

                    <div className="field">
                        <label>
                            Postal Code
                        </label>

                        <span>
                            {vendorData?.PostalCode}
                        </span>
                    </div>

                </div>

            </div>

            {/* ========================= */}
            {/* CONTACT DETAILS */}
            {/* ========================= */}

            <div className="section">

                <h3>
                    Contact Information
                </h3>

                <div className="grid-4">

                    <div className="field">
                        <label>
                            Contact Person Name
                        </label>

                        <span>
                            {vendorData?.ContactPersonName}
                        </span>
                    </div>

                    <div className="field">
                        <label>
                            Email Id
                        </label>

                        <span>
                            {vendorData?.EmailId}
                        </span>
                    </div>

                    <div className="field">
                        <label>
                            Phone Number
                        </label>

                        <span>
                            {vendorData?.PhoneNumber}
                        </span>
                    </div>

                    <div className="field">
                        <label>
                            Alternate Contact
                        </label>

                        <span>
                            {vendorData?.AlternateContact}
                        </span>
                    </div>

                </div>

            </div>

            {/* ========================= */}
            {/* BANK DETAILS */}
            {/* ========================= */}

            <div className="section">

                <h3>
                    Banking Details
                </h3>

                <div className="grid-4">

                    <div className="field">
                        <label>
                            Beneficiary Name
                        </label>

                        <span>
                            {vendorData?.BeneficiaryName}
                        </span>
                    </div>

                    <div className="field">
                        <label>
                            Bank Name
                        </label>

                        <span>
                            {vendorData?.BankName}
                        </span>
                    </div>

                    <div className="field">
                        <label>
                            Bank Address
                        </label>

                        <span>
                            {vendorData?.BankAddress}
                        </span>
                    </div>

                    <div className="field">
                        <label>
                            Account Number / IBAN
                        </label>

                        <span>
                            {vendorData?.AccountNumberIBAN}
                        </span>
                    </div>

                    <div className="field">
                        <label>
                            SWIFT / BIC Code
                        </label>

                        <span>
                            {vendorData?.SWIFTBICCode}
                        </span>
                    </div>

                    <div className="field">
                        <label>
                            Routing Number / ABA
                        </label>

                        <span>
                            {vendorData?.RoutingNumberABA}
                        </span>
                    </div>

                    <div className="field">
                        <label>
                            IFSC Code
                        </label>

                        <span>
                            {vendorData?.IFSCCode}
                        </span>
                    </div>

                    <div className="field">
                        <label>
                            Intermediary Bank
                        </label>

                        <span>
                            {vendorData?.IntermediaryBank}
                        </span>
                    </div>

                    <div className="field">
                        <label>
                            Intermediary SWIFT Code
                        </label>

                        <span>
                            {vendorData?.IntermediarySWIFTCode}
                        </span>
                    </div>

                </div>

            </div>

            {/* ========================= */}
            {/* TAX DETAILS */}
            {/* ========================= */}

            <div className="section">

                <h3>
                    Tax & Regulatory Information
                </h3>

                <div className="grid-4">

                    <div className="field">
                        <label>
                            Nature Of Payment
                        </label>

                        {/* <span>
                            {
                                vendorData?.NatureOfPayment?.Title
                            }
                        </span> */}
                    </div>

                    <div className="field">
                        <label>
                            Tax Document Available
                        </label>

                        <span>
                            {
                                vendorData?.TaxDocumentAvailable
                            }
                        </span>
                    </div>

                    <div className="field">
                        <label>
                            DTAA Applicable
                        </label>

                        <span>
                            {
                                vendorData?.DTAAApplicable
                            }
                        </span>
                    </div>

                    <div className="field">
                        <label>
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
                                        <label>
                                            Declaration Type
                                        </label>

                                        <span>
                                            {
                                                item.DeclarationType
                                            }
                                        </span>
                                    </div>

                                    <div className="field">
                                        <label>
                                            Document Available
                                        </label>

                                        <span>
                                            {
                                                item.DocumentAvailable
                                            }
                                        </span>
                                    </div>

                                    <div className="field">
                                        <label>
                                            Document Number
                                        </label>

                                        <span>
                                            {
                                                item.DocumentNumber
                                            }
                                        </span>
                                    </div>

                                    <div className="field">
                                        <label>
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
                                        <label>
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
                                        <label>
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
                                        <label>
                                            Tax Identification Number
                                        </label>

                                        <span>
                                            {
                                                item.TaxIdentificationNumber
                                            }
                                        </span>
                                    </div>

                                    <div className="field">
                                        <label>
                                            Acknowledgment Number
                                        </label>

                                        <span>
                                            {
                                                item.AcknowledgmentNumber
                                            }
                                        </span>
                                    </div>

                                    <div className="field">
                                        <label>
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
            <div className="section">
                <h3>Permanent Establishment Declaration</h3>

                <div className="grid-4">
                    <div className="field">
                        <label>Document Available</label>
                        <span>{peData?.DocumentAvailable}</span>
                    </div>

                    <div className="field">
                        <label>Document Number</label>
                        <span>{peData?.DocumentNumber}</span>
                    </div>

                    <div className="field">
                        <label>Document Date</label>
                        <span>
                            {peData?.DocumentDate}
                        </span>
                    </div>

                    <div className="field">
                        <label>SEP Clause</label>
                        <span>{peData?.SEPClause ? "Yes" : "No"}</span>
                    </div>
                </div>
            </div>
            <div className="section">
                <h3>Tax Residency Certificate</h3>

                <div className="grid-4">
                    <div className="field">
                        <label>Document Number</label>
                        <span>{trcData?.DocumentNumber}</span>
                    </div>

                    <div className="field">
                        <label>Country Of Tax Residence</label>
                        <span>{trcData?.CountryOfTaxResidence}</span>
                    </div>

                    <div className="field">
                        <label>TIN</label>
                        <span>{trcData?.TaxIdentificationNumber}</span>
                    </div>

                    <div className="field">
                        <label>Document Available</label>
                        <span>{trcData?.DocumentAvailable}</span>
                    </div>
                </div>
            </div>
            <div className="section">
                <h3>Form 10F</h3>

                <div className="grid-4">
                    <div className="field">
                        <label>Document Number</label>
                        <span>{form10FData?.DocumentNumber}</span>
                    </div>

                    <div className="field">
                        <label>Acknowledgment Number</label>
                        <span>{form10FData?.AcknowledgmentNumber}</span>
                    </div>

                    <div className="field">
                        <label>Document Available</label>
                        <span>{form10FData?.DocumentAvailable}</span>
                    </div>

                    <div className="field">
                        <label>Document Date</label>
                        <span>
                            {form10FData?.DocumentDate}
                        </span>
                    </div>
                </div>
            </div>

            {/* ========================= */}
            {/* THRESHOLD DETAILS */}
            {/* ========================= */}

            {/* <div className="section">

                <h3>
                    Threshold Details
                </h3>

                <div className="grid-4">

                    <div className="field">
                        <label>
                            Eligible Amount Without WHT
                        </label>

                        <span>
                            {
                                vendorData?.EligibleAmountWithoutWHT
                            }
                        </span>
                    </div>

                    <div className="field">
                        <label>
                            Approved / Paid Amount
                        </label>

                        <span>
                            {
                                vendorData?.ApprovedAmountPaidAmount
                            }
                        </span>
                    </div>

                    <div className="field">
                        <label>
                            Balance Eligible Amount
                        </label>

                        <span>{vendorData?.BalanceEligibleAmount} </span>
                    </div>

                    <div className="field">
                        <label>
                            From Date
                        </label>

                        <span>
                            {
                                vendorData?.FromDate
                            }
                        </span>
                    </div>

                    <div className="field">
                        <label>
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

            {/* ========================= */}
            {/* ATTACHMENTS */}
            {/* ========================= */}

            {/* <div className="section">

                <h3>
                    Compliance & Supporting Documents
                </h3>

                <div className="attachment-grid">

                    {
                        vendorData?.AttachmentFiles?.map(
                            (
                                file: any,
                                index: number
                            ) => (

                                <a
                                    key={index}
                                    href={file.ServerRelativeUrl}
                                    target="_blank"
                                    rel="noreferrer"
                                    className="attachment-link"
                                >
                                    {file.FileName}
                                </a>
                            )
                        )
                    }

                </div>

            </div> */}

            {/* ========================= */}
            {/* REMARKS */}
            {/* ========================= */}

            <div className="section">

                <h3>
                    Checker Remarks
                </h3>

                <textarea
                    className="remarks-box"
                    value={remarks}
                    onChange={(e) =>
                        setRemarks(e.target.value)
                    }
                />

            </div>

            {/* ========================= */}
            {/* BUTTONS */}
            {/* ========================= */}

            <div className="button-section">

                <button
                    className="approve-btn"
                    onClick={handleApprove}
                >
                    Approve
                </button>

                <button
                    className="reject-btn"
                    onClick={handleReject}
                >
                    Reject
                </button>

                <button
                    className="cancel-btn"
                    onClick={() =>
                        window.history.back()
                    }
                >
                    Exit
                </button>

            </div>

        </div>
    );
};

export default VendorApprovalForm;