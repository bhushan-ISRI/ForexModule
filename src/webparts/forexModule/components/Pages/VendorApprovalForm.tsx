import * as React from "react";
import { IForexModuleProps } from "../IForexModuleProps";
import SPCRUDOPS from "../../service/BAL/spcrud";
import { useParams } from "react-router-dom";
import { useHistory } from "react-router-dom";
import logo from "../../assets/sona-comstarlogo.png";
import view from "../../assets/Eye.png";

const VendorApprovalFormFirst: React.FC<IForexModuleProps> = (props) => {
    const history = useHistory();
    const spCrudOps = SPCRUDOPS();

    const [item, setItem] = React.useState<any>(null);
    const [comments, setComments] = React.useState("");
    const [approvalMatrix, setApprovalMatrix] = React.useState<any[]>([]);
    // const itemId = new URLSearchParams(window.location.href).get("Id");
    const itemId = useParams<{ Id: string }>();
    const [checkDocumentUpload,setCheckDocumentUpload] = React.useState("");
    React.useEffect(() => {
        loadRequest(itemId.Id);
        loadApprovalMatrix();
    }, []);


    const loadRequest = async (itemId: string) => {

        const sp = await spCrudOps;
        console.log("itemId =", itemId);
        console.log("Number(itemId) =", Number(itemId));

        const data = await sp.getData(
            "VendorMaster",
            "*,Currency/Currency,Country/Country,AttachmentFiles",
            "AttachmentFiles,Currency,Country",
            `Id eq ${Number(itemId)}`,
            { column: "Id", isAscending: false },
            1000,
            props
        );

        // if (data.length > 0) {
        //     setItem(data[0]);
        // }
        if (data.length > 0) {
            console.log("Loaded Item", data[0]);
            setCheckDocumentUpload(data[0].DocumentUpload);
            setItem(data[0]);
        }
    };
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
    const approveRequest = async () => {

        const sp = await spCrudOps;
        let ApprovalStatus;
        if(checkDocumentUpload == "Yes"){
            ApprovalStatus = "Pending With IDT"
        }else{
            ApprovalStatus = "Pending For Document Upload"
        }


        await sp.updateData(
            "VendorMaster",
            Number(itemId.Id),
            {
                RequestStatus:ApprovalStatus,
                ApproverComments: comments,
                    CurrentApproverId: approvalMatrix.length > 0 ? approvalMatrix[0].Approver.Id : null,
                Status: "Active"
            },
            props
        );

        alert("Approved Successfully");
        history.push("/VendorApprovalDashboard");

    };

    const rejectRequest = async () => {

        const sp = await spCrudOps;

        await sp.updateData(
            "VendorMaster",
            Number(itemId.Id),
            {
                RequestStatus: "Rejected",
                ApproverComments: comments
            },
            props
        );

        alert("Rejected Successfully");
        history.push("/VendorApprovalDashboard");
    };

    if (!item) {
        return <div>Loading...</div>;
    }

    return (
        <>

            <div className='MainUplodForm' style={{ margin: "5px 0px" }}>
                <div className='row'>
                    <div className='col-md-12'>
                        <div className='Main-Boxpoup'>
                            <div className="bordered">
                                <a><img src={logo} /></a>
                                <h1>Vendor Approval Form</h1>
                            </div>
                            <div className='borderedbox'>
                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label> Vendor Basic Details</label>
                                </div>
                                <div className='main-formcontainer'>
                                    <div className="row mb-20">
                                        <div className="col-md-3">
                                            <label>Oracle Vendor Code</label>
                                            <input
                                                className="form-control"
                                                value={item.VendorCode || ""}
                                                disabled
                                            />
                                        </div>

                                        <div className="col-md-3">
                                            <label>Oracle Vendor Name</label>
                                            <input
                                                className="form-control"
                                                value={item.VendorName || ""}
                                                disabled
                                            />
                                        </div>

                                        <div className="col-md-3">
                                            <label>Vendor Name (Legal)</label>
                                            <input
                                                className="form-control"
                                                value={item.VendorNameLegal || ""}
                                                disabled
                                            />
                                        </div>

                                        {/* <div className="col-md-3">
                                            <label>Vendor Short Name</label>
                                            <input
                                                className="form-control"
                                                value={item.VendorShortName || ""}
                                                disabled
                                            />
                                        </div> */}
                                        <div className="col-md-3">
                                            <label className='font'>Currency  <span style={{ color: "red" }}>*</span></label>
                                            <input
                                                className="form-control"
                                                value={item.Currency?.Currency || ""}
                                                disabled
                                            />
                                        </div>

                                    </div>
                                    <div className="row mb-20">
                                        
                                        <div className="col-md-3">
                                            <label className='font'>Country  <span style={{ color: "red" }}>*</span></label>
                                            <input
                                                className="form-control"
                                                value={item.Country?.Country || ""}
                                                disabled
                                            />
                                        </div>



                                        <div className="col-md-3">
                                            <label className='font'>State  <span style={{ color: "red" }}>*</span></label>

                                            <input
                                                type="text"
                                                className="form-controltext"
                                                value={item.state0 || ""}
                                                // onChange={(e) => setStateName(e.target.value)}
                                                // placeholder="Enter State"
                                                disabled
                                            />
                                        </div>
                                        <div className="col-md-3">
                                            <label className="font">City  <span style={{ color: "red" }}>*</span></label>

                                            <input
                                                type="text"
                                                className="form-controltext"
                                                value={item.city0 || ""}
                                                // onChange={(e) => setCityName(e.target.value)}
                                                // placeholder="Enter City"
                                                disabled
                                            />
                                        </div>
                                         <div className="col-md-3">
                                            <label className='font'>Postal Code  <span style={{ color: "red" }}>*</span></label>
                                            <input
                                                type="number"
                                                className="form-control"
                                                value={item.Pincode || ""}
                                                // onChange={(e) => setPostalCode(e.target.value)}
                                                disabled
                                            />
                                        </div>

                                            <div className="col-md-3">
                                            <label className='font'>Address Line 1  <span style={{ color: "red" }}>*</span></label>
                                            <textarea
                                                className="form-control"
                                                rows={2}
                                                value={item.VendorAddress || ""}
                                                // onChange={(e) => setAddress(e.target.value)}
                                                disabled
                                            />
                                        </div>
                                    </div>
                                    <div className="row mb-20">


                                       

                                        
                                    </div>

                                </div>
                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label> Contact Information</label>
                                </div>
                                <div className='main-formcontainer'>
                                    <div className="row mb-20">
                                        <div className="col-md-3">
                                            <label>Contact Person Name</label>
                                            <input
                                                className="form-control"
                                                value={item.ContactPersonName || ""}
                                                disabled
                                            />
                                        </div>

                                        <div className="col-md-3">
                                            <label>Email Id</label>
                                            <input
                                                className="form-control"
                                                value={item.EmailId || ""}
                                                disabled
                                            />
                                        </div>

                                        <div className="col-md-3">
                                            <label>Phone Number</label>
                                            <input
                                                className="form-control"
                                                value={item.PhoneNumber || ""}
                                                disabled
                                            />
                                        </div>

                                        <div className="col-md-3">
                                            <label>Alternate Contact</label>
                                            <input
                                                className="form-control"
                                                value={item.AlternateContact || ""}
                                                disabled
                                            />
                                        </div>

                                    </div>
                                </div>
                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Banking Details</label>
                                </div>
                                <div className='main-formcontainer'>
                                    <div className="row mb-20">

                                        <div className="col-md-3">
                                            <label>Beneficiary Name</label>
                                            <input
                                                className="form-control"
                                                value={item.BeneficiaryName || ""}
                                                disabled
                                            />
                                        </div>

                                        <div className="col-md-3">
                                            <label>Bank Name</label>
                                            <input
                                                className="form-control"
                                                value={item.BankName || ""}
                                                disabled
                                            />
                                        </div>

                                        <div className="col-md-3">
                                            <label>Bank Address</label>
                                            <input
                                                className="form-control"
                                                value={item.BankAddress || ""}
                                                disabled
                                            />
                                        </div>

                                        <div className="col-md-3">
                                            <label>Account Number / IBAN</label>
                                            <input
                                                className="form-control"
                                                value={item.AccountNumberIBAN || ""}
                                                disabled
                                            />
                                        </div>

                                    </div>
                                    <div className="row mb-20">
                                        <div className="col-md-3">
                                            <label className='font'>SWIFT / BIC Code  <span style={{ color: "red" }}>*</span></label>
                                            <input
                                                className="form-control"
                                                value={item.SWIFTBICCode}
                                                // onChange={(e) => setSwiftCode(e.target.value)}
                                                disabled
                                            />
                                        </div>

                                        <div className="col-md-3">
                                            <label className='font'>Routing Number / ABA</label>
                                            <input
                                                className="form-control"
                                                value={item.RoutingNumberABA}
                                                // onChange={(e) => setRoutingNumberABA(e.target.value)}
                                                disabled
                                            />
                                        </div>

                                        {/* <div className="col-md-3">
                                            <label className='font'>IFSC Code</label>
                                            <input
                                                className="form-control"
                                                value={item.IFSCCode}
                                                // onChange={(e) => setIfscCode(e.target.value)}
                                                disabled
                                            />
                                        </div> */}

                                        {/* <div className="col-md-3">
                                            <label className='font'>Intermediary Bank</label>
                                            <input
                                                className="form-control"
                                                value={item.IntermediaryBank}
                                                // onChange={(e) => setIntermediaryBank(e.target.value)}
                                                disabled
                                            />
                                        </div> */}
                                    </div>
                                </div>
                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Tax & Regulatory Information</label>
                                </div>
                                <div className='main-formcontainer'>
                                    <div className="row mb-20">

                                        <div className="col-md-3">
                                            <label>Purpose Code</label>
                                            <input
                                                className="form-control"
                                                value={item.PurposeCodeRBI || ""}
                                                disabled
                                            />
                                        </div>

                                        <div className="col-md-3">
                                            <label>TRC</label>
                                            <input
                                                className="form-control"
                                                value={item.TaxResidencyCertificateTRC || ""}
                                                disabled
                                            />
                                        </div>

                                        <div className="col-md-3">
                                            <label>PE Declaration</label>
                                            <input
                                                className="form-control"
                                                value={item.PEDeclaration || ""}
                                                disabled
                                            />
                                        </div>

                                        <div className="col-md-3">
                                            <label>DTAA Applicable</label>
                                            <input
                                                className="form-control"
                                                value={item.DTAAApplicable || ""}
                                                disabled
                                            />
                                        </div>

                                    </div>
                                    <div className="row mb-20">
                                        <div className="col-md-3">
                                            <label className='font'>
                                                Withholding Tax Applicable
                                                
                                            </label>
                                            <input
                                                className="form-control"
                                                value={item.withholdingTaxApplicable}
                                                // onChange={(e) => setIntermediaryBank(e.target.value)}
                                                disabled
                                            />

                                        </div>
                                        {/* <div className="col-md-3">
                                            <label className='font'>
                                                Country of Tax Residence
                                               
                                            </label>
                                            <input
                                                className="form-control"
                                                value={item.CountryOfTaxResidence}
                                                // onChange={(e) => setIntermediaryBank(e.target.value)}
                                                disabled
                                            />


                                        </div> */}
                                    </div>
                                </div>
                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Attachments</label>
                                </div>
                                <div className='main-formcontainer'>
                                    <div className="row mb-20">
                                        <div className="col-md-12">
                                            <table className="custom-table">

                                                <thead>
                                                    <tr>
                                                        <th>Document Type</th>
                                                        <th>File Name</th>
                                                        <th>Action</th>
                                                    </tr>
                                                </thead>

                                                <tbody>

                                                    {(item.AttachmentFiles || []).map((f: any) => (

                                                        <tr key={f.FileName}>

                                                            <td>
                                                                {
                                                                    f.FileName.startsWith("VendorInvoice_")
                                                                        ? "Vendor Invoice / Proforma Invoice"
                                                                        : f.FileName.startsWith("TRC_")
                                                                            ? "TRC"
                                                                            : f.FileName.startsWith("KYC_")
                                                                                ? "KYC Documents"
                                                                                : f.FileName.startsWith("BankConfirmation_")
                                                                                    ? "Bank Confirmation / Cancelled Cheque"
                                                                                    : f.FileName.startsWith("OtherDocument_")
                                                                                        ? "Other Documents"
                                                                                        : "Attachment"
                                                                }
                                                            </td>

                                                            <td>{f.FileName}</td>

                                                            <td>
                                                                <a
                                                                    href={f.ServerRelativeUrl}
                                                                    target="_blank"
                                                                    rel="noreferrer"
                                                                >
                                                                    <img src={view} width="15" />
                                                                </a>
                                                            </td>

                                                        </tr>

                                                    ))}

                                                </tbody>

                                            </table>
                                        </div>
                                    </div>
                                </div>
                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Approver Comments</label>
                                </div>
                                <div className='main-formcontainer'>
                                    <div className="row mb20">
                                        <div className="col-md-12">
                                            <textarea
                                                className="form-control"
                                                rows={4}
                                                value={comments}
                                                onChange={(e) =>
                                                    setComments(e.target.value)
                                                }
                                            />
                                        </div>
                                    </div>
                                </div>
                                <div style={{ margin: "10px", display: "flex", justifyContent: "center", gap: "5px", alignItems: "center" }}>
                                    <a className="Submit-btn" onClick={approveRequest}> Approve </a>
                                    <a className="Reject-btn" onClick={rejectRequest}>Reject</a>
                                    <a className="Exit-btn" onClick={()=>history.push('/VendorCreationDashboard')}>Exit</a>

                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div >
        </>
    );
};

export default VendorApprovalFormFirst;