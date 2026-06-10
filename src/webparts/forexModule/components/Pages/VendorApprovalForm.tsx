import * as React from "react";
import { IForexModuleProps } from "../IForexModuleProps";
import SPCRUDOPS from "../../service/BAL/spcrud";
import { useParams } from "react-router-dom";
import { useHistory } from "react-router-dom";

const VendorApprovalFormFirst: React.FC<IForexModuleProps> = (props) => {
    const history = useHistory();
    const spCrudOps = SPCRUDOPS();

    const [item, setItem] = React.useState<any>(null);
    const [comments, setComments] = React.useState("");

    // const itemId = new URLSearchParams(window.location.href).get("Id");
    const itemId = useParams<{ Id: string }>();

    React.useEffect(() => {
        loadRequest(itemId.Id);
    }, []);

    const loadRequest = async (itemId: string) => {

        const sp = await spCrudOps;
        console.log("itemId =", itemId);
        console.log("Number(itemId) =", Number(itemId));

        const data = await sp.getData(
            "VendorMaster",
            "*,AttachmentFiles",
            "AttachmentFiles",
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
            setItem(data[0]);
        }
    };

    const approveRequest = async () => {

        const sp = await spCrudOps;

        await sp.updateData(
            "VendorMaster",
            Number(itemId.Id),
            {
                RequestStatus: "Approved",
                ApproverComments: comments,
                CurrentApproverId: null,
                Status:"Active"
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
        <div className="container-fluid mt-3">

            <h4 className="text-center text-primary mb-4">
                Vendor Approval Form
            </h4>

            {/* Vendor Basic Details */}

            <fieldset className="border p-3 mb-3">

                <legend className="w-auto px-2">
                    Vendor Basic Details
                </legend>

                <div className="row mb-2">

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

                    <div className="col-md-3">
                        <label>Vendor Short Name</label>
                        <input
                            className="form-control"
                            value={item.VendorShortName || ""}
                            disabled
                        />
                    </div>

                </div>

            </fieldset>

            {/* Contact Information */}

            <fieldset className="border p-3 mb-3">

                <legend className="w-auto px-2">
                    Contact Information
                </legend>

                <div className="row">

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

            </fieldset>

            {/* Banking Details */}

            <fieldset className="border p-3 mb-3">

                <legend className="w-auto px-2">
                    Banking Details
                </legend>

                <div className="row">

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

            </fieldset>

            {/* Tax Information */}

            <fieldset className="border p-3 mb-3">

                <legend className="w-auto px-2">
                    Tax Information
                </legend>

                <div className="row">

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

            </fieldset>

            {/* Attachments */}

            <fieldset className="border p-3 mb-3">

                <legend className="w-auto px-2">
                    Attachments
                </legend>

                <table className="table table-bordered">

                    <thead>
                        <tr>
                            <th>File Name</th>
                            <th>Action</th>
                        </tr>
                    </thead>

                    <tbody>

                        {(item.AttachmentFiles || []).map((f: any) => (

                            <tr key={f.FileName}>

                                <td>{f.FileName}</td>

                                <td>
                                    <a
                                        href={f.ServerRelativeUrl}
                                        target="_blank"
                                        rel="noreferrer"
                                    >
                                        View
                                    </a>
                                </td>

                            </tr>

                        ))}

                    </tbody>

                </table>

            </fieldset>

            {/* Comments */}

            <fieldset className="border p-3 mb-3">

                <legend className="w-auto px-2">
                    Approver Comments
                </legend>

                <textarea
                    className="form-control"
                    rows={4}
                    value={comments}
                    onChange={(e) =>
                        setComments(e.target.value)
                    }
                />

            </fieldset>

            <div className="text-center">

                <button
                    className="btn btn-success me-2"
                    onClick={approveRequest}
                >
                    Approve
                </button>

                <button
                    className="btn btn-danger"
                    onClick={rejectRequest}
                >
                    Reject
                </button>

            </div>

        </div>
    );
};

export default VendorApprovalFormFirst;