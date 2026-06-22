import * as React from "react";
import { IForexModuleProps } from "../IForexModuleProps";
import SPCRUDOPS from "../../service/BAL/spcrud";
import { useHistory } from "react-router-dom";

import edit from '../../assets/Pencil.png';
import view from '../../assets/Eye.png';
const IDTVendorApprovalDashboard: React.FC<IForexModuleProps> = (props) => {
    const history = useHistory();
    const spCrudOps = SPCRUDOPS();

    const [activeTab, setActiveTab] = React.useState("My");
    const [myRequests, setMyRequests] = React.useState<any[]>([]);
    const [pendingRequests, setPendingRequests] = React.useState<any[]>([]);
    const [allRequests, setAllRequests] = React.useState<any[]>([]);
    const [searchText, setSearchText] = React.useState("");
    const [loading, setLoading] = React.useState(false);
    const [currentPage, setCurrentPage] = React.useState(1);
    const recordsPerPage = 10;

    React.useEffect(() => {
        loadDashboard();
    }, []);

    const loadDashboard = async () => {

        try {

            setLoading(true);

            const sp = await spCrudOps;

            const currentUserId =
                props.context.pageContext.legacyPageContext.userId;

            const allData = await sp.getData(
                "VendorMaster",
                "*,Author/Title,Country/Country,CurrentApprover/Title,CurrentApprover/ID",
                "Author,Country,CurrentApprover",
                `RequestStatus ne 'Rejected'`,//and ApprovedByIDTChecker eq 'No'
                { column: "Id", isAscending: false },
                5000,
                props
            );

            console.log("Current User Id:", currentUserId);
            console.log("VendorMaster Data:", allData);
            console.log("Record Count:", allData?.length);

            setAllRequests(allData || []);

            setMyRequests(
                (allData || []).filter(
                    (x: any) =>
                        Number(x.AuthorId) === Number(currentUserId)
                )
            );

            setPendingRequests(
                (allData || []).filter(
                    (x: any) =>
                        Number(x.CurrentApproverId) === Number(currentUserId)
                )
            );

        } catch (e) {

            console.error("Dashboard Error:", e);

        } finally {

            setLoading(false);

        }
    };

    const totalRequests = allRequests.length;

    const approvedRequests =
        allRequests.filter(
            (x: any) => x.RequestStatus === "Approved"
        ).length;

    const rejectedRequests = allRequests.filter((x: any) => x.RequestStatus === "Rejected").length;

    const pendingCount = allRequests.filter((x: any) => x.RequestStatus?.includes("Pending")).length;

    const currentData = allRequests;

    const filteredData = currentData
        .filter((x: any) => {

            if (!searchText)
                return true;

            return (
                x.VendorName?.toLowerCase().includes(searchText.toLowerCase()) ||
                x.VendorCode?.toLowerCase().includes(searchText.toLowerCase())
            );
        })
        .sort((a: any, b: any) => b.Id - a.Id);

    const indexOfLastRecord = currentPage * recordsPerPage;
    const indexOfFirstRecord = indexOfLastRecord - recordsPerPage;

    const currentRecords = filteredData.slice(
        indexOfFirstRecord,
        indexOfLastRecord
    );

    const totalPages = Math.ceil(
        filteredData.length / recordsPerPage
    );
    const openRequest = (item: any) => {

        if (item.RequestStatus === "Pending With IDT") {

            history.push(
                `/VendorApprovalForm/${item.Id}`
            );

        } else {

            history.push(
                `/VendorApprovalFormFirst/${item.Id}`
            );

        }

    };
    return (

        <>



            {/* <h3 className="mb-3">
                Vendor Approval Dashboard
            </h3> */}

            {/* Summary Cards */}

            {/* <div className="row mb-4">

                <div className="col-md-3">
                    <div className="card bg-primary text-white">
                        <div className="card-body text-center">
                            <h2>{totalRequests}</h2>
                            <h6>Total Requests</h6>
                        </div>
                    </div>
                </div>

                <div className="col-md-3">
                    <div className="card bg-warning">
                        <div className="card-body text-center">
                            <h2>{pendingCount}</h2>
                            <h6>Pending</h6>
                        </div>
                    </div>
                </div>

                <div className="col-md-3">
                    <div className="card bg-success text-white">
                        <div className="card-body text-center">
                            <h2>{approvedRequests}</h2>
                            <h6>Approved</h6>
                        </div>
                    </div>
                </div>

                <div className="col-md-3">
                    <div className="card bg-danger text-white">
                        <div className="card-body text-center">
                            <h2>{rejectedRequests}</h2>
                            <h6>Rejected</h6>
                        </div>
                    </div>
                </div>

            </div> */}



            <div className="header">
                <h2> Vendor Approval Dashboard </h2>
            </div>

            <div className="mainsecondapprove">
                <div>
                    <input
                        className="form-control"
                        style={{ width: "250px;" }}
                        placeholder="Search Vendor Name / Code"
                        value={searchText}
                        onChange={(e) =>
                            setSearchText(
                                e.target.value
                            )
                        }
                    />
                </div>
            </div>


            <div className="card">

                {/* <div className="card-header">
                    Vendor Requests
                </div> */}

                <div className="card-body">

                    {loading && (
                        <h5>
                            Loading...
                        </h5>
                    )}

                    <div style={{ overflowX: "auto" }}>

                        <table className="custom-table">

                            <thead>

                                <tr>

                                    <th>ID</th>

                                    <th>Vendor Code</th>

                                    <th>Vendor Name</th>

                                    <th>Status</th>

                                    <th>Pending At</th>

                                    <th>Created By</th>

                                    <th>Created</th>

                                    <th>Action</th>

                                </tr>

                            </thead>

                            <tbody>

                                {currentRecords.length === 0 && (
                                    <tr>
                                        <td colSpan={8} className="text-center text-danger">
                                            No Records Found
                                        </td>
                                    </tr>
                                )}

                                {currentRecords.map((item: any) => (

                                    <tr key={item.Id}>

                                        <td>{item.Id}</td>

                                        <td>{item.VendorCode}</td>

                                        <td>{item.VendorName}</td>

                                        <td>
                                            <span
                                            // className={
                                            //     item.RequestStatus === "Approved"
                                            //         ? ""
                                            //         : item.RequestStatus === "Rejected"
                                            //             ? ""
                                            //             : "badge bg-warning"
                                            // }
                                            >
                                                {item.RequestStatus}
                                            </span>
                                        </td>

                                        <td>{item.CurrentApprover?.Title}</td>

                                        <td>{item.Author?.Title}</td>

                                        <td>
                                            {item.Created
                                                ? new Date(item.Created).toLocaleDateString()
                                                : ""}
                                        </td>

                                        <td>
                                            {Number(item.CurrentApproverId) ===
                                                Number(props.context.pageContext.legacyPageContext.userId) &&
                                                item.RequestStatus !== "Sent Back" &&
                                                item.RequestStatus !== "Send Back" && item.DocumentUpload === "Yes" && (
                                                    <a
                                                        title="Edit"
                                                        onClick={() => openRequest(item)}
                                                        style={{ marginRight: "10px", cursor: "pointer" }}
                                                    >
                                                        <img src={edit} alt="Edit" width={15} />
                                                    </a>
                                                )}
                                            <a onClick={() => history.push(`/VendorViewForm/${item.Id}`)}>
                                                <img src={view} alt="" width={15} />
                                            </a>
                                            {/* <button
                                                className="btn btn-primary btn-sm"
                                                onClick={() => openRequest(item)}
                                            >
                                                <i className="bi bi-pencil-square me-1"></i>
                                                Edit
                                            </button> */}
                                        </td>

                                    </tr>

                                ))}

                            </tbody>

                        </table>
                        <div className="d-flex justify-content-center mt-3">

                            <button
                                className="btn btn-secondary me-2"
                                disabled={currentPage === 1}
                                onClick={() =>
                                    setCurrentPage(currentPage - 1)
                                }
                            >
                                Previous
                            </button>

                            <span
                                style={{
                                    padding: "8px 15px"
                                }}
                            >
                                Page {currentPage} of {totalPages}
                            </span>

                            <button
                                className="btn btn-secondary ms-2"
                                disabled={currentPage === totalPages}
                                onClick={() =>
                                    setCurrentPage(currentPage + 1)
                                }
                            >
                                Next
                            </button>

                        </div>

                    </div>

                </div>

            </div>

        </>
    );
};

export default IDTVendorApprovalDashboard;