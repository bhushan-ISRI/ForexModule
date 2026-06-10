import * as React from "react";
import { IForexModuleProps } from "../IForexModuleProps";
import SPCRUDOPS from "../../service/BAL/spcrud";
import { useHistory } from "react-router-dom";

const VendorApprovalDashboard: React.FC<IForexModuleProps> = (props) => {
    const history = useHistory();
    const spCrudOps = SPCRUDOPS();

    const [activeTab, setActiveTab] = React.useState("My");
    const [myRequests, setMyRequests] = React.useState<any[]>([]);
    const [pendingRequests, setPendingRequests] = React.useState<any[]>([]);
    const [allRequests, setAllRequests] = React.useState<any[]>([]);
    const [searchText, setSearchText] = React.useState("");
    const [loading, setLoading] = React.useState(false);

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
                "CurrentApproverId eq " + currentUserId,
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

    const filteredData =
        currentData.filter((x: any) => {

            if (!searchText)
                return true;

            return (
                x.VendorName?.toLowerCase().includes(searchText.toLowerCase()) || x.VendorCode?.toLowerCase().includes(searchText.toLowerCase())
            );
        });

    // const openRequest = (item: any) => {
    //     if (Number(item.CurrentApproverId) === Number(props.context.pageContext.legacyPageContext.userId)) {
    //         window.location.href = `${props.context.pageContext.web.absoluteUrl}/SitePages/VendorApproval.aspx?Id=${item.Id}`;
    //     } else {
    //         window.location.href = `${props.context.pageContext.web.absoluteUrl}/SitePages/VendorCreation.aspx?Id=${item.Id}`;
    //     }
    // };
    const openRequest = (item: any) => {

        if (item.RequestStatus === "Approved") {

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

        <div className="container-fluid mt-3">

            <h3 className="mb-3">
                Vendor Approval Dashboard
            </h3>

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

            {/* Tabs */}

            <ul className="nav nav-tabs mb-3">

                <li className="nav-item">
                    <button
                        className={`nav-link ${activeTab === "My"
                            ? "active"
                            : ""
                            }`}
                        onClick={() =>
                            setActiveTab("My")
                        }
                    >
                        My Requests
                    </button>
                </li>

                <li className="nav-item">
                    <button
                        className={`nav-link ${activeTab === "Pending"
                            ? "active"
                            : ""
                            }`}
                        onClick={() =>
                            setActiveTab("Pending")
                        }
                    >
                        Pending For Me
                    </button>
                </li>

                <li className="nav-item">
                    <button
                        className={`nav-link ${activeTab === "All"
                            ? "active"
                            : ""
                            }`}
                        onClick={() =>
                            setActiveTab("All")
                        }
                    >
                        All Requests
                    </button>
                </li>

            </ul>

            {/* Search */}

            <div className="card mb-3">

                <div className="card-body">

                    <div className="row">

                        <div className="col-md-4">

                            <input
                                className="form-control"
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

                </div>

            </div>

            {/* Grid */}

            <div className="card">

                <div className="card-header">
                    Vendor Requests
                </div>

                <div className="card-body">

                    {loading && (
                        <h5>
                            Loading...
                        </h5>
                    )}

                    <div className="table-responsive">

                        <table className="table table-bordered table-striped">

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

                                {filteredData.length === 0 && (
                                    <tr>
                                        <td colSpan={8} className="text-center text-danger">
                                            No Records Found
                                        </td>
                                    </tr>
                                )}

                                {filteredData.map((item: any) => (

                                    <tr key={item.Id}>

                                        <td>{item.Id}</td>

                                        <td>{item.VendorCode}</td>

                                        <td>{item.VendorName}</td>

                                        <td>
                                            <span
                                                className={
                                                    item.RequestStatus === "Approved"
                                                        ? "badge bg-success"
                                                        : item.RequestStatus === "Rejected"
                                                            ? "badge bg-danger"
                                                            : "badge bg-warning"
                                                }
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
                                            <button
                                                className="btn btn-primary btn-sm"
                                                onClick={() => openRequest(item)}
                                            >
                                                <i className="bi bi-pencil-square me-1"></i>
                                                Edit
                                            </button>
                                        </td>

                                    </tr>

                                ))}

                            </tbody>

                        </table>

                    </div>

                </div>

            </div>

        </div>
    );
};

export default VendorApprovalDashboard;