import * as React from "react";
import { IForexModuleProps } from "../IForexModuleProps";
import SPCRUDOPS from "../../service/BAL/spcrud";
import { useHistory } from "react-router-dom";

const VendorDashboard: React.FC<IForexModuleProps> = (props) => {

    const spCrudOps = SPCRUDOPS();

    const [vendorData, setVendorData] = React.useState<any[]>([]);
    const [searchText, setSearchText] = React.useState("");
    const [statusFilter, setStatusFilter] = React.useState("");
    const history = useHistory();

    React.useEffect(() => {
        loadData();
    }, []);
    //     const createNewRequest = () => {

    //     history.push("/VendorCreation");

    // };

    const loadData = async () => {

        try {

            const sp = await spCrudOps;
            const currentUserId = props.context.pageContext.legacyPageContext.userId;

            const data = await sp.getData(
                "VendorMaster",
                "*,Author/Title,Country/Country",
                "Author,Country",
                `AuthorId eq ${currentUserId}`,
                { column: "Id", isAscending: true },
                5000,
                props
            );

            setVendorData(data);

        } catch (error) {
            console.log(error);
        }
    };

    const totalRequests = vendorData.length;

    const pendingRequests =
        vendorData.filter(
            (x: any) =>
                x.Status === "Pending"
        ).length;

    const approvedRequests =
        vendorData.filter(
            (x: any) =>
                x.Status === "Approved"
        ).length;

    const rejectedRequests =
        vendorData.filter(
            (x: any) =>
                x.Status === "Rejected"
        ).length;

    const draftRequests =
        vendorData.filter(
            (x: any) =>
                x.Status === "Draft"
        ).length;

    const filteredData =
        vendorData.filter((x: any) => {

            const vendorMatch =
                !searchText ||
                x.VendorName
                    ?.toLowerCase()
                    .includes(
                        searchText.toLowerCase()
                    );

            const statusMatch =
                !statusFilter ||
                x.Status === statusFilter;

            return vendorMatch && statusMatch;

        });

    // const openRequest = (id: number) => {

    //     window.location.href =
    //         `${window.location.origin}${props.context.pageContext.web.serverRelativeUrl}/SitePages/VendorCreation.aspx?Id=${id}`;

    // };
    const openRequest = (id: number) => {

        history.push(
            `/VendorCreationForm/${id}`
        );

    };

    return (
        <div className="container-fluid mt-3">
            {/* 
            <div className="row mb-3">

                <div className="col-md-3">
                    <div className="card bg-primary text-white">
                        <div className="card-body text-center">
                            <h2>{totalRequests}</h2>
                            <h6>Total Requests</h6>
                        </div>
                    </div>
                </div>

                <div className="col-md-3">
                    <div className="card bg-warning text-dark">
                        <div className="card-body text-center">
                            <h2>{pendingRequests}</h2>
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

            {/* <div className="row mb-3">

                <div className="col-md-3">
                    <div className="card bg-secondary text-white">
                        <div className="card-body text-center">
                            <h2>{draftRequests}</h2>
                            <h6>Draft</h6>
                        </div>
                    </div>
                </div>

            </div> */}

            <div className="d-flex justify-content-between align-items-center mb-3">

                <h3 className="mb-0">
                    Vendor Dashboard
                </h3>

                <button
                    className="btn-danger px-4"
                    onClick={() => history.push("/CreationForm")}
                    style={{ width: "261px !important" }}
                >
                    <i className="fa fa-plus me-2"></i>
                    Vendor Creation Request
                </button>

            </div>

            <div className="card shadow-sm mb-3">

                <div className="card-header bg-light">
                    <strong>Search Vendor Requests</strong>
                </div>

                <div className="card-body">

                    <div className="row g-3">

                        <div className="col-md-6">

                            <input
                                type="text"
                                className="form-control"
                                placeholder="Search Vendor Name..."
                                value={searchText}
                                onChange={(e) =>
                                    setSearchText(e.target.value)
                                }
                            />

                        </div>

                        <div className="col-md-3">

                            <select
                                className="form-select"
                                value={statusFilter}
                                onChange={(e) =>
                                    setStatusFilter(e.target.value)
                                }
                            >
                                <option value="">
                                    All Status
                                </option>

                                <option value="Draft">
                                    Draft
                                </option>

                                <option value="Pending">
                                    Pending
                                </option>

                                <option value="Approved">
                                    Approved
                                </option>

                                <option value="Rejected">
                                    Rejected
                                </option>

                            </select>

                        </div>

                    </div>

                </div>

            </div>

            <div className="card">

                <div className="card-header bg-primary text-white d-flex justify-content-between">

                    <span>Vendor Requests</span>

                    <span>
                        Total Records : {filteredData.length}
                    </span>

                </div>

                <div className="card-body">

                    <div className="table-responsive">

                        <table className="table table-bordered table-striped">

                            <thead>

                                <tr>

                                    <th>ID</th>

                                    <th>Vendor Code</th>

                                    <th>Vendor Name</th>

                                    <th>Vendor Type</th>

                                    <th>Country</th>

                                    <th>Status</th>

                                    <th>Created By</th>

                                    <th>Created</th>

                                    <th>Action</th>

                                </tr>

                            </thead>

                            <tbody>

                                {filteredData.length === 0 ? (

                                    <tr>
                                        <td
                                            colSpan={9}
                                            className="text-center"
                                        >
                                            No Records Found
                                        </td>
                                    </tr>

                                ) : (

                                    filteredData.map(
                                        (item: any) => (

                                            <tr key={item.Id}>

                                                <td>
                                                    {item.Id}
                                                </td>

                                                <td>
                                                    {item.VendorCode}
                                                </td>

                                                <td>
                                                    {item.VendorName}
                                                </td>

                                                <td>
                                                    {item.VendorType}
                                                </td>

                                                <td>
                                                    {item.Country?.Country}
                                                </td>

                                                <td>

                                                    <span
                                                        className={
                                                            item.Status === "Approved"
                                                                ? "badge bg-success"
                                                                : item.Status === "Rejected"
                                                                    ? "badge bg-danger"
                                                                    : item.Status === "Pending"
                                                                        ? "badge bg-warning"
                                                                        : "badge bg-secondary"
                                                        }
                                                    >
                                                        {item.RequestStatus}
                                                    </span>

                                                </td>

                                                <td>
                                                    {item.Author?.Title}
                                                </td>

                                                <td>
                                                    {
                                                        new Date(
                                                            item.Created
                                                        ).toLocaleDateString()
                                                    }
                                                </td>

                                                <td>

                                                    <button
                                                        className="btn btn-primary btn-sm me-1"
                                                        onClick={() =>
                                                            openRequest(
                                                                item.Id
                                                            )
                                                        }
                                                    >
                                                        <i className="bi bi-pencil-square me-1"></i>
                                                        Edit
                                                    </button>

                                                </td>

                                            </tr>

                                        )
                                    )

                                )}

                            </tbody>

                        </table>

                    </div>

                </div>

            </div>

        </div>
    );
};

export default VendorDashboard;