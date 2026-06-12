import * as React from "react";
import { IForexModuleProps } from "../IForexModuleProps";
import SPCRUDOPS from "../../service/BAL/spcrud";
import { useHistory, Link } from "react-router-dom";
import edit from '../../assets/Pencil.png';

const VendorDashboard: React.FC<IForexModuleProps> = (props) => {

    const spCrudOps = SPCRUDOPS();

    const [vendorData, setVendorData] = React.useState<any[]>([]);
    const [searchText, setSearchText] = React.useState("");
    const [statusFilter, setStatusFilter] = React.useState("");
    const [currentPage, setCurrentPage] = React.useState(1);
    const recordsPerPage = 10;
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
                { column: "Created", isAscending: false },
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
                x.RequestStatus === "Pending"
        ).length;

    const approvedRequests =
        vendorData.filter(
            (x: any) =>
                x.RequestStatus === "Approved"
        ).length;

    const rejectedRequests =
        vendorData.filter(
            (x: any) =>
                x.RequestStatus === "Rejected"
        ).length;

    const draftRequests =
        vendorData.filter(
            (x: any) =>
                x.RequestStatus === "Draft"
        ).length;

    const filteredData = vendorData
        .filter((x: any) => {

            const vendorMatch =
                !searchText ||
                x.VendorName?.toLowerCase().includes(
                    searchText.toLowerCase()
                );

            const statusMatch =
                !statusFilter ||
                x.RequestStatus === statusFilter;

            return vendorMatch && statusMatch;
        })
        .sort((a: any, b: any) => b.Id - a.Id);

    const indexOfLastRecord = currentPage * recordsPerPage;

    const indexOfFirstRecord = indexOfLastRecord - recordsPerPage;

    const currentRecords = filteredData.slice(
        indexOfFirstRecord,
        indexOfLastRecord
    );

    const totalPages = Math.ceil(filteredData.length / recordsPerPage);
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


        <>

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



            <div className="header">
                <h2> Vendor Dashboard </h2>
            </div>
            <div className="mainsecondapprove">
                <div className="mainsecondsmall">
                    <div>
                        <input
                            type="text"
                            className="form-control"
                            style={{ width: "250px;" }}
                            placeholder="Search Vendor Name..."
                            value={searchText}
                            onChange={(e) =>
                                setSearchText(e.target.value)
                            }
                        />

                    </div>
                    <div>
                        <select className="form-controltext" style={{ width: "250px" }}
                            value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)}>
                            <option value=""> All Status </option>
                            <option value="Draft"> Draft </option>
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
                <div>
                    <Link to="/CreationForm" className="create-button"> + Vendor Creation Request </Link>
                </div>
            </div>





            {/* <div className="d-flex justify-content-between align-items-center mb-3">

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

            </div> */}

            {/* <div className="card shadow-sm mb-3"> */}

            {/* <div className="card-header bg-light">
                    <strong>Search Vendor Requests</strong>
                </div> */}

            {/* <div className="card-body">

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

                </div> */}

            {/* </div> */}

            <div className="card">


                {/* <div className="card-header bg-primary text-white d-flex justify-content-between">

                    <span>Vendor Requests</span>

                    <span>
                        Total Records : {filteredData.length}
                    </span>

                </div> */}

                <div className="card-body">

                    <div style={{ overflowX: "auto" }}>

                        <table className="custom-table">

                            <thead>

                                <tr>

                                    <th>ID</th>

                                    <th>Vendor Code</th>

                                    <th>Vendor Name</th>

                                    {/* <th>Vendor Type</th> */}

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

                                    currentRecords.map(
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

                                                {/* <td>
                                                    {item.VendorType}
                                                </td> */}

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
                                                                        : "badge"
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
                                                    {(item.RequestStatus === "Approved" && item.TaxDocumentAvailable === "No") && (

                                                    <a  onClick={() => openRequest(item.Id)}>
                                                        <img src={edit} alt="" width={15} />
                                                    </a>
                                                   
                                                    )}
                                                </td>

                                            </tr>

                                        )
                                    )

                                )}

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

export default VendorDashboard;