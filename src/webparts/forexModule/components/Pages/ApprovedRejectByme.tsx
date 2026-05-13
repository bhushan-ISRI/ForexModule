import * as React from "react";
import { useEffect, useMemo, useState } from "react";
import { Link } from "react-router-dom";
import type { IForexModuleProps } from "../IForexModuleProps";

import "../../components/Pages/Css/ApprovedRejectedbyme.scss";
import SPCRUDOPS from "../../service/BAL/spcrud";
import { useHistory } from "react-router-dom";

const ModernForexDashboard: React.FC<IForexModuleProps> = (props) => {

  const [listData, setListData] = useState<any[]>([]);
  const [filteredData, setFilteredData] = useState<any[]>([]);

  const [loading, setLoading] = useState(false);

  const [searchTerm, setSearchTerm] = useState("");
  const [statusFilter, setStatusFilter] = useState("All");

  // =========================================
  // PAGINATION STATES
  // =========================================
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 10;

  const currentUser = props.context.pageContext.user.displayName
    ?.toLowerCase()
    ?.trim();

  // =========================================
  // FORMAT DATE
  // =========================================
  const history = useHistory();
  const formatDate = (dateString: string) => {

    if (!dateString) return "-";

    return new Date(dateString)
      .toLocaleDateString("en-GB", {
        day: "2-digit",
        month: "2-digit",
        year: "numeric"
      })
      .replace(/\//g, "-");
  };

  // =========================================
  // GET CURRENT APPROVER
  // =========================================
  const getCurrentApprover = (allApprovers: any) => {

    let matrix: any[] = [];

    try {

      matrix =
        typeof allApprovers === "string"
          ? JSON.parse(allApprovers)
          : allApprovers || [];

    } catch {

      matrix = [];
    }

    const current = matrix.find(
      (x: any) => x.Status === "Pending"
    );

    if (!current) return "Completed";

    return `${current.Role} - ${current.Approver}`;
  };

  // =========================================
  // LOAD DATA
  // =========================================
  const loadData = async () => {

    try {

      setLoading(true);

      const spCrudOps = await SPCRUDOPS();

      const data = await spCrudOps.getRootData(
        "ForexRequest",
        "ID,Title,ForexNumber,ForexType,EmployeeName,EmployeeCode,RequestedOn,Status,Location,VendorName,TotalAmount,AllApprovers,WorkFlowHistory",
        "",
        "",
        { column: "ID", isAscending: false },
        5000,
        props
      );

      console.log("🔥 Dashboard Data:", data);

      setListData(Array.isArray(data) ? data : []);
      setFilteredData(Array.isArray(data) ? data : []);

    } catch (e) {

      console.error("❌ Dashboard Error:", e);

      alert("Error loading dashboard");

    } finally {

      setLoading(false);
    }
  };

  useEffect(() => {
    loadData();
  }, []);

  // =========================================
  // DASHBOARD COUNTS
  // =========================================
  const dashboardCounts = useMemo(() => {

    const approvedByMe = listData.filter((item) => {

      let history: any[] = [];

      try {

        history =
          typeof item.WorkFlowHistory === "string"
            ? JSON.parse(item.WorkFlowHistory)
            : item.WorkFlowHistory || [];

      } catch {

        history = [];
      }

      return history.some(
        (h: any) =>
          h.ActionTaken === "Approved" &&
          h.CurrentApprover?.toLowerCase()?.trim() === currentUser
      );

    }).length;

    const rejectedByMe = listData.filter((item) => {

      let history: any[] = [];

      try {

        history =
          typeof item.WorkFlowHistory === "string"
            ? JSON.parse(item.WorkFlowHistory)
            : item.WorkFlowHistory || [];

      } catch {

        history = [];
      }

      return history.some(
        (h: any) =>
          h.ActionTaken === "Rejected" &&
          h.CurrentApprover?.toLowerCase()?.trim() === currentUser
      );

    }).length;

    const sentBackByMe = listData.filter((item) => {

      let history: any[] = [];

      try {

        history =
          typeof item.WorkFlowHistory === "string"
            ? JSON.parse(item.WorkFlowHistory)
            : item.WorkFlowHistory || [];

      } catch {

        history = [];
      }

      return history.some(
        (h: any) =>
          h.ActionTaken === "Sent Back" &&
          h.CurrentApprover?.toLowerCase()?.trim() === currentUser
      );

    }).length;

    return {
      total: listData.length,
      approvedByMe,
      rejectedByMe,
      sentBackByMe
    };

  }, [listData]);

  // =========================================
  // FILTERS
  // =========================================
 useEffect(() => {

  // =====================================
  // DEFAULT DATA = ONLY MY ACTIONED ITEMS
  // =====================================
  let data = listData.filter((item) => {

    let history: any[] = [];

    try {

      history =
        typeof item.WorkFlowHistory === "string"
          ? JSON.parse(item.WorkFlowHistory)
          : item.WorkFlowHistory || [];

    } catch {

      history = [];
    }

    return history.some(
      (h: any) =>
        ["Approved", "Rejected", "Sent Back"].includes(h.ActionTaken) &&
        (h.CurrentApprover || "")
          .toLowerCase()
          .trim() === currentUser
    );
  });

  // =====================================
  // STATUS FILTER
  // =====================================
  if (statusFilter !== "All") {

    if (statusFilter === "ApprovedByMe") {

      data = data.filter((item) => {

        let history: any[] = [];

        try {

          history =
            typeof item.WorkFlowHistory === "string"
              ? JSON.parse(item.WorkFlowHistory)
              : item.WorkFlowHistory || [];

        } catch {

          history = [];
        }

        return history.some(
          (h: any) =>
            h.ActionTaken === "Approved" &&
            (h.CurrentApprover || "")
              .toLowerCase()
              .trim() === currentUser
        );
      });
    }

    else if (statusFilter === "RejectedByMe") {

      data = data.filter((item) => {

        let history: any[] = [];

        try {

          history =
            typeof item.WorkFlowHistory === "string"
              ? JSON.parse(item.WorkFlowHistory)
              : item.WorkFlowHistory || [];

        } catch {

          history = [];
        }

        return history.some(
          (h: any) =>
            h.ActionTaken === "Rejected" &&
            (h.CurrentApprover || "")
              .toLowerCase()
              .trim() === currentUser
        );
      });
    }

    else if (statusFilter === "SentBackByMe") {

      data = data.filter((item) => {

        let history: any[] = [];

        try {

          history =
            typeof item.WorkFlowHistory === "string"
              ? JSON.parse(item.WorkFlowHistory)
              : item.WorkFlowHistory || [];

        } catch {

          history = [];
        }

        return history.some(
          (h: any) =>
            h.ActionTaken === "Sent Back" &&
            (h.CurrentApprover || "")
              .toLowerCase()
              .trim() === currentUser
        );
      });
    }
  }

  // =====================================
  // SEARCH FILTER
  // =====================================
  if (searchTerm) {

    const lowerSearch = searchTerm.toLowerCase();

    data = data.filter((item) =>
      Object.values({
        requestNo: item.ForexNumber,
        employee: item.EmployeeName,
        vendor: item.VendorName,
        status: item.Status,
        location: item.Location,
        amount: item.TotalAmount,
        //status: item.Status,
        requestDate: formatDate(item.RequestedOn),
        forexType: item.ForexType

      })
        .join(" ")
        .toLowerCase()
        .includes(lowerSearch)
    );
  }

  setFilteredData(data);

}, [listData, searchTerm, statusFilter]);

  // =========================================
  // RESET PAGE ON FILTER CHANGE
  // =========================================
  useEffect(() => {
    setCurrentPage(1);
  }, [searchTerm, statusFilter]);

  // =========================================
  // PAGINATION LOGIC
  // =========================================
  const totalPages = Math.ceil(filteredData.length / itemsPerPage);

  const startIndex = (currentPage - 1) * itemsPerPage;
  const endIndex = startIndex + itemsPerPage;

  const paginatedData = filteredData.slice(startIndex, endIndex);

  return (

    <div className="approved-rejected-dashboard">   

      <div className="dashboard-header">

        <div>
          <h2>Forex Workflow Dashboard</h2>
          <p>Track approvals, rejections and workflow status</p>
        </div>

        {/* <Link to="/NewRequest" className="new-request-btn">
          + New Request
        </Link> */}

      </div>

      {/* =============================== */}
      {/* CARDS */}
      {/* =============================== */}
      <div className="dashboard-cards">

        <div className="dashboard-card total">
          <h3>Total Requests</h3>
          <h1>{dashboardCounts.total}</h1>
        </div>

        <div className="dashboard-card approved">
          <h3>Approved By Me</h3>
          <h1>{dashboardCounts.approvedByMe}</h1>
        </div>

        <div className="dashboard-card rejected">
          <h3>Rejected By Me</h3>
          <h1>{dashboardCounts.rejectedByMe}</h1>
        </div>

        <div className="dashboard-card sentback">
          <h3>Sent Back By Me</h3>
          <h1>{dashboardCounts.sentBackByMe}</h1>
        </div>

      </div>

      {/* =============================== */}
      {/* FILTERS */}
      {/* =============================== */}
      <div className="filter-section">

        <input
          type="text"
          placeholder="Search request..."
          className="form-control"
          value={searchTerm}
          onChange={(e) => setSearchTerm(e.target.value)}
        />

        <select
          className="form-control"
          value={statusFilter}
          onChange={(e) => setStatusFilter(e.target.value)}
        >
          <option value="All">All</option>
          <option value="Pending">Pending</option>
          <option value="ApprovedByMe">Approved By Me</option>
          <option value="RejectedByMe">Rejected By Me</option>
          <option value="SentBackByMe">Sent Back By Me</option>
          <option value="Rejected">Rejected</option>
          <option value="Sent Back">Sent Back</option>
          <option value="Paid">Paid</option>
        </select>

      </div>

      {/* =============================== */}
      {/* TABLE */}
      {/* =============================== */}
      <div className="table-wrapper">

        <table className="custom-table">

          <thead>
            <tr>
              <th>Request No</th>
              <th>Type</th>
              <th>Employee</th>
              <th>Vendor</th>
              <th>Location</th>
              <th>Amount</th>
              <th>Status</th>
              {/* <th>Pending At</th> */}
              <th>Date</th>
              <th>Action</th>
            </tr>
          </thead>

          <tbody>

            {loading ? (

              <tr>
                <td colSpan={10}>Loading...</td>
              </tr>

            ) : paginatedData.length === 0 ? (

              <tr>
                <td colSpan={10}>No Records Found</td>
              </tr>

            ) : (

              paginatedData.map((item) => (

                <tr key={item.ID}>

                  <td>{item.ForexNumber}</td>

                  <td>{item.ForexType}</td>

                  <td>{item.EmployeeName}</td>

                  <td>{item.VendorName}</td>

                  <td>{item.Location}</td>

                  <td>₹ {item.TotalAmount}</td>

                  <td>
                    <span
                      className={`status-badge ${item.Status?.replace(/\s/g, '-')}`}
                    >
                      {item.Status}
                    </span>
                  </td>

                  {/* <td>
                     {getCurrentApprover(item.AllApprovers)} 
                  </td> */}

                  <td>
                    {formatDate(item.RequestedOn)}
                  </td>

                  <td>

                    <Link to={`/ViewRequest/${item.ID}`}>
                      View
                    </Link>

                  </td>

                </tr>

              ))
            )}

          </tbody>

        </table>

        {/* =============================== */}
        {/* PAGINATION */}
        {/* =============================== */}
        {!loading && totalPages > 1 && (

          <div className="pagination-wrapper">

            <button
              className="page-btn"
              disabled={currentPage === 1}
              onClick={() => setCurrentPage((prev) => prev - 1)}
            >
              Previous
            </button>

            {[...Array(totalPages)].map((_, index) => (

              <button
                key={index}
                className={`page-number ${
                  currentPage === index + 1 ? "active-page" : ""
                }`}
                onClick={() => setCurrentPage(index + 1)}
              >
                {index + 1}
              </button>

            ))}

            <button
              className="page-btn"
              disabled={currentPage === totalPages}
              onClick={() => setCurrentPage((prev) => prev + 1)}
            >
              Next
            </button>

          </div>

        )}

      </div>
      <a
    onClick={() => history.push("/")}
    className="exit-btn"
>
    Exit
</a>

    </div>
  );
};

export default ModernForexDashboard;