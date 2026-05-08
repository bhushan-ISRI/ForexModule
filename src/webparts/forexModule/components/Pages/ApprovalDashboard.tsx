import * as React from "react";
import { useState, useEffect } from "react";
import { useHistory, Link } from "react-router-dom";
import type { IForexModuleProps } from "../IForexModuleProps";
import "../../components/Pages/Css/InitiatorDashboard.scss";
import SPCRUDOPS from "../../service/BAL/spcrud";

import Left from "../../assets/LeftArrow.png";
import Right from "../../assets/RightArrow.png";
import View from "../../assets/Eye.png";
import Edit from "../../assets/Pencil.png";
import Renew from "../../assets/Renew.png";

export const ApprovalDashboard: React.FC<IForexModuleProps> = (
  props: IForexModuleProps
) => {
  const [searchTerm, setSearchTerm] = useState("");
  const [statusFilter, setStatusFilter] = useState("All");
  const [listData, setListData] = useState<any[]>([]);
  const [filteredData, setFilteredData] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);

  // Pagination
  const itemsPerPage = 10;
  const [currentPage, setCurrentPage] = useState(1);

  const history = useHistory();

  const formatDate = (dateString: string) => {
    if (!dateString) return "";
    const date = new Date(dateString);

    return date
      .toLocaleDateString("en-GB", {
        day: "2-digit",
        month: "2-digit",
        year: "numeric",
      })
      .replace(/\//g, "-");
  };

  // 🔹 Fetch Data
  const GetListData = async () => {
    setLoading(true);

    const spCrudOps = await SPCRUDOPS();

    const parentItems = await spCrudOps.getRootData(
      "ForexRequest",
      "ID,ForexNumber,EmployeeName,EmployeeCode,RequestedOn,Status,Location,VendorName,TotalAmount,Author/Id,CurrentApprover/Id,NextApprovers/Id,ForexType",
      "Author,CurrentApprover,NextApprovers",
      "CurrentApproverId eq " + props.id,
      { column: "ID", isAscending: false },
      5000,
      props
    );

    setListData(parentItems);
    setFilteredData(parentItems);
    setLoading(false);
  };

  useEffect(() => {
    GetListData();
  }, []);

  // 🔹 Search + Status Filter Combined
  useEffect(() => {
    let data = [...listData];

    if (statusFilter !== "All") {
      data = data.filter(
        (item) =>
          item.Status?.toLowerCase().trim() ===
          statusFilter.toLowerCase().trim()
      );
    }

    if (searchTerm) {
      const lowerSearch = searchTerm.toLowerCase();

      data = data.filter((item) =>
        Object.values({
          requestNo: item.ForexNumber,
          requesttype: item.ForexType,
          // requestDate: formatDate(item.RequestDate),
          vendorName: item.VendorName,
          EmployeeCode: item.EmployeeCode,
          EmployeeName: item.EmployeeName,
          requestDate: formatDate(item.RequestedOn),
          status: item.Status,
          amount: item.TotalAmount,
        })
          .join(" ")
          .toLowerCase()
          .includes(lowerSearch)
      );
    }

    setFilteredData(data);
    setCurrentPage(1);
  }, [searchTerm, statusFilter, listData]);

  // 🔹 Pagination
  const totalPages = Math.ceil(filteredData.length / itemsPerPage);

  const handlePageChange = (page: number) => {
    if (page >= 1 && page <= totalPages) {
      setCurrentPage(page);
    }
  };

  const sortedData = [...filteredData].sort((a, b) => b.ID - a.ID);

  const paginatedData = sortedData.slice(
    (currentPage - 1) * itemsPerPage,
    currentPage * itemsPerPage
  );

  return (
    <div className="dashboard-wrapper">

      {/* Header (UNCHANGED) */}
      <div className="header">
        <div className="left-banner">
          <div className="logo-text">
            <h2>Approval Dashboard</h2>
          </div>
        </div>
      </div>

      {/* Filters */}
      <div className="filter-section">
        <div className="filter-left">
          <input
            type="text"
            placeholder="Search..."
            className="form-control"
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
          />

          <select
            className="form-control"
            value={statusFilter}
            onChange={(e) => setStatusFilter(e.target.value)}
          >
            <option value="All">All Status</option>

            {/* <option value="Draft">Draft</option> */}
            <option value="Pending">Pending</option>
            <option value="Pending with Treasury for Verification">Pending with Treasury for Verification</option>
            <option value="Pending for Vouching">Pending for Vouching</option>

            {/* <option value="Approved">Approved</option> */}
            <option value="Sent Back">Sent Back</option>
            <option value="Rejected">Rejected</option>

            <option value="Paid And Closed">Paid and Closed</option>
            <option value="Paid and Pending for Settlement">Paid and Pending for Settlement</option>
            {/* <option value="Verified and Closed with Bank">Verified and Closed with Bank</option> */}
            {/* <option value="Closed">Closed</option> */}
          </select>

        </div>

        {/* <Link to="/NewRequest" className="create-button">
        + New Request
      </Link> */}
      </div>

      {/* Table */}
      <div className="table-section">
        <div className="table-vert-scroll">
          <table className="custom-table">
            <thead>
              <tr>
                <th>Request No.</th>
                <th>Request Type</th>
                <th>EmployeeCode</th>
                <th>Employee Name</th>
                <th>Request Date</th>
                <th>Location</th>
                <th>Vendor Name</th>
                <th>Amount</th>
                <th>Status</th>
                <th>Action</th>
              </tr>
            </thead>

            <tbody>
              {loading ? (
                <tr>
                  <td colSpan={8} className="text-center">
                    Loading data...
                  </td>
                </tr>
              ) : paginatedData.length === 0 ? (
                <tr>
                  <td colSpan={8} className="text-center">
                    No records found
                  </td>
                </tr>
              ) : (
                paginatedData.map((item) => (
                  <tr key={item.ID}>
                    <td>{item.ForexNumber}</td>
                    <td>{item.ForexType}</td>
                    <td>{item.EmployeeCode}</td>
                    <td>{item.EmployeeName}</td>
                    <td>{formatDate(item.RequestedOn)}</td>
                    <td>{item.Location}</td>
                    <td>{item.VendorName}</td>
                    <td>₹ {item.TotalAmount || "-"}</td>
                    <td>
                      <span
                        className={`status-badge ${item.Status?.replace(" ", "-")}`}
                      >
                        {item.Status}
                      </span>
                    </td>
                    <td>
                      {item.Status === "Paid and Pending for Settlement" ? (
                        <Link to={`/TrackerApprovalForm/${item.ID}`}>
                          <img src={Edit} width={16} alt="Tracker Approval" />
                        </Link>
                      ) : (
                        <Link to={`/ApprovalRequest/${item.ID}`}>
                          <img src={Edit} width={16} alt="Approval" />
                        </Link>
                      )}
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>

        {/* Pagination */}
        <div className="pagination">
          <button
            onClick={() => handlePageChange(currentPage - 1)}
            disabled={currentPage === 1}
          >
            <img src={Left} width={14} />
          </button>

          {Array.from({ length: totalPages }, (_, i) => i + 1)
            .filter((page) => Math.abs(page - currentPage) <= 2)
            .map((page) => (
              <button
                key={page}
                onClick={() => handlePageChange(page)}
                className={currentPage === page ? "active" : ""}
              >
                {page}
              </button>
            ))}

          <button
            onClick={() => handlePageChange(currentPage + 1)}
            disabled={currentPage === totalPages}
          >
            <img src={Right} width={14} />
          </button>
        </div>
      </div>
    </div>
  );


};

export default ApprovalDashboard;
