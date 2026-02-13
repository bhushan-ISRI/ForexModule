import * as React from "react";
import { useState, useEffect } from "react";
import { useHistory, Link } from "react-router-dom";
import type { IForexModuleProps } from '../IForexModuleProps';
import "../../components/Pages/Css/InitiatorDashboard.scss";
import SPCRUDOPS from "../../service/BAL/spcrud";

import Left from "../../assets/LeftArrow.png";
import Right from "../../assets/RightArrow.png";

import View from "../../assets/Eye.png";
import Edit from "../../assets/Pencil.png";
import Renew from "../../assets/Renew.png";
export const InitiatorDashboard: React.FC<IForexModuleProps> = (props: IForexModuleProps) => {
    const [searchTerm, setSearchTerm] = useState("");
    const [listData, setListData] = useState<any[]>([]);
    const [filteredData, setFilteredData] = useState<any[]>([]);
    const [loading, setLoading] = useState(true);

    // Pagination
    const itemsPerPage = 10;
    const [currentPage, setCurrentPage] = useState(1);
    const totalPages = Math.ceil(filteredData.length / itemsPerPage);

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
            .replace(/\//g, "-"); // dd-MM-yyyy
    };

    //Latest 
    const GetListData = async () => {
        setLoading(true);

        const spCrudOps = await SPCRUDOPS();

        const parentItems = await spCrudOps.getRootData(
            "NDADetail",
            "ID,ReqtNo,RequestorName,RequestorEmail,RequestDate,VendorName,ProductName,NDAStartDate,NDAExpiryDate,VendorContactName,VendorContactEmail,AdditionalNotes,Status,ServiceType/ID,ServiceType/Title,Author/ID",
            "ServiceType,Author",
            "AuthorId eq " + props.id,
            { column: "ID", isAscending: true },
            5000,
            props);
        setListData(parentItems);
        setFilteredData(parentItems);
        setLoading(false);
    };

    useEffect(() => {
        GetListData();
    }, []);

    // 🔹 Filter data by search term

    useEffect(() => {
        if (!searchTerm) {
            setFilteredData(listData);
            setCurrentPage(1);
            return;
        }

        const lowerSearch = searchTerm.toLowerCase();

        const filtered = listData.filter((item) =>
            Object.values({
                id: item.Id,
                requestDate: formatDate(item.RequestDate),
                vendorName: item.VendorName,
                contactPerson: item.VendorContactName,
                serviceType: item.ServiceType,
                status: item.Status,
                requestorName: item.RequestorName,
                requestorEmail: item.RequestorEmail,
                vendorEmail: item.VendorContactEmail,
                notes: item.AdditionalNotes
            })
                .join(" ")
                .toLowerCase()
                .includes(lowerSearch)
        );

        setFilteredData(filtered);
        setCurrentPage(1);
    }, [searchTerm, listData]);


    // 🔹 Pagination
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
        <div className="min-h-screen bg-gray-100">
            {/* Header */}
            <div className="header">
                <div className="left-banner">
                    <div className="logo-text">
                        <h2> Forex Initiation Dashboard </h2>
                    </div>
                </div>
            </div>
            <div className='col-md-12 px-2 d-flex justify-content-between align-items-center flex-wrap'>
                <div>
                    <input type="text" placeholder="Search..."
                        className="form-control" style={{ width: "250px" }}
                        value={searchTerm}
                        onChange={(e) => setSearchTerm(e.target.value)}
                    />
                </div>
                <div className='Dashbaordcreatebutton'>
                    <Link to="/NewRequest" className='create-button'>New Request</Link>
                </div>
            </div>
            <main className="Main-Dash mx-2">
                <div className="overflow-x-auto">
                    <div className="table-vert-scroll">

                        <table className="custom-table min-w-full bg-white rounded-2xl shadow-md">
                            <thead
                                style={{ backgroundColor: "#3c3e45" }}
                                className="text-white"
                            >
                                <tr>
                                    <th className="px-4 py-2">Req No.</th>
                                    <th className="px-4 py-2">Request Date</th>
                                    <th className="px-4 py-2">Vendor Name</th>
                                    <th className="px-4 py-2">Contact Person</th>
                                    <th className="px-4 py-2">Service Type</th>
                                    <th className="px-4 py-2">Status</th>
                                    <th className="px-4 py-2">View</th>
                                </tr>
                            </thead>
                            <tbody>
                                {paginatedData.map((item, index) => (
                                    <tr key={index} className="border-t">
                                        <td className="px-4 py-2">{item.ReqtNo}</td>
                                        <td className="px-4 py-2">{formatDate(item.RequestDate)}</td>
                                        <td className="px-4 py-2">{item.VendorName}</td>
                                        <td className="px-4 py-2">{item.VendorContactName}</td>
                                        <td className="px-4 py-2">{item.ServiceType?.Title}</td>
                                        <td className="px-4 py-2">{item.Status}</td>
                                        <td className="px-4 py-2">
                                            {item.Status === "Send back" || item.Status === "Draft" ? (
                                                <Link to={`/NDAEditform/${item.Id}`}>
                                                    <img src={Edit} width={15} alt="Edit" />
                                                </Link>
                                            ) : (
                                                <>
                                                    <Link to={`/NDAViewmore/${item.Id}`}>
                                                        <img src={View} width={15} alt="View" />
                                                    </Link>
                                                    {item.Status === "Approved" && (
                                                        <Link to={`/RenewalForm/${item.Id}`} style={{marginLeft : "10px"}}>
                                                            <img src={Renew} width={15} alt="New Action" />
                                                        </Link>
                                                    )}
                                                </>
                                            )}
                                        </td>

                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>

                    {/* Pagination */}
                    <div className="flex justify-center mt-6 overflow-x-auto">
                        <div className="flex space-x-2 flex-nowrap px-4 py-2 bg-#2149d5 rounded shadow" style={{ textAlign: "end" }}>
                            {/* Previous Button */}
                            <button
                                onClick={() => handlePageChange(currentPage - 1)}
                                disabled={currentPage === 1}
                                style={{
                                    backgroundColor: "#fff",
                                    border: "1px solid #000 !important",
                                    marginRight: "5px",
                                    opacity: currentPage === 1 ? 0.5 : 1,
                                }}
                                className="px-3 py-1 border rounded"
                            >
                                <img src={Left} alt="" width={15} />
                            </button>
                            {/* Main Page Numbers */}
                            {Array.from({ length: totalPages }, (_, i) => i + 1)
                                .filter((page) => Math.abs(page - currentPage) <= 2)
                                .map((page) => (
                                    <button
                                        key={page}
                                        onClick={() => handlePageChange(page)}
                                        style={{
                                            backgroundColor: currentPage === page ? "#3c3e45" : "#fff",
                                            color: currentPage === page ? "#fff" : "#000",
                                            fontWeight: currentPage === page ? "bold" : "normal",
                                            margin: currentPage === page ? "5px" : "5px",
                                        }}
                                        className="px-3 py-1 border rounded"
                                    >
                                        {page}
                                    </button>
                                ))}

                            {/* Next Button */}
                            <button
                                onClick={() => handlePageChange(currentPage + 1)}
                                disabled={currentPage === totalPages}
                                style={{
                                    backgroundColor: "#fff",
                                    border: "1px solid #000 !important",
                                    marginLeft: "5px",
                                    opacity: currentPage === totalPages ? 0.5 : 1,
                                }}
                                className="px-3 py-1 border rounded"
                            >
                                <img src={Right} alt="" width={15} />
                            </button>
                        </div>
                    </div>
                </div>
            </main>
        </div>
    );
};

export default InitiatorDashboard;
