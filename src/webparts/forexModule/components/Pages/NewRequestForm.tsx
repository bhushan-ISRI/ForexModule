import React, { useState } from "react";
import "../Pages/Css/NewRequest.scss";
import { IForexModuleProps } from "../IForexModuleProps";
import { useHistory } from 'react-router-dom';

interface InvoiceRow {
    invoiceNo: string;
    invoiceDate: string;
    boeNo: string;
    boeDate: string;
    mrnNo: string;
    blNo: string;
    blDate: string;
    invoiceAmount: string;
    mrnDate?: string; // Added for Service-Bill Payment
}

const NewRequest = (props: IForexModuleProps) => {
        const history = useHistory();
    const [paymentType, setPaymentType] = useState("Goods-Bill Payment");
    const [taxDocumentView, setTaxDocumentView] = useState("Yes");

    const [rows, setRows] = React.useState([
        {
            invoiceNo: "",
            invoiceDate: "",
            boeNo: "",
            boeDate: "",
            mrnNo: "",
            blNo: "",
            blDate: "",
            invoiceAmount: "",
            mrnDate: "" // Added for Service-Bill Payment
        }
    ]);
    const addRow = () => {
        setRows([
            ...rows,
            {
                invoiceNo: "",
                invoiceDate: "",
                boeNo: "",
                boeDate: "",
                mrnNo: "",
                blNo: "",
                blDate: "",
                invoiceAmount: "",
                mrnDate: "" // Added for Service-Bill Payment
            }
        ]);
    };

    const deleteRow = (index: number) => {
        const updatedRows = rows.filter((_, i) => i !== index);
        setRows(updatedRows);
    };

    const handleChange = (
        index: number,
        field: keyof InvoiceRow,
        value: string
    ) => {
        const updatedRows = [...rows];
        updatedRows[index][field] = value;
        setRows(updatedRows);
    };

    const totalInvoiceAmount = rows.reduce((sum, row) => {
        return sum + (parseFloat(row.invoiceAmount) || 0);
    }, 0);
    const uniqueBoeNumbers = rows
        .map(r => r.boeNo)
        .filter((value, index, self) =>
            value && self.indexOf(value) === index
        );

    const uniqueBlNumbers = rows
        .map(r => r.blNo)
        .filter((value, index, self) =>
            value && self.indexOf(value) === index
        );
    return (
        <div className="forex-wrapper">

            {/* ================= HEADER ================= */}
            <div className="forex-header">
                <h2>Forex Payment Request Form</h2>
            </div>

            <div className="forex-card">

                {/* ================= REQUESTOR ================= */}
                <Section title="Requestor Information">
                    <Grid>
                        <Field label="Type">
                            <select value={paymentType} onChange={(e) => setPaymentType(e.target.value)}>
                                <option value="Goods-Bill Payment">Goods-Bill Payment</option>
                                <option value="Service-Bill Payment">Service-Bill Payment</option>
                                <option value="Goods-Advance Payment">Goods-Advance Payment</option>
                                <option value="Service-Advance Payment">Service-Advance Payment</option>
                            </select>
                        </Field>
                    </Grid>
                    <Grid style={{ marginTop: "20px" }}>
                        <Field label="Employee Code"><input type="text" /></Field>
                        <Field label="Employee Name"><input type="text" /></Field>
                        <Field label="Division"><input type="text" /></Field>
                        <Field label="Location"><input type="text" /></Field>
                        <Field label="RM"><input type="text" /></Field>
                        <Field label="HOD"><input type="text" /></Field>
                        <Field label="Contact No"><input type="text" /></Field>
                        <Field label="Employee Status"><input type="text" /></Field>
                        <Field label="Email" full><input type="email" /></Field>
                    </Grid>
                </Section>

                {/* ================= VENDOR ================= */}
                <Section title="Vendor / Beneficiary Details">
                    <Grid>
                        <Field label="Vendor Code"><input /></Field>
                        <Field label="Vendor Name"><input /></Field>
                        <Field label="Address" full><input /></Field>
                        <Field label="City"><input /></Field>
                        <Field label="Country"><input /></Field>
                        <Field label="Pincode"><input /></Field>
                        <Field label="Bank Name"><input /></Field>
                        <Field label="Bank Country"><input /></Field>
                        <Field label="Bank Swift Code"><input /></Field>
                        <Field label="Bank Branch Address"><input /></Field>
                        <Field label="Bank IBAN / Account No" full><input /></Field>
                    </Grid>
                </Section>

                {/* ================= TAX INFO ================= */}
                <Section title="Tax & Regulatory Information">
                    <Grid>
                        <Field label="Nature of Payment"><input /></Field>
                        <Field label="Tax Document Available?">
                            <select onChange={(e) => { setTaxDocumentView(e.target.value) }}>
                                <option>Yes</option>
                                <option>No</option>
                            </select>
                        </Field>
                        {taxDocumentView === "No" && (
                            <Field >
                                <span style={{ color: "red" }}>
                                    (if No, withholding tax will be applicable)
                                </span>
                            </Field>
                        )}

                        {taxDocumentView === "Yes" && (
                            <Field label="DTAA Applicable?">
                                <select><option>Yes</option><option>No</option></select>
                            </Field>
                        )}
                    </Grid>
                </Section>
                {taxDocumentView === "Yes" && (
                    <><Section title="Permanent Establishment Declaration">
                        <Grid>
                            <Field label="Document Available"><select><option>Yes</option><option>No</option></select></Field>
                            <Field label="Document Number"><input /></Field>
                            <Field label="Document Date"><input type="date" /></Field>
                            <Field label="Validity Start Date"><input type="date" /></Field>
                            <Field label="Validity End Date"><input type="date" /></Field>
                            <Field label="View Document"><input type="file" /></Field>

                        </Grid>
                    </Section><Section title="Tax Residency Certificate">
                            <Grid>
                                <Field label="Document Available"><select><option>Yes</option><option>No</option></select></Field>
                                <Field label="Document Number"><input /></Field>
                                <Field label="Country of Tax Residence"><input /></Field>
                                <Field label="Tax Identification Number"><input /></Field>
                                <Field label="Validity Start Date"><input type="date" /></Field>
                                <Field label="Validity End Date"><input type="date" /></Field>
                                <Field label="View Document"><input type="file" /></Field>

                            </Grid>
                        </Section><Section title="Form 10F">
                            <Grid>
                                <Field label="Document Available"><select><option>Yes</option><option>No</option></select></Field>
                                <Field label="Document Number"><input /></Field>
                                <Field label="Acknowledgment Number"><input /></Field>
                                <Field label="Document Date"><input type="date" /></Field>
                                <Field label="Validity Start Date"><input type="date" /></Field>
                                <Field label="Validity End Date"><input type="date" /></Field>
                                <Field label="View Document"><input type="file" /></Field>
                            </Grid>
                        </Section></>
                )}
                <Section >
                    <Grid>
                        <Field label="From"><input type="date" /></Field>

                        <Field label="To"><input type="date" /></Field>
                    </Grid>
                    <Grid>

                        <Field label="Eligible amount that can be transmitted without WHT">
                            <input type="number" />
                        </Field>


                        <Field label="Paid Amount">
                            <input type="number" />
                        </Field>

                        <Field label="Balance eligible amount(Without with holding Tax)">
                            <input type="number" />
                        </Field>
                    </Grid>

                </Section>

                {/* ================= FOREX DETAILS ================= */}
                {/* ==============================Goods-Bill Payment============================== */}
                {paymentType === "Goods-Bill Payment" && (
                    <Section title="Forex Payment Request Details">
                        <Grid>
                            <Field label="Request Number"><input /></Field>
                            <Field label="Requested On"><input type="date" /></Field>
                            <Field label="Currency"><input /></Field>
                            <Field label="Total Amount"><input type="number" /></Field>
                            <Field label="Foreign Bank Charges"><input type="number" /></Field>
                            {/* <Field label="PO/Contract No"><input /></Field>
                            <Field label="PO Date"><input type="date" /></Field>
                            <Field label="Expected Settlement Date"><input type="date" /></Field> */}
                        </Grid>

                        <table className="data-table" style={{ marginTop: "10px" }}>
                            <thead>
                                <tr>
                                    <th>Sr.No</th>
                                    <th>Invoice No</th>
                                    <th>Invoice Date</th>
                                    <th>BOE No</th>
                                    <th>BOE Date</th>
                                    <th>MRN No</th>
                                    <th>Bill of Lading No</th>
                                    <th>Bill of Lading Date</th>
                                    <th>Invoice Amount</th>
                                    <th>Attach Invoice</th>
                                    <th>Attach Other Docs</th>
                                    <th>Add/Delete Entry</th>
                                </tr>
                            </thead>
                            <tbody>
                                {rows.map((row, index) => (
                                    <tr key={index}>
                                        <td>{index + 1}</td>

                                        <td>
                                            <input
                                                value={row.invoiceNo}
                                                onChange={(e) =>
                                                    handleChange(index, "invoiceNo", e.target.value)
                                                }
                                            />
                                        </td>

                                        <td>
                                            <input
                                                type="date"
                                                value={row.invoiceDate}
                                                onChange={(e) =>
                                                    handleChange(index, "invoiceDate", e.target.value)
                                                }
                                            />
                                        </td>

                                        <td>
                                            <input
                                                value={row.boeNo}
                                                onChange={(e) =>
                                                    handleChange(index, "boeNo", e.target.value)
                                                }
                                            />
                                        </td>

                                        <td>
                                            <input
                                                type="date"
                                                value={row.boeDate}
                                                onChange={(e) =>
                                                    handleChange(index, "boeDate", e.target.value)
                                                }
                                            />
                                        </td>

                                        <td>
                                            <input
                                                value={row.mrnNo}
                                                onChange={(e) =>
                                                    handleChange(index, "mrnNo", e.target.value)
                                                }
                                            />
                                        </td>

                                        <td>
                                            <input
                                                value={row.blNo}
                                                onChange={(e) =>
                                                    handleChange(index, "blNo", e.target.value)
                                                }
                                            />
                                        </td>

                                        <td>
                                            <input
                                                type="date"
                                                value={row.blDate}
                                                onChange={(e) =>
                                                    handleChange(index, "blDate", e.target.value)
                                                }
                                            />
                                        </td>

                                        <td>
                                            <input
                                                type="number"
                                                value={row.invoiceAmount}
                                                onChange={(e) =>
                                                    handleChange(index, "invoiceAmount", e.target.value)
                                                }
                                            />
                                        </td>

                                        <td>
                                            <input type="file" />
                                        </td>

                                        <td>
                                            <input type="file" />
                                        </td>

                                        <td style={{ textAlign: "center" }}>
                                            {/* Show PLUS only on last row */}
                                            {index === rows.length - 1 && (
                                                <button
                                                    type="button"
                                                    onClick={addRow}
                                                    style={{
                                                        background: "#28a745",
                                                        color: "white",
                                                        marginRight: "5px",
                                                        border: "none",
                                                        padding: "5px 10px",
                                                        cursor: "pointer",
                                                        borderRadius: "4px"
                                                    }}
                                                >
                                                    +
                                                </button>
                                            )}

                                            {/* Show DELETE if more than 1 row */}
                                            {rows.length > 1 && (
                                                <button
                                                    type="button"
                                                    onClick={() => deleteRow(index)}
                                                    style={{
                                                        background: "#dc3545",
                                                        color: "white",
                                                        border: "none",
                                                        padding: "5px 10px",
                                                        cursor: "pointer",
                                                        borderRadius: "4px"
                                                    }}
                                                >
                                                    ✖
                                                </button>
                                            )}
                                        </td>

                                    </tr>
                                ))}
                            </tbody>
                            <tfoot>
                                <tr>
                                    <td colSpan={8} style={{ textAlign: "right", fontWeight: "bold" }}>
                                        Total Invoice Amount:
                                    </td>
                                    <td style={{ fontWeight: "bold" }}>
                                        {totalInvoiceAmount.toFixed(2)}
                                    </td>
                                    <td colSpan={3}></td>
                                </tr>
                            </tfoot>
                        </table>
                        <div style={{ display: "flex", gap: "40px", marginTop: "30px" }}>

                            {/* BOE TABLE */}
                            <div style={{ flex: 1 }}>
                                <p style={{ color: "red", fontSize: "13px" }}>
                                    Unique BOE no will be listed below
                                </p>

                                <table className="data-table">
                                    <thead>
                                        <tr>
                                            <th>BOE No</th>
                                            <th>Attach Documents</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {uniqueBoeNumbers.map((boe, index) => (
                                            <tr key={index}>
                                                <td>{boe}</td>
                                                <td>
                                                    <input type="file" />
                                                </td>
                                            </tr>
                                        ))}
                                    </tbody>
                                </table>
                            </div>
                            {/* BILL OF LADING TABLE */}
                            <div style={{ flex: 1 }}>
                                <p style={{ color: "red", fontSize: "13px" }}>
                                    Unique Bill of lading no will be listed below
                                </p>

                                <table className="data-table">
                                    <thead>
                                        <tr>
                                            <th>Bill of Lading Number</th>
                                            <th>Attach Documents</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {uniqueBlNumbers.map((bl, index) => (
                                            <tr key={index}>
                                                <td>{bl}</td>
                                                <td>
                                                    <input type="file" />
                                                </td>
                                            </tr>
                                        ))}
                                    </tbody>
                                </table>
                            </div>

                        </div>

                    </Section>
                )}
                {/* ==============================service-Bill Payment============================== */}
                {paymentType === "Service-Bill Payment" && (
                    <Section title="Forex Payment Request Details">
                        <Grid>
                            <Field label="Request Number"><input /></Field>
                            <Field label="Requested On"><input type="date" /></Field>
                            <Field label="Currency"><input /></Field>
                            <Field label="Total Amount"><input type="number" /></Field>
                            <Field label="Foreign Bank Charges"><input type="number" /></Field>
                            {/* <Field label="PO/Contract No"><input /></Field>
                            <Field label="PO Date"><input type="date" /></Field>
                            <Field label="Expected Settlement Date"><input type="date" /></Field> */}
                        </Grid>

                        <table className="data-table" style={{ marginTop: "10px" }}>
                            <thead>
                                <tr>
                                    <th>Sr No</th>
                                    <th>Invoice No</th>
                                    <th>Invoice Date</th>
                                    <th>Invoice Amount</th>
                                    <th>MRN No</th>
                                    <th>MRN Date</th>
                                    <th>Attach Invoice</th>
                                    <th>Attach Other Document</th>
                                    <th>Add/Delete Entry</th>
                                </tr>
                            </thead>
                            <tbody>
                                {rows.map((row, index) => (
                                    <tr key={index}>
                                        <td>{index + 1}</td>

                                        <td>
                                            <input
                                                value={row.invoiceNo}
                                                onChange={(e) =>
                                                    handleChange(index, "invoiceNo", e.target.value)
                                                }
                                            />
                                        </td>

                                        <td>
                                            <input
                                                type="date"
                                                value={row.invoiceDate}
                                                onChange={(e) =>
                                                    handleChange(index, "invoiceDate", e.target.value)
                                                }
                                            />
                                        </td>

                                        <td>
                                            <input
                                                value={row.invoiceAmount}
                                                onChange={(e) =>
                                                    handleChange(index, "invoiceAmount", e.target.value)
                                                }
                                            />
                                        </td>

                                        <td>
                                            <input
                                                value={row.mrnNo}
                                                onChange={(e) =>
                                                    handleChange(index, "mrnNo", e.target.value)
                                                }
                                            />
                                        </td>
                                        <td>
                                            <input
                                                type="date"
                                                value={row.mrnDate}
                                                onChange={(e) =>
                                                    handleChange(index, "mrnDate", e.target.value)
                                                }
                                            />
                                        </td>
                                        <td>
                                            <input
                                                type="file"


                                            />
                                        </td>

                                        <td>
                                            <input
                                                type="file"

                                            />
                                        </td>




                                        <td style={{ textAlign: "center" }}>
                                            {/* Show PLUS only on last row */}
                                            {index === rows.length - 1 && (
                                                <button
                                                    type="button"
                                                    onClick={addRow}
                                                    style={{
                                                        background: "#28a745",
                                                        color: "white",
                                                        marginRight: "5px",
                                                        border: "none",
                                                        padding: "5px 10px",
                                                        cursor: "pointer",
                                                        borderRadius: "4px"
                                                    }}
                                                >
                                                    +
                                                </button>
                                            )}

                                            {/* Show DELETE if more than 1 row */}
                                            {rows.length > 1 && (
                                                <button
                                                    type="button"
                                                    onClick={() => deleteRow(index)}
                                                    style={{
                                                        background: "#dc3545",
                                                        color: "white",
                                                        border: "none",
                                                        padding: "5px 10px",
                                                        cursor: "pointer",
                                                        borderRadius: "4px"
                                                    }}
                                                >
                                                    ✖
                                                </button>
                                            )}
                                        </td>

                                    </tr>
                                ))}
                            </tbody>
                            <tfoot>
                                <tr>
                                    <td colSpan={3} style={{ textAlign: "right", fontWeight: "bold" }}>
                                        Total Invoice Amount:
                                    </td>
                                    <td style={{ fontWeight: "bold" }}>
                                        {totalInvoiceAmount.toFixed(2)}
                                    </td>
                                    <td colSpan={5}></td>
                                </tr>
                            </tfoot>
                        </table>
                    </Section>
                )}

                {/* ==============================Goods - Advance Payment || Service - Advance Payment============================== */}
                {paymentType === "Service-Advance Payment" && (

                    <Section title="Forex Payment Request Details">
                        <Grid>
                            <Field label="Request Number"><input /></Field>
                            <Field label="Requested On"><input type="date" /></Field>
                            <Field label="Currency"><input /></Field>
                            <Field label="Total Amount"><input type="number" /></Field>
                            <Field label="Foreign Bank Charges"><input type="number" /></Field>
                            <Field label="PO/Contract No"><input /></Field>
                            <Field label="PO Date"><input type="date" /></Field>
                            <Field label="Expected Settlement Date"><input type="date" /></Field>
                        </Grid>

                        <table className="data-table" style={{ marginTop: "10px" }}>
                            <thead>
                                <tr>
                                    <th>Sr No</th>
                                    <th>Performa Invoice No</th>
                                    <th>Performa Invoice Date</th>
                                    <th>Performa Invoice Amount</th>
                                    <th>Attach PO</th>
                                    <th>Attach PI</th>
                                    <th>Attach Other Document</th>
                                    <th>Add/Delete Entry</th>
                                </tr>
                            </thead>
                            <tbody>
                                {rows.map((row, index) => (
                                    <tr key={index}>
                                        <td>{index + 1}</td>

                                        <td>
                                            <input
                                                value={row.invoiceNo}
                                                onChange={(e) =>
                                                    handleChange(index, "invoiceNo", e.target.value)
                                                }
                                            />
                                        </td>

                                        <td>
                                            <input
                                                type="date"
                                                value={row.invoiceDate}
                                                onChange={(e) =>
                                                    handleChange(index, "invoiceDate", e.target.value)
                                                }
                                            />
                                        </td>

                                        <td>
                                            <input
                                                value={row.invoiceAmount}
                                                onChange={(e) =>
                                                    handleChange(index, "invoiceAmount", e.target.value)
                                                }
                                            />
                                        </td>


                                        <td>
                                            <input
                                                type="file"

                                            />
                                        </td>
                                        <td>
                                            <input
                                                type="file"


                                            />
                                        </td>

                                        <td>
                                            <input
                                                type="file"

                                            />
                                        </td>




                                        <td style={{ textAlign: "center" }}>
                                            {/* Show PLUS only on last row */}
                                            {index === rows.length - 1 && (
                                                <button
                                                    type="button"
                                                    onClick={addRow}
                                                    style={{
                                                        background: "#28a745",
                                                        color: "white",
                                                        marginRight: "5px",
                                                        border: "none",
                                                        padding: "5px 10px",
                                                        cursor: "pointer",
                                                        borderRadius: "4px"
                                                    }}
                                                >
                                                    +
                                                </button>
                                            )}

                                            {/* Show DELETE if more than 1 row */}
                                            {rows.length > 1 && (
                                                <button
                                                    type="button"
                                                    onClick={() => deleteRow(index)}
                                                    style={{
                                                        background: "#dc3545",
                                                        color: "white",
                                                        border: "none",
                                                        padding: "5px 10px",
                                                        cursor: "pointer",
                                                        borderRadius: "4px"
                                                    }}
                                                >
                                                    ✖
                                                </button>
                                            )}
                                        </td>

                                    </tr>
                                ))}
                            </tbody>
                            <tfoot>
                                <tr>
                                    <td colSpan={3} style={{ textAlign: "right", fontWeight: "bold" }}>
                                        Total Invoice Amount:
                                    </td>
                                    <td style={{ fontWeight: "bold" }}>
                                        {totalInvoiceAmount.toFixed(2)}
                                    </td>
                                    <td colSpan={4}></td>
                                </tr>
                            </tfoot>
                        </table>
                    </Section>
                )}
                {paymentType === "Goods-Advance Payment" && (

                    <Section title="Forex Payment Request Details">
                        <Grid>
                            <Field label="Request Number"><input /></Field>
                            <Field label="Requested On"><input type="date" /></Field>
                            <Field label="Currency"><input /></Field>
                            <Field label="Total Amount"><input type="number" /></Field>
                            <Field label="Foreign Bank Charges"><input type="number" /></Field>
                            <Field label="PO/Contract No"><input /></Field>
                            <Field label="PO Date"><input type="date" /></Field>
                            <Field label="Expected Settlement Date"><input type="date" /></Field>
                        </Grid>

                        <table className="data-table" style={{ marginTop: "10px" }}>
                            <thead>
                                <tr>
                                    <th>Sr No</th>
                                    <th>Performa Invoice No</th>
                                    <th>Performa Invoice Date</th>
                                    <th>Performa Invoice Amount</th>
                                    <th>Attach PO</th>
                                    <th>Attach PI</th>
                                    <th>Attach Other Document</th>
                                    <th>Add/Delete Entry</th>
                                </tr>
                            </thead>
                            <tbody>
                                {rows.map((row, index) => (
                                    <tr key={index}>
                                        <td>{index + 1}</td>

                                        <td>
                                            <input
                                                value={row.invoiceNo}
                                                onChange={(e) =>
                                                    handleChange(index, "invoiceNo", e.target.value)
                                                }
                                            />
                                        </td>

                                        <td>
                                            <input
                                                type="date"
                                                value={row.invoiceDate}
                                                onChange={(e) =>
                                                    handleChange(index, "invoiceDate", e.target.value)
                                                }
                                            />
                                        </td>

                                        <td>
                                            <input
                                                value={row.invoiceAmount}
                                                onChange={(e) =>
                                                    handleChange(index, "invoiceAmount", e.target.value)
                                                }
                                            />
                                        </td>


                                        <td>
                                            <input
                                                type="file"

                                            />
                                        </td>
                                        <td>
                                            <input
                                                type="file"


                                            />
                                        </td>

                                        <td>
                                            <input
                                                type="file"

                                            />
                                        </td>




                                        <td style={{ textAlign: "center" }}>
                                            {/* Show PLUS only on last row */}
                                            {index === rows.length - 1 && (
                                                <button
                                                    type="button"
                                                    onClick={addRow}
                                                    style={{
                                                        background: "#28a745",
                                                        color: "white",
                                                        marginRight: "5px",
                                                        border: "none",
                                                        padding: "5px 10px",
                                                        cursor: "pointer",
                                                        borderRadius: "4px"
                                                    }}
                                                >
                                                    +
                                                </button>
                                            )}

                                            {/* Show DELETE if more than 1 row */}
                                            {rows.length > 1 && (
                                                <button
                                                    type="button"
                                                    onClick={() => deleteRow(index)}
                                                    style={{
                                                        background: "#dc3545",
                                                        color: "white",
                                                        border: "none",
                                                        padding: "5px 10px",
                                                        cursor: "pointer",
                                                        borderRadius: "4px"
                                                    }}
                                                >
                                                    ✖
                                                </button>
                                            )}
                                        </td>

                                    </tr>
                                ))}
                            </tbody>
                            <tfoot>
                                <tr>
                                    <td colSpan={3} style={{ textAlign: "right", fontWeight: "bold" }}>
                                        Total Invoice Amount:
                                    </td>
                                    <td style={{ fontWeight: "bold" }}>
                                        {totalInvoiceAmount.toFixed(2)}
                                    </td>
                                    <td colSpan={4}></td>
                                </tr>
                            </tfoot>
                        </table>
                    </Section>
                )}



                {/* ================= CORRESPONDENT ================= */}
                <Section title="Correspondent Bank Details">
                    <Grid>
                        <Field label="Bank Name"><input /></Field>
                        <Field label="Swift Code"><input /></Field>
                        <Field label="Bank Account No"><input /></Field>
                        <Field label="Remarks" full><textarea rows={3}></textarea></Field>
                    </Grid>
                </Section>

                <div className="button-row">
                    <button className="btn-submit">Submit</button>
                    <button className="btn-exit" onClick={() => history.push("/")}>Exit</button>
                </div>

            </div>
        </div>
    );
};

export default NewRequest;


/* Helper Components */
const Section = ({ title, children }: any) => (
    <div className="form-section">
        <h3>{title}</h3>
        {children}
    </div>
);

const Grid = ({ children }: any) => (
    <div className="form-grid">{children}</div>
);

const Field = ({ label, children, full }: any) => (
    <div className={full ? "form-field full" : "form-field"}>
        <label>{label}</label>
        {children}
    </div>
);
