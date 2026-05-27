import * as React from "react";
import "./Css/VendorReviewForm.scss";
import { IForexModuleProps } from "../IForexModuleProps";

const VendorReviewForm: React.FC<IForexModuleProps> = (props) => {
  return (
    <div className="vendor-review-container">
      <h2 className="title">Vendor Creation - Review Form</h2>

      {/* Vendor Basic Details */}
      <div className="section">
        <h3>Vendor Basic Details</h3>

        <div className="form-grid">
          <div>
            <label>Oracle Vendor Code:</label>
            <span>VE001213</span>
          </div>

          <div>
            <label>Oracle Vendor Name:</label>
            <span>ABC Private Ltd</span>
          </div>

          <div>
            <label>Vendor Type:</label>
            <span>Domestic</span>
          </div>

          <div>
            <label>Country:</label>
            <span>India</span>
          </div>

          <div>
            <label>Currency:</label>
            <span>USD</span>
          </div>

          <div>
            <label>Address Line 1:</label>
            <span>img road</span>
          </div>

          <div>
            <label>Address Line 2:</label>
            <span>Mumbai</span>
          </div>

          <div>
            <label>State:</label>
            <span>Maharashtra</span>
          </div>

          <div>
            <label>Postal Code:</label>
            <span>123456</span>
          </div>
        </div>
      </div>

      {/* Contact Information */}
      <div className="section">
        <h3>Contact Information</h3>

        <div className="form-grid">
          <div>
            <label>Contact Person Name:</label>
            <span>XYZ</span>
          </div>

          <div>
            <label>Email Id:</label>
            <span>xyz@gmail.com</span>
          </div>

          <div>
            <label>Phone Number:</label>
            <span>8976568222</span>
          </div>

          <div>
            <label>Alternate Contact:</label>
            <span>-</span>
          </div>
        </div>
      </div>

      {/* Banking Details */}
      <div className="section">
        <h3>Banking Details</h3>

        <div className="form-grid">
          <div>
            <label>Beneficiary Name:</label>
            <span>XYZ</span>
          </div>

          <div>
            <label>Bank Name:</label>
            <span>HDFC Bank</span>
          </div>

          <div>
            <label>Bank Address:</label>
            <span>ABC</span>
          </div>

          <div>
            <label>Account Number / IBAN:</label>
            <span>1234567890</span>
          </div>

          <div>
            <label>SWIFT / BIC Code:</label>
            <span>HDFC00012</span>
          </div>

          <div>
            <label>Routing Number / ABA:</label>
            <span>785412</span>
          </div>

          <div>
            <label>IFSC Code:</label>
            <span>HDFC0001</span>
          </div>

          <div>
            <label>Intermediary Bank:</label>
            <span>-</span>
          </div>
        </div>
      </div>

      {/* Tax & Regulatory Information */}
      <div className="section">
        <h3>Tax & Regulatory Information</h3>

        <div className="form-grid">
          <div>
            <label>Nature of Payment:</label>
            <span>Service</span>
          </div>

          <div>
            <label>Purpose Code (RBI):</label>
            <span>P0802</span>
          </div>

          <div>
            <label>Tax Residency Certificate (TRC):</label>
            <span>Yes</span>
          </div>

          <div>
            <label>PE Declaration:</label>
            <span>Available</span>
          </div>

          <div>
            <label>Withholding Tax Applicable:</label>
            <span>Yes</span>
          </div>

          <div>
            <label>DTAA Applicable:</label>
            <span>No</span>
          </div>

          <div>
            <label>Country of Tax Residence:</label>
            <span>India</span>
          </div>
        </div>
      </div>

      {/* Compliance Documents */}
      <div className="section">
        <h3>Compliance and Supporting Documents</h3>

        <div className="documents-grid">
          <div className="doc-item">
            <label>Vendor Invoice / Proforma Invoice</label>
            <button className="download-btn">⬇</button>
          </div>

          <div className="doc-item">
            <label>TRC</label>
            <button className="download-btn">⬇</button>
          </div>

          <div className="doc-item">
            <label>Bank Confirmation / Cancelled Cheque</label>
            <button className="download-btn">⬇</button>
          </div>

          <div className="doc-item">
            <label>KYC Documents</label>
            <button className="download-btn">⬇</button>
          </div>

          <div className="doc-item">
            <label>Other Documents</label>
            <button className="download-btn">⬇</button>
          </div>
        </div>

        <div className="remarks-section">
          <label>Remarks</label>
          <textarea rows={3}></textarea>
        </div>
      </div>

      {/* Buttons */}
      <div className="button-group">
        <button className="approve-btn">Approve</button>
        <button className="reject-btn">Reject</button>
        <button className="exit-btn">Exit</button>
      </div>
    </div>
  );
};

export default VendorReviewForm;