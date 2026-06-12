import * as React from "react";
import { IForexModuleProps } from "../IForexModuleProps";
import SPCRUDOPS from "../../service/BAL/spcrud";
import { SPHttpClient } from "@microsoft/sp-http";
import { useHistory } from "react-router-dom";
import logo from "../../assets/sona-comstarlogo.png";

const CreationForm: React.FC<IForexModuleProps> = (props) => {
    const history = useHistory();
    const spCrudOps = SPCRUDOPS();

    const [currencies, setCurrencies] = React.useState<any[]>([]);
    const [countries, setCountries] = React.useState<any[]>([]);
    const [states, setStates] = React.useState<any[]>([]);
    const [cities, setCities] = React.useState<any[]>([]);
    const [natureOfPayments, setNatureOfPayments] = React.useState<any[]>([]);

    const [vendorCode, setVendorCode] = React.useState("");
    const [vendorName, setVendorName] = React.useState("");
    const [vendorShortName, setVendorShortName] = React.useState("");
    const [vendorType, setVendorType] = React.useState("");

    const [currencyId, setCurrencyId] = React.useState<number>();
    const [countryId, setCountryId] = React.useState<number>();
    const [stateId, setStateId] = React.useState<number>();
    const [cityId, setCityId] = React.useState<number>();
    const [natureOfPaymentId, setNatureOfPaymentId] = React.useState<number>();

    const [address, setAddress] = React.useState("");
    const [postalCode, setPostalCode] = React.useState("");
    const [contactPersonName, setContactPersonName] = React.useState("");
    const [emailId, setEmailId] = React.useState("");
    const [phoneNumber, setPhoneNumber] = React.useState("");
    const [alternateContact, setAlternateContact] = React.useState("");

    const [beneficiaryName, setBeneficiaryName] = React.useState("");
    const [bankName, setBankName] = React.useState("");
    const [bankAddress, setBankAddress] = React.useState("");
    const [accountNumberIBAN, setAccountNumberIBAN] = React.useState("");
    const [swiftCode, setSwiftCode] = React.useState("");
    const [routingNumberABA, setRoutingNumberABA] = React.useState("");
    const [ifscCode, setIfscCode] = React.useState("");
    const [intermediaryBank, setIntermediaryBank] = React.useState("");

    const [purposeCodeRBI, setPurposeCodeRBI] = React.useState("");
    const [taxResidencyCertificateTRC, setTaxResidencyCertificateTRC] = React.useState("");
    const [peDeclaration, setPEDeclaration] = React.useState("");
    const [countryOfTaxResidence, setCountryOfTaxResidence] = React.useState("");
    const [dtaaApplicable, setDTAAApplicable] = React.useState("");
    const [vendorNameLegal, setVendorNameLegal] = React.useState("");
    const [withholdingTaxApplicable, setWithholdingTaxApplicable] = React.useState("");
    const [intermediarySwiftCode, setIntermediarySwiftCode] = React.useState("");
    // const [address, setAddress] = React.useState("");
    // const [postalCode, setPostalCode] = React.useState("");
    const [vendorInvoice, setVendorInvoice] = React.useState<File | null>(null);
    const [trcFile, setTrcFile] = React.useState<File | null>(null);
    const [kycFile, setKycFile] = React.useState<File | null>(null);
    const [bankConfirmationFile, setBankConfirmationFile] = React.useState<File | null>(null);
    const [otherDocumentFile, setOtherDocumentFile] = React.useState<File | null>(null);
    const [cityName, setCityName] = React.useState("");
    const [stateName, setStateName] = React.useState("");
    const loadLookups = async () => {
        const sp = await spCrudOps;

        const currencyData = await sp.getData(
            "CurrencyMaster",
            "Id,Title",
            "",
            "",
            { column: "Title", isAscending: true },
            5000,
            props
        );

        const countryData = await sp.getData(
            "CountryMaster",
            "Id,Country",
            "",
            "",
            { column: "Country", isAscending: true },
            5000,
            props
        );

        const stateData = await sp.getData(
            "State",
            "Id,Title",
            "",
            "",
            { column: "Title", isAscending: true },
            5000,
            props
        );

        const cityData = await sp.getData(
            "CityMaster",
            "Id,City",
            "",
            "",
            { column: "City", isAscending: true },
            5000,
            props
        );

        const nopData = await sp.getData(
            "NatureofPaymentMaster",
            "Id,Title",
            "",
            "",
            { column: "Title", isAscending: true },
            5000,
            props
        );

        setCurrencies(currencyData);
        setCountries(countryData);
        setStates(stateData);
        setCities(cityData);
        setNatureOfPayments(nopData);
    };
    let approvalMatrix: any[] = [];
    React.useEffect(() => {
        loadLookups();
        Approvalrequest();
    }, []);

    const Approvalrequest = async () => {
        const sp = await spCrudOps;

        const approvalUsers = await sp.getData(
            "ForexApprovalMAtrix",
            "ID,Approver/ID,Approver/Title",
            "Approver",
            "RequestType eq 'Treasury Approval' and Status eq 'Active'",
            { column: "ID", isAscending: true },
            5000,
            props
        );

        // let approvalMatrix: any[] = [];

        approvalUsers.forEach((x: any, index: number) => {

            approvalMatrix.push({
                Seq: index + 1,
                Role: x.Role?.RoleName,
                Approver: x.Approver?.Title,
                ApproverID: x.Approver?.ID,
                Status: index === 0 ? "Pending" : "Not Started"
            });

        });
    }


    const saveVendorRequest = async () => {
        const sp = await spCrudOps;
        if (!vendorCode?.trim()) {
            alert("Oracle Vendor Code is mandatory.");
            return;
        }

        if (!vendorName?.trim()) {
            alert("Oracle Vendor Name is mandatory.");
            return;
        }

        if (!vendorNameLegal?.trim()) {
            alert("Vendor Name (Legal) is mandatory.");
            return;
        }

        if (!address?.trim()) {
            alert("Address Line 1 is mandatory.");
            return;
        }
      if (!vendorShortName?.trim()) {
    alert("Vendor Short Name is mandatory.");
    return;
}

if (!currencyId) {
    alert("Currency is mandatory.");
    return;
}

if (!countryId) {
    alert("Country is mandatory.");
    return;
}

if (!stateName?.trim()) {
    alert("State is mandatory.");
    return;
}

if (!cityName?.trim()) {
    alert("City is mandatory.");
    return;
}

if (!postalCode?.trim()) {
    alert("Postal Code is mandatory.");
    return;
}

if (!beneficiaryName?.trim()) {
    alert("Beneficiary Name is mandatory.");
    return;
}

if (!bankName?.trim()) {
    alert("Bank Name is mandatory.");
    return;
}

if (!bankAddress?.trim()) {
    alert("Bank Address is mandatory.");
    return;
}

if (!accountNumberIBAN?.trim()) {
    alert("Account Number / IBAN is mandatory.");
    return;
}

if (!swiftCode?.trim()) {
    alert("SWIFT / BIC Code is mandatory.");
    return;
}
  if (!dtaaApplicable) {
            alert("DTAA Applicable is mandatory.");
            return;
        }

        if (!withholdingTaxApplicable) {
            alert("Withholding Tax Applicable is mandatory.");
            return;
        }

        if (!countryOfTaxResidence) {
            alert("Country Of Tax Residence is mandatory.");
            return;
        }

if (!purposeCodeRBI?.trim()) {
    alert("Purpose Code (FIR) is mandatory.");
    return;
}

if (!taxResidencyCertificateTRC?.trim()) {
    alert("Tax Residency Certificate is mandatory.");
    return;
}

if (!peDeclaration?.trim()) {
    alert("PE Declaration is mandatory.");
    return;
}
// if (!vendorInvoice) {
//     alert("Vendor Invoice / Proforma Invoice is mandatory.");
//     return;
// }

// if (!trcFile) {
//     alert("TRC document is mandatory.");
//     return;
// }

// if (!kycFile) {
//     alert("KYC document is mandatory.");
//     return;
// }

// if (!bankConfirmationFile) {
//     alert("Bank Confirmation / Cancelled Cheque is mandatory.");
//     return;
// }

        if (
            !validateFile(vendorInvoice) ||
            !validateFile(trcFile) ||
            !validateFile(kycFile) ||
            !validateFile(bankConfirmationFile) ||
            !validateFile(otherDocumentFile)
        ) {
            return;
        }

        const existingVendor = await sp.getData(
            "VendorMaster",
            "Id,VendorCode,VendorName",
            "",
            `VendorCode eq '${vendorCode.trim()}' or VendorName eq '${vendorName.trim()}'`,
            { column: "Id", isAscending: false },
            5000,
            props
        );

        if (existingVendor.length > 0) {

            const duplicateCode = existingVendor.find(
                (x: any) =>
                    x.VendorCode?.toLowerCase() === vendorCode.trim().toLowerCase()
            );

            const duplicateName = existingVendor.find(
                (x: any) =>
                    x.VendorName?.toLowerCase() === vendorName.trim().toLowerCase()
            );

            if (duplicateCode) {
                alert("Oracle Vendor Code already exists.");
                return;
            }

            if (duplicateName) {
                alert("Vendor Name already exists.");
                return;
            }
        }

        await Approvalrequest();

        const firstApprover = approvalMatrix.length > 0 ? approvalMatrix[0] : null;
        const payload = {
            Title: vendorName,

            VendorCode: vendorCode,
            VendorName: vendorName,
            VendorShortName: vendorShortName,
            VendorType: vendorType,

            VendorAddress: address,
            Pincode: postalCode,

            CurrencyId: currencyId,
            CountryId: countryId,
            // StateId: stateId,
            // CityId: cityId,
            city0: cityName,
            state0: stateName,

            ContactPersonName: contactPersonName,
            EmailId: emailId,
            PhoneNumber: phoneNumber,
            AlternateContact: alternateContact,

            BeneficiaryName: beneficiaryName,
            BankName: bankName,
            BankAddress: bankAddress,
            AccountNumberIBAN: accountNumberIBAN,
            SWIFTBICCode: swiftCode,
            RoutingNumberABA: routingNumberABA,
            IFSCCode: ifscCode,
            IntermediaryBank: intermediaryBank,

            NatureOfPaymentId: natureOfPaymentId,
            PurposeCodeRBI: purposeCodeRBI,

            TaxResidencyCertificateTRC: taxResidencyCertificateTRC,
            PEDeclaration: peDeclaration,
            CountryOfTaxResidence: countryOfTaxResidence,
            DTAAApplicable: dtaaApplicable,
            RequestStatus: "Pending",
            Status: "Inactive",
            CurrentApproverId: firstApprover?.ApproverID || null,
            VendorNameLegal: vendorNameLegal,
            TaxDocumentAvailable: "No",
            withholdingTaxApplicable:withholdingTaxApplicable,
            ApprovedByIDTChecker:"No"
        };

        const result = await sp.insertData(
            "VendorMaster",
            payload,
            props
        );
        const itemId = result.data.Id;


        if (vendorInvoice) {
            await uploadAttachment(
                "VendorMaster",
                itemId,
                vendorInvoice,
                "VendorInvoice",
                props
            );
        }

        if (trcFile) {
            await uploadAttachment(
                "VendorMaster",
                itemId,
                trcFile,
                "TRC",
                props
            );
        }

        if (kycFile) {
            await uploadAttachment(
                "VendorMaster",
                itemId,
                kycFile,
                "KYC",
                props
            );
        }

        if (bankConfirmationFile) {
            await uploadAttachment(
                "VendorMaster",
                itemId,
                bankConfirmationFile,
                "BankConfirmation",
                props
            );
        }

        if (otherDocumentFile) {
            await uploadAttachment(
                "VendorMaster",
                itemId,
                otherDocumentFile,
                "OtherDocument",
                props
            );
        }
        alert("Saved Successfully");
        history.push("/VendorCreationDashboard");

    };

    const uploadAttachment = async (
        listName: string,
        itemId: number,
        file: File,
        documentType: string,
        props: any
    ) => {

        const webUrl =
            props.context.pageContext.web.absoluteUrl;

        const fileName =
            `${documentType}_${Date.now()}_${file.name}`;

        await props.context.spHttpClient.post(
            `${webUrl}/_api/web/lists/getbytitle('${listName}')/items(${itemId})/AttachmentFiles/add(FileName='${fileName}')`,
            SPHttpClient.configurations.v1,
            {
                body: file
            }
        );
    };
    const validateFile = (file: File | null): boolean => {

        if (!file) return true;

        const allowedExtensions = [
            ".pdf",
            ".doc",
            ".docx",
            ".xls",
            ".xlsx",
            ".jpg",
            ".jpeg",
            ".png"
        ];

        const fileName = file.name.toLowerCase();

        const isValid = allowedExtensions.some(
            ext => fileName.endsWith(ext)
        );

        if (!isValid) {

            alert(
                "Only PDF, Word, Excel, JPG, JPEG and PNG files are allowed."
            );

            return false;
        }
        const maxSize = 20 * 1024 * 1024;

        if (file.size > maxSize) {

            alert(
                `File '${file.name}' exceeds the maximum size limit of 20 MB.`
            );

            return false;
        }

        return true;
    };
    return (
        <>

            <div className='MainUplodForm' style={{ margin: "5px 0px" }}>
                <div className='row'>
                    <div className='col-md-12'>
                        <div className='Main-Boxpoup'>
                            <div className="bordered">
                                <a><img src={logo} /></a>
                                <h1>Vendor Creation - Request Form</h1>
                            </div>
                            <div className='borderedbox'>
                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Vendor Basic Details</label>
                                </div>
                                <div className='main-formcontainer'>
                                    <div className="row mb-20">
                                        <div className="col-md-3">
                                            <label className='font'>Oracle Vendor Code  <span style={{ color: "red" }}>*</span></label>
                                            <input
                                                className="form-control"
                                                value={vendorCode}
                                                onChange={(e) => setVendorCode(e.target.value)}
                                            />
                                        </div>

                                        <div className="col-md-3">
                                            <label className='font'>Oracle Vendor Name  <span style={{ color: "red" }}>*</span></label>
                                            <input
                                                className="form-control"
                                                value={vendorName}
                                                onChange={(e) => setVendorName(e.target.value)}
                                            />
                                        </div>

                                        <div className="col-md-3">
                                            <label className='font'>Vendor Name (Legal)  <span style={{ color: "red" }}>*</span></label>
                                            <input
                                                className="form-control"
                                                value={vendorNameLegal}
                                                onChange={(e) => setVendorNameLegal(e.target.value)}
                                            />
                                        </div>

                                        <div className="col-md-3">
                                            <label className='font'>Vendor Short Name  <span style={{ color: "red" }}>*</span></label>
                                            <input
                                                className="form-control"
                                                value={vendorShortName}
                                                onChange={(e) => setVendorShortName(e.target.value)}
                                            />
                                        </div>
                                    </div>

                                    <div className="row mb-20">
                                        <div className="col-md-3">
                                            <label className='font'>Currency  <span style={{ color: "red" }}>*</span></label>
                                            <select
                                                className="form-controltext"
                                                value={currencyId}
                                                onChange={(e) => setCurrencyId(Number(e.target.value))}
                                            >
                                                <option value="">Select Currency</option>

                                                {currencies.map((x) => (
                                                    <option key={x.Id} value={x.Id}>
                                                        {x.Title}
                                                    </option>
                                                ))}
                                            </select>
                                        </div>
                                        <div className="col-md-3">
                                            <label className='font'>Country  <span style={{ color: "red" }}>*</span></label>
                                            <select
                                                className="form-controltext"
                                                value={countryId}
                                                onChange={(e) => setCountryId(Number(e.target.value))}
                                            >
                                                <option value="">Select Country</option>

                                                {countries.map((x) => (
                                                    <option key={x.Id} value={x.Id}>
                                                        {x.Country}
                                                    </option>
                                                ))}
                                            </select>
                                        </div>
                                        {/* 
                                        

                                        <div className="col-md-2">
                                            <label className='font'>City</label>
                                            <select
                                                className="form-controltext"
                                                value={cityId}
                                                onChange={(e) => setCityId(Number(e.target.value))}
                                            >
                                                <option value="">Select City</option>

                                                {cities.map((x) => (
                                                    <option key={x.Id} value={x.Id}>
                                                        {x.City}
                                                    </option>
                                                ))}
                                            </select>
                                        </div> */}
                                        <div className="col-md-3">
                                            <label className='font'>State  <span style={{ color: "red" }}>*</span></label>
                                            {/* <select
                                                className="form-controltext"
                                                value={stateId}
                                                onChange={(e) => setStateId(Number(e.target.value))}
                                            >
                                                <option value="">Select State</option>

                                                {states.map((x) => (
                                                    <option key={x.Id} value={x.Id}>
                                                        {x.Title}
                                                    </option>
                                                ))}
                                                    
                                            </select> */}
                                            <input
                                                type="text"
                                                className="form-controltext"
                                                value={stateName}
                                                onChange={(e) => setStateName(e.target.value)}
                                                placeholder="Enter State"
                                            />
                                        </div>
                                        <div className="col-md-3">
                                            <label className="font">City  <span style={{ color: "red" }}>*</span></label>

                                            <input
                                                type="text"
                                                className="form-controltext"
                                                value={cityName}
                                                onChange={(e) => setCityName(e.target.value)}
                                                placeholder="Enter City"
                                            />
                                        </div>


                                    </div>

                                    <div className="row mb-20">


                                        <div className="col-md-4">
                                            <label className='font'>Postal Code  <span style={{ color: "red" }}>*</span></label>
                                            <input
                                                type="number"
                                                className="form-control"
                                                value={postalCode}
                                                onChange={(e) => setPostalCode(e.target.value)}
                                            />
                                        </div>
                                        {/* <div className="col-md-4">
                                            <label className='font'>Vendor Type</label>
                                            <input
                                                className="form-control"
                                                value={vendorType}
                                                onChange={(e) => setVendorType(e.target.value)}
                                            />
                                        </div> */}
                                        <div className="col-md-5">
                                            <label className='font'>Address Line 1  <span style={{ color: "red" }}>*</span></label>
                                            <textarea
                                                className="form-control"
                                                rows={2}
                                                value={address}
                                                onChange={(e) => setAddress(e.target.value)}
                                            />
                                        </div>
                                    </div>
                                </div>
                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Contact Information</label>
                                </div>
                                <div className='main-formcontainer'>
                                    <div className="row mb-20">
                                        <div className="col-md-3">
                                            <label className='font'>Contact Person Name</label>
                                            <input
                                                className="form-control"
                                                value={contactPersonName}
                                                onChange={(e) => setContactPersonName(e.target.value)}
                                            />
                                        </div>

                                        <div className="col-md-3">
                                            <label className='font'>Email Id</label>
                                            <input
                                                className="form-control"
                                                value={emailId}
                                                onChange={(e) => setEmailId(e.target.value)}
                                            />
                                        </div>

                                        <div className="col-md-3">
                                            <label className='font'>Phone Number</label>
                                            <input
                                                type="number"
                                                className="form-control"
                                                value={phoneNumber}
                                                onChange={(e) => setPhoneNumber(e.target.value)}
                                            />
                                        </div>

                                        <div className="col-md-3">
                                            <label className='font'>Alternate Contact</label>
                                            <input
                                                type="number"
                                                className="form-control"
                                                value={alternateContact}
                                                onChange={(e) => setAlternateContact(e.target.value)}
                                            />
                                        </div>
                                    </div>
                                </div>
                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Banking Details</label>
                                </div>
                                <div className='main-formcontainer'>
                                    <div className="row mb-20">
                                        <div className="col-md-3">
                                            <label className='font'>Beneficiary Name  <span style={{ color: "red" }}>*</span></label>
                                            <input
                                                className="form-control"
                                                value={beneficiaryName}
                                                onChange={(e) => setBeneficiaryName(e.target.value)}
                                            />
                                        </div>

                                        <div className="col-md-3">
                                            <label className='font'>Bank Name  <span style={{ color: "red" }}>*</span></label>
                                            <input
                                                className="form-control"
                                                value={bankName}
                                                onChange={(e) => setBankName(e.target.value)}
                                            />
                                        </div>

                                        <div className="col-md-3">
                                            <label className='font'>Bank Address  <span style={{ color: "red" }}>*</span></label>
                                            <input
                                                className="form-control"
                                                value={bankAddress}
                                                onChange={(e) => setBankAddress(e.target.value)}
                                            />
                                        </div>

                                        <div className="col-md-3">
                                            <label className='font'>Account Number / IBAN  <span style={{ color: "red" }}>*</span></label>
                                            <input
                                                className="form-control"
                                                value={accountNumberIBAN}
                                                onChange={(e) => setAccountNumberIBAN(e.target.value)}
                                            />
                                        </div>
                                    </div>
                                    <div className="row mb-20">
                                        <div className="col-md-3">
                                            <label className='font'>SWIFT / BIC Code  <span style={{ color: "red" }}>*</span></label>
                                            <input
                                                className="form-control"
                                                value={swiftCode}
                                                onChange={(e) => setSwiftCode(e.target.value)}
                                            />
                                        </div>

                                        <div className="col-md-3">
                                            <label className='font'>Routing Number / ABA</label>
                                            <input
                                                className="form-control"
                                                value={routingNumberABA}
                                                onChange={(e) => setRoutingNumberABA(e.target.value)}
                                            />
                                        </div>

                                        <div className="col-md-3">
                                            <label className='font'>IFSC Code</label>
                                            <input
                                                className="form-control"
                                                value={ifscCode}
                                                onChange={(e) => setIfscCode(e.target.value)}
                                            />
                                        </div>

                                        <div className="col-md-3">
                                            <label className='font'>Intermediary Bank</label>
                                            <input
                                                className="form-control"
                                                value={intermediaryBank}
                                                onChange={(e) => setIntermediaryBank(e.target.value)}
                                            />
                                        </div>
                                    </div>
                                </div>
                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Tax & Regulatory Information</label>
                                </div>
                                <div className='main-formcontainer'>
                                    <div className="row mb-20">
                                        {/* <div className="col-md-3">
                                            <label className='font'>Nature of Payment</label>
                                            <select
                                                className="form-control"
                                                value={natureOfPaymentId}
                                                onChange={(e) => setNatureOfPaymentId(Number(e.target.value))}
                                            >
                                                <option value="">Select Nature Of Payment</option>

                                                {natureOfPayments.map((x) => (
                                                    <option key={x.Id} value={x.Id}>
                                                        {x.Title}
                                                    </option>
                                                ))}
                                            </select>
                                        </div> */}

                                        <div className="col-md-3">
                                            <label className='font'>Purpose Code (FIR)  <span style={{ color: "red" }}>*</span></label>
                                            <input
                                                className="form-control"
                                                value={purposeCodeRBI}
                                                onChange={(e) => setPurposeCodeRBI(e.target.value)}
                                            />
                                        </div>

                                        <div className="col-md-3">
                                            <label className='font'>Tax Residency Certificate  <span style={{ color: "red" }}>*</span></label>
                                            <input
                                                className="form-control"
                                                value={taxResidencyCertificateTRC}
                                                onChange={(e) => setTaxResidencyCertificateTRC(e.target.value)}
                                            />
                                        </div>

                                        <div className="col-md-3">
                                            <label className='font'>PE Declaration  <span style={{ color: "red" }}>*</span> </label>
                                            <input
                                                className="form-control"
                                                value={peDeclaration}
                                                onChange={(e) => setPEDeclaration(e.target.value)}
                                            />
                                        </div>
                                        <div className="col-md-3">
                                            <label className='font'>
                                                Withholding Tax Applicable
                                                <span style={{ color: "red" }}>*</span>
                                            </label>

                                            <select
                                                className="form-control"
                                                value={withholdingTaxApplicable}
                                                onChange={(e) => setWithholdingTaxApplicable(e.target.value)}
                                            >
                                                <option value="">Select</option>
                                                <option value="Yes">Yes</option>
                                                <option value="No">No</option>
                                            </select>
                                        </div>
                                    </div>
                                    <div className="row mb-20">
                                        

                                        <div className="col-md-3">
                                            <label className='font'>
                                                DTAA Applicable <span style={{ color: "red" }}>*</span>
                                            </label>

                                            <select
                                                className="form-control"
                                                value={dtaaApplicable}
                                                onChange={(e) => setDTAAApplicable(e.target.value)}
                                            >
                                                <option value="">Select</option>
                                                <option value="Yes">Yes</option>
                                                <option value="No">No</option>
                                            </select>
                                        </div>

                                        <div className="col-md-3">
                                            <label className='font'>
                                                Country of Tax Residence
                                                <span style={{ color: "red" }}>*</span>
                                            </label>

                                            <select
                                                className="form-control"
                                                value={countryOfTaxResidence}
                                                onChange={(e) => setCountryOfTaxResidence(e.target.value)}
                                            >
                                                <option value="">Select Country</option>

                                                {countries.map((x: any) => (
                                                    <option
                                                        key={x.Id}
                                                        value={x.Country}
                                                    >
                                                        {x.Country}
                                                    </option>
                                                ))}
                                            </select>
                                        </div>
                                    </div>
                                </div>
                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Compliance & Supporting Documents</label>
                                </div>
                                <div className='main-formcontainer'>
                                    <div className="row mb-20">
                                        <div className="col-md-4">
                                            <label className='font'>Vendor Invoice / Proforma Invoice </label>
                                            <input
                                                type="file"
                                                className="form-control"
                                                accept=".pdf,.doc,.docx,.xls,.xlsx,.jpg,.jpeg"
                                                onChange={(e) =>
                                                    setVendorInvoice(
                                                        e.target.files?.length
                                                            ? e.target.files[0]
                                                            : null
                                                    )
                                                }
                                            />
                                        </div>

                                        <div className="col-md-4">
                                            <label className='font'>TRC </label>
                                            <input
                                                type="file"
                                                className="form-control"
                                                accept=".pdf,.doc,.docx,.xls,.xlsx,.jpg,.jpeg"
                                                onChange={(e) =>
                                                    setTrcFile(
                                                        e.target.files?.length
                                                            ? e.target.files[0]
                                                            : null
                                                    )
                                                }
                                            />
                                        </div>

                                        <div className="col-md-4">
                                            <label className='font'>KYC Documents </label>
                                            <input
                                                type="file"
                                                className="form-control"
                                                accept=".pdf,.doc,.docx,.xls,.xlsx,.jpg,.jpeg"
                                                onChange={(e) =>
                                                    setKycFile(
                                                        e.target.files?.length
                                                            ? e.target.files[0]
                                                            : null
                                                    )
                                                }
                                            />
                                        </div>
                                    </div>
                                    <div className="row mb-20">
                                        <div className="col-md-4">
                                            <label className='font'>Bank Confirmation / Cancelled Cheque </label>
                                            <input
                                                type="file"
                                                className="form-control"
                                                accept=".pdf,.doc,.docx,.xls,.xlsx,.jpg,.jpeg"
                                                onChange={(e) =>
                                                    setBankConfirmationFile(
                                                        e.target.files?.length
                                                            ? e.target.files[0]
                                                            : null
                                                    )
                                                }
                                            />
                                        </div>

                                        <div className="col-md-4">
                                            <label className='font'>Other Documents </label>
                                            <input
                                                type="file"
                                                className="form-control"
                                                accept=".pdf,.doc,.docx,.xls,.xlsx,.jpg,.jpeg"
                                                onChange={(e) =>
                                                    setOtherDocumentFile(
                                                        e.target.files?.length
                                                            ? e.target.files[0]
                                                            : null
                                                    )
                                                }
                                            />
                                        </div>
                                    </div>
                                </div>
                                <div style={{ margin: "10px", display: "flex", justifyContent: "center", gap: "5px", alignItems: "center" }}>
                                    <a className="Submit-btn" onClick={saveVendorRequest}> Submit </a>
                                    <a className="Exit-btn" onClick={history.goBack}>Exit</a>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </>
    );
};

export default CreationForm;