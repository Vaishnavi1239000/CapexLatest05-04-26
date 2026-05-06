import * as React from "react";
import "./advanced.scss";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { useEffect, useState } from "react";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/PeoplePicker";
// interface IVendor {
//     VendorCode: string;
//     VendorName: string;
// }
interface IVendor {
  Id: number;
  VendorCode: string;
  VendorName: string;
}

const NewAdvanceform = ({ context }: any) => {
  const sp = spfi().using(SPFx(context));
  const [attachments, setAttachments] = useState<any[]>([]);
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);

  const [employee, setEmployee] = React.useState<any>({});
  //const [selectedUser, setSelectedUser] = useState<any>(null);
  //const [attachments, setAttachments] = useState<File[]>([]);
  const [previousAdvances, setPreviousAdvances] = useState<any[]>([]);

  const [employeeName, setEmployeeName] = React.useState("");
  const [pickerKey, setPickerKey] = React.useState<number>(0);
  const [vendors, setVendors] = useState<IVendor[]>([]);

  const [selectedUser, setSelectedUser] = useState<any[]>([]);
  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  const [selectedVendorName, setSelectedVendorName] = useState("");

  const [poAmount, setPoAmount] = useState("");
  const [advanceAmount, setAdvanceAmount] = useState("");
  const [paidAmount, setPaidAmount] = useState("");
  const [poNumber, setPoNumber] = useState("");
  const [poDate, setPoDate] = useState("");
  const [poTerms, setPoTerms] = useState("");

  const [expectedDate, setExpectedDate] = useState("");

  const [glCode, setGlCode] = useState("390111003");

  const [costCenter, setCostCenter] = useState("");

  const [remarks, setRemarks] = useState("");
  const [projectDesc, setProjectDesc] = useState("");
  const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);
  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: context.pageContext.web.absoluteUrl,
    msGraphClientFactory: context.msGraphClientFactory,
    spHttpClient: context.spHttpClient,
  };
  const handleNumberChange = (value: string, setter: any) => {
    // Allow only numbers and decimal (max one dot)
    const regex = /^\d*\.?\d*$/;

    if (regex.test(value)) {
      void setter(value);
    }
  };
  const getPreviousAdvances = async (vendorId: number) => {
    try {
      debugger;
      console.log("Fetching for Vendor:", vendorId);

      const data = await sp.web.lists
        .getByTitle("CapexAdvance")
        .items.select(
          "PONumber",
          "RequestAdvanceAmount",
          "Created",
          "VoucherDate",

          "PaidAmount",
          "Status",
          "VendorCode/Id",
        )
        .expand("VendorCode")
        .filter(`VendorCode/Id eq ${vendorId} and Status eq 'Paid'`)
        .orderBy("Created", false)();

      console.log("DATA:", data);

      void setPreviousAdvances(data);
    } catch (error) {
      console.error("Error fetching previous advances:", error);
      void setPreviousAdvances([]);
    }
  };
  const handleRemoveFile = (index: number) => {
    const updatedFiles = [...selectedFiles];
    updatedFiles.splice(index, 1);
    setSelectedFiles(updatedFiles);
  };

  const handleExit = () => {
    window.location.href =
      "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexForm.aspx?page=User";
  };

  const getLoggedInUser = async () => {
    try {
      const currentUser = await sp.web.currentUser();
      const email = currentUser.Email;

      const user = await sp.web.lists
        .getByTitle("EmployeeMaster")
        .items.select(
          "EmployeeCode",
          "EmployeeName",
          "Division",
          "Location",
          "EmployeeEmail",
          "ReportingManager/Title",
          "ReportingManager/Id",
          "HOD/Title",
          "HOD/Id",
          "ContactNo",
          "EmployeeStatus",
          "CostCenter",
        )
        .expand("ReportingManager", "HOD")
        .filter(`EmployeeEmail eq '${email}'`)
        .top(1)();

      if (user.length > 0) {
        void setEmployee(user[0]);
      }
    } catch (error) {
      console.log("Error fetching user:", error);
    }
  };
  const getVendors = async () => {
    const data = await sp.web.lists
      .getByTitle("VendorMaster")
      .items.select("Id", "VendorCode", "VendorName")(); // ✅ Id required

    void setVendors(data);
  };

  const getFinancialYear = () => {
    const today = new Date();
    const year = today.getFullYear();
    const month = today.getMonth() + 1;

    if (month >= 4) {
      // April to March
      return `${year.toString().slice(-2)}-${(year + 1).toString().slice(-2)}`;
    } else {
      return `${(year - 1).toString().slice(-2)}-${year.toString().slice(-2)}`;
    }
  };
  const generateCapexId = async () => {
    try {
      const fy = getFinancialYear(); // 25-26

      const items = await sp.web.lists
        .getByTitle("CapexAdvance")
        .items.select("CapexID", "ID")
        .filter(`startswith(CapexID,'CPX/${fy}/')`)
        .orderBy("ID", false)
        .top(1)();

      let nextNumber = 1;

      if (items.length > 0 && items[0].CapexID) {
        const lastId = items[0].CapexID;

        const parts = lastId.split("/");

        if (parts.length === 3) {
          const lastNumber = parseInt(parts[2]);

          if (!isNaN(lastNumber)) {
            nextNumber = lastNumber + 1;
          }
        }
      }

      const formattedNumber = nextNumber.toString().padStart(5, "0");

      return `CPX/${fy}/${formattedNumber}`;
    } catch (error) {
      alert(error);
      console.error("Capex ID Error:", error);
      return `CPX/${getFinancialYear()}/00001`;
    }
  };
  const uploadAttachments = async (capexId: string) => {
    try {
      if (!selectedFiles || selectedFiles.length === 0) return;

      const safeCapexId = capexId.replace(/\//g, "_");

      const libraryName = "CapexAdvanceDocs";
      const webUrl = context.pageContext.web.serverRelativeUrl;

      const folderPath = `${webUrl}/${libraryName}/${safeCapexId}`;

      // ✅ Ensure folder
      await sp.web.folders.addUsingPath(`${libraryName}/${safeCapexId}`);

      // ✅ Upload files properly
      for (const file of selectedFiles) {
        await sp.web
          .getFolderByServerRelativePath(folderPath)
          .files.addUsingPath(file.name, file, { Overwrite: true });
      }

      console.log("✅ Files uploaded successfully");
    } catch (error) {
      console.error("❌ Upload error:", error);
    }
  };
  const buildApprovalFlow = async () => {
    try {
      const flow: any[] = [];

      // =========================
      // 🔹 RM
      // =========================
      if (employee.ReportingManager?.Id) {
        flow.push({
          Id: employee.ReportingManager.Id,
          Name: employee.ReportingManager?.Title,
          Role: "RM",
          Level: 1,
          Status: "Pending",
        });
      }

      // =========================
      // 🔹 HOD
      // =========================
      if (employee.HOD?.Id) {
        flow.push({
          Id: employee.HOD?.Id,
          Name: employee.HOD?.Title,
          Role: "HOD",
          Level: 2,
          Status: "Pending",
        });
      }

      // =========================
      // 🔥 MATRIX APPROVERS
      // =========================
      const matrixData = await sp.web.lists
        .getByTitle("CapexApprovalMatrix")
        .items.select("Role/RoleName,Level/Level,Approver/Id,Approver/Title")
        .expand("Approver,Role,Level")
        .filter("Status eq 'Active'")
        .orderBy("Level", true)();

      const matrixApprovers = matrixData.map((item: any, index: number) => ({
        Id: item.Approver?.Id,
        Name: item.Approver?.Title,
        Role: item.Role.RoleName,
        Level: flow.length + index + 1,
        Status: "Pending",
      }));

      // =========================
      // 🔹 MERGE FLOW
      // =========================
      const fullFlow = [...flow, ...matrixApprovers];

      // =========================
      // 🔹 REMOVE DUPLICATES
      // =========================
      const uniqueFlow = fullFlow.filter(
        (v, i, self) => self.findIndex((x) => x.Id === v.Id) === i,
      );

      return uniqueFlow;
    } catch (error) {
      console.error("Approval Flow Error:", error);
      return [];
    }
  };

  useEffect(() => {
    const loadFlow = async () => {
      if (employee && employee.ReportingManager) {
        try {
          const flow = await buildApprovalFlow();
          setApprovalMatrix(flow);
        } catch (error) {
          console.error(error);
        }
      }
    };

    void loadFlow();
  }, [employee]);

  // useEffect(() => {
  //   if (employee && employee.ReportingManager) {
  //     buildApprovalFlow().then(setApprovalMatrix);
  //   }
  // }, [employee]);

  const validateForm = () => {
    const errors: string[] = [];

    if (!selectedVendorId) {
      errors.push("Please select the Vendor code");
    }

    if (!poNumber) {
      errors.push("Please update PO Number");
    }

    if (!poDate) {
      errors.push("Please update PO date");
    }

    if (!poTerms) {
      errors.push("Please update PO Terms");
    }

    if (!poAmount) {
      errors.push("Please update PO Amount");
    }

    if (!advanceAmount || Number(advanceAmount) <= 0) {
      errors.push("Please update Advance Amount");
    }

    if (!paidAmount || Number(paidAmount) <= 0) {
      errors.push("Please update Paid Amount");
    }

    // 🔥 NEW VALIDATION
    if (
      advanceAmount &&
      paidAmount &&
      Number(paidAmount) > Number(advanceAmount)
    ) {
      errors.push("Paid Amount cannot be greater than Advance Amount");
    }

    if (!expectedDate) {
      errors.push("Please update Settlement Date");
    }

    if (!selectedUser || selectedUser.length === 0) {
      errors.push("Please select PIC Name");
    }

    if (!projectDesc) {
      errors.push("Please enter Project Description");
    }

    if (!selectedFiles || selectedFiles.length === 0) {
      errors.push("Please upload at least one attachment");
    }

    return errors;
  };

  const handleSubmit = async () => {
    try {
      debugger;
      const errors = validateForm();

      if (errors.length > 0) {
        alert(errors.join("\n")); // 👈 shows exactly like your screenshot
        return;
      }

      const capexId = await generateCapexId();

      // ✅ Validate Vendor

      // ✅ Get Email from PeoplePicker
      const userEmail = selectedUser[0]?.secondaryText;

      if (!userEmail) {
        alert("User email not found");
        return;
      }

      // ✅ Ensure User (FIX ERROR)
      const ensuredUser = await sp.web.ensureUser(userEmail);

      const flow = await buildApprovalFlow();

      // 🔥 Set first approver as current
      if (flow.length > 0) {
        flow[0].Status = "In Progress";
      }

      const currentApprover = flow.length > 0 ? flow[0].Id : null;
      const currentUser = context.pageContext?.user?.displayName || "User";
      const wfHistory = [
        {
          CurrentApprover: currentUser,
          ActionTaken: "Submitted",
          Comment: remarks || "",
          Date: new Date().toISOString(),
        },
      ];

      await sp.web.lists.getByTitle("CapexAdvance").items.add({
        Title: capexId,
        CapexID: capexId,

        // Employee
        EmployeeCode: employee.EmployeeCode,
        EmployeeName: employee.EmployeeName,
        Division: employee.Division,
        Location: employee.Location,
        Email: employee.EmployeeEmail,
        RM: employee.ReportingManager?.Title,
        HOD: employee.HOD?.Title,
        ContactNo: employee.ContactNo,
        EmployeeStatus: employee.EmployeeStatus,

        // Vendor (LOOKUP)
        VendorCodeId: selectedVendorId, // ✅ FIX
        VendorName: selectedVendorName,

        // PO
        PONumber: poNumber,
        PODate: poDate ? new Date(poDate) : null,
        POAdvanceTerms: poTerms,

        // Amount
        POAmtGST: poAmount,
        RequestAdvanceAmount: advanceAmount,
        PaidAmount: paidAmount,

        // Advance
        ExpectedDateofSettlement: expectedDate ? new Date(expectedDate) : null,

        // Person field
        PICNameId: ensuredUser.Id, // ✅ FIX

        // Other
        GL: glCode,
        CostCenter: employee.CostCenter,
        Remarks: remarks,
        ProjectDescription: projectDesc,

        Status: "Pending for Approver",

        ApprovalMatrix: JSON.stringify(flow),

        CurrentApproverId: currentApprover,

        WorkFlowHistory: JSON.stringify(wfHistory),
      });
      debugger;
      await uploadAttachments(capexId); // 🔥 FIXED

      console.log("Attachments:", selectedFiles);
      alert("Submitted successfully ✅");

      // 🔥 REDIRECT
      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexForm.aspx?page=User";
    } catch (error) {
      console.error("ERROR:", error);
      alert("Error while saving ❌");
    }
  };

  const handledraft = async () => {
    try {
      const capexId = await generateCapexId();

      let ensuredUserId: number | null = null;

      // ✅ Only process if user selected
      if (selectedUser && selectedUser.length > 0) {
        const userEmail = selectedUser[0]?.secondaryText;

        if (userEmail) {
          const ensuredUser = await sp.web.ensureUser(userEmail);
          ensuredUserId = ensuredUser.Id;
        }
      }
      const userEmail = selectedUser[0]?.secondaryText;

      if (!userEmail) {
        alert("User email not found");
        return;
      }

      // ✅ Ensure User (FIX ERROR)
      const ensuredUser = await sp.web.ensureUser(userEmail);

      const flow = await buildApprovalFlow();

      // 🔥 Set first approver as current
      if (flow.length > 0) {
        flow[0].Status = "In Progress";
      }

      const currentApprover = flow.length > 0 ? flow[0].Id : null;
      const currentUser = context.pageContext?.user?.displayName || "User";
      const wfHistory = [
        {
          CurrentApprover: currentUser,
          ActionTaken: "Submitted",
          Comment: remarks || "",
          Date: new Date().toISOString(),
        },
      ];

      await sp.web.lists.getByTitle("CapexAdvance").items.add({
        Title: capexId,
        CapexID: capexId,

        // Employee
        EmployeeCode: employee.EmployeeCode,
        EmployeeName: employee.EmployeeName,
        Division: employee.Division,
        Location: employee.Location,
        Email: employee.EmployeeEmail,
        RM: employee.ReportingManager?.Title,
        HOD: employee.HOD?.Title,
        ContactNo: employee.ContactNo,
        EmployeeStatus: employee.EmployeeStatus,

        // Vendor
        VendorCodeId: selectedVendorId,
        VendorName: selectedVendorName,

        // PO
        PONumber: poNumber,
        PODate: poDate ? new Date(poDate) : null,
        POAdvanceTerms: poTerms,

        // Amount
        POAmtGST: poAmount,
        RequestAdvanceAmount: advanceAmount,
        PaidAmount: paidAmount,

        // Advance
        ExpectedDateofSettlement: expectedDate ? new Date(expectedDate) : null,

        // ✅ PIC (OPTIONAL)
        ...(ensuredUserId && { PICNameId: ensuredUserId }),

        // Other
        GL: glCode,
        CostCenter: employee.CostCenter,
        Remarks: remarks,
        ProjectDescription: projectDesc,

        Status: "Draft",

        ApprovalMatrix: JSON.stringify(flow),

        CurrentApproverId: currentApprover,

        WorkFlowHistory: JSON.stringify(wfHistory),
      });

      const safeCapexId = capexId.replace(/\//g, "_");
      void uploadAttachments(safeCapexId);

      alert("Draft saved successfully ✅");

      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexForm.aspx?page=User";
    } catch (error) {
      console.error("ERROR:", error);
      alert("Error while saving ❌");
    }
  };

  React.useEffect(() => {
    if (!context) return;

    void getLoggedInUser();
    void getVendors(); // 👈 ADD THIS
  }, [context]);

  return (
    <div className="form-container">
      <h2 className="form-title">(New advanceed)Advance Payment</h2>
      <div className="approval-flow">
        {approvalMatrix.map((a, index) => (
          <div key={index} className="approval-step">
            <div>
              <b>{a.Role}</b>
            </div>
            <div>{a.Name}</div>

            {/* 🔥 STATUS */}
            <div
              style={{
                color:
                  a.Status === "Approved"
                    ? "green"
                    : a.Status === "Rejected"
                      ? "red"
                      : a.Status === "In Progress"
                        ? "orange"
                        : "gray",
              }}
            >
              {a.Status}
            </div>
          </div>
        ))}
      </div>
      {/* Employee Info */}
      <div className="section">
        <h3>Requestor Information</h3>

        <div className="form-row">
          <label>Employee Code</label>
          <input value={employee.EmployeeCode || ""} readOnly />

          <label>Employee Name</label>
          <input value={employee.EmployeeName || ""} readOnly />

          <label>Division</label>
          <input value={employee.Division || ""} readOnly />

          <label>Location</label>
          <input value={employee.Location || ""} readOnly />
        </div>

        <div className="form-row">
          <label>Email</label>
          <input value={employee.EmployeeEmail || ""} readOnly />

          <label>RM</label>
          <input value={employee.ReportingManager?.Title || ""} readOnly />

          <label>HOD</label>
          <input value={employee.HOD?.Title || ""} readOnly />

          <label>Contact No</label>
          <input value={employee.ContactNo || ""} readOnly />
        </div>

        <div className="form-row">
          <label>Employee Status</label>
          <input value={employee.EmployeeStatus || ""} readOnly />
        </div>
      </div>

      {/* Vendor Section */}
      <div className="section">
        <h3>Vendor & PO Details</h3>

        <div className="form-row">
          <label>Vendor Code</label>
          <select
            value={selectedVendorId || ""}
            onChange={(e) => {
              const id = Number(e.target.value);
              const vendor = vendors.find((v) => v.Id === id);

              setSelectedVendorId(id);
              setSelectedVendorName(vendor?.VendorName || "");

              // 🔥 FETCH DATA
              if (id) {
                void getPreviousAdvances(id);
              }
            }}
          >
            <option value="">Select Vendor</option>
            {vendors.map(
              (
                v, // ✅ FIX
              ) => (
                <option key={v.Id} value={v.Id}>
                  {v.VendorCode}
                </option>
              ),
            )}
          </select>

          <br></br>
          <label>Vendor Name</label>
          <input value={selectedVendorName} readOnly />

          <label>PO Number</label>
          <input
            value={poNumber}
            onChange={(e) => setPoNumber(e.target.value)}
          />

          <label>PO Date</label>
          <input
            type="date"
            value={poDate}
            onChange={(e) => setPoDate(e.target.value)}
          />
        </div>

        <div className="form-row">
          <label>PO Advance Terms</label>
          <input value={poTerms} onChange={(e) => setPoTerms(e.target.value)} />

          <label>PO Amount (GST)</label>
          <input
            value={poAmount}
            onChange={(e) => handleNumberChange(e.target.value, setPoAmount)}
          />

          <label>Request Advance Amount</label>
          <input
            value={advanceAmount}
            onChange={(e) =>
              handleNumberChange(e.target.value, setAdvanceAmount)
            }
          />

          <label style={{ color: "red" }}>Paid Amount</label>
          <input
            value={paidAmount}
            onChange={(e) => handleNumberChange(e.target.value, setPaidAmount)}
          />
        </div>
      </div>

      {/* Advance Section */}
      <div className="section">
        <h3>Advance Details</h3>

        <div className="form-row">
          <label>Expected Settlement Date</label>
          <input
            type="date"
            value={expectedDate}
            onChange={(e) => setExpectedDate(e.target.value)}
          />

          <label>PIC Name</label>
          <PeoplePicker
            context={peoplePickerContext}
            personSelectionLimit={1}
            ensureUser={true}
            principalTypes={[PrincipalType.User]}
            onChange={(items) => setSelectedUser(items)}
          />

          <label>GL Code</label>
          <input value={glCode} readOnly />

          <label>Cost Center</label>

          <input value={employee.CostCenter || ""} readOnly />
        </div>
      </div>

      {/* Remarks */}
      <div className="section">
        <h3>Remarks</h3>

        <div className="form-row">
          <label>Remarks</label>
          <textarea
            value={remarks}
            onChange={(e) => setRemarks(e.target.value)}
          />

          <label>Project Description</label>
          <textarea
            value={projectDesc}
            onChange={(e) => setProjectDesc(e.target.value)}
          />
        </div>
      </div>

      {/* Table */}
      <div className="section">
        <h3>Previous Advances</h3>

        <table className="data-table">
          <thead>
            <tr>
              <th>PO Number</th>
              <th>Previous Advance</th>
              <th>Requested Date</th>
              <th>Paid Date</th>
              <th>MRN No</th>
              <th>Settled Amount</th>
              <th>Pending Advance</th>
            </tr>
          </thead>
          <tbody>
            {previousAdvances.length === 0 ? (
              <tr>
                <td colSpan={7} style={{ textAlign: "center" }}>
                  No Data
                </td>
              </tr>
            ) : (
              previousAdvances.map((item: any, index: number) => {
                const pending = Math.max(
                  0,
                  Number(item.RequestAdvanceAmount || 0) -
                    Number(item.PaidAmount || 0),
                );

                return (
                  <tr key={index}>
                    <td>{item.PONumber}</td>
                    <td>{item.RequestAdvanceAmount}</td>

                    <td>
                      {item.Created
                        ? new Date(item.Created).toLocaleDateString()
                        : ""}
                    </td>

                    <td>
                      {item.VoucherDate
                        ? new Date(item.VoucherDate).toLocaleDateString()
                        : ""}
                    </td>

                    <td>{item.VoucherNumber}</td>
                    <td>{item.PaidAmount}</td>
                    <td>{pending}</td>
                  </tr>
                );
              })
            )}
          </tbody>
        </table>
      </div>
      <div className="col-md-4">
        <label className="font">Attach</label>

        <input
          type="file"
          multiple
          onChange={(e) => {
            const files = e.target.files;
            if (!files) return;
            setSelectedFiles((prev) => [...prev, ...Array.from(files)]);
          }}
        />

        {/* 🔥 Show selected files */}
        {selectedFiles.length > 0 && (
          <ul style={{ marginTop: "10px" }}>
            {selectedFiles.map((file, index) => (
              <li key={index}>
                {file.name}

                {/* 🔥 Remove button */}
                <button
                  type="button"
                  style={{
                    marginLeft: "10px",
                    color: "red",
                    cursor: "pointer",
                  }}
                  onClick={() => handleRemoveFile(index)}
                >
                  Remove
                </button>
              </li>
            ))}
          </ul>
        )}
      </div>
      {/* Attachment
      <input
        type="file"
        multiple
        onChange={(e) => {
          if (e.target.files) {
            setAttachments(Array.from(e.target.files));
          }
        }}
      /> */}

      {/* Buttons */}

      <div className="actions">
        <button className="primary" onClick={handleSubmit}>
          Submit
        </button>

        <button className="secondary" onClick={handledraft}>
          Save Draft
        </button>

        <button className="exit" onClick={handleExit}>
          Exit
        </button>
      </div>
      {/* <div className="actions">
        <button className="primary" onClick={handleSubmit}>
          Submit
        </button>

        <button className="secondary" onClick={handledraft}>
          Save Draft
        </button>
        <button className="exit" onClick={handleExit}>
          Exit
        </button>
      </div> */}
    </div>
  );
};

export default NewAdvanceform;
