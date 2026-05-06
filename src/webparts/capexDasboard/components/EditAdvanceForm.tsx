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

interface IVendor {
  Id: number;
  VendorCode: string;
  VendorName: string;
}

const EditAdvanceForm = ({ context, formData, onClose }: any) => {
  const sp = spfi().using(SPFx(context));

  // =========================
  // STATES
  // =========================
  const [vendors, setVendors] = useState<IVendor[]>([]);
  const [employee, setEmployee] = useState<any>({});
  const [attachments, setAttachments] = useState<any[]>([]);
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [selectedUser, setSelectedUser] = useState<any[]>([]);

  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  const [selectedVendorName, setSelectedVendorName] = useState("");

  const [poNumber, setPoNumber] = useState("");
  const [poDate, setPoDate] = useState("");
  const [poTerms, setPoTerms] = useState("");
  const [poAmount, setPoAmount] = useState("");
  const [advanceAmount, setAdvanceAmount] = useState("");
  const [paidAmount, setPaidAmount] = useState("");
  const [expectedDate, setExpectedDate] = useState("");
  const [vendorName, setVendorName] = useState("");
  const [glCode, setGlCode] = useState("390111003");
  const [costCenter, setCostCenter] = useState("");
  const [remarks, setRemarks] = useState("");
  const [projectDesc, setProjectDesc] = useState("");

  const [approverRemarks, setApproverRemarks] = useState("");
  const [voucherDate, setVoucherDate] = useState("");
  const [voucherNumber, setVoucherNumber] = useState("");
  const [UTRDate, setUTRDate] = useState("");
  const [UTRNumber, setUTRNumber] = useState("");
  const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);
  const [workflowHistory, setWorkflowHistory] = useState<any[]>([]);
  const [previousAdvances, setPreviousAdvances] = useState<any[]>([]);
  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: context.pageContext.web.absoluteUrl,
    msGraphClientFactory: context.msGraphClientFactory,
    spHttpClient: context.spHttpClient,
  };

  // =========================
  // LOAD DATA
  // =========================

  const handleDeleteExistingFile = async (file: any) => {
    try {
      if (!window.confirm(`Delete ${file.Name}?`)) return;

      await sp.web.getFileByServerRelativePath(file.ServerRelativeUrl).delete();

      // update UI
      setAttachments((prev) =>
        prev.filter((f) => f.ServerRelativeUrl !== file.ServerRelativeUrl),
      );

      alert("File deleted ✅");
    } catch (error) {
      console.error("Delete error:", error);
    }
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
  const getVendors = async () => {
    const data = await sp.web.lists
      .getByTitle("VendorMaster")
      .items.select("Id", "VendorCode", "VendorName")();
    void setVendors(data);
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
          "HOD/Title",
          "ContactNo",
          "EmployeeStatus",
          "CostCenter",
        )
        .expand("ReportingManager", "HOD")
        .filter(`EmployeeEmail eq '${email}'`)
        .top(1)();

      if (user.length > 0) {
        setEmployee(user[0]);
      }
    } catch (error) {
      console.log("Error fetching user:", error);
    }
  };

  const getAttachments = async (capexId: string) => {
    try {
      const safe = capexId.replace(/\//g, "_");
      const path = `/sites/SonaFinance/CapexAdvanceDocs/${safe}`;

      const files = await sp.web
        .getFolderByServerRelativePath(path)
        .files();

      void setAttachments(files);
    } catch {
      void setAttachments([]);
    }
  };

  // =========================
  // UPLOAD FILES
  // =========================

  // =========================
  // VALIDATION
  // =========================
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

    if (
      (!attachments || attachments.length === 0) &&
      (!selectedFiles || selectedFiles.length === 0)
    ) {
      errors.push("Please upload at least one attachment");
    }

    return errors;
  };




  const handleExit = () => {
    if (onClose) {
      onClose();
    } else {
      window.location.reload();
    }
  };
  const ensureFolder = async (folderPath: string) => {
    try {
      await sp.web.getFolderByServerRelativePath(folderPath)();
    } catch {
      // create folder if not exists
      const parentPath = folderPath.substring(0, folderPath.lastIndexOf("/"));
      const folderName = folderPath.substring(folderPath.lastIndexOf("/") + 1);

      await sp.web
        .getFolderByServerRelativePath(parentPath)
        .folders.addUsingPath(folderName);
    }
  };
  const uploadFiles = async () => {
    try {
      if (!formData?.CapexID || selectedFiles.length === 0) return;

      const safe = formData.CapexID.replace(/\//g, "_");
      const folderPath = `/sites/SonaFinance/CapexAdvanceDocs/${safe}`;

      // ✅ Ensure folder exists
      await ensureFolder(folderPath);

      for (const file of selectedFiles) {
        await sp.web
          .getFolderByServerRelativePath(folderPath)
          .files.addUsingPath(file.name, file, { Overwrite: true });
      }

      setSelectedFiles([]);
      await getAttachments(formData.CapexID);

    } catch (error) {
      console.error("Upload error:", error);
      alert("File upload failed ❌");
    }
  };


  // =========================
  // UPDATE
  // =========================
  const handleSubmit = async () => {
    try {
      const errors = validateForm();
      if (errors.length > 0) {
        alert(errors.join("\n"));
        return;
      }

      const ensuredUser = await sp.web.ensureUser(
        selectedUser[0]?.secondaryText
      );
      // 🔥 PRESERVE FLOW
      const existingFlow = formData.ApprovalMatrix
        ? JSON.parse(formData.ApprovalMatrix)
        : [];
const updatedFlow = existingFlow.map((a: any, index: number) => ({
  ...a,
  Status: index === 0 ? "In Progress" : "Pending"
}));
const currentApprover = updatedFlow.length > 0 ? updatedFlow[0].Id : null;
      // 🔥 PRESERVE HISTORY + ADD EDIT ENTRY
      const history = formData.WorkFlowHistory
        ? JSON.parse(formData.WorkFlowHistory)
        : [];

      history.push({
        CurrentApprover: employee.EmployeeName,
        ActionTaken: "Edited",
        Comment: remarks,
        Date: new Date().toISOString()
      });
      await sp.web.lists
        .getByTitle("CapexAdvance")
        .items.getById(formData.ID)
        .update({
          Title: formData.CapexID,
          CapexID: formData.CapexID,

          EmployeeCode: employee.EmployeeCode,
          EmployeeName: employee.EmployeeName,
          Division: employee.Division,
          Location: employee.Location,
          Email: employee.EmployeeEmail,
          RM: employee.ReportingManager?.Title,
          HOD: employee.HOD?.Title,
          ContactNo: employee.ContactNo,
          EmployeeStatus: employee.EmployeeStatus,

          VendorCodeId: selectedVendorId,
          VendorName: selectedVendorName,

          PONumber: poNumber,
          PODate: poDate ? new Date(poDate) : null,
          POAdvanceTerms: poTerms,

          POAmtGST: poAmount,
          RequestAdvanceAmount: advanceAmount,
          PaidAmount: paidAmount,

          ExpectedDateofSettlement: expectedDate
            ? new Date(expectedDate)
            : null,

          PICNameId: ensuredUser.Id,

          GL: glCode,
          CostCenter: costCenter,
          Remarks: remarks,
          ProjectDescription: projectDesc,



          Status: "Pending for Approver",
          // Status: formData.Status, 

  ApprovalMatrix: JSON.stringify(updatedFlow),

  WorkFlowHistory: JSON.stringify(history),

  CurrentApproverId: currentApprover
        });

      if (selectedFiles.length > 0) {
        await uploadFiles();
      }

      alert("Updated successfully ✅");
      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexForm.aspx?page=User";
      void handleExit();
    } catch (e) {
      console.error(e);
      alert("Error ❌");
    }
  };


  // =========================
  // BIND DATA
  // =========================
  useEffect(() => {
    if (!formData) return;

    setPoNumber(formData.PONumber || "");
    setPoDate(formData.PODate?.split("T")[0] || "");
    setPoTerms(formData.POAdvanceTerms || "");
    setPoAmount(formData.POAmtGST || "");
    setAdvanceAmount(formData.RequestAdvanceAmount || "");
    setPaidAmount(formData.PaidAmount || "");
    setExpectedDate(formData.ExpectedDateofSettlement?.split("T")[0] || "");

    setVendorName(formData.VendorName || "");
    setSelectedVendorId(formData.VendorCodeId || null); // ✅ ADD THIS
                    void getPreviousAdvances(formData.VendorCodeId || null);

    setSelectedVendorName(formData.VendorName || "");   // ✅ ADD THIS

    setGlCode(formData.GL || "");
    setCostCenter(formData.CostCenter || "");
    setRemarks(formData.Remarks || "");
    setProjectDesc(formData.ProjectDescription || "");

    setApproverRemarks(formData.ApproverRemarks || "");
    setVoucherDate(formData.VoucherDate?.split("T")[0] || "");
    setVoucherNumber(formData.VoucherNumber || "");
    setUTRDate(formData.UTRDate?.split("T")[0] || "");
    setUTRNumber(formData.UTRNumber || "");

    // ✅ PIC FIX
    if (formData?.PICName?.Title) {
      setSelectedUser([
        {
          text: formData.PICName.Title,
          secondaryText: formData.PICName.EMail,
        },
      ]);
    }
    // ✅ Approval Matrix
    if (formData?.ApprovalMatrix) {
      try {
        const parsed =
          typeof formData.ApprovalMatrix === "string"
            ? JSON.parse(formData.ApprovalMatrix)
            : formData.ApprovalMatrix;

        setApprovalMatrix(Array.isArray(parsed) ? parsed : []);
      } catch (e) {
        console.error("ApprovalMatrix parse error", e);
        setApprovalMatrix([]);
      }
    }

    // ✅ Workflow History
    if (formData?.WorkFlowHistory) {
      try {
        const parsed =
          typeof formData.WorkFlowHistory === "string"
            ? JSON.parse(formData.WorkFlowHistory)
            : formData.WorkFlowHistory;

        setWorkflowHistory(Array.isArray(parsed) ? parsed : []);
      } catch (e) {
        console.error("WorkFlowHistory parse error", e);
        setWorkflowHistory([]);
      }
    }

    if (formData.CapexID) {
      void getAttachments(formData.CapexID);
    }

  }, [formData]);

  useEffect(() => {
    void getLoggedInUser();
    void getVendors();
  }, []);

  // =========================
  // UI
  // =========================
  return (
    <div className="form-container">
      <h2 className="form-title">Edit Advance Payment</h2>
      <div className="section">
        <h3>Approval Matrix</h3>

        {approvalMatrix.length === 0 ? (
          <p>No approval data</p>
        ) : (
          <div className="approval-flow">
            {approvalMatrix.map((a, index) => (
              <div
                key={index}
                className={`approval-step ${a.Status === "In Progress"
                  ? "active"
                  : a.Status === "Approved"
                    ? "approved"
                    : a.Status === "Rejected"
                      ? "rejected"
                      : a.Status === "Send Back"
                        ? "sendback"
                        : ""
                  }`}
              >
                <div><b>{a.Role}</b></div>
                <div>{a.Name}</div>
                <div>{a.Status}</div>
              </div>
            ))}
          </div>
        )}
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
              if (id) {
                void getPreviousAdvances(id);
              }
            }}
          >
            <option value="">Select Vendor</option>
            {vendors.map((v) => (
              <option key={v.Id} value={v.Id}>
                {v.VendorCode}
              </option>
            ))}
          </select>

          <label>Vendor Name</label>
          <input value={selectedVendorName || vendorName} readOnly />

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
          <input
            value={poTerms}
            onChange={(e) => setPoTerms(e.target.value)}
          />

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
            defaultSelectedUsers={
              selectedUser.length > 0
                ? [selectedUser[0].secondaryText]  // ✅ use email
                : []
            }
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


      {/* Previous Advances Table */}
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
            <tr>
              <td colSpan={7} style={{ textAlign: "center" }}>
                No Data
              </td>
            </tr>
          </tbody>
        </table>
      </div>
      <div className="section">
        <h3>Workflow History</h3>

        {workflowHistory.length === 0 ? (
          <p>No history available</p>
        ) : (
          <div className="workflow-history">
            {workflowHistory.map((h, index) => (
              <div key={index} className="history-item">
                <div>
                  {h.ActionTaken === "Submitted" && "📩 "}
                  {h.ActionTaken === "Approved" && "✅ "}
                  {h.ActionTaken === "Rejected" && "❌ "}
                  {h.ActionTaken === "Send Back" && "↩ "}
                  {h.ActionTaken}
                </div>

                <div><b>{h.CurrentApprover}</b></div>
                <div>{h.Comment}</div>
                <div>{new Date(h.Date).toLocaleString()}</div>
              </div>
            ))}
          </div>
        )}
      </div>
      {/* Attachments */}
      <div className="section">
        <div className="col-md-4">
                    <label className="font">Attachments</label>
                    {attachments.length > 0 && (
                      <ul>
                        {attachments.map((file: any, index: number) => (
                          <li key={index}>
                            <a
                              href={file.ServerRelativeUrl}
                              target="_blank"
                              rel="noopener noreferrer"
                            >
                              {file.Name}
                            </a>

                            {/* 🔥 DELETE FROM SHAREPOINT */}
                            <button
                              type="button"
                              style={{ marginLeft: "10px", color: "red" }}
                              onClick={() => handleDeleteExistingFile(file)}
                            >
                              Delete
                            </button>
                          </li>
                        ))}
                      </ul>
                    )}

                    <input
                      type="file"
                      multiple
                      className="form-control"
                      onChange={(e) => {
                        if (e.target.files) {
                          setSelectedFiles(Array.from(e.target.files));
                        }
                      }}
                    />
                  </div>
        {/* <h3>Attachments</h3>

        {attachments.length > 0 && (
          <ul>
            {attachments.map((file: any, index: number) => (
              <li key={index}>
                <a
                  href={file.ServerRelativeUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  {file.Name}
                </a>
              </li>
            ))}
          </ul>
        )}

        <input
          type="file"
          multiple
          onChange={(e) => {
            if (e.target.files) {
              setSelectedFiles(Array.from(e.target.files));
            }
          }}
        /> */}
      </div>

      {/* Buttons */}
      <div className="actions">
        <button className="primary" onClick={handleSubmit}>
          Submit
        </button>



        <button className="exit" onClick={handleExit}>
          Exit
        </button>
      </div>
    </div>
  );
};
export default EditAdvanceForm;
