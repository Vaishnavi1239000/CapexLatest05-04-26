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
const ViewAdvanceForm = ({ context, formData, onClose }: any) => {
  const [attachments, setAttachments] = useState<any[]>([]);
  const sp = spfi().using(SPFx(context));
  const [employee, setEmployee] = useState<any>({});
  // 🔹 Employee
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [selectedUser, setSelectedUser] = useState<any[]>([]);
  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  const [selectedVendorName, setSelectedVendorName] = useState("");

  const [vendors, setVendors] = useState<IVendor[]>([]);
  const [employeeName, setEmployeeName] = useState("");
  const [employeeEmail, setEmployeeEmail] = useState("");
  const [VendorCode, setVendorCode] = useState<number | null>(null);
  // 🔹 Form fields
  const [poNumber, setPoNumber] = useState("");
  const [poDate, setPoDate] = useState("");
  const [poTerms, setPoTerms] = useState("");
  const [poAmount, setPoAmount] = useState("");
  const [advanceAmount, setAdvanceAmount] = useState("");
  const [paidAmount, setPaidAmount] = useState("");
  const [expectedDate, setExpectedDate] = useState("");
  const [vendorName, setVendorName] = useState("");
  const [glCode, setGlCode] = useState("");
  const [costCenter, setCostCenter] = useState("");
  const [remarks, setRemarks] = useState("");
  const [projectDesc, setProjectDesc] = useState("");

  // 🔹 Extra fields
  const [approverRemarks, setApproverRemarks] = useState("");
  const [voucherDate, setVoucherDate] = useState("");
  const [VouchingNumber, setVouchingNumber] = useState("");
  const [UTRDate, setUTRDate] = useState("");
  const [UTRNumber, setUTRNumber] = useState("");
  const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);
  const [workflowHistory, setWorkflowHistory] = useState<any[]>([]);
  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: context.pageContext.web.absoluteUrl,
    msGraphClientFactory: context.msGraphClientFactory,
    spHttpClient: context.spHttpClient,
  };
  const getAttachments = async (capexId: string) => {
    try {
      if (!capexId) return;

      const safeCapexId = capexId.replace(/\//g, "_");

      const folderPath = `/sites/SonaFinance/CapexAdvanceDocs/${safeCapexId}`;

      console.log("Folder Path:", folderPath); // ✅ DEBUG

      const files = await sp.web
        .getFolderByServerRelativePath(folderPath)
        .files();

      console.log("Files:", files); // ✅ DEBUG

      setAttachments(files || []);
    } catch (error) {
      console.log("Attachment fetch error:", error);
      setAttachments([]);
    }
  };

  const uploadFiles = async () => {
    if (!formData?.CapexID || selectedFiles.length === 0) return;

    const safe = formData.CapexID.replace(/\//g, "_");
    const path = `/sites/SonaFinance/CapexAdvanceDocs/${safe}`;

    for (const file of selectedFiles) {
      await sp.web
        .getFolderByServerRelativePath(path)
        .files.addUsingPath(file.name, file, { Overwrite: true });
    }

    void setSelectedFiles([]);
    void getAttachments(formData.CapexID);
  };

  // ✅ Fetch Item by ID
  const getVendors = async () => {
    const data = await sp.web.lists
      .getByTitle("VendorMaster")
      .items.select("Id", "VendorCode", "VendorName")();
    void setVendors(data);
  };

  // ✅ Bind SharePoint Data
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
    setSelectedVendorName(formData.VendorName || ""); // ✅ ADD THIS

    setGlCode(formData.GL || "");
    setCostCenter(formData.CostCenter || "");
    setRemarks(formData.Remarks || "");
    setProjectDesc(formData.ProjectDescription || "");

    setApproverRemarks(formData.ApproverRemarks || "");
    setVoucherDate(formData.VoucherDate?.split("T")[0] || "");
    setVouchingNumber(formData.VouchingNumber || "");
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

    if (formData.CapexID) {
      void getAttachments(formData.CapexID);
    }
    // ✅ Approval Matrix (SAFE)
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
    } else {
      setApprovalMatrix([]);
    }

    // ✅ Workflow History (SAFE)
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
    } else {
      setWorkflowHistory([]);
    }
  }, [formData]);


  const handleExit = () => {
    if (onClose) onClose();
    else window.location.reload();
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
  useEffect(() => {
    void getLoggedInUser();
    void getVendors();
  }, []);
  return (
    <div className="form-container">
      <h2>Advance Payment (View)</h2>
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
      {/* Employee */}

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

      {/* Vendor */}
      <div className="section">
        <h3>Vendor & PO</h3>

        <label>Vendor Code</label>

        <select
          value={selectedVendorId ?? ""}
          disabled={true}
          onChange={(e) => {

            const id = Number(e.target.value);
            const vendor = vendors.find((v) => v.Id === id);
            setSelectedVendorId(id);
            setSelectedVendorName(vendor?.VendorName || "");
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
        <input value={vendorName} readOnly />

        <label>PO Number</label>
        <input value={poNumber} readOnly />

        <label>PO Date</label>
        <input type="date" value={poDate} readOnly />

        <label>PO Terms</label>
        <input value={poTerms} readOnly />
      </div>

      {/* Amount */}
      <div className="section">
        <h3>Amount</h3>

        <label>PO Amount</label>
        <input value={poAmount} readOnly />

        <label>Advance Amount</label>
        <input value={advanceAmount} readOnly />

        <label>Paid Amount</label>
        <input value={paidAmount} readOnly />
      </div>

      {/* Advance */}
      <div className="section">
        <h3>Advance Details</h3>

        <label>Expected Date</label>
        <input type="date" value={expectedDate} readOnly />

        <label>PIC Name</label>
        <PeoplePicker
          context={peoplePickerContext}
          personSelectionLimit={1}
          disabled={true}
          principalTypes={[PrincipalType.User]}
          defaultSelectedUsers={
            formData?.PICName?.Title ? [formData.PICName.Title] : []
          }
        />

        <label>GL Code</label>
        <input value={glCode} readOnly />

        <label>Cost Center</label>
        <input value={costCenter} readOnly />
      </div>

      {/* Remarks */}
      <div className="section">
        <h3>Remarks</h3>

        <textarea value={remarks} readOnly />
        <h3>Project Description</h3>
        <textarea value={projectDesc} readOnly />
      </div>

      {/* Approver */}
      <div className="section">
        <h3>Approver Details</h3>

        <label>Approver Remarks</label>
        <textarea value={approverRemarks} readOnly />

        <label>Voucher Date</label>
        <input type="date" value={voucherDate} readOnly />

        <label>Voucher Number</label>
        <input value={VouchingNumber} readOnly />

        <label>UTR Date</label>
        <input type="date" value={UTRDate} readOnly />

        <label>UTR Number</label>
        <input value={UTRNumber} readOnly />
      </div>

      {/* Attachment */}
      <div className="section">
        <h3>Attachments</h3>

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
                    {h.ActionTaken === "Vouched" && "💰 "}
                    {h.ActionTaken === "Paid" && "💸 "}
                    {h.ActionTaken}
                  </div>

                  <div><b>{h.CurrentApprover}</b></div>
                  <div>{h.Comment}</div>
                  <div className="date">
                    {new Date(h.Date).toLocaleString()}
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>

      </div>
      {/* Button */}
      <div style={{ textAlign: "center", marginTop: "20px" }}>
        <button onClick={handleExit}>Back</button>
      </div>
    </div>
  );
};

export default ViewAdvanceForm;
