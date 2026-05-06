import * as React from "react";
import "./advanced.scss";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { useEffect, useState } from "react";
import { IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
interface IProps {
  context: any;
  itemId: number; // 👈 IMPORTANT
}

interface IVendor {
  Id: number;
  VendorCode: string;
  VendorName: string;
}
const ApproverAdvanceForm: React.FC<IProps> = ({ context, itemId }) => {
  const sp = spfi().using(SPFx(context));
  const [selectedVendorName, setSelectedVendorName] = useState("");
  const [itemData, setItemData] = useState<any>(null);
  const [approverRemarks, setApproverRemarks] = useState("");
  const [attachments, setAttachments] = useState<any[]>([]);
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [vendors, setVendors] = useState<IVendor[]>([]);
  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);
  const [workflowHistory, setWorkflowHistory] = useState<any[]>([]);
  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: context.pageContext.web.absoluteUrl,
    msGraphClientFactory: context.msGraphClientFactory,
    spHttpClient: context.spHttpClient,
  };
  const getAttachments = async (capexId: string) => {
    debugger;
    try {
      if (!capexId) return;

      const safeCapexId = capexId.replace(/\//g, "_");

      // ✅ FIXED PATH (MOST IMPORTANT)
      const folderPath = `CapexAdvanceDocs/${safeCapexId}`;

      console.log("Folder Path:", folderPath);

      const files = await sp.web
        .getFolderByServerRelativePath(folderPath)
        .files();

      console.log("Files:", files);

      setAttachments(files || []);
    } catch (error) {
      console.log("Attachment fetch error:", error);
      setAttachments([]);
    }
  };

  const uploadAttachments = async (capexId: string) => {
    try {
      if (attachments.length === 0) return;

      const safeCapexId = capexId.replace(/\//g, "_"); // ✅ move here

      const folderPath = `CapexAdvanceDocs/${safeCapexId}`;

      await sp.web.folders.addUsingPath(folderPath);

      for (const file of attachments) {
        await sp.web
          .getFolderByServerRelativePath(folderPath)
          .files.addUsingPath(file.name, file, { Overwrite: true });
      }
    } catch (error) {
      console.error("Upload error:", error);
    }
  };
  const getVendors = async () => {
    try {
      const data = await sp.web.lists
        .getByTitle("VendorMaster")
        .items.select("Id", "VendorCode", "VendorName")();

      setVendors(data);
    } catch (error) {
      console.error("Vendor fetch error:", error);
    }
  };


  // ✅ Fetch Item by ID
  const getItemById = async () => {
    try {


      const item = await sp.web.lists
        .getByTitle("CapexAdvance")
        .items.getById(itemId)
        .select("*", "PICName/Title", "VendorCode/Id", "VendorCode/VendorCode")
        .expand("PICName", "VendorCode")
        // 👈 ADD
        ();


      setItemData(item);
      //setApproverRemarks(item.ApproverRemarks || "");

      // ✅ FIX: Set VendorId + Name
      setSelectedVendorId(item.VendorCode?.Id || null);
      // 🔥 IMPORTANT
      setSelectedVendorName(item.VendorName); // optional

      // ✅ FETCH ATTACHMENTS
      if (item.CapexID) {
        await getAttachments(item.CapexID);
      }
      if (item.ApprovalMatrix) {
        try {
          const parsed =
            typeof item.ApprovalMatrix === "string"
              ? JSON.parse(item.ApprovalMatrix)
              : item.ApprovalMatrix;

          setApprovalMatrix(Array.isArray(parsed) ? parsed : []);
        } catch (e) {
          console.error("ApprovalMatrix parse error", e);
          setApprovalMatrix([]);
        }
      } else {
        setApprovalMatrix([]);
      }
      if (item.WorkFlowHistory) {
        try {
          const parsed =
            typeof item.WorkFlowHistory === "string"
              ? JSON.parse(item.WorkFlowHistory)
              : item.WorkFlowHistory;

          setWorkflowHistory(Array.isArray(parsed) ? parsed : []);
        } catch (e) {
          console.error("WorkFlowHistory parse error", e);
          setWorkflowHistory([]);
        }
      } else {
        setWorkflowHistory([]);
      }
    } catch (error) {
      console.error("Fetch error:", error);
    }
  };

  useEffect(() => {
    if (!context || !itemId) return;

    const loadData = async () => {
      await getItemById();    // 👈 FIRST load item to get VendorCode
      await getVendors();     // 👈 FIRST load vendors
      await getItemById();    // 👈 THEN item
    };

    void loadData();
  }, [context, itemId]);


  // ✅ Approve
  const buildApprovalFlow = async () => {

    try {

      // =========================
      // 🔹 EXISTING FLOW (SOURCE OF TRUTH)
      // =========================
      const existingFlow = itemData.ApprovalMatrix
        ? JSON.parse(itemData.ApprovalMatrix)
        : [];

      // =========================
      // 🔹 KEEP RM + HOD FROM OLD FLOW
      // =========================
      const baseApprovers = existingFlow.filter(
        (a: any) => a.Role === "RM" || a.Role === "HOD"
      );

      // =========================
      // 🔥 GET LATEST MATRIX APPROVERS
      // =========================
      const matrixData = await sp.web.lists
        .getByTitle("CapexApprovalMatrix")
        .items
        .select("Role/RoleName,Approver/Id,Approver/Title,Level/Level")
        .expand("Approver,Role,Level")
        .filter("Status eq 'Active'")
        .orderBy("Level", true)();

      const matrixApprovers = matrixData.map((item: any, index: number) => ({
        Id: item.Approver?.Id,
        Name: item.Approver?.Title,
        Role: item.Role?.RoleName,
        Level: baseApprovers.length + index + 1,
        Status: "Pending"
      }));

      // =========================
      // 🔹 MERGE FLOW
      // =========================
      const fullFlow = [...baseApprovers, ...matrixApprovers];

      // =========================
      // 🔹 REMOVE DUPLICATES
      // =========================
      const uniqueFlow = fullFlow.filter(
        (v, i, self) => self.findIndex(x => x.Id === v.Id) === i
      );

      return uniqueFlow;

    } catch (error) {
      console.error("Approval Flow Error:", error);
      return [];
    }
  };

  const handleApprove = async () => {
    try {
      if (!approverRemarks) {
        alert("Please enter Remarks");
        return;
      }

      // =========================
      // 🔹 EXISTING FLOW
      // =========================
      let existingFlow = itemData.ApprovalMatrix
        ? JSON.parse(itemData.ApprovalMatrix)
        : [];

      // =========================
      // 🔥 REBUILD FLOW (LATEST FROM MATRIX LIST)
      // =========================
      const latestFlow = await buildApprovalFlow();

      // =========================
      // 🔥 CHECK IF MATRIX CHANGED
      // =========================
      let isChanged = false;

      if (latestFlow.length !== existingFlow.length) {
        isChanged = true;
      } else {
        for (let i = 0; i < latestFlow.length; i++) {
          if (latestFlow[i].Id !== existingFlow[i]?.Id) {
            isChanged = true;
            break;
          }
        }
      }

      // =========================
      // 🔥 IF CHANGED → RESET FLOW
      // =========================
      if (isChanged) {
        existingFlow = latestFlow;

        // reset all
        existingFlow.forEach((a: any) => (a.Status = "Pending"));

        if (existingFlow.length > 0) {
          existingFlow[0].Status = "In Progress";
        }
      }

      // =========================
      // 🔹 FIND CURRENT APPROVER
      // =========================
      const userEmail = context.pageContext.user.email;

      //const ensuredUser = await sp.web.ensureUser(userEmail);

      const currentUserId = context.pageContext.legacyPageContext.userId;
      const currentIndex = existingFlow.findIndex(
        (a: any) => a.Id === currentUserId
      );

      if (currentIndex === -1) {
        alert("You are not the current approver");
        return;
      }

      // =========================
      // 🔹 UPDATE CURRENT APPROVER
      // =========================
      existingFlow[currentIndex].Status = "Approved";

      // =========================
      // 🔹 MOVE NEXT
      // =========================
      let nextApproverId = null;

      if (existingFlow[currentIndex + 1]) {
        existingFlow[currentIndex + 1].Status = "In Progress";
        nextApproverId = existingFlow[currentIndex + 1].Id;
      }

      // =========================
      // 🔥 STATUS LOGIC
      // =========================
      let finalStatus = itemData.Status;

      const currentRole = existingFlow[currentIndex]?.Role;

      if (currentRole === "RM") {
        finalStatus = "Pending for Approver";
      } else if (currentRole === "HOD") {
        finalStatus = "Pending for PF Approver";
      } else {
        finalStatus = nextApproverId ? "Pending" : "Approved";
      }

      // =========================
      // 🔥 WORKFLOW HISTORY
      // =========================
      const history = itemData.WorkFlowHistory
        ? JSON.parse(itemData.WorkFlowHistory)
        : [];

      history.push({
        CurrentApprover: context.pageContext.user.displayName,
        ActionTaken: "Approved",
        Comment: approverRemarks,
        Date: new Date().toISOString()
      });

      // =========================
      // 🔥 UPDATE SHAREPOINT
      // =========================
      await sp.web.lists
        .getByTitle("CapexAdvance")
        .items.getById(itemId)
        .update({
          ApproverRemarks: approverRemarks,

          Status: finalStatus, // ✅ FIXED

          ApprovalMatrix: JSON.stringify(existingFlow),

          CurrentApproverId: nextApproverId,

          WorkFlowHistory: JSON.stringify(history)
        });

      alert("Approved successfully ✅");

      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexForm.aspx?page=Approver";

    } catch (error) {
      console.error("Approve error:", error);
      alert("Error ❌");
    }
  };

  // ✅ Sent Back
  const handleSendBack = async () => {
    try {
      if (!approverRemarks) {
        alert("Please enter Remarks");
        return;
      }

      const flow = itemData.ApprovalMatrix
        ? JSON.parse(itemData.ApprovalMatrix)
        : [];

      const currentUserId = context.pageContext.legacyPageContext.userId;

      const currentIndex = flow.findIndex(
        (a: any) => a.Id === currentUserId
      );

      if (currentIndex !== -1) {
        flow[currentIndex].Status = "Send Back";
      }

      const history = itemData.WorkFlowHistory
        ? JSON.parse(itemData.WorkFlowHistory)
        : [];

      history.push({
        CurrentApprover: context.pageContext.user.displayName,
        ActionTaken: "Send Back",
        Comment: approverRemarks,
        Date: new Date().toISOString()
      });

      await sp.web.lists
        .getByTitle("CapexAdvance")
        .items.getById(itemId)
        .update({
          ApproverRemarks: approverRemarks,
          Status: "Send Back",

          ApprovalMatrix: JSON.stringify(flow),
          WorkFlowHistory: JSON.stringify(history),

          CurrentApproverId: itemData.CurrentApproverId
        });

      alert("Send Back ✅");
      window.location.href = "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexForm.aspx?page=Approver";

    } catch (error) {
      console.error(error);
    }
  };

  // ✅ Reject
  const handleReject = async () => {
    try {
      if (!approverRemarks) {
        alert("Please enter Remarks");
        return;
      }

      const flow = itemData.ApprovalMatrix
        ? JSON.parse(itemData.ApprovalMatrix)
        : [];

      const currentUserId = context.pageContext.legacyPageContext.userId;

      const currentIndex = flow.findIndex(
        (a: any) => a.Id === currentUserId
      );

      if (currentIndex !== -1) {
        flow[currentIndex].Status = "Rejected";
      }

      const history = itemData.WorkFlowHistory
        ? JSON.parse(itemData.WorkFlowHistory)
        : [];

      history.push({
        CurrentApprover: context.pageContext.user.displayName,
        ActionTaken: "Rejected",
        Comment: approverRemarks,
        Date: new Date().toISOString()
      });

      await sp.web.lists
        .getByTitle("CapexAdvance")
        .items.getById(itemId)
        .update({
          ApproverRemarks: approverRemarks,
          Status: "Rejected",

          ApprovalMatrix: JSON.stringify(flow),
          WorkFlowHistory: JSON.stringify(history),

          CurrentApproverId: null
        });

      alert("Rejected ❌");
      window.location.href = "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexForm.aspx?page=Approver";

    } catch (error) {
      console.error(error);
    }
  };

  const handleExit = () => {
    window.location.href = `https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexForm.aspx?page=Approver`;
  };

  // ⛔ Wait until data loads
  if (!itemData) return <div>Loading...</div>;

  return (
    <div className="form-container">
      <h2 className="form-title">Advance Payment (Approver)</h2>
      <div className="section">
        <h3>Approval Matrix</h3>

        <div className="approval-flow">
          {approvalMatrix.length === 0 ? (
            <p>No approval data</p>
          ) : (
            approvalMatrix.map((a, index) => (
              <div
                key={index}
                className={`approval-step ${a.Status === "In Progress"
                  ? "active"
                  : a.Status === "Approved"
                    ? "approved"
                    : ""
                  }`}
              >
                <div><b>{a.Role}</b></div>
                <div>{a.Name}</div>

                <div className="status">{a.Status}</div>
              </div>
            ))
          )}
        </div>
      </div>

      {/* Employee Info */}
      <div className="section">
        <h3>Requestor Information</h3>

        <div className="form-row">
          <label>Employee Code</label>
          <input value={itemData.EmployeeCode || ""} readOnly />

          <label>Employee Name</label>
          <input value={itemData.EmployeeName || ""} readOnly />

          <label>Division</label>
          <input value={itemData.Division || ""} readOnly />

          <label>Location</label>
          <input value={itemData.Location || ""} readOnly />
        </div>

        <div className="form-row">
          <label>Email</label>
          <input value={itemData.Email || ""} readOnly />

          <label>RM</label>
          <input value={itemData.RM || ""} readOnly />

          <label>HOD</label>
          <input value={itemData.HOD || ""} readOnly />

          <label>Contact No</label>
          <input value={itemData.ContactNo || ""} readOnly />
        </div>
      </div>

      {/* Vendor */}
      <div className="section">
        <h3>Vendor & PO Details</h3>

        <div className="form-row">
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
          <input value={itemData.VendorName || ""} readOnly />

          <label>PO Number</label>
          <input value={itemData.PONumber || ""} readOnly />

          <label>PO Date</label>
          <input
            value={
              itemData.PODate
                ? new Date(itemData.PODate).toLocaleDateString("en-GB")
                : ""
            }
            readOnly
          />
        </div>

        <div className="form-row">
          <label>PO Terms</label>
          <input value={itemData.POAdvanceTerms || ""} readOnly />

          <label>PO Amount</label>
          <input value={itemData.POAmtGST || ""} readOnly />

          <label>Advance Amount</label>
          <input value={itemData.RequestAdvanceAmount || ""} readOnly />

          <label>Paid Amount</label>
          <input value={itemData.PaidAmount || ""} readOnly />
        </div>
      </div>

      {/* Advance */}
      <div className="section">
        <h3>Advance Details</h3>

        <div className="form-row">
          <label>Expected Settlement</label>
          <input
            value={
              itemData.ExpectedDateofSettlement
                ? new Date(
                  itemData.ExpectedDateofSettlement,
                ).toLocaleDateString("en-GB")
                : ""
            }
            readOnly
          />
          <label>PIC Name</label>
          <PeoplePicker
            context={peoplePickerContext}
            personSelectionLimit={1}
            disabled={true}
            principalTypes={[PrincipalType.User]}
            defaultSelectedUsers={
              itemData?.PICName?.Title ? [itemData.PICName.Title] : []
            }

          />
          <label>GL Code</label>
          <input value={itemData.GL || ""} readOnly />

          <label>Cost Center</label>
          <input value={itemData.CostCenter || ""} readOnly />
        </div>
      </div>

      {/* Remarks */}
      <div className="section">
        <h3>Remarks</h3>

        <div className="form-row">
          <label>User Remarks</label>
          <textarea value={itemData.Remarks || ""} readOnly />

          <label>Project Description</label>
          <textarea value={itemData.ProjectDescription || ""} readOnly />
        </div>
      </div>
      <div className="section">
        <h3>Workflow History</h3>

        {workflowHistory.length === 0 ? (
          <p>No history available</p>
        ) : (
          <div className="workflow-history">
            {workflowHistory.map((h, index) => (
              <div key={index} className="history-item">
                <div><b>{h.ActionTaken}</b></div>
                <div>{h.CurrentApprover}</div>
                <div>{h.Comment}</div>
                <div className="date">
                  {new Date(h.Date).toLocaleString()}
                </div>
              </div>
            ))}
          </div>
        )}
      </div>

      {/* ✅ ONLY EDITABLE FIELD */}
      <div className="section">
        <h3>Approver Action</h3>

        <label>Approver Remarks</label>
        <textarea
          value={approverRemarks}
          onChange={(e) => setApproverRemarks(e.target.value)}
        />
      </div>

      <div className="section">
        <h3>Attachments</h3>

        {attachments.length === 0 ? (
          <p>No attachments found</p>
        ) : (
          <ul>
            {attachments.map((file: any, index: number) => (
              <li key={index}>
                <a href={file.ServerRelativeUrl} target="_blank">
                  {file.Name}
                </a>
              </li>
            ))}
          </ul>
        )}
      </div>

      {/* Buttons */}
      <div className="actions">
        <button className="primary" onClick={handleApprove}>
          Approve
        </button>

        <button className="secondary" onClick={handleSendBack}>
          Sent Back
        </button>

        <button className="secondary" onClick={handleReject}>
          Reject
        </button>

        <button className="exit" onClick={handleExit}>
          Exit
        </button>
      </div>
    </div>
  );
};

export default ApproverAdvanceForm;
