import * as React from "react";
import "./advanced.scss";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { useEffect, useState } from "react";
import { IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/PeoplePicker";
interface IProps {
  context: any;
  itemId: number; // 👈 IMPORTANT
}
interface IVendor {
  Id: number;
  VendorCode: string;
  VendorName: string;
}
const APperformerAdvanceFormForUTR: React.FC<IProps> = ({
  context,
  itemId,
}) => {
  const sp = spfi().using(SPFx(context));
  const [attachments, setAttachments] = useState<any[]>([]);
  const [itemData, setItemData] = useState<any>(null);
  const [approverRemarks, setApproverRemarks] = useState("");
  const [voucherDate, setVoucherDate] = useState("");
  const [selectedVendorName, setSelectedVendorName] = useState("");
  const [vendors, setVendors] = useState<IVendor[]>([]);
  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  const [voucherNumber, setVoucherNumber] = useState("");
  const [UTRDate, setUTRDate] = useState("");
  const [UTRNumber, setUTRNumber] = useState("");
  const [UTRRemarks, setUTRRemarks] = useState("");
  const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);
  const [workflowHistory, setWorkflowHistory] = useState<any[]>([]);
  // ✅ Fetch Item by ID
  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: context.pageContext.web.absoluteUrl,
    msGraphClientFactory: context.msGraphClientFactory,
    spHttpClient: context.spHttpClient,
  };

  const getAttachments = async (capexId: string) => {
    try {
      const safe = capexId.replace(/\//g, "_");
      const path = `/sites/SonaFinance/CapexAdvanceDocs/${safe}`;

      const files = await sp.web.getFolderByServerRelativePath(path).files();

      void setAttachments(files);
    } catch {
      void setAttachments([]);
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
      // ✅ Approval Matrix
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

      // ✅ Workflow History
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
 const handleApprove = async () => {
  try {
    if (!UTRDate || !UTRNumber || !UTRRemarks) {
      alert("All fields required");
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
      flow[currentIndex].Status = "Approved";
    }

    const history = itemData.WorkFlowHistory
      ? JSON.parse(itemData.WorkFlowHistory)
      : [];

    history.push({
      CurrentApprover: context.pageContext.user.displayName,
      ActionTaken: "Paid",
      Comment: UTRRemarks,
      Date: new Date().toISOString()
    });

    await sp.web.lists
      .getByTitle("CapexAdvance")
      .items.getById(itemId)
      .update({
        ApproverRemarks: approverRemarks,
        UTRDate: new Date(UTRDate),
        UTRNumber: UTRNumber,
        UTRRemarks: UTRRemarks,

        Status: "Paid",

        ApprovalMatrix: JSON.stringify(flow),
        WorkFlowHistory: JSON.stringify(history),

        CurrentApproverId: null // 🔥 final stage
      });

    alert("Paid ✅");
 window.location.href =
      "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexForm.aspx?page=Performer";
  } catch (error) {
    console.error(error);
  }
};
  // ✅ Sent Back
const handleSendBack = async () => {
  try {
    if (!UTRRemarks) {
      alert("Enter remarks");
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
      Comment: UTRRemarks,
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
     window.location.href =
      "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexForm.aspx?page=Performer";

  } catch (error) {
    console.error(error);
  }
};
  // ✅ Reject
  const handleReject = async () => {
  try {
    if (!UTRRemarks) {
      alert("Enter remarks");
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
      Comment: UTRRemarks,
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
     window.location.href =
      "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexForm.aspx?page=Performer";

  } catch (error) {
    console.error(error);
  }
};

  const handleExit = () => {
    window.location.href = `https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexForm.aspx?page=Performer`;
  };

  // ⛔ Wait until data loads
  if (!itemData) return <div>Loading...</div>;

  return (
    <div className="form-container">
      <h2 className="form-title">Advance Payment (Approver)</h2>
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
                    : ""
                  }`}
              >
                <div><b>{a.Role}</b></div>
                <div>{a.Name}</div>
                <div className="status">{a.Status}</div>
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
                <div>
                  {h.ActionTaken === "Approved" && "✅ "}
                  {h.ActionTaken === "Rejected" && "❌ "}
                  {h.ActionTaken === "Send Back" && "↩ "}
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

      {/* ✅ ONLY EDITABLE FIELD */}
      <div className="section">
        <h3>Approver Action</h3>

        <label>Approver Remarks</label>
        <textarea
          value={approverRemarks}
          onChange={(e) => setApproverRemarks(e.target.value)}
          readOnly
        />
      </div>
      <label>Voucher Date</label>
      <input
        value={
          itemData.VoucherDate
            ? new Date(itemData.VoucherDate).toLocaleDateString("en-GB")
            : ""
        }
        readOnly
      />

      <label>Voucher Number</label>
      <input value={itemData.VouchingNumber || ""} readOnly />


      {/* Attachment */}
      <div className="section">
        <div className="section">
          <h3>Attachments</h3>

          {attachments.length === 0 ? (
            <p>No attachments</p>
          ) : (
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
        </div>
      </div>
      <label>UTR Date</label>
      <input
        type="date"
        value={UTRDate}
        onChange={(e) => setUTRDate(e.target.value)}
      />
      <label>UTR Number</label>
      <input value={UTRNumber} onChange={(e) => setUTRNumber(e.target.value)} />
      <label>UTR Remarks</label>
      <input onChange={(e) => setUTRRemarks(e.target.value)} />
      {/* Buttons */}
      <div className="actions">
        <button className="primary" onClick={handleApprove}>
          Paid
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

export default APperformerAdvanceFormForUTR;
