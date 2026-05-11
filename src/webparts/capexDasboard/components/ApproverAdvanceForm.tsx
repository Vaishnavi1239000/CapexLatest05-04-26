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
import logo from "../assets/sona-comstarlogo.png";
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
    <div className='MainUplodForm' style={{ margin: "5px 0px" }}>
      <div className='row'>
        <div className='col-md-12'>
          <div className='Main-Boxpoup'>
            <div className="bordered">
              <img src={logo} />
              <h1> Approval Matrix </h1>
            </div>
            {approvalMatrix.length === 0 ? (
              <p>No approval data</p>
            ) : (
              <div className="displayWF">
                <ul className="approval-flow">
                  {approvalMatrix.map((a, index) => (
                    <li
                      key={index}
                      className={`approval-step ${a.Status === "In Progress"
                        ? "active"
                        : a.Status === "Approved"
                          ? "approved"
                          : ""
                        }`}
                    >
                      {a.Role} - {a.Name}
                    </li>
                  ))}
                </ul>
              </div>
            )}
            {/* <div className='borderedbox'>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Requestor Information</label>
              </div>
              <div className='main-formcontainer'>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label htmlFor="Employee Code" className='font'>Employee Code</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.EmployeeCode}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="Employee Name" className='font'>Employee Name </label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.EmployeeName}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="Employee Email" className='font'>Employee Email </label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.EmployeeEmail}</label>
                  </div>
                </div>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label htmlFor="Contact No" className='font'>Contact No</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.ContactNo}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="Employee Status" className='font'>Employee Status</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.EmployeeStatus}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="Division" className='font'>Division</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.Division}</label>
                  </div>
                </div>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label htmlFor="Location" className='font'>Location</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.Location}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="RM" className='font'>RM</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.RM}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="HOD" className='font'>HOD</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.HOD}</label>
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Vendor & PO Details</label>
              </div>
              <div className='main-formcontainer'>
                <div className='row mb-20'>
                  <div className="col-md-4">
                    <label className="font">Vendor Code</label>
                    <select
                      value={selectedVendorId ?? ""}
                      disabled={true}
                      onChange={(e) => {
                        const id = Number(e.target.value);
                        const vendor = vendors.find((v) => v.Id === id);
                        setSelectedVendorId(id);
                        setSelectedVendorName(vendor?.VendorName || "");
                      }} className="formtext-control"
                    >
                      <option value="">Select Vendor</option>
                      {vendors.map((v) => (
                        <option key={v.Id} value={v.Id}>
                          {v.VendorCode}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div className="col-md-4">
                    <label>Vendor Name</label>
                    <input value={itemData.VendorName || ""} className="form-control readonly" />
                  </div>
                  <div className="col-md-4">
                    <label>PO Number</label>
                    <input value={itemData.PONumber || ""} className="form-control readonly" />
                  </div>
                </div>
                <div className='row mb-20'>
                  <div className="col-md-4">
                    <label className="font">PO Date</label>
                    <input value={itemData.PODate ? new Date(itemData.PODate).toLocaleDateString("en-GB") : ""} className="font-control readonly" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Terms</label>
                    <input value={itemData.POAdvanceTerms || ""} className="font-control readonly" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Amount</label>
                    <input value={itemData.POAmtGST || ""} className="font-control readonly" />
                  </div>
                </div>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label className="font">Advance Amount</label>
                    <input value={itemData.RequestAdvanceAmount || ""} className="font-control readonly" />
                  </div>
                  <div className='col-md-4'>
                    <label className="font">Paid Amount</label>
                    <input value={itemData.PaidAmount || ""} className="font-control readonly" />
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Advance Details</label>
              </div>
              <div className='main-formcontainer'>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Expected Settlement</label>
                    <input value={itemData.ExpectedDateofSettlement ? new Date(itemData.ExpectedDateofSettlement,).toLocaleDateString("en-GB") : ""}
                      className="font-control readonly" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PIC Name</label>
                    <PeoplePicker
                      context={peoplePickerContext}
                      personSelectionLimit={1}
                      disabled={true}
                      principalTypes={[PrincipalType.User]}
                      defaultSelectedUsers={
                        itemData?.PICName?.Title ? [itemData.PICName.Title] : []
                      }
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">GL Code</label>
                    <input value={itemData.GL || ""} className="font-control readonly" />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Cost Center</label>
                    <input value={itemData.CostCenter || ""} className="font-control readonly" />
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Remarks</label>
              </div>
              <div className='main-formcontainer'>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">User Remarks</label>
                    <textarea value={itemData.Remarks || ""} className="font-control readonly" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Project Description</label>
                    <textarea value={itemData.ProjectDescription || ""} className="font-control readonly" />
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Workflow History</label>
              </div>
              <div className='main-formcontainer'>
                <div className='row mb-20'>
                  <div className='col-md-12'>
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
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Approver Action</label>
              </div>
              <div className='main-formcontainer'>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Approver Remarks</label>
                    <textarea className="font-control"
                      value={approverRemarks}
                      onChange={(e) => setApproverRemarks(e.target.value)}
                    />
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Attachments</label>
              </div>
              <div className='main-formcontainer'>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Attachments</label>
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
                </div>
              </div>
              <div style={{ display: "flex", justifyContent: "center", gap: "5px", marginBottom: "1rem", marginTop: "1rem" }}>
                <a onClick={handleApprove} className="submit-btn">
                  Approve
                </a>
                <a onClick={handleSendBack} className="Rework-btn">
                  Sent Back
                </a>
                <a onClick={handleReject} className="Reject-btn">
                  Reject
                </a>
                <a href="#" onClick={handleExit} className="reset-btn">
                  Exit
                </a>
              </div>
            </div> */}
            <div className='borderedbox'>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Requestor Information</label>
              </div>
              <div className='main-formcontainer'>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label htmlFor="Employee Code" className='font'>Employee Code</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.EmployeeCode}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="Employee Name" className='font'>Employee Name </label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.EmployeeName}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="Employee Email" className='font'>Employee Email </label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.EmployeeEmail}</label>
                  </div>
                </div>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label htmlFor="Contact No" className='font'>Contact No</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.ContactNo}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="Employee Status" className='font'>Employee Status</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.EmployeeStatus}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="Division" className='font'>Division</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.Division}</label>
                  </div>
                </div>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label htmlFor="Location" className='font'>Location</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.Location}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="RM" className='font'>RM</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.RM}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="HOD" className='font'>HOD</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.HOD}</label>
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Vendor & PO Details</label>
              </div>
              <div className='main-formcontainer'>
                <div className='row mb-20'>
                  <div className="col-md-4">
                    <label className="font">Vendor Code</label>
                    <select
                      value={selectedVendorId ?? ""}
                      disabled={true}
                      onChange={(e) => {
                        const id = Number(e.target.value);
                        const vendor = vendors.find((v) => v.Id === id);
                        setSelectedVendorId(id);
                        setSelectedVendorName(vendor?.VendorName || "");
                      }} className="formtext-control"
                    >
                      <option value="">Select Vendor</option>
                      {vendors.map((v) => (
                        <option key={v.Id} value={v.Id}>
                          {v.VendorCode}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div className="col-md-4">
                    <label>Vendor Name</label>
                    <input value={itemData.VendorName || ""} className="form-control readonly" />
                  </div>
                  <div className="col-md-4">
                    <label>PO Number</label>
                    <input value={itemData.PONumber || ""} className="form-control readonly" />
                  </div>
                </div>
                <div className='row mb-20'>
                  <div className="col-md-4">
                    <label className="font">PO Date</label>
                    <input value={itemData.PODate ? new Date(itemData.PODate).toLocaleDateString("en-GB") : ""} className="font-control readonly" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Terms</label>
                    <input value={itemData.POAdvanceTerms || ""} className="font-control readonly" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Amount</label>
                    <input value={itemData.POAmtGST || ""} className="font-control readonly" />
                  </div>
                </div>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label className="font">Advance Amount</label>
                    <input value={itemData.RequestAdvanceAmount || ""} className="font-control readonly" />
                  </div>
                  <div className='col-md-4'>
                    <label className="font">Paid Amount</label>
                    <input value={itemData.PaidAmount || ""} className="font-control readonly" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Expected Settlement</label>
                    <input value={itemData.ExpectedDateofSettlement ? new Date(itemData.ExpectedDateofSettlement,).toLocaleDateString("en-GB") : ""}
                      className="font-control readonly" />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">PIC Name</label>
                    <PeoplePicker
                      context={peoplePickerContext}
                      personSelectionLimit={1}
                      disabled={true}
                      principalTypes={[PrincipalType.User]}
                      defaultSelectedUsers={
                        itemData?.PICName?.Title ? [itemData.PICName.Title] : []
                      }
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">GL Code</label>
                    <input value={itemData.GL || ""} className="font-control readonly" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Cost Center</label>
                    <input value={itemData.CostCenter || ""} className="font-control readonly" />
                  </div>
                </div>
                <div className='row mb-20'>
                  <div className='col-md-12'>
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
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">User Remarks</label>
                    <textarea value={itemData.Remarks || ""} className="font-control readonly" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Project Description</label>
                    <textarea value={itemData.ProjectDescription || ""} className="font-control readonly" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Approver Remarks</label>
                    <textarea className="font-control"
                      value={approverRemarks}
                      onChange={(e) => setApproverRemarks(e.target.value)}
                    />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Attachments</label>
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
                </div>
                <div style={{ display: "flex", justifyContent: "center", gap: "5px", marginBottom: "1rem", marginTop: "1rem" }}>
                  <a onClick={handleApprove} className="submit-btn">
                    Approve
                  </a>
                  <a onClick={handleSendBack} className="Rework-btn">
                    Sent Back
                  </a>
                  <a onClick={handleReject} className="Reject-btn">
                    Reject
                  </a>
                  <a href="#" onClick={handleExit} className="reset-btn">
                    Exit
                  </a>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>

    </div >
  );
};

export default ApproverAdvanceForm;
