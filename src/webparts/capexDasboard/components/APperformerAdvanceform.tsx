import * as React from "react";
import "./advanced.scss";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { useEffect, useState } from "react";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import logo from "../assets/sona-comstarlogo.png";

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
const APperformerAdvanceform: React.FC<IProps> = ({ context, itemId }) => {
  const sp = spfi().using(SPFx(context));
  const [previousAdvances, setPreviousAdvances] = useState<any[]>([]);
  const [attachments, setAttachments] = useState<any[]>([]);
  const today = new Date();

const localDate: string = new Date(
  today.getTime() - today.getTimezoneOffset() * 60000
)
  .toISOString()
  .split("T")[0];
  const actionLock = React.useRef(false);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [itemData, setItemData] = useState<any>(null);
  const [approverRemarks, setApproverRemarks] = useState("");
  const [voucherDate, setVoucherDate] = useState("");
  const [voucherNumber, setVoucherNumber] = useState("");
  const [selectedVendorName, setSelectedVendorName] = useState("");
  const [vendors, setVendors] = useState<IVendor[]>([]);
  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);
  const [workflowHistory, setWorkflowHistory] = useState<any[]>([]);
  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: context.pageContext.web.absoluteUrl,
    msGraphClientFactory: context.msGraphClientFactory,
    spHttpClient: context.spHttpClient,
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
  // ✅ Fetch Item by ID
  const getItemById = async () => {
    try {
      const item = await sp.web.lists
        .getByTitle("CapexAdvance")
        .items.getById(itemId)
        .select("*", "PICName/Title", "VendorCode/Id", "VendorCode/VendorCode", "Author/Id","Author/Title","Author/EMail")
        .expand("PICName", "VendorCode", "Author")();
      // 👈 ADD

      setItemData(item);
      //  setApproverRemarks(item.ApproverRemarks || "");

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
      await getItemById(); // 👈 FIRST load item to get VendorCode
      await getVendors(); // 👈 FIRST load vendors
     // await getItemById(); // 👈 THEN item
      if (selectedVendorId) {
      void getPreviousAdvances(selectedVendorId);
    }
    };

    void loadData();
  }, [context, itemId]);

  const buildApprovalFlow = async () => {

    const existingFlow = itemData.ApprovalMatrix
      ? JSON.parse(itemData.ApprovalMatrix)
      : [];

    const baseApprovers = existingFlow.filter(
      (a: any) => a.Role === "RM" || a.Role === "HOD"
    );

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

    const fullFlow = [...baseApprovers, ...matrixApprovers];

    return fullFlow;
  };
  const mergeFlowWithStatus = (oldFlow: any[], newFlow: any[]) => {
    return newFlow.map((newItem) => {
      const oldItem = oldFlow.find((o) => o.Id === newItem.Id);

      return {
        ...newItem,
        Status: oldItem?.Status || newItem.Status // 🔥 preserve status
      };
    });
  };
  // ✅ Approve
  const handleApprove = async () => {
    if (actionLock.current) return;

  actionLock.current = true;
    if (isSubmitting) return;
    try {
      setIsSubmitting(true);
      if (!voucherDate || voucherDate.trim() === "") {
      alert("Please enter Voucher Date");
      actionLock.current = false;
        setIsSubmitting(false);
      return;
    }
      if (voucherDate > localDate) {
      alert("VoucherDate cannot be a future date");
      // return;
       actionLock.current = false;
        setIsSubmitting(false);
        return;
    }
    if (!voucherNumber || voucherNumber.trim() === "") {
      alert("Please enter Voucher Number");
      actionLock.current = false;
        setIsSubmitting(false);
      return;
    }
     if (!approverRemarks || approverRemarks.trim() === "") {
      alert("Please enter Remarks");
      actionLock.current = false;
        setIsSubmitting(false);
      return;
    }

      // =========================
      // 🔹 OLD FLOW
      // =========================
      const oldFlow = itemData.ApprovalMatrix
        ? JSON.parse(itemData.ApprovalMatrix)
        : [];

      // =========================
      // 🔥 REBUILD NEW FLOW
      // =========================
      const latestFlow = await buildApprovalFlow();

      // =========================
      // 🔥 PRESERVE STATUS
      // =========================
      const finalFlow = mergeFlowWithStatus(oldFlow, latestFlow);

      // =========================
      // 🔥 CURRENT USER
      // =========================
      const currentUserId = context.pageContext.legacyPageContext.userId;

      const currentIndex = finalFlow.findIndex(
        (a: any) => a.Id === currentUserId
      );

      if (currentIndex === -1) {
        alert("You are not current approver");
        return;
      }

      // =========================
      // 🔥 MARK CURRENT APPROVED
      // =========================
      finalFlow[currentIndex].Status = "Approved";

      // =========================
      // 🔥 FIND NEXT APPROVER
      // =========================
      let nextApproverId = null;

      if (finalFlow[currentIndex + 1]) {
        finalFlow[currentIndex + 1].Status = "In Progress";
        nextApproverId = finalFlow[currentIndex + 1].Id;
      } else {
        // 🔥 NO NEXT → KEEP SAME (or null if final)
        nextApproverId = finalFlow[currentIndex].Id;
      }

      // =========================
      // 🔥 UPDATE SHAREPOINT
      // =========================
      await sp.web.lists
        .getByTitle("CapexAdvance")
        .items.getById(itemId)
        .update({
          ApproverRemarks: approverRemarks,

          VoucherDate: voucherDate ? new Date(voucherDate) : null,
          VouchingNumber: voucherNumber,

          Status: "Pending for PF Approver UTR",

          ApprovalMatrix: JSON.stringify(finalFlow),

          // 🔥 IMPORTANT
          CurrentApproverId: nextApproverId,
          ApproverStatus: "Pending for PF Approver UTR"
        });

      alert("Approved successfully ✅");
      actionLock.current = false;
      setIsSubmitting(false);

      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexForm.aspx?page=Performer";

    } catch (error) {
      console.error("Approve error:", error);
      alert("Error ❌");
      actionLock.current = false;
        setIsSubmitting(false);
    }
  };

  // ✅ Sent Back
  const handleSendBack = async () => {
    if (actionLock.current) return;
    actionLock.current = true;
    setIsSubmitting(true);

    try {
     if (!approverRemarks || approverRemarks.trim() === "") {
      alert("Please enter Remarks");
      actionLock.current = false;
      setIsSubmitting(false);
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

          CurrentApproverId: null
        });

      alert("Send Back ✅");
      actionLock.current = false;
      setIsSubmitting(false);
      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexForm.aspx?page=Performer";
    } catch (error) {
      console.error(error);
      actionLock.current = false;
      setIsSubmitting(false);
    }
  };

  // ✅ Reject
  const handleReject = async () => {
    if (actionLock.current) return;
    actionLock.current = true;
    setIsSubmitting(true);

    try {
      if (!approverRemarks || approverRemarks.trim() === "") {
      alert("Please enter Remarks");
      actionLock.current = false;
      setIsSubmitting(false);
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
      actionLock.current = false;
      setIsSubmitting(false);
      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexForm.aspx?page=Performer";

    } catch (error) {
      console.error(error);
      actionLock.current = false;
      setIsSubmitting(false);
    }
  };

  const handleExit = () => {
    window.location.href = `https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexForm.aspx?page=Performer`;
  };

  // ⛔ Wait until data loads
  if (!itemData) return <div>Loading...</div>;

  return (

    <>
      <div className='MainUplodForm' style={{ margin: "5px 0px" }}>
        <div className='row'>
          <div className='col-md-12'>
            <div className='Main-Boxpoup'>
              <div className="bordered">
                <img src={logo} />
                <h1>Capex Advance Payment (Approver) </h1>
              </div>
              {approvalMatrix.length === 0 ? (
                <p>No approval data</p>
              ) : (
                <div className="displayWF">
                  <ul className="approval-flow">
                     <li className={`approval-step`}>
                      
                           {`Initiator`} - {itemData.Author.Title}
                        
                    </li>
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
                  <label>Capex Details</label>
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
                      <input value={itemData.PODate ? new Date(itemData.PODate).toLocaleDateString("en-GB") : ""} className="form-control readonly" />
                    </div>
                    <div className="col-md-4">
                      <label className="font">PO Terms</label>
                      <input value={itemData.POAdvanceTerms || ""} className="form-control readonly" />
                    </div>
                    <div className="col-md-4">
                      <label className="font">PO Amount</label>
                      <input value={itemData.POAmtGST || ""} className="form-control readonly" />
                    </div>
                  </div>
                  <div className='row mb-20'>
                    <div className='col-md-4'>
                      <label className="font">Advance Amount</label>
                      <input value={itemData.RequestAdvanceAmount || ""} className="form-control readonly" />
                    </div>
                    <div className='col-md-4'>
                      <label className="font">Paid Amount</label>
                      <input value={itemData.PaidAmount || ""} className="form-control readonly" />
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
                        className="form-control readonly" />
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
                      <input value={itemData.GL || ""} className="form-control readonly" />
                    </div>
                  </div>
                  <div className="row mb-20">
                    <div className="col-md-4">
                      <label className="font">Cost Center</label>
                      <input value={itemData.CostCenter || ""} className="form-control readonly" />
                    </div>
                  </div>
                </div>
                <div className="heading1" style={{ marginTop: "10px" }}>
                  <label>Remarks</label>
                </div>
                <div className='main-formcontainer'>
                  <div className="row mb-20">
                    <div className="col-md-4">
                      <label className="font" style={{ display: "block" }}>User Remarks</label>
                      <label className='fonttext textbox readonly'>{itemData.Remarks || ""}</label>
                    </div>
                    <div className="col-md-4">
                      <label className="font" style={{ display: "block" }}>Project Description</label>
                      <label className='fonttext textbox readonly'>{itemData.ProjectDescription || ""}</label>
                    </div>
                  </div>
                </div>
                <div className="heading1" style={{ marginTop: "10px" }}>
                  <label>Workflow History</label>
                </div>
                <div className='main-formcontainer'>
                  <div className="row mb-20">
                    <div className="col-md-12">
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
                                {h.ActionTaken === "Vouched" && "💰 "}
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
                </div>
                <div className='main-formcontainer'>
                  <div className="row mb-20">
                    <div className="col-md-4">
                      <label className='font'>Voucher Date</label>
                      <input
                        type="date"
                        value={voucherDate}
                        onChange={(e) => setVoucherDate(e.target.value)} className='form-control'
                      />
                    </div>
                    <div className="col-md-4">
                      <label className='font'>Voucher Number</label>
                      <input
                        value={voucherNumber} className='form-control'
                        onChange={(e) => setVoucherNumber(e.target.value)}
                      />
                    </div>
                  </div>
                </div>
                <div className="heading1" style={{ marginTop: "10px" }}>
                  <label>Attachments</label>
                </div>
                <div className='main-formcontainer'>
                  <div className="row mb-20">
                    <div className="col-md-12">
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
                </div>
                <div className="heading1" style={{ marginTop: "10px" }}>
                  <label>Approver Action</label>
                </div>
                <div className='main-formcontainer'>
                  <div className="row mb-20">
                    <div className="col-md-4">
                      <label className='font'>Approver Remarks</label>
                      <textarea
                        value={approverRemarks} className='form-control'
                        onChange={(e) => setApproverRemarks(e.target.value)}
                      />
                    </div>
                  </div>
                </div>
                <div style={{ display: "flex", justifyContent: "center", gap: "5px", marginBottom: "1rem", marginTop: "1rem" }}>
                  <a onClick={handleApprove} className="submit-btn">
                    Vouch
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
                  <label>Capex Details</label>
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
                      <input value={itemData.PODate ? new Date(itemData.PODate).toLocaleDateString("en-GB") : ""} className="form-control readonly" />
                    </div>
                    <div className="col-md-4">
                      <label className="font">PO Terms</label>
                      <input value={itemData.POAdvanceTerms || ""} className="form-control readonly" />
                    </div>
                    <div className="col-md-4">
                      <label className="font">PO Amount</label>
                      <input value={itemData.POAmtGST || ""} className="form-control readonly" />
                    </div>
                  </div>
                  <div className='row mb-20'>
                    <div className='col-md-4'>
                      <label className="font">Advance Amount</label>
                      <input value={itemData.RequestAdvanceAmount || ""} className="form-control readonly" />
                    </div>
                    <div className='col-md-4'>
                      <label className="font">Paid Amount</label>
                      <input value={itemData.PaidAmount || ""} className="form-control readonly" />
                    </div>
                    <div className="col-md-4">
                      <label className="font">Expected Settlement</label>
                      <input value={itemData.ExpectedDateofSettlement ? new Date(itemData.ExpectedDateofSettlement,).toLocaleDateString("en-GB") : ""}
                        className="form-control readonly" />
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
                      <input value={itemData.GL || ""} className="form-control readonly" />
                    </div>
                    <div className="col-md-4">
                      <label className="font">Cost Center</label>
                      <input value={itemData.CostCenter || ""} className="form-control readonly" />
                    </div>
                  </div>
                  <div className="row mb-20">
                    <div className="col-md-12">
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
                                {h.ActionTaken === "Vouched" && "💰 "}
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
                  <div className="row mb-20">
                    <div className="col-md-4">
                      <label className="font" style={{ display: "block" }}>User Remarks</label>
                      <label className='fonttext textbox readonly'>{itemData.Remarks || ""}</label>
                    </div>
                    <div className="col-md-4">
                      <label className="font" style={{ display: "block" }}>Project Description</label>
                      <label className='fonttext textbox readonly'>{itemData.ProjectDescription || ""}</label>
                    </div>
                    <div className="col-md-4">
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
                  <div className="row mb-20">
                    <div className="col-md-4">
                      <label className='font'>Voucher Date</label>
                      <input
                        type="date"
                        value={voucherDate}
                        onChange={(e) => setVoucherDate(e.target.value)}
                        max={new Date().toISOString().split("T")[0]}  className='form-control'
                      />
                    </div>
                    <div className="col-md-4">
                      <label className='font'>Voucher Number</label>
                      <input
                        value={voucherNumber} className='form-control'
                        onChange={(e) => setVoucherNumber(e.target.value)}
                      />
                    </div>
                    <div className="col-md-4">
                      <label className='font'>Approver Remarks</label>
                      <textarea
                        value={approverRemarks} className='form-control'
                        onChange={(e) => setApproverRemarks(e.target.value)}
                      />
                    </div>

                  </div>
                   <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Previous Advances</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-12">
                    <div style={{ overflowX: "auto" }}>
                      <div className="table-vert-scroll">
                        <table className="custom-table min-w-full bg-white rounded-2xl shadow-md">
                          <thead
                            className="text-white"
                            style={{ backgroundColor: "rgb(60, 62, 69)" }}
                          >
                            <tr>
                              <th className="px-4 py-2">PO Number</th>
                              <th className="px-4 py-2">Previous Advance</th>
                              <th className="px-4 py-2">Requested Date</th>
                              <th className="px-4 py-2">Paid Date</th>
                              <th className="px-4 py-2">MRN No</th>
                              <th className="px-4 py-2">Settled Amount</th>
                              <th className="px-4 py-2">Pending Advance</th>
                            </tr>
                          </thead>
                          <tbody>
                            {previousAdvances.length === 0 ? (
                              <tr>
                                <td colSpan={7} style={{ textAlign: "center" }}>
                                  No previous advances available
                                </td>
                              </tr>
                            ) : (
                              previousAdvances.map(
                                (item: any, index: number) => {
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
                                          ? new Date(
                                              item.Created,
                                            ).toLocaleDateString("en-GB")
                                          : ""}
                                      </td>

                                      <td>
                                          {item.VoucherDate
                                            ? new Date(
                                                item.VoucherDate,
                                              ).toLocaleDateString('en-GB')
                                            : ""}
                                        </td>

                                      <td>{item.VoucherNumber}</td>
                                      <td>{item.PaidAmount}</td>
                                      <td>{pending}</td>
                                    </tr>
                                  );
                                },
                              )
                            )}
                          </tbody>
                        </table>
                         
                      </div>
                    </div>
                  </div>
                </div>
              </div>
                  <div style={{ display: "flex", justifyContent: "center", gap: "5px", marginBottom: "1rem", marginTop: "1rem" }}>
                  <a
                      className={`submit-btn ${isSubmitting ? "disabled-btn" : ""}`}
                      onClick={!isSubmitting ? handleApprove : undefined}
                    >
                      {isSubmitting ? "Processing..." : "Approve"}
                    </a>

                    <a
                      className={`Rework-btn ${isSubmitting ? "disabled-btn" : ""}`}
                      onClick={!isSubmitting ? handleSendBack : undefined}
                    >
                      {isSubmitting ? "Processing..." : "Send Back"}
                    </a>

                    <a
                      className={`Reject-btn ${isSubmitting ? "disabled-btn" : ""}`}
                      onClick={!isSubmitting ? handleReject : undefined}
                    >
                      {isSubmitting ? "Processing..." : "Reject"}
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
      </div>

    </>

  );
};

export default APperformerAdvanceform;
