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
  itemId: number; 
}

interface IVendor {
  Id: number;
  VendorCode: string;
  VendorName: string;
}
const ApproverAdvanceForm: React.FC<IProps> = ({ context, itemId }) => {
  const sp = spfi().using(SPFx(context));
  const actionLock = React.useRef(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [previousAdvances, setPreviousAdvances] = useState<any[]>([]);
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
    debugger;
    try {
      if (!capexId) return;

      const safeCapexId = capexId.replace(/\//g, "_");

      
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

      const safeCapexId = capexId.replace(/\//g, "_");

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
 const getItemById = async () => {
    try {
      debugger;

      const item = await sp.web.lists
        .getByTitle("CapexAdvance")
        .items.getById(itemId)
        .select(
          "*",
          "PICName/Title",
          "VendorCode/Id",
          "VendorCode/VendorCode",
          "Author/Id",
          "Author/Title",
          "Author/EMail",
        )
        .expand("PICName", "VendorCode", "Author")();

      console.log("ITEM DATA:", item);

      setItemData(item);

     
      const vendorId = item?.VendorCode?.Id || null;

      console.log("Vendor Id:", vendorId);

      setSelectedVendorId(vendorId);

      
      setSelectedVendorName(item?.VendorName || "");

     
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
  
  const getItemById1 = async () => {
    try {
      const item = await sp.web.lists
        .getByTitle("CapexAdvance")
        .items.getById(itemId)
        .select(
          "*",
          "PICName/Title",
          "VendorCode/Id",
          "VendorCode/VendorCode",
          "Author/Id",
          "Author/Title",
          "Author/EMail",
        )
        .expand("PICName", "VendorCode", "Author")();
      

      setItemData(item);
      //setApproverRemarks(item.ApproverRemarks || "");

     
      setSelectedVendorId(item.VendorCode?.Id || null);
      
      setSelectedVendorName(item.VendorName); 

      
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
      debugger;

      //await getLoggedInUser();

    
      await getVendors();

      
      await getItemById();
    };

    void loadData();
  }, [context, itemId]);

  useEffect(() => {
    if (selectedVendorId) {
      console.log("Calling Previous Advances:", selectedVendorId);

      void getPreviousAdvances(selectedVendorId);
    }
  }, [selectedVendorId]);


 
 
  const buildApprovalFlow = async () => {
    try {
      
      const existingFlow = itemData.ApprovalMatrix
        ? JSON.parse(itemData.ApprovalMatrix)
        : [];

      
      const baseApprovers = existingFlow.filter(
        (a: any) => a.Role === "RM" || a.Role === "HOD",
      );

      
      const matrixData = await sp.web.lists
        .getByTitle("CapexApprovalMatrix")
        .items.select("Role/RoleName,Approver/Id,Approver/Title,Level/Level")
        .expand("Approver,Role,Level")
        .filter("Status eq 'Active'")
        .orderBy("Level", true)();

      const matrixApprovers = matrixData.map((item: any, index: number) => ({
        Id: item.Approver?.Id,
        Name: item.Approver?.Title,
        Role: item.Role?.RoleName,
        Level: baseApprovers.length + index + 1,
        Status: "Pending",
      }));

      
      const fullFlow = [...baseApprovers, ...matrixApprovers];

      
      const uniqueFlow = fullFlow.filter(
        (v, i, self) => self.findIndex((x) => x.Id === v.Id) === i,
      );

      return uniqueFlow;
    } catch (error) {
      console.error("Approval Flow Error:", error);
      return [];
    }
  };

  const handleApprove = async () => {
    if (actionLock.current) return;

    //actionLock.current = true;
    try {
     if (!approverRemarks || approverRemarks.trim() === "") {
      alert("Please enter Remarks");
      //actionLock.current = false;
      return;
    }
    if (isProcessing) return;

      setIsProcessing(true);
      
      let existingFlow = itemData.ApprovalMatrix
        ? JSON.parse(itemData.ApprovalMatrix)
        : [];

     
      const latestFlow = await buildApprovalFlow();

      
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

      if (isChanged) {
        existingFlow = latestFlow;

        // reset all
        existingFlow.forEach((a: any) => (a.Status = "Pending"));

        if (existingFlow.length > 0) {
          existingFlow[0].Status = "In Progress";
        }
      }

     
      const userEmail = context.pageContext.user.email;

      //const ensuredUser = await sp.web.ensureUser(userEmail);

      const currentUserId = context.pageContext.legacyPageContext.userId;
      const currentIndex = existingFlow.findIndex(
        (a: any) => a.Id === currentUserId,
      );

      if (currentIndex === -1) {
        alert("You are not the current approver");
        return;
      }

     
      existingFlow[currentIndex].Status = "Approved";

      
      let nextApproverId = null;

      if (existingFlow[currentIndex + 1]) {
        existingFlow[currentIndex + 1].Status = "In Progress";
        nextApproverId = existingFlow[currentIndex + 1].Id;
      }

      
      let finalStatus = itemData.Status;

      const currentRole = existingFlow[currentIndex]?.Role;

      if (currentRole === "RM") {
        finalStatus = "Pending for Approver";
      } else if (currentRole === "HOD") {
        finalStatus = "Pending for PF Approver";
      } else {
        finalStatus = nextApproverId ? "Pending" : "Approved";
      }

      
      const history = itemData.WorkFlowHistory
        ? JSON.parse(itemData.WorkFlowHistory)
        : [];

      history.push({
        CurrentApprover: context.pageContext.user.displayName,
        ActionTaken: "Approved",
        Comment: approverRemarks,
        Date: new Date().toISOString(),
      });

      let ApproverStatus;

      if (currentRole === "RM") {
        ApproverStatus = "Pending for HOD";
      } else if (currentRole === "HOD") {
        ApproverStatus = "Pending for PF Approver";
      } else {
        ApproverStatus = nextApproverId ? "Pending" : "Approved";
      }

      await sp.web.lists
        .getByTitle("CapexAdvance")
        .items.getById(itemId)
        .update({
          ApproverRemarks: approverRemarks,

          Status: finalStatus, 

          ApprovalMatrix: JSON.stringify(existingFlow),

          CurrentApproverId: nextApproverId,

          WorkFlowHistory: JSON.stringify(history),
          ApproverStatus: ApproverStatus,
        });

      alert("Approved successfully ");

      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexForm.aspx?page=Approver";
    } catch (error) {
      console.error("Approve error:", error);
   
      alert("Error occurred while approving the request.");
    }
    finally {
    //actionLock.current = false;
      setIsProcessing(false);
  }
  };

 
  const handleSendBack = async () => {
    if (actionLock.current) return;
   // actionLock.current = true;
    setIsProcessing(true);
    try {
      if (!approverRemarks || approverRemarks.trim() === "") {
        alert("Please enter Remarks");
          //actionLock.current = false;
          setIsProcessing(false);
        return;
      }

      const flow = itemData.ApprovalMatrix
        ? JSON.parse(itemData.ApprovalMatrix)
        : [];

      const currentUserId = context.pageContext.legacyPageContext.userId;

      const currentIndex = flow.findIndex((a: any) => a.Id === currentUserId);

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
        Date: new Date().toISOString(),
      });

      await sp.web.lists
        .getByTitle("CapexAdvance")
        .items.getById(itemId)
        .update({
          ApproverRemarks: approverRemarks,
          Status: "Send Back",

          ApprovalMatrix: JSON.stringify(flow),
          WorkFlowHistory: JSON.stringify(history),

          CurrentApproverId: itemData.CurrentApproverId,
        });

      alert("Send Back ✅");
      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexForm.aspx?page=Approver";
    } catch (error) {
      console.error(error);
      
    }
    finally {
    //actionLock.current = false;
      setIsProcessing(false);
  }
  };

  
  const handleReject = async () => {
    if (actionLock.current) return;
   // actionLock.current = true;
    setIsProcessing(true);
    try {
      if (!approverRemarks || approverRemarks.trim() === "") {
        alert("Please enter Remarks");
        //actionLock.current = false;
        setIsProcessing(false);
        return;
      }

      const flow = itemData.ApprovalMatrix
        ? JSON.parse(itemData.ApprovalMatrix)
        : [];

      const currentUserId = context.pageContext.legacyPageContext.userId;

      const currentIndex = flow.findIndex((a: any) => a.Id === currentUserId);

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
        Date: new Date().toISOString(),
      });

      await sp.web.lists
        .getByTitle("CapexAdvance")
        .items.getById(itemId)
        .update({
          ApproverRemarks: approverRemarks,
          Status: "Rejected",

          ApprovalMatrix: JSON.stringify(flow),
          WorkFlowHistory: JSON.stringify(history),

          CurrentApproverId: null,
        });

      alert("Rejected ❌");
      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexForm.aspx?page=Approver";
    } catch (error) {
      console.error(error);
      
    }
    finally {
    //actionLock.current = false;
      setIsProcessing(false);
  }
  };

  const handleExit = () => {
    window.location.href = `https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexForm.aspx?page=Approver`;
  };

  
  if (!itemData) return <div>Loading...</div>;

  return (
    <div className="MainUplodForm" style={{ margin: "5px 0px" }}>
      <div className="row">
        <div className="col-md-12">
          <div className="Main-Boxpoup">
            <div className="bordered">
              <img src={logo} />
              <h1>  Approval Form </h1>
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
                      className={`approval-step ${
                        a.Status === "In Progress"
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
            <div className="borderedbox">
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Requestor Information</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Employee Code" className="font">
                      Employee Code
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.EmployeeCode}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Name" className="font">
                      Employee Name{" "}
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.EmployeeName}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Email" className="font">
                      Employee Email{" "}
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext">
                      {" "}
                      {itemData.EmployeeEmail}
                    </label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Contact No" className="font">
                      Contact No
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.ContactNo}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Status" className="font">
                      Employee Status
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext">
                      {" "}
                      {itemData.EmployeeStatus}
                    </label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Division" className="font">
                      Division
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.Division}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Location" className="font">
                      Location
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.Location}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="RM" className="font">
                      RM
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.RM}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="HOD" className="font">
                      HOD
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.HOD}</label>
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Capex Details</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
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
                      }}
                      className="formtext-control"
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
                    <input
                      value={itemData.VendorName || ""}
                      className="form-control readonly"
                    />
                  </div>
                  <div className="col-md-4">
                    <label>PO Number</label>
                    <input
                      value={itemData.PONumber || ""}
                      className="form-control readonly"
                    />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">PO Date</label>
                    <input
                      value={
                        itemData.PODate
                          ? new Date(itemData.PODate).toLocaleDateString(
                              "en-GB",
                            )
                          : ""
                      }
                      className="font-control readonly"
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Terms</label>
                    <textarea
    value={itemData.POAdvanceTerms || ""}
    className="form-control readonly"
    rows={3}
    readOnly
  />
                    {/* <input
                      value={itemData.POAdvanceTerms || ""}
                      className="font-control readonly"
                    /> */}
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Amount</label>
                    <input
                      value={itemData.POAmtGST || ""}
                      className="font-control readonly"
                    />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Advance Amount</label>
                    <input
                      value={itemData.RequestAdvanceAmount || ""}
                      className="font-control readonly"
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Paid Amount</label>
                    <input
                      value={itemData.PaidAmount || ""}
                      className="font-control readonly"
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Expected Settlement</label>
                    <input
                      value={
                        itemData.ExpectedDateofSettlement
                          ? new Date(
                              itemData.ExpectedDateofSettlement,
                            ).toLocaleDateString("en-GB")
                          : ""
                      }
                      className="font-control readonly"
                    />
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
                    <input
                      value={itemData.GL || ""}
                      className="font-control readonly"
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Cost Center</label>
                    <input
                      value={itemData.CostCenter || ""}
                      className="font-control readonly"
                    />
                  </div>
                </div>
               
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">User Remarks</label>
                    <textarea
                      value={itemData.Remarks || ""}
                      className="font-control readonly"
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Project Description</label>
                    <textarea
                      value={itemData.ProjectDescription || ""}
                      className="font-control readonly"
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Approver Remarks</label>
                    <textarea
                      className="font-control"
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
                            })
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              </div>
            </div>
                <div className="row mb-20">
                    <div className="col-md-12">
                      {workflowHistory.length === 0 ? (
                        <p>No history available</p>
                      ) : (
                        <div className="workflow-history">
                          <table
                          className="workflow-table"
                          style={{ width: "100%" }}
                        >
                          <thead>
                            <tr>
                              <th style={{ padding: "8px", textAlign: "left" }}>
                                Action By
                              </th>
                              <th style={{ padding: "8px", textAlign: "left" }}>
                                Action Taken
                              </th>
                              <th style={{ padding: "8px", textAlign: "left" }}>
                                Date
                              </th>
                              <th style={{ padding: "8px", textAlign: "left" }}>
                                Comment
                              </th>
                            </tr>
                          </thead>
                          <tbody>
                            {workflowHistory
                              .filter(
                                (h: any) =>
                                  h.ActionTaken &&
                                  h.ActionTaken !== "Draft Saved" && h.ActionTaken !== "Edited",
                              )
                              .map((h: any, idx: number) => (
                                <tr key={idx}>
                                  <td style={{ padding: "8px" }}>
                                    {h.CurrentApprover || ""}
                                  </td>

                                  <td style={{ padding: "8px" }}>
                                    {h.ActionTaken || ""}
                                  </td>

                                  <td style={{ padding: "8px" }}>
                                    {h.Date
                                      ? new Date(h.Date).toLocaleDateString(
                                          "en-GB",
                                        )
                                      : ""}
                                  </td>

                                  <td style={{ padding: "8px" }}>
                                    {h.Comment || ""}
                                  </td>
                                </tr>
                              ))}
                          </tbody>
                        </table>
                        
                        </div>
                      )}
                    </div>
                  </div>
               
                <div
                  style={{
                    display: "flex",
                    justifyContent: "center",
                    gap: "5px",
                    marginBottom: "1rem",
                    marginTop: "1rem",
                  }}
                >
                  <a
                      className={`submit-btn ${isProcessing ? "disabled-btn" : ""}`}
                      onClick={!isProcessing ? handleApprove : undefined}
                    >
                      {isProcessing ? "Processing..." : "Approve"}
                    </a>

                    <a
                      className={`Rework-btn ${isProcessing ? "disabled-btn" : ""}`}
                      onClick={!isProcessing ? handleSendBack : undefined}
                    >
                      {isProcessing ? "Processing..." : "Send Back"}
                    </a>

                    <a
                      className={`Reject-btn ${isProcessing ? "disabled-btn" : ""}`}
                      onClick={!isProcessing ? handleReject : undefined}
                    >
                      {isProcessing ? "Processing..." : "Reject"}
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
  );
};

export default ApproverAdvanceForm;
