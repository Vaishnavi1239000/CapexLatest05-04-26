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
import logo from "../assets/sona-comstarlogo.png";
const ViewAdvanceForm = ({ context, formData, onClose }: any) => {
  const [attachments, setAttachments] = useState<any[]>([]);
  const sp = spfi().using(SPFx(context));
  const [employee, setEmployee] = useState<any>({});
  const actionLock = React.useRef(false);
  const [isSubmitting, setIsSubmitting] = useState(false);

  const [previousAdvances, setPreviousAdvances] = useState<any[]>([]);
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [selectedUser, setSelectedUser] = useState<any[]>([]);
  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  const [selectedVendorName, setSelectedVendorName] = useState("");

  const [vendors, setVendors] = useState<IVendor[]>([]);
  const [employeeName, setEmployeeName] = useState("");
  const [employeeEmail, setEmployeeEmail] = useState("");
  const [VendorCode, setVendorCode] = useState<number | null>(null);
  
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

  
  const getVendors = async () => {
    const data = await sp.web.lists
      .getByTitle("VendorMaster")
      .items.select("Id", "VendorCode", "VendorName")();
    void setVendors(data);
  };

 
  useEffect(() => {
    debugger;
    if (!formData) return;

    setPoNumber(formData.PONumber || "");
    setPoDate(formData.PODate?.split("T")[0] || "");
    setPoTerms(formData.POAdvanceTerms || "");
    setPoAmount(formData.POAmtGST || "");
    setAdvanceAmount(formData.RequestAdvanceAmount || "");
    setPaidAmount(formData.PaidAmount || "");
    setExpectedDate(formData.ExpectedDateofSettlement?.split("T")[0] || "");

    setVendorName(formData.VendorName || "");
    setSelectedVendorId(formData.VendorCodeId || null); 

     if (formData.VendorCodeId) {
      void getPreviousAdvances(formData.VendorCodeId);
    }
    setSelectedVendorName(formData.VendorName || ""); 

    setGlCode(formData.GL || "");
    setCostCenter(formData.CostCenter || "");
    setRemarks(formData.Remarks || "");
    setProjectDesc(formData.ProjectDescription || "");

    setApproverRemarks(formData.ApproverRemarks || "");
    setVoucherDate(formData.VoucherDate?.split("T")[0] || "");
    setVouchingNumber(formData.VouchingNumber || "");
    setUTRDate(formData.UTRDate?.split("T")[0] || "");
    setUTRNumber(formData.UTRNumber || "");

   
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
  const getEmployeeDetails = async () => {
    try {
      debugger;

      if (!formData?.Email) return;

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
        .filter(`EmployeeEmail eq '${formData.Email}'`)
        .top(1)();

      if (user.length > 0) {
        setEmployee(user[0]);
      }
    } catch (error) {
      console.log("Error fetching employee:", error);
    }
  };
  useEffect(() => {
   // void getLoggedInUser();
    void getEmployeeDetails();
    void getVendors();
     if (selectedVendorId) {
      void getPreviousAdvances(selectedVendorId);
    }
  }, []);
  return (
    <>

      <div className='MainUplodForm' style={{ margin: "5px 0px" }}>
        <div className='row'>
          <div className='col-md-12'>
            <div className='Main-Boxpoup'>
            
              <div className="bordered">
                <img src={logo} />
                <h1> Capex Advance Payment (View) </h1>
              </div>
              {approvalMatrix.length === 0 ? (
                <p>No approval data</p>
              ) : (

                <div className="displayWF">
                  <ul className="approval-flow">
                    <li className={`approval-step`}>
                      
                           {`Initiator`} - {employee.EmployeeName}
                        
                    </li>
                    {approvalMatrix.map((a, index) => (
                      <li
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
                        {a.Role} - {a.Name}
                      </li>
                    ))}
                  </ul>
                </div>
              )}
              {/* <div className='borderedbox'>
                <div className="heading1">
                  <label>Requestor Information</label>
                </div>
                <div className='main-formcontainer'>
                  <div className='row mb-20'>
                    <div className='col-md-4'>
                      <label htmlFor="Employee Code" className='font'>Employee Code</label> : &nbsp;&nbsp;
                      <label className='fonttext'>  {employee.EmployeeCode}</label>
                    </div>
                    <div className='col-md-4'>
                      <label htmlFor="Employee Name" className='font'>Employee Name </label> : &nbsp;&nbsp;
                      <label className='fonttext'>  {employee.EmployeeName}</label>
                    </div>
                    <div className='col-md-4'>
                      <label htmlFor="Employee Email" className='font'>Employee Email </label> : &nbsp;&nbsp;
                      <label className='fonttext'>  {employee.EmployeeEmail}</label>
                    </div>
                  </div>
                  <div className='row mb-20'>
                    <div className='col-md-4'>
                      <label htmlFor="Contact No" className='font'>Contact No</label> : &nbsp;&nbsp;
                      <label className='fonttext'>  {employee.ContactNo}</label>
                    </div>
                    <div className='col-md-4'>
                      <label htmlFor="Employee Status" className='font'>Employee Status</label> : &nbsp;&nbsp;
                      <label className='fonttext'>  {employee.EmployeeStatus}</label>
                    </div>
                    <div className='col-md-4'>
                      <label htmlFor="Division" className='font'>Division</label> : &nbsp;&nbsp;
                      <label className='fonttext'>  {employee.Division}</label>
                    </div>
                  </div>
                  <div className='row mb-20'>
                    <div className='col-md-4'>
                      <label htmlFor="Location" className='font'>Location</label> : &nbsp;&nbsp;
                      <label className='fonttext'>  {employee.Location}</label>
                    </div>
                    <div className='col-md-4'>
                      <label htmlFor="RM" className='font'>RM</label> : &nbsp;&nbsp;
                      <label className='fonttext'>  {employee.ReportingManager?.Title}</label>
                    </div>
                    <div className='col-md-4'>
                      <label htmlFor="HOD" className='font'>HOD</label> : &nbsp;&nbsp;
                      <label className='fonttext'>  {employee.HOD?.Title}</label>
                    </div>
                  </div>
                </div>
                <div className="heading1" style={{ marginTop: "10px" }}>
                  <label>Vendor & PO</label>
                </div>
                <div className='main-formcontainer'>
                  <div className='row mb-20'>
                    <div className='col-md-4'>
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
                      <label className="font">Vendor Name</label>
                      <input className="form-control readonly" value={vendorName} />
                    </div>
                    <div className="col-md-4">
                      <label className="font">PO Number</label>
                      <input className="form-control readonly" value={poNumber} />
                    </div>
                  </div>
                  <div className='row mb-20'>
                    <div className='col-md-4'>
                      <label className="font">PO Date</label>
                      <input type="date" className="form-control readonly" value={poDate} />
                    </div>
                    <div className='col-md-4'>
                      <label className="font">PO Terms</label>
                      <input className="form-control readonly" value={poTerms} />
                    </div>
                  </div>
                </div>
                <div className="heading1" style={{ marginTop: "10px" }}>
                  <label>Amount</label>
                </div>
                <div className='main-formcontainer'>
                  <div className='row mb-20'>
                    <div className="col-md-4">
                      <label className="font">PO Amount</label>
                      <input className="form-control readonly" value={poAmount} />
                    </div>
                    <div className="col-md-4">
                      <label className="font">Advance Amount</label>
                      <input className="form-control readonly" value={advanceAmount} />
                    </div>
                    <div className="col-md-4">
                      <label className="font">Paid Amount</label>
                      <input className="form-control readonly" value={paidAmount} />
                    </div>
                  </div>
                </div>
                <div className="heading1" style={{ marginTop: "10px" }}>
                  <label>Advance Details</label>
                </div>
                <div className='main-formcontainer'>
                  <div className='row mb-20'>
                    <div className="col-md-4">
                      <label className="font">Expected Date</label>
                      <input type="date" className="form-control readonly" value={expectedDate} />
                    </div>
                    <div className="col-md-4">
                      <label className="font">PIC Name</label>
                      <PeoplePicker
                        context={peoplePickerContext}
                        personSelectionLimit={1}
                        disabled={true}
                        principalTypes={[PrincipalType.User]}
                        defaultSelectedUsers={
                          formData?.PICName?.Title ? [formData.PICName.Title] : []
                        }
                      />
                    </div>
                    <div className="col-md-4">
                      <label className="font">GL Code</label>
                      <input className="form-control readonly" value={glCode} />
                    </div>
                  </div>
                  <div className='row mb-20'>
                    <div className="col-md-4">
                      <label className="font">Cost Center</label>
                      <input className="form-control readonly" value={costCenter} />
                    </div>
                  </div>
                </div>
                <div className="heading1" style={{ marginTop: "10px" }}>
                  <label>Remark</label>
                </div>
                <div className='main-formcontainer'>
                  <div className='row mb-20'>
                    <div className='col-md-4'>
                      <label className="font" style={{ display: "block" }}>Remark</label>
                      <label className='fonttext textbox readonly' style={{ height: "auto", width: "100%" }}>{remarks}</label>
                    </div>
                    <div className='col-md-4'>
                      <label className="font" style={{ display: "block" }}>Project Description</label>
                      <label className='fonttext textbox readonly' style={{ height: "auto", width: "100%" }}>{projectDesc}</label>
                    </div>
                  </div>
                </div>
                <div className="heading1" style={{ marginTop: "10px" }}>
                  <label>Approver Details</label>
                </div>
                <div className='main-formcontainer'>
                  <div className='row mb-20'>
                    <div className='col-md-4'>
                      <label className="font" style={{ display: "block" }}>Approver Remarks</label>
                      <label className='fonttext textbox readonly' style={{ height: "auto", width: "100%" }}>{approverRemarks}</label>
                    </div>
                    <div className='col-md-4'>
                      <label className="font">Voucher Date</label>
                      <input type="date" className="form-control readonly" value={voucherDate} />
                    </div>
                    <div className='col-md-4'>
                      <label className="font">Voucher Number</label>
                      <input className="form-control readonly" value={VouchingNumber} />
                    </div>
                  </div>
                  <div className='row mb-20'>
                    <div className='col-md-4'>
                      <label className="font">UTR Date</label>
                      <input type="date" className="form-control readonly" value={UTRDate} />
                    </div>
                    <div className='col-md-4'>
                      <label className="font">UTR Number</label>
                      <input className="form-control readonly" value={UTRNumber} />
                    </div>
                  </div>
                </div>
                <div className="heading1" style={{ marginTop: "10px" }}>
                  <label>Upload Document</label>
                </div>
                <div className='main-formcontainer'>
                  <div className='row mb-20'>
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
                            </li>
                          ))}
                        </ul>
                      )}
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
                </div>
                <div className='row' style={{margin : "15px 0px"}}>
                  <div className='col-md-12'>
                    <div className='text-center'>
                      <a href="#" onClick={handleExit} className="reset-btn">
                        Exit
                      </a>
                    </div>
                  </div>
                </div>
              </div> */}
              <div className='borderedbox'>
                <div className="heading1">
                  <label>Requestor Information</label>
                </div>
                <div className='main-formcontainer'>
                  <div className='row mb-20'>
                    <div className='col-md-4'>
                      <label htmlFor="Employee Code" className='font'>Employee Code</label> : &nbsp;&nbsp;
                      <label className='fonttext'>  {employee.EmployeeCode}</label>
                    </div>
                    <div className='col-md-4'>
                      <label htmlFor="Employee Name" className='font'>Employee Name </label> : &nbsp;&nbsp;
                      <label className='fonttext'>  {employee.EmployeeName}</label>
                    </div>
                    <div className='col-md-4'>
                      <label htmlFor="Employee Email" className='font'>Employee Email </label> : &nbsp;&nbsp;
                      <label className='fonttext'>  {employee.EmployeeEmail}</label>
                    </div>
                  </div>
                  <div className='row mb-20'>
                    <div className='col-md-4'>
                      <label htmlFor="Contact No" className='font'>Contact No</label> : &nbsp;&nbsp;
                      <label className='fonttext'>  {employee.ContactNo}</label>
                    </div>
                    <div className='col-md-4'>
                      <label htmlFor="Employee Status" className='font'>Employee Status</label> : &nbsp;&nbsp;
                      <label className='fonttext'>  {employee.EmployeeStatus}</label>
                    </div>
                    <div className='col-md-4'>
                      <label htmlFor="Division" className='font'>Division</label> : &nbsp;&nbsp;
                      <label className='fonttext'>  {employee.Division}</label>
                    </div>
                  </div>
                  <div className='row mb-20'>
                    <div className='col-md-4'>
                      <label htmlFor="Location" className='font'>Location</label> : &nbsp;&nbsp;
                      <label className='fonttext'>  {employee.Location}</label>
                    </div>
                    <div className='col-md-4'>
                      <label htmlFor="RM" className='font'>RM</label> : &nbsp;&nbsp;
                      <label className='fonttext'>  {employee.ReportingManager?.Title}</label>
                    </div>
                    <div className='col-md-4'>
                      <label htmlFor="HOD" className='font'>HOD</label> : &nbsp;&nbsp;
                      <label className='fonttext'>  {employee.HOD?.Title}</label>
                    </div>
                  </div>
                </div>
                <div className="heading1" style={{ marginTop: "10px" }}>
                  <label>Vendor & PO</label>
                </div>
                <div className='main-formcontainer'>
                  <div className='row mb-20'>
                    <div className='col-md-4'>
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
                      <label className="font">Vendor Name</label>
                      <input className="form-control readonly" value={vendorName} />
                    </div>
                    <div className="col-md-4">
                      <label className="font">PO Number</label>
                      <input className="form-control readonly" value={poNumber} />
                    </div>
                  </div>
                  <div className='row mb-20'>
                    <div className='col-md-4'>
                      <label className="font">PO Date</label>
                      <input type="date" className="form-control readonly" value={poDate} />
                    </div>
                    <div className='col-md-4'>
                      <label className="font">PO Terms</label>
                      <textarea
    value={poTerms || ""}
    className="form-control readonly"
    rows={3}
    readOnly
  />
                      {/* <input className="form-control readonly" value={poTerms} /> */}
                    </div>
                    <div className="col-md-4">
                      <label className="font">PO Amount</label>
                      <input className="form-control readonly" value={poAmount} />
                    </div>
                  </div>
                  <div className='row mb-20'>
                    <div className="col-md-4">
                      <label className="font">Advance Amount</label>
                      <input className="form-control readonly" value={advanceAmount} />
                    </div>
                    <div className="col-md-4">
                      <label className="font">Paid Amount</label>
                      <input className="form-control readonly" value={paidAmount} />
                    </div>
                    <div className="col-md-4">
                      <label className="font">Expected Date</label>
                      <input type="date" className="form-control readonly" value={expectedDate} />
                    </div>
                  </div>
                  <div className='row mb-20'>
                    <div className="col-md-4">
                      <label className="font">PIC Name</label>
                      <PeoplePicker
                        context={peoplePickerContext}
                        personSelectionLimit={1}
                        disabled={true}
                        principalTypes={[PrincipalType.User]}
                        defaultSelectedUsers={
                          formData?.PICName?.Title ? [formData.PICName.Title] : []
                        }
                      />
                    </div>
                    <div className="col-md-4">
                      <label className="font">GL Code</label>
                      <input className="form-control readonly" value={glCode} />
                    </div>
                    <div className="col-md-4">
                      <label className="font">Cost Center</label>
                      <input className="form-control readonly" value={costCenter} />
                    </div>
                  </div>
                  <div className='row mb-20'>
                    <div className='col-md-4'>
                      <label className="font" style={{ display: "block" }}>Remark</label>
                      <label className='fonttext textbox readonly' style={{ height: "auto", width: "100%" }}>{remarks}</label>
                    </div>
                    <div className='col-md-4'>
                      <label className="font" style={{ display: "block" }}>Project Description</label>
                      <label className='fonttext textbox readonly' style={{ height: "auto", width: "100%" }}>{projectDesc}</label>
                    </div>
                    <div className='col-md-4'>
                      <label className="font" style={{ display: "block" }}>Approver Remarks</label>
                      <label className='fonttext textbox readonly' style={{ height: "auto", width: "100%" }}>{approverRemarks}</label>
                    </div>
                  </div>
                  <div className='row mb-20'>
                    <div className='col-md-4'>
                      <label className="font">Voucher Date</label>
                      <input type="date" className="form-control readonly" value={voucherDate} />
                    </div>
                    <div className='col-md-4'>
                      <label className="font">Voucher Number</label>
                      <input className="form-control readonly" value={VouchingNumber} />
                    </div>
                    <div className='col-md-4'>
                      <label className="font">UTR Date</label>
                      <input type="date" className="form-control readonly" value={UTRDate} />
                    </div>
                  </div>
                  <div className='row mb-20'>
                    <div className='col-md-4'>
                      <label className="font">UTR Number</label>
                      <input className="form-control readonly" value={UTRNumber} />
                    </div>
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
                  <div className='row mb-20'>
                    <div className='col-md-12'>
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
                          {/* {workflowHistory.map((h, index) => (
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
                          ))} */}
                        </div>
                      )}
                    </div>
                  </div>
                  <div className='row' style={{ margin: "15px 0px" }}>
                    <div className='col-md-12'>
                      <div className='text-center'>
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
        </div>
      </div>

    </>
  );
};

export default ViewAdvanceForm;
