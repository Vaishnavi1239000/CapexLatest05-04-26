import * as React from "react";
import "./advanced.scss";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { useEffect, useRef, useState } from "react";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import logo from "../assets/sona-comstarlogo.png";
interface IVendor {
  Id: number;
  VendorCode: string;
  VendorName: string;
}

const EditAdvanceForm = ({ context, formData, onClose }: any) => {
  const sp = spfi().using(SPFx(context));
  const today = new Date();

  const localDate: string = new Date(
    today.getTime() - today.getTimezoneOffset() * 60000
  )
    .toISOString()
    .split("T")[0];

  const [isDraftSaving, setIsDraftSaving] = useState(false);
  const submitRef = useRef(false);
  const [previousAdvances, setPreviousAdvances] = useState<any[]>([]);
  const [vendors, setVendors] = useState<IVendor[]>([]);
  const [employee, setEmployee] = useState<any>({});
  const [attachments, setAttachments] = useState<any[]>([]);
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [selectedUser, setSelectedUser] = useState<any[]>([]);
  const [isSubmitting, setIsSubmitting] = useState(false);
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
  //const [previousAdvances, setPreviousAdvances] = useState<any[]>([]);
  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: context.pageContext.web.absoluteUrl,
    msGraphClientFactory: context.msGraphClientFactory,
    spHttpClient: context.spHttpClient,
  };



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
  const handleDeleteAttachment = async (fileName: string) => {
    try {
      if (!formData?.CapexID) return;

      const safeCapexId = formData.CapexID.replace(/\//g, "_");

      const folderPath = `/sites/SonaFinance/CapexAdvanceDocs/${safeCapexId}`;

      await sp.web
        .getFolderByServerRelativePath(folderPath)
        .files.getByUrl(fileName)
        .recycle();

      // Update UI after delete
      const updatedFiles = attachments.filter(
        (file: any) => file.Name !== fileName,
      );

      setAttachments(updatedFiles);

      alert("Attachment deleted successfully ✅");
    } catch (error) {
      console.error("Delete attachment error:", error);
      alert("Error deleting attachment ❌");
    }
  };
  const handleRemoveFile = (index: number) => {
    const updatedFiles = [...selectedFiles];
    updatedFiles.splice(index, 1);
    setSelectedFiles(updatedFiles);
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
      if (!vendorId) {
        setPreviousAdvances([]);
        return;
      }
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
    try {
      const data = await sp.web.lists
        .getByTitle("VendorMaster")
        .items.select("Id", "VendorCode", "VendorName", "Status")
        .filter("Status eq 'Active'")()
        ;

      setVendors(data);
    } catch (error) {
      console.error("Vendor fetch error:", error);
    }
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

      const files = await sp.web.getFolderByServerRelativePath(path).files();

      void setAttachments(files);
    } catch {
      void setAttachments([]);
    }
  };

  const validateForm = () => {
    const errors: string[] = [];

    if (!selectedVendorId || selectedVendorId === 0) {
      errors.push("Please select the Vendor code");
    }

    if (!poNumber || poNumber.trim() === "") {
      errors.push("Please update PO Number");
    }

    if (!poDate || poDate.trim() === "") {
      errors.push("Please update PO date");
    }
    if (poDate > localDate) {
      errors.push("PO Date cannot be a future date");
      // return;
    }

    if (!poTerms || poTerms.trim() === "") {
      errors.push("Please update PO Terms");
    }

    if (!poAmount || Number(poAmount) <= 0) {
      errors.push("Please update PO Amount");
    }

    if (!advanceAmount || Number(advanceAmount) <= 0) {
      errors.push("Please update Advance Amount");
    }

    if (!paidAmount || Number(paidAmount) <= 0) {
      errors.push("Please update Paid Amount");
    }

    if (poAmount && advanceAmount && Number(advanceAmount) > Number(poAmount)) {
      errors.push(
        "The requested advance amount cannot be greater than the PO Amount (Including GST)",
      );
    }


    if (
      advanceAmount &&
      paidAmount &&
      Number(paidAmount) > Number(advanceAmount)
    ) {
      errors.push("Paid Amount cannot be greater than Advance Amount");
    }





    if (expectedDate) {
      const today = new Date().setHours(0, 0, 0, 0);
      const selected = new Date(expectedDate).setHours(0, 0, 0, 0);

      if (selected < today) {
        errors.push("Settlement date cannot be a past date");
      }
    }

    if (!selectedUser || selectedUser.length === 0) {
      errors.push("Please select PIC Name");
    }

    if (!projectDesc || projectDesc.trim() === "") {
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

  const uploadAttachments = async (capexId: string) => {
    try {
      if (!selectedFiles || selectedFiles.length === 0) return;

      const safeCapexId = capexId.replace(/\//g, "_");

      const libraryName = "CapexAdvanceDocs";
      const webUrl = context.pageContext.web.serverRelativeUrl;

      const folderPath = `${webUrl}/${libraryName}/${safeCapexId}`;


      await sp.web.folders.addUsingPath(`${libraryName}/${safeCapexId}`);


      for (const file of selectedFiles) {
        await sp.web
          .getFolderByServerRelativePath(folderPath)
          .files.addUsingPath(file.name, file, { Overwrite: true });
      }

      console.log("Files uploaded successfully");
    } catch (error) {
      console.error("Upload error:", error);
    }
  };
  const handleDraft = async () => {
    try {
      if (isSubmitting) return;
      setIsSubmitting(true);
      // const capexId = await generateCapexId();

      let ensuredUserId: number | null = null;


      if (selectedUser && selectedUser.length > 0) {
        const userEmail = selectedUser[0]?.secondaryText;

        if (userEmail) {
          const ensuredUser = await sp.web.ensureUser(userEmail);
          ensuredUserId = ensuredUser.Id;
        }
      }



      // const currentApprover = flow.length > 0 ? flow[0].Id : null;
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


        ...(ensuredUserId && { PICNameId: ensuredUserId }),

        // Other
        GL: glCode,
        CostCenter: employee.CostCenter,
        Remarks: remarks,
        ProjectDescription: projectDesc,

        Status: "Draft",

        //  ApprovalMatrix: JSON.stringify(flow),

        // CurrentApproverId: currentApprover,

        //WorkFlowHistory: JSON.stringify(wfHistory),
      });

      const safeCapexId = formData.capexId.replace(/\//g, "_");
      void uploadAttachments(safeCapexId);

      alert("Draft saved successfully ✅");

      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexForm.aspx?page=User";
    } catch (error) {
      console.error("ERROR:", error);
      alert("Error while saving ❌");
      // setIsSubmitting(false);
    }
  };
  const handledraft = async () => {
    try {
      debugger;
      let ensuredUserId: number | null = null;
      if (isSubmitting) return;
      setIsSubmitting(true);
      // const flow = await buildApprovalFlow();
      // const currentApprover = flow.length > 0 ? flow[0].Id : null;

      //  const loginName =
      //  selectedUser[0]?.secondaryText;
      ////const ensuredUser = await sp.web.ensureUser(loginName);

      //const ensuredUserId = ensuredUser.Id;


      const loginName = selectedUser[0]?.secondaryText;
      if (loginName != null || loginName != undefined) {
        const ensuredUser = await sp.web.ensureUser(
          selectedUser[0]?.secondaryText,
        );

        ensuredUserId = ensuredUser.Id;
      } else {
        ensuredUserId = null;
      }

      const existingFlow = formData.ApprovalMatrix
        ? JSON.parse(formData.ApprovalMatrix)
        : [];

      // 🔥 preserve history
      const history = formData.WorkFlowHistory
        ? JSON.parse(formData.WorkFlowHistory)
        : [];

      history.push({
        CurrentApprover: employee.EmployeeName,
        ActionTaken: "Edited",
        Comment: remarks,
        Date: new Date().toISOString(),
      });

      await sp.web.lists
        .getByTitle("CapexAdvance")
        .items.getById(formData.ID)
        .update({
          Title: formData.CapexID,
          CapexID: formData.CapexID,

          EmployeeCode: employee.EmployeeCode,
          EmployeeName: employee.EmployeeName,

          VendorCodeId: selectedVendorId,
          VendorName: selectedVendorName,

          PONumber: poNumber,
          PODate: poDate ? new Date(poDate) : null,

          POAmtGST: poAmount,
          RequestAdvanceAmount: advanceAmount,
          PaidAmount: paidAmount,

          ExpectedDateofSettlement: expectedDate
            ? new Date(expectedDate)
            : null,

          PICNameId: ensuredUserId,

          GL: glCode,
          CostCenter: costCenter,
          Remarks: remarks,
          ProjectDescription: projectDesc,

          Status: "Draft",
          //ApprovalMatrix: JSON.stringify(flow),

          // CurrentApproverId: currentApprover, // 🔥 not started

          WorkFlowHistory: JSON.stringify(history),
        });

      if (selectedFiles.length > 0) {
        await uploadFiles();
      }

      alert("Draft saved successfully ✅");
      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexForm.aspx?page=User";
      void handleExit();
      // window.location.reload();
    } catch (error) {
      console.error("ERROR:", error);
      alert(error);
    }
    finally {
      // submitRef.current = false;
      setIsSubmitting(false);
    }
  };

  // =========================
  // UPDATE
  // =========================
  const handleSubmit = async () => {
    if (submitRef.current) return;

    //submitRef.current = true;
    setIsSubmitting(true);
    try {
      const errors = validateForm();
      if (errors.length > 0) {
        alert(errors.join("\n"));
        submitRef.current = false;
        setIsSubmitting(false);
        return;
      }

      const ensuredUser = await sp.web.ensureUser(
        selectedUser[0]?.secondaryText,
      );
      // 🔥 PRESERVE FLOW
      const existingFlow = formData.ApprovalMatrix
        ? JSON.parse(formData.ApprovalMatrix)
        : [];
      const updatedFlow = existingFlow.map((a: any, index: number) => ({
        ...a,
        Status: index === 0 ? "In Progress" : "Pending",
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
        Date: new Date().toISOString(),
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

          CurrentApproverId: currentApprover,
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
    finally {
      //submitRef.current = false;
      setIsSubmitting(false);
    }
  };

  // =========================
  // BIND DATA
  // =========================
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
    //void getPreviousAdvances(formData.VendorCodeId || null);

    setSelectedVendorName(formData.VendorName || "");

    setGlCode(formData.GL || "");
    setCostCenter(formData.CostCenter || "");
    setRemarks(formData.Remarks || "");
    setProjectDesc(formData.ProjectDescription || "");

    setApproverRemarks(formData.ApproverRemarks || "");
    setVoucherDate(formData.VoucherDate?.split("T")[0] || "");
    setVoucherNumber(formData.VoucherNumber || "");
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
    debugger;
    void getLoggedInUser();
    void getVendors();
    if (selectedVendorId) {
      void getPreviousAdvances(selectedVendorId);
    }
  }, []);



  // =========================
  // UI
  // =========================
  return (
    <>
      <div className="MainUplodForm" style={{ margin: "5px 0px" }}>
        <div className="row">
          <div className="col-md-12">
            <div className="Main-Boxpoup">

              <div className="bordered">
                <img src={logo} />
                <h1>Capex Edit Advance Payment </h1>
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
                            : ""
                          }`}
                      >
                        {a.Role} - {a.Name}
                      </li>
                    ))}
                  </ul>
                </div>
              )}

              <div className="borderedbox">
                <div className="heading1">
                  <label>Requestor Information</label>
                </div>
                <div className="main-formcontainer">
                  <div className="row mb-20">
                    <div className="col-md-4">
                      <label htmlFor="Employee Code" className="font">
                        Employee Code
                      </label>{" "}
                      : &nbsp;&nbsp;
                      <label className="fonttext">
                        {" "}
                        {employee.EmployeeCode}
                      </label>
                    </div>
                    <div className="col-md-4">
                      <label htmlFor="Employee Name" className="font">
                        Employee Name{" "}
                      </label>{" "}
                      : &nbsp;&nbsp;
                      <label className="fonttext">
                        {" "}
                        {employee.EmployeeName}
                      </label>
                    </div>
                    <div className="col-md-4">
                      <label htmlFor="Employee Email" className="font">
                        Employee Email{" "}
                      </label>{" "}
                      : &nbsp;&nbsp;
                      <label className="fonttext">
                        {" "}
                        {employee.EmployeeEmail}
                      </label>
                    </div>
                  </div>
                  <div className="row mb-20">
                    <div className="col-md-4">
                      <label htmlFor="Contact No" className="font">
                        Contact No
                      </label>{" "}
                      : &nbsp;&nbsp;
                      <label className="fonttext"> {employee.ContactNo}</label>
                    </div>
                    <div className="col-md-4">
                      <label htmlFor="Employee Status" className="font">
                        Employee Status
                      </label>{" "}
                      : &nbsp;&nbsp;
                      <label className="fonttext">
                        {" "}
                        {employee.EmployeeStatus}
                      </label>
                    </div>
                    <div className="col-md-4">
                      <label htmlFor="Division" className="font">
                        Division
                      </label>{" "}
                      : &nbsp;&nbsp;
                      <label className="fonttext"> {employee.Division}</label>
                    </div>
                  </div>
                  <div className="row mb-20">
                    <div className="col-md-4">
                      <label htmlFor="Location" className="font">
                        Location
                      </label>{" "}
                      : &nbsp;&nbsp;
                      <label className="fonttext"> {employee.Location}</label>
                    </div>
                    <div className="col-md-4">
                      <label htmlFor="RM" className="font">
                        RM
                      </label>{" "}
                      : &nbsp;&nbsp;
                      <label className="fonttext">
                        {" "}
                        {employee.ReportingManager?.Title}
                      </label>
                    </div>
                    <div className="col-md-4">
                      <label htmlFor="HOD" className="font">
                        HOD
                      </label>{" "}
                      : &nbsp;&nbsp;
                      <label className="fonttext"> {employee.HOD?.Title}</label>
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
                      <span className="required" style={{ color: "red" }}>
                        *
                      </span>
                      <select
                        value={selectedVendorId || ""}
                        onChange={(e) => {
                          const id = Number(e.target.value);
                          const vendor = vendors.find((v) => v.Id === id);
                          setSelectedVendorId(id);
                          setSelectedVendorName(vendor?.VendorName || "");
                          if (id > 0) {
                            void getPreviousAdvances(id);
                          } else {
                            // Clear table data
                            setPreviousAdvances([]);
                          }
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
                      <label className="font">Vendor Name</label>
                      <span className="required" style={{ color: "red" }}>
                        *
                      </span>
                      <input
                        value={selectedVendorName || vendorName}
                        className="form-control readonly"
                      />
                    </div>
                    <div className="col-md-4">
                      <label className="font">PO Number</label>
                      <span className="required" style={{ color: "red" }}>
                        *
                      </span>
                      <input
                        value={poNumber}
                        onChange={(e) => setPoNumber(e.target.value)}
                        className="form-control"
                      />
                    </div>
                  </div>
                  <div className="row mb-20">
                    <div className="col-md-4">
                      <label className="font">PO Date</label>
                      <span className="required" style={{ color: "red" }}>
                        *
                      </span>
                      <input
                        type="date"
                        value={poDate}
                        max={new Date().toISOString().split("T")[0]}
                        onChange={(e) => setPoDate(e.target.value)}
                        className="form-control"
                      />
                    </div>
                    <div className="col-md-4">
                      <label className="font">PO Advance Terms</label>
                      <span className="required" style={{ color: "red" }}>
                        *
                      </span>
                      <textarea
                        value={poTerms}
                        onChange={(e) => setPoTerms(e.target.value)}
                        className="form-control"
                        rows={3}
                      />
                    </div>
                    <div className="col-md-4">
                      <label className="font">PO Amount (GST)</label>
                      <span className="required" style={{ color: "red" }}>
                        *
                      </span>
                      <input
                        value={poAmount}
                        onChange={(e) =>
                          handleNumberChange(e.target.value, setPoAmount)
                        }
                        className="form-control"
                      />
                    </div>
                  </div>
                  <div className="row mb-20">
                    <div className="col-md-4">
                      <label className="font">Request Advance Amount</label>
                      <span className="required" style={{ color: "red" }}>
                        *
                      </span>
                      <input
                        value={advanceAmount}
                        onChange={(e) =>
                          handleNumberChange(e.target.value, setAdvanceAmount)
                        }
                        className="form-control"
                      />
                    </div>
                    <div className="col-md-4">
                      <label className="font" style={{ color: "red" }}>
                        Paid Amount
                      </label>
                      <span className="required" style={{ color: "red" }}>
                        *
                      </span>
                      <input
                        value={paidAmount}
                        onChange={(e) =>
                          handleNumberChange(e.target.value, setPaidAmount)
                        }
                        className="form-control"
                      />
                    </div>
                    <div className="col-md-4">
                      <label className="font">Expected Settlement Date</label>
                      <span className="required" style={{ color: "red" }}>
                        *
                      </span>
                      <input
                        type="date"
                        value={expectedDate}
                        min={localDate}
                        onChange={(e) => setExpectedDate(e.target.value)}
                        className="form-control"
                      />
                    </div>
                  </div>
                  <div className="row mb-20">
                    <div className="col-md-4">
                      <label className="font">PIC Name</label>
                      <span className="required" style={{ color: "red" }}>
                        *
                      </span>
                      <PeoplePicker
                        context={peoplePickerContext}
                        personSelectionLimit={1}
                        ensureUser={true}
                        principalTypes={[PrincipalType.User]}
                        defaultSelectedUsers={
                          selectedUser.length > 0
                            ? [selectedUser[0].secondaryText]
                            : []
                        }
                        onChange={(items) => setSelectedUser(items)}
                      />
                    </div>
                    <div className="col-md-4">
                      <label className="font">GL Code</label>
                      <span className="required" style={{ color: "red" }}>
                        *
                      </span>
                      <input
                        value={glCode}
                        // onChange={(e) => setExpectedDate(e.target.value)}
                        className="form-control"
                        readOnly
                      />
                    </div>
                    <div className="col-md-4">
                      <label className="font">Cost Center</label>
                      <span className="required" style={{ color: "red" }}>
                        *
                      </span>
                      <input
                        value={employee.CostCenter || ""}
                        className="form-control"
                        readOnly
                      />
                    </div>
                  </div>
                  <div className="row mb-20">
                    <div className="col-md-4">
                      <label className="font">Project Description</label>
                      <span className="required" style={{ color: "red" }}>
                        *
                      </span>
                      <textarea
                        value={projectDesc}
                        onChange={(e) => setProjectDesc(e.target.value)}
                        className="form-control"
                      />
                    </div>

                    <div className="col-md-4">
                      <label className="font">
                        Attachments
                        <span className="required" style={{ color: "red" }}>
                          *
                        </span>
                      </label>

                      {/* Existing Attachments */}
                      {/* Existing Attachments */}
                      {attachments.length > 0 && (
                        <ul className="mt-2">
                          {attachments.map((file: any, index: number) => (
                            <li
                              key={index}
                              style={{
                                display: "flex",
                                alignItems: "center",
                                gap: "10px",
                                marginBottom: "5px",
                              }}
                            >
                              <a
                                href={file.ServerRelativeUrl}
                                target="_blank"
                                rel="noopener noreferrer"
                              >
                                {file.Name}
                              </a>

                              {/* Delete Existing File */}
                              <a onClick={() => handleDeleteAttachment(file.Name)}>
                                X
                              </a>
                            </li>
                          ))}
                        </ul>
                      )}

                      {/* Newly Selected Files */}
                      {selectedFiles.length > 0 && (
                        <ul className="mt-2">
                          {selectedFiles.map((file: File, index: number) => (
                            <li
                              key={index}
                              style={{
                                display: "flex",
                                alignItems: "center",
                                gap: "10px",
                                marginBottom: "5px",
                              }}
                            >
                              <a
                                href={URL.createObjectURL(file)}
                                target="_blank"
                                rel="noopener noreferrer"
                              >
                                {file.name}
                              </a>

                              {/* Remove Selected File */}
                              <button
                                type="button"
                                className="btn btn-sm btn-danger"
                                onClick={() => {
                                  const updatedFiles = selectedFiles.filter(
                                    (_: File, i: number) => i !== index,
                                  );

                                  setSelectedFiles(updatedFiles);
                                }}
                              >
                                Remove
                              </button>
                            </li>
                          ))}
                        </ul>
                      )}

                      <input
                        type="file"
                        multiple style={{ height: "34px", padding: "3px" }}
                        className="form-control"
                        onChange={(e) => {
                          if (e.target.files) {
                            setSelectedFiles(Array.from(e.target.files));
                          }
                        }}
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
                  <div className="heading1" style={{ marginTop: "10px" }}>
                    <label>Remarks</label>
                  </div>
                  <div className="main-formcontainer">
                    <div className="row mb-20">
                      <div className="col-md-12">
                        <label className="font">Remarks</label>
                        <textarea
                          value={remarks}
                          onChange={(e) => setRemarks(e.target.value)}
                          className="form-control"
                        />
                      </div>
                    </div>
                  </div>
                  <div className="main-formcontainer" style={{ marginTop: "10px" }}>
                    <div className="row mb-20">
                      <div className="col-md-12">
                        {workflowHistory.length === 0 ? (
                          <p>No history available</p>
                        ) : (
                          <div style={{overflowX : "auto"}}>
                            <table
                              className="custom-table"
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
                                      h.ActionTaken !== "Save as Draft" && h.ActionTaken !== "Edited",
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
                  </div>
                  <div
                    style={{
                      display: "flex",
                      justifyContent: "center",
                      gap: "10px",
                      margin: "10px",
                    }}
                  >
                    <button
                      type="button"
                      onClick={!isSubmitting ? handleSubmit : undefined}
                      disabled={isSubmitting}
                      className="submit-btn"
                      style={{
                        pointerEvents: isSubmitting ? "none" : "auto",
                        opacity: isSubmitting ? 0.6 : 1,
                        cursor: isSubmitting ? "not-allowed" : "pointer",
                      }}
                    >
                      {isSubmitting ? "Submitting..." : "Submit"}
                    </button>
                    <button
                      type="button"
                      onClick={!isDraftSaving ? handledraft : undefined}
                      disabled={isDraftSaving}
                      className="Rework-btn"
                      style={{
                        pointerEvents: isDraftSaving ? "none" : "auto",
                        opacity: isDraftSaving ? 0.6 : 1,
                        cursor: isDraftSaving ? "not-allowed" : "pointer",
                      }}
                    >
                      {isDraftSaving ? "Saving..." : "Save as Draft"}
                    </button>

                    <a className="reset-btn" onClick={handleExit}>
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
export default EditAdvanceForm;
