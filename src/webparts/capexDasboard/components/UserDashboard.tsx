import * as React from "react";
import "./userDashboardsc.scss";
import NewAdvanceform from "./NewAdvanceform";
import ViewAdvanceForm from "./ViewAdvanceForm";
import EditAdvanceForm from "./EditAdvanceForm";

import { useState } from "react";

import sonalogo from "../assets/SonaPNGLogo.png";
import userlogo from "../assets/userlogo.png";
import "../assets/bootstrap/css/bootstrap.css";

import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";

interface UserDashboardProps {
  context: any;
}

const UserDashboard: React.FC<UserDashboardProps> = ({ context }) => {
  const sp = spfi().using(SPFx(context));
  const [formType, setFormType] = useState<"new" | "view" | "Edit" | null>(
    null,
  );

  const [activeMenu, setActiveMenu] = React.useState("My Request");
  const [searchText, setSearchText] = React.useState("");
  const [statusFilter, setStatusFilter] = React.useState("");
  const [showForm, setShowForm] = React.useState(false);
  const [data, setData] = React.useState<any[]>([]);
  const [currentUserName, setCurrentUserName] = React.useState("");
  const [selectedItem, setSelectedItem] = React.useState<any>(null);
    const [CurrentUserId, setCurrentUserId] = React.useState<any>(null);
  // ✅ GET CURRENT USER
  const getLoggedInUser = async () => {
    try {
      const user = await sp.web.currentUser();
      setCurrentUserName(user.Title);
      setCurrentUserId(user.Id);

    } catch (error) {
      console.error("User error:", error);
    }
  };

  // ✅ GET LIST DATA
  const getCapexData = async () => {
    debugger;
    try {
            const user = await sp.web.currentUser();
      const userId = user.Id;
      const items = await sp.web.lists
  .getByTitle("CapexAdvance")
  .items.select(
    "ID",
    "Title",
    "Created",
    "EmployeeName",
    "VendorName",
    "VendorCode/Id",
    "VendorCode/VendorCode", // 👈 IMPORTANT
    "PONumber",
    "RequestAdvanceAmount",
    "Status"
  )
  .expand("VendorCode")
  .filter(`AuthorId eq '${userId}'`) // 👈 MUST
  .orderBy("ID", false)();


      const formatted = items.map((item: any) => ({
  ID: item.ID,
  id: item.Title,
  date: item.Created
    ? new Date(item.Created).toLocaleDateString("en-GB")
    : "",
  EmployeeName: item.EmployeeName,

  vendor: item.VendorName || "",
  vendorCode: item.VendorCode?.VendorCode || "", // 👈 FIX

  po: item.PONumber || "",
  amount: item.RequestAdvanceAmount || 0,
  status: item.Status || "",
}));


      setData(formatted);
    } catch (error) {
      console.error("Data error:", error);
    }
  };

  // ✅ VIEW CLICK
  const handleViewClick = async (item: any) => {
    try {
      const fullItem = await sp.web.lists
        .getByTitle("CapexAdvance")
        .items.getById(item.ID)
        .select("*", "PICName/Title", "PICName/Id")
        .expand("PICName")();

      setSelectedItem(fullItem);
      setFormType("view");
      setShowForm(true);
    } catch (error) {
      console.error("View error:", error);
    }
  };
  const handleEditClick = async (item: any) => {
    try {
      const fullItem = await sp.web.lists
        .getByTitle("CapexAdvance")
        .items.getById(item.ID)
        .select("*", "PICName/Title","PICName/EMail")
        .expand("PICName")();

      setSelectedItem(fullItem);
      setFormType("Edit");
      setShowForm(true);
    } catch (error) {
      console.error("View error:", error);
    }
  };

  const filteredData = data.filter((item) => {
    const text = searchText.toLowerCase();
    const status = statusFilter.toLowerCase();

    let menuFilter = true;

    if (activeMenu === "Paid") {
      menuFilter = item.status?.toLowerCase() === "paid";
    } else if (activeMenu === "Rejected") {
      menuFilter = item.status?.toLowerCase() === "rejected";
    } else if (activeMenu === "My Request") {
      menuFilter = true;
    }

    return (
      menuFilter &&
      (item.id?.toLowerCase().includes(text) ||
        item.vendor?.toLowerCase().includes(text) ||
        item.po?.toLowerCase().includes(text)) &&
      (!status || item.status?.toLowerCase().includes(status))
    );
  });

  // ✅ LOAD DATA
  React.useEffect(() => {
    if (!context) return;
    void getLoggedInUser();
    void getCapexData();
  }, [context]);

  // ✅ OPEN VIEW PAGE
  if (showForm) {
    if (formType === "view") {
      return (
        <ViewAdvanceForm
          context={context}
          formData={selectedItem}
          onClose={() => {
            setShowForm(false);
            setFormType(null);
            void getCapexData();
          }}
        />
      );
    }

    if (formType === "new") {
      return (
        <NewAdvanceform
          context={context}
          onClose={() => {
            setShowForm(false);
            setFormType(null);
            void getCapexData();
          }}
        />
      );
    }

    if (formType === "Edit") {
      return (
        <EditAdvanceForm
          context={context}
          formData={selectedItem} // ✅ THIS LINE IS MISSING
          onClose={() => {
            setShowForm(false);
            setFormType(null);
            void getCapexData();
          }}
        />
      );
    }
  }

  return (
    <div style={{ display: "flex" }}>
      {/* Sidebar */}
      <div
        className="sidebarmenu"
        style={{
          width: "200px",
          background: "black",
          color: "white",
          height: "100vh",
          paddingTop: "20px",
          textAlign: "center",
        }}
      >
        <h3
          className={activeMenu === "My Request" ? "active" : ""}
          onClick={() => setActiveMenu("My Request")}
        >
          My Request
        </h3>

        <h3
          className={activeMenu === "Paid" ? "active" : ""}
          onClick={() => setActiveMenu("Paid")}
        >
          Paid
        </h3>

        <h3
          className={activeMenu === "Rejected" ? "active" : ""}
          onClick={() => setActiveMenu("Rejected")}
        >
          Rejected
        </h3>
      </div>

      {/* Main */}
      <div style={{ flex: 1 }}>
        {/* Header */}
        <div className="row hederbox">
          <div className="d-left">
            <img src={sonalogo} />
          </div>

          <div className="userinfo">{currentUserName}</div>

          <div className="d-right d-left">
            <img src={userlogo} />
          </div>
        </div>

        <div className="subsection">
          {/* Top */}
          <div className="row">
            <div className="col-md-5">
              <div className="titlebox">CAPEX User Dashboard</div>
            </div>

            <div className="col-md-7">
              <div className="row">
                {/* SEARCH */}
                <div className="col-md-4">
                  <input
                    placeholder="Search"
                    value={searchText}
                    onChange={(e) => setSearchText(e.target.value)}
                  />
                </div>

                {/* STATUS */}
                <div className="col-md-4">
                  <select
                    value={statusFilter}
                    onChange={(e) => setStatusFilter(e.target.value)}
                  >
                    <option value="">All</option>
                    <option value="Submitted">Submitted</option>
                    <option value="Approved">Approved</option>
                    <option value="Rejected">Rejected</option>
                    <option value="Draft">Draft</option>
                  </select>
                </div>

                {/* NEW REQUEST */}
                <div className="col-md-4">
                  <button
                    onClick={() => {
                      setSelectedItem(null);
                      setFormType("new");
                      setShowForm(true);
                    }}
                  >
                    New Request
                  </button>
                </div>
              </div>
            </div>
          </div>

          {/* TABLE */}
          <table className="data-table-section">
            <thead>
              <tr>
                <th>Action</th>
                <th>Payment ID</th>
                <th>Requestor Date</th>
                <th>Requestor Name</th>
                <th>Requestor Type</th>
                <th>Vendor Code</th>
                <th>Vendor Name</th>
                <th>PO Number</th>
                <th>Advance Amount</th>
                <th>Pending With</th>
                <th>Status</th>
              </tr>
            </thead>

            <tbody>
              {filteredData.length === 0 ? (
                <tr>
                  <td colSpan={7} style={{ textAlign: "center" }}>
                    No Data
                  </td>
                </tr>
              ) : (
                filteredData.map((item, i) => (
                  <tr key={i}>
                    <td
                      style={{
                        display: "flex",
                        gap: "10px",
                        cursor: "pointer",
                      }}
                    >
                      {/* View Icon */}
                      <span
                        onClick={() => handleViewClick(item)}
                        title="View"
                        style={{ cursor: "pointer" }}
                      >
                        👁
                      </span>

                      {/* Edit Icon */}
                      {(item.status === "Draft" ||
                        item.status === "Send Back") && (
                        <span
                          onClick={() => handleEditClick(item)}
                          style={{ cursor: "pointer" }}
                          title="Edit"
                        >
                          ✏️
                        </span>
                      )}
                    </td>

                    <td>{item.id}</td>
                    <td>{item.date}</td>
                    <td>{item.EmployeeName}</td>
                    <td>Capex Advance</td>
                    <td>{item.vendorCode}</td>

                    <td>{item.vendor}</td>
                    <td>{item.po}</td>
                    <td>₹ {item.amount}</td>
                    <td>Approver</td>

                    <td>{item.status}</td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

export default UserDashboard;
