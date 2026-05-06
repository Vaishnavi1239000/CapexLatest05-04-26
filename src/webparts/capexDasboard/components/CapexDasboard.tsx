import * as React from 'react';
import './usersite.scss';

import UserDashboard from "./UserDashboard";
import ApproverDashboard from "./ApproverDashboard";
import APperformerDashboard from './APperformerDashboard';
import APperformerdashboardforutr from './APperformerdashboardforutr';
import { ICapexDasboardProps } from './ICapexDasboardProps';

export default function CapexDasboard(props: ICapexDasboardProps) {

  const [page, setPage] = React.useState<string>("home");

  // ✅ Navigation function (updates URL + state)
  const navigate = (pageName: string) => {
    const url = `${props.context.pageContext.web.absoluteUrl}/SitePages/CapexForm.aspx?page=${pageName}`;
    window.history.pushState({}, "", url);
    setPage(pageName);
  };

  // ✅ On page load → read URL
  React.useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const pageParam = params.get("page");

    if (pageParam) {
      setPage(pageParam);
    }
  }, []);

  // ✅ Handle browser back/forward
  React.useEffect(() => {
    const handlePopState = () => {
      const params = new URLSearchParams(window.location.search);
      const pageParam = params.get("page") || "home";
      setPage(pageParam);
    };

    window.addEventListener("popstate", handlePopState);

    return () => {
      window.removeEventListener("popstate", handlePopState);
    };
  }, []);

  return (
    <div>

      {/* ✅ HOME PAGE */}
      {page === "home" && (
        <div className="main-container">
          <h1 className="title">Capex Advance form</h1>

          <div className="buttonGrid">
            <button className="btn" onClick={() => navigate("User")}>
              User Dashboard
            </button>

            <button className="btn" onClick={() => navigate("Approver")}>
              Approver Dashboard
            </button>

            <button className="btn fullWidth" onClick={() => navigate("Performer")}>
              AP Performer Dashboard
            </button>

            {/* <button className="btn fullWidth" onClick={() => navigate("apPerformer")}>
              AP Performer Dashboard for UTR
            </button> */}
          </div>
        </div>
      )}

      {/* ✅ USER DASHBOARD */}
      {page === "User" && (
        <UserDashboard context={props.context} />
      )}

      {/* ✅ APPROVER DASHBOARD */}
      {page === "Approver" && (
        <ApproverDashboard context={props.context} />
      )}

      {/* ✅ AP PERFORMER */}
      {page === "Performer" && (
        <APperformerDashboard context={props.context} />
      )}

      {/* ✅ AP PERFORMER UTR */}
      {/* {page === "apPerformer" && (
        <APperformerdashboardforutr context={props.context} />
      )} */}

    </div>
  );
}
