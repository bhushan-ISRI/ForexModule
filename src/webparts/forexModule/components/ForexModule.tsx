import * as React from 'react';
import styles from './ForexModule.module.scss';
import type { IForexModuleProps } from './IForexModuleProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { HashRouter as Router, Switch, Route, useLocation } from 'react-router-dom';

import Sidebar from '../components/Pages/Sidebar';
import { InitiatorDashboard } from '../components/Pages/InitiatorDashboard';
import NewRequest from './Pages/NewRequestForm';
import ApprovalDashboard from './Pages/ApprovalDashboard';
import APTeamLP from './Pages/APTeamLandingPage';
import TreasuryLandingPage from './Pages/TreasuryLandingPage';
import Editrequest from './Pages/NewRequestEditForm';
import ViewRequestForm from './Pages/NewRequestViewForm';
import ApprovalRequestForm from './Pages/ApprovalRequestForm';

import TrackerForm from './Pages/InitiatorPaymentForm';
import TrackerApprovalForm from './Pages/TrackerApprovalForm';
import ModernForexDashboard from './Pages/ApprovedRejectByme';
import VendorReviewForm from './Pages/VendorReviewForm';
import VendorCreationForm from './Pages/VendorCreationForm';
import VendorApprovalForm from './Pages/IDCApprovalForm';
import CreationForm from './Pages/Creationform';
import VendorDashboard from './Pages/VendorCreationDashboard';
import VendorApprovalDashboard from './Pages/VendorApprovalDashboard';
import VendorApprovalFormFirst from './Pages/VendorApprovalForm';
import VendorViewForm from './Pages/VendorViewForm';

const ForexModule: React.FC<IForexModuleProps> = (props) => {
  const {
    description,
    isDarkTheme,
    environmentMessage,
    hasTeamsContext,
    userDisplayName
  } = props;

  const location = useLocation(); // ✅ Now inside Router

  const hideSidebar =
    location.pathname === "/NewRequest" ||
    location.pathname.startsWith("/ViewRequest/") ||
    location.pathname.startsWith("/EditRequest/") ||
    location.pathname.startsWith("/ApprovalRequest/") || location.pathname.startsWith("/VendorCreationForm ") || location.pathname.startsWith("/AdvancePaymentTracker/") || location.pathname.startsWith('/TrackerApprovalForm/') || location.pathname.startsWith('/ApprovedRejectedByMe');

  return (

    <div className="container-fluid" style={{ display: 'flex', width: '100%' }}>
      {!hideSidebar && <Sidebar {...props} />}
      <div className="main" style={{
        width: hideSidebar ? "100%" : "calc(100% - 200px)",
        transition: "width 0.3s ease"
      }}>
        <Switch>
          <Route exact path="/" render={() => <InitiatorDashboard {...props} />} />
          <Route exact path="/NewRequest" render={() => <NewRequest {...props} />} />
          <Route exact path="/ViewRequest/:Id" render={() => <ViewRequestForm {...props} />} />
          <Route exact path="/EditRequest/:Id" render={() => <Editrequest {...props} />} />
          <Route exact path="/AdvancePaymentTracker/:Id" render={() => <TrackerForm {...props} />} />
          <Route exact path="/ApprovalRequest/:Id" render={() => <ApprovalRequestForm {...props} />} />
          <Route exact path="/ApprovalDashboard" render={() => <ApprovalDashboard {...props} />} />
          <Route exact path="/APTeamLandingPage" render={() => <APTeamLP {...props} />} />
          <Route exact path="/TreasuryLandingPage" render={() => <TreasuryLandingPage {...props} />} />
          <Route exact path="/TrackerApprovalForm/:Id" render={() => <TrackerApprovalForm {...props} />} />
          <Route exact path="/ApprovedRejectedByMe" render={() => <ModernForexDashboard {...props} />} />
          <Route exact path="/VendorReviewForm" render={() => <VendorReviewForm {...props} />} />
          <Route exact path="/VendorCreationForm/:Id" render={() => <VendorCreationForm {...props} />} />
          <Route exact path="/VendorApprovalForm/:Id" render={() => <VendorApprovalForm {...props} />} />
          <Route exact path="/CreationForm" render={() => <CreationForm {...props} />} />
          <Route exact path="/VendorCreationDashboard" render={() => <VendorDashboard {...props} />} />
          <Route exact path="/VendorApprovalDashboard" render={() => <VendorApprovalDashboard {...props} />} />
          <Route exact path="/VendorApprovalFormFirst/:Id" render={() => <VendorApprovalFormFirst {...props} />} />
          <Route exact path="/VendorViewForm/:Id" render={() => <VendorViewForm {...props} />} />
        </Switch>
      </div>
    </div>

  );
}

const Drr: React.FC<IForexModuleProps> = (props) => {
  return (
    <Router>
      <ForexModule {...props} />
    </Router>
  );
};

export default Drr;
