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
    location.pathname.startsWith("/ApprovalRequest/");
    
    return (
   
      <div className="container-fluid" style={{ display: 'flex', width: '100%' }}>
        {!hideSidebar && <Sidebar {...props} />}
        <div className="main">
          <Switch>
            <Route exact path="/" render={() => <InitiatorDashboard {...props} />} />
            <Route exact path="/NewRequest" render={() => <NewRequest {...props} />} />
            <Route exact path="/ViewRequest/:Id" render={() => <ViewRequestForm {...props} />} />
            <Route exact path="/EditRequest/:Id" render={() => <Editrequest {...props} />} />
            <Route exact path="/ApprovalRequest/:Id" render={() => <ApprovalRequestForm {...props} />} />
            <Route exact path="/ApprovalDashboard" render={() => <ApprovalDashboard {...props} />} />
            <Route exact path="/APTeamLandingPage" render={() => <APTeamLP {...props} />} />
            <Route exact path="/TreasuryLandingPage" render={() => <TreasuryLandingPage {...props} />} />

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
