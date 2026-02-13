import * as React from 'react';
import styles from './ForexModule.module.scss';
import type { IForexModuleProps } from './IForexModuleProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { HashRouter as Router, Switch, Route } from 'react-router-dom';

import Sidebar from '../components/Pages/Sidebar';
import { InitiatorDashboard } from '../components/Pages/InitiatorDashboard';
import NewRequest from './Pages/NewRequestForm';
import ApprovalDashboard from './Pages/ApprovalDashboard';
import APTeamLP from './Pages/APTeamLandingPage';
import TreasuryLandingPage from './Pages/TreasuryLandingPage';

export default class ForexModule extends React.Component<IForexModuleProps> {
  public render(): React.ReactElement<IForexModuleProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
    <Router>
      <div className="container-fluid" style={{ display: 'flex', width: '100%' }}>
        <Sidebar {...this.props} />
        <div className="main">
          <Switch>
            <Route exact path="/" render={() => <InitiatorDashboard {...this.props} />} />
            <Route exact path="/NewRequest" render={() => <NewRequest {...this.props} />} />
            <Route exact path="/ApprovalDashboard" render={() => <ApprovalDashboard {...this.props} />} />
            <Route exact path="/APTeamLandingPage" render={() => <APTeamLP {...this.props} />} />
            <Route exact path="/TreasuryLandingPage" render={() => <TreasuryLandingPage {...this.props} />} />

          </Switch>
        </div>
      </div>
    </Router>
  );
  }
}
