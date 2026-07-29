import * as React from 'react';
import styles from './CustomerSatisfaction.module.scss';
import { ICustomerSatisfactionProps } from './ICustomerSatisfactionProps';
import { ICustomerSatisfactionState, ViewName } from './ICustomerSatisfactionState';
import { UserRoleService } from '../services/UserRole_Service';
import { CSPUserRole_Service } from '../services/CSPUserRole_Service';
import Sidebar from './SideBar/Sidebar';
import Home from './Home/Home';
import About from './About/About';
import ActionPlan from './ActionPlan/ActionPlan';
import Results from './Results/Results';
import ServiceResultDirectory from './Results/ServiceResultDirectory';
import Dashboard from './Dashboard/Dashboard';
import RoleManagement from './RoleManagement/RoleManagement';

export default class CustomerSatisfaction
  extends React.Component<ICustomerSatisfactionProps, ICustomerSatisfactionState> {

  private userRoleService: UserRoleService;
  private cspUserRoleService: CSPUserRole_Service;

  constructor(props: ICustomerSatisfactionProps) {
    super(props);
    this.state = {
      isTeamLeader: false,
      userRole: 'visitor' as 'visitor' | 'leader' | 'manager' | 'admin',
      isLoading: true,
      sentToday: 0,
      receivedToday: 0,
      isSidebarCollapsed: false,
      currentView: 'home',
      userService: '',
      selectedResultDepartment: undefined,
      selectedResultService: undefined,
      selectedResultLabel: undefined
    };
    this.userRoleService = new UserRoleService(props.context, 'RoleInService');
    this.cspUserRoleService = new CSPUserRole_Service(props.context);
  }

  public async componentDidMount(): Promise<void> {
    console.log('CustomerSatisfaction componentDidMount - currentUserEmail:', this.props.currentUserEmail);

    // Get user role from CSP_UserRole list
    const userRoleType = await this.cspUserRoleService.getUserRoleType(this.props.currentUserEmail);
    console.log('CustomerSatisfaction componentDidMount - userRoleType:', userRoleType);

    const isTeamLeader = userRoleType === 'leader' || userRoleType === 'manager' || userRoleType === 'admin';
    console.log('CustomerSatisfaction componentDidMount - isTeamLeader:', isTeamLeader);

    // Get user service from RoleInService list (for ActionPlan)
    const userRole = await this.userRoleService.getUserRole(this.props.currentUserEmail);
    const userService = userRole?.Title || '';

    console.log('CustomerSatisfaction componentDidMount - setting state with userRole:', userRoleType);
    this.setState({ userRole: userRoleType, isTeamLeader, isLoading: false, userService });
  }

  private navigateTo = (view: ViewName): void => {
    this.setState({
      currentView: view,
      selectedResultDepartment: undefined,
      selectedResultService: undefined,
      selectedResultLabel: undefined,
    });
  };

  private openServiceResult = (department: string, service: string, label: string): void => {
    this.setState({
      currentView: 'serviceResult',
      selectedResultDepartment: department,
      selectedResultService: service,
      selectedResultLabel: label,
    });
  };

  private renderHome(): JSX.Element {
    const { currentUserDisplayName } = this.props;

    return (
      <Home
        context={this.props.context}
        currentUserDisplayName={currentUserDisplayName}
      />
    );
  }

  private renderAbout(): JSX.Element {
    return (
      <main className={styles.mainPlain}>
        <div className={styles.pageBreadcrumb}>
          Home › <strong>Learn More</strong>
        </div>
        <About context={this.props.context} />
      </main>
    );
  }

  private renderActionPlan(): JSX.Element {
    const { userRole, userService } = this.state;

    if (userRole === 'admin') {
      return <ActionPlan context={this.props.context} userService="" filterMode="admin" />;
    }

    if (userRole === 'manager') {
      return <ActionPlan context={this.props.context} userService={userService} filterMode="manager" />;
    }

    // leader
    if (!userService) {
      return (
        <main className={styles.mainPlain}>
          <div className={styles.pageBreadcrumb}>Home › <strong>Action Plan</strong></div>
          <div style={{ padding: 40 }}>
            <h2>Action Plan</h2>
            <p>Service information not available.</p>
          </div>
        </main>
      );
    }

    return <ActionPlan context={this.props.context} userService={userService} filterMode="leader" />;
  }

  private renderCompany(): JSX.Element {
    return <Results context={this.props.context} department="Company" />;
  }

  private renderISS(): JSX.Element {
    return <Results context={this.props.context} department="ISS" hideActionPlan={true} />;
  }

  private renderIS(): JSX.Element {
    return (
      <ServiceResultDirectory
        context={this.props.context}
        title="Internal Support"
        resultDepartment="IS"
        divisions={['Internal Support']}
        onShowResult={this.openServiceResult}
      />
    );
  }

  private renderSS(): JSX.Element {
    return (
      <ServiceResultDirectory
        context={this.props.context}
        title="Sales Support"
        resultDepartment="SS"
        divisions={['Sales Support']}
        onShowResult={this.openServiceResult}
      />
    );
  }

  private renderDTS(): JSX.Element {
    return <Results context={this.props.context} department="DTS" hideActionPlan={true} />;
  }

  private renderDTSComponentManufacturing(): JSX.Element {
    return (
      <ServiceResultDirectory
        context={this.props.context}
        title="Component Manufacturing"
        resultDepartment="DTS"
        divisions={['Component Manufacturing']}
        onShowResult={this.openServiceResult}
      />
    );
  }

  private renderDTSEngineeringTechnology(): JSX.Element {
    return (
      <ServiceResultDirectory
        context={this.props.context}
        title="Engineering Technology"
        resultDepartment="DTS"
        divisions={['Engineering Technology']}
        onShowResult={this.openServiceResult}
      />
    );
  }

  private renderDTSEnterpriseApplications(): JSX.Element {
    return (
      <ServiceResultDirectory
        context={this.props.context}
        title="Enterprise Applications"
        resultDepartment="DTS"
        divisions={['Enterprise Applications']}
        onShowResult={this.openServiceResult}
      />
    );
  }

  private renderDTSESTechnology(): JSX.Element {
    return (
      <ServiceResultDirectory
        context={this.props.context}
        title="ES Technology"
        resultDepartment="DTS"
        divisions={['ES Technology']}
        onShowResult={this.openServiceResult}
      />
    );
  }

  private renderDTSResidential(): JSX.Element {
    return (
      <ServiceResultDirectory
        context={this.props.context}
        title="Residential"
        resultDepartment="DTS"
        divisions={['Residential']}
        onShowResult={this.openServiceResult}
      />
    );
  }

  private renderSelectedServiceResult(): JSX.Element {
    const { selectedResultDepartment, selectedResultService, selectedResultLabel } = this.state;
    if (!selectedResultDepartment || !selectedResultService) {
      return this.renderPlaceholder('Result');
    }

    return (
      <Results
        context={this.props.context}
        department={selectedResultDepartment}
        serviceFilter={selectedResultService}
        viewLabel={selectedResultLabel}
      />
    );
  }

  private renderRoleManagement(): JSX.Element {
    return <RoleManagement context={this.props.context} />;
  }

  private renderDashboard(): JSX.Element {
    const { currentUserDisplayName } = this.props;
    return <Dashboard currentUserDisplayName={currentUserDisplayName} />;
  }

  private renderPlaceholder(title: string): JSX.Element {
    return (
      <main className={styles.mainPlain}>
        <div className={styles.pageBreadcrumb}>Home › <strong>{title}</strong></div>
        <div style={{ padding: 40 }}>
          <h2>{title}</h2>
          <p>This page is under construction.</p>
        </div>
      </main>
    );
  }

  private renderCurrentView(): JSX.Element {
    switch (this.state.currentView) {
      case 'home':       return this.renderHome();
      case 'about':      return this.renderAbout();
      case 'actionplan': return this.renderActionPlan();
      case 'dashboard':  return this.renderDashboard();
      case 'company':    return this.renderCompany();
      case 'iss':        return this.renderISS();
      case 'is':         return this.renderIS();
      case 'ss':         return this.renderSS();
      case 'dts':        return this.renderDTS();
      case 'dtsComponentManufacturing': return this.renderDTSComponentManufacturing();
      case 'dtsEngineeringTechnology': return this.renderDTSEngineeringTechnology();
      case 'dtsEnterpriseApplications': return this.renderDTSEnterpriseApplications();
      case 'dtsESTechnology': return this.renderDTSESTechnology();
      case 'dtsResidential': return this.renderDTSResidential();
      case 'serviceResult': return this.renderSelectedServiceResult();
      case 'admin':      return this.renderRoleManagement();
      default:           return this.renderHome();
    }
  }

  public render(): React.ReactElement<ICustomerSatisfactionProps> {
    const { isSidebarCollapsed, currentView, isTeamLeader, userRole } = this.state;

    return (
      <section className={`${styles.customerSatisfaction} ${isSidebarCollapsed ? styles.collapsed : ''}`}>
        <Sidebar
          isTeamLeader={isTeamLeader}
          userRole={userRole}
          isSidebarCollapsed={isSidebarCollapsed}
          currentView={currentView}
          onToggleSidebar={() => this.setState(prev => ({ isSidebarCollapsed: !prev.isSidebarCollapsed }))}
          onNavigate={this.navigateTo}
        />
        {this.renderCurrentView()}
      </section>
    );
  }
}