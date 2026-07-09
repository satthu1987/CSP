import * as React from 'react';
import styles from './CustomerSatisfaction.module.scss';
import { ICustomerSatisfactionProps } from './ICustomerSatisfactionProps';
import { ICustomerSatisfactionState, ViewName } from './ICustomerSatisfactionState';
import { UserRoleService } from '../services/UserRole_Service';
import Sidebar from './SideBar/Sidebar';
import Home from './Home/Home';
import About from './About/About';
import ActionPlan from './ActionPlan/ActionPlan';
import Dashboard from './Dashboard/Dashboard';

export default class CustomerSatisfaction
  extends React.Component<ICustomerSatisfactionProps, ICustomerSatisfactionState> {

  private userRoleService: UserRoleService;

  constructor(props: ICustomerSatisfactionProps) {
    super(props);
    this.state = {
      isTeamLeader: false,
      isLoading: true,
      sentToday: 0,
      receivedToday: 0,
      isSidebarCollapsed: false,
      currentView: 'home',
      userService: ''
    };
    this.userRoleService = new UserRoleService(props.context, 'RoleInService');
  }

  public async componentDidMount(): Promise<void> {
    const userRole = await this.userRoleService.getUserRole(this.props.currentUserEmail);
    const isTeamLeader = userRole !== undefined;
    const userService = userRole?.Title || '';
    this.setState({ isTeamLeader, isLoading: false, userService });
  }

  private navigateTo = (view: ViewName): void => {
    this.setState({ currentView: view });
  };

  private renderHome(): JSX.Element {
    const { currentUserDisplayName } = this.props;

    return (
      <Home
        currentUserDisplayName={currentUserDisplayName}
      />
    );
  }

  private renderAbout(): JSX.Element {
    return (
      <main className={styles.mainPlain}>
        <div className={styles.pageBreadcrumb}>
          Home › <strong>About</strong>
        </div>
        <About context={this.props.context} />
      </main>
    );
  }

  private renderActionPlan(): JSX.Element {
    const { userService } = this.state;
    
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

    return <ActionPlan context={this.props.context} userService={userService} />;
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
      case 'admin':      return this.renderPlaceholder('Admin');
      default:           return this.renderHome();
    }
  }

  public render(): React.ReactElement<ICustomerSatisfactionProps> {
    const { isSidebarCollapsed, currentView, isTeamLeader } = this.state;

    return (
      <section className={`${styles.customerSatisfaction} ${isSidebarCollapsed ? styles.collapsed : ''}`}>
        <Sidebar
          isTeamLeader={isTeamLeader}
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