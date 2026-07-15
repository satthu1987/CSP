import * as React from 'react';
import styles from './Sidebar.module.scss';
import { Icon, TooltipHost } from '@fluentui/react';

export type ViewName = 'home' | 'about' | 'actionplan' | 'customerfeedback' | 'dashboard' | 'admin';

export interface ISidebarProps {
  isTeamLeader: boolean;
  isSidebarCollapsed: boolean;
  currentView: ViewName;
  onToggleSidebar: () => void;
  onNavigate: (view: ViewName) => void;
}

export default class Sidebar extends React.Component<ISidebarProps> {
  private renderNavItem = (
    iconName: string,
    label: string,
    view: ViewName
  ): JSX.Element => {
    const { isSidebarCollapsed, currentView, onNavigate } = this.props;
    const active = currentView === view;
    const itemClass = `${styles.navItem} ${active ? styles.active : ''}`;

    const content = (
      <a
        className={itemClass}
        key={label}
        onClick={() => onNavigate(view)}
      >
        <Icon iconName={iconName} />
        {!isSidebarCollapsed && <span>{label}</span>}
      </a>
    );

    return isSidebarCollapsed
      ? <TooltipHost content={label} key={label}>{content}</TooltipHost>
      : content;
  };

  public render(): JSX.Element {
    const { isSidebarCollapsed, isTeamLeader, onToggleSidebar } = this.props;

    return (
      <aside className={`${styles.sidebar} ${isSidebarCollapsed ? styles.collapsed : ''}`}>
        <div className={styles.brand}>
          <div className={styles.brandLogo}>CSP</div>
          {!isSidebarCollapsed && (
            <span className={styles.brandTitle}>Customer Satisfaction Program</span>
          )}
        </div>

        <button
          className={styles.toggleBtn}
          onClick={onToggleSidebar}
          title={isSidebarCollapsed ? 'Expand menu' : 'Collapse menu'}
        >
          <Icon iconName={isSidebarCollapsed ? 'ChevronRight' : 'ChevronLeft'} />
        </button>

        <nav className={styles.nav}>
          <div className={styles.navGroup}>
            {!isSidebarCollapsed && <div className={styles.navGroupTitle}>Home</div>}
            {this.renderNavItem('Home', 'Home', 'home')}
            {this.renderNavItem('Info', 'About', 'about')}
          </div>

          {isTeamLeader && (
            <>
              <div className={styles.navGroup}>
                {!isSidebarCollapsed && <div className={styles.navGroupTitle}>Personal</div>}
                {this.renderNavItem('Edit', 'Action Plan', 'actionplan')}
              {this.renderNavItem('Feedback', 'Customer Feedback', 'customerfeedback')}
              </div>
              <div className={styles.navGroup}>
                {!isSidebarCollapsed && <div className={styles.navGroupTitle}>Administration</div>}
                {this.renderNavItem('ViewDashboard', 'Dashboard', 'dashboard')}
                {this.renderNavItem('Admin', 'Admin', 'admin')}
              </div>
            </>
          )}
        </nav>

        <button className={styles.signOutBtn}>
          <Icon iconName="SignOut" />
          {!isSidebarCollapsed && <span>Sign out</span>}
        </button>
      </aside>
    );
  }
}
