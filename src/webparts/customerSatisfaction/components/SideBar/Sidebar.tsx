import * as React from 'react';
import styles from './Sidebar.module.scss';
import { Icon, TooltipHost } from '@fluentui/react';

export type ViewName = 'home' | 'about' | 'actionplan' | 'customerfeedback' | 'dashboard' | 'admin' | 'company' | 'iss' | 'is' | 'ss' | 'dts';

export interface ISidebarProps {
  isTeamLeader: boolean;
  isSidebarCollapsed: boolean;
  currentView: ViewName;
  onToggleSidebar: () => void;
  onNavigate: (view: ViewName) => void;
}

interface ISidebarState {
  expandedMenus: string[];
}

export default class Sidebar extends React.Component<ISidebarProps, ISidebarState> {
  constructor(props: ISidebarProps) {
    super(props);
    this.state = {
      expandedMenus: []
    };
  }

  private toggleSubmenu = (menuName: string): void => {
    const expanded = [...this.state.expandedMenus];
    const index = expanded.indexOf(menuName);
    if (index > -1) {
      expanded.splice(index, 1);
    } else {
      expanded.push(menuName);
    }
    this.setState({ expandedMenus: expanded });
  };

  private renderSubmenuItem = (
    label: string,
    view: ViewName
  ): JSX.Element => {
    const { isSidebarCollapsed, currentView, onNavigate } = this.props;
    const active = currentView === view;
    const itemClass = `${styles.submenuItem} ${active ? styles.active : ''}`;

    return (
      <a
        className={itemClass}
        key={label}
        onClick={() => onNavigate(view)}
      >
        {!isSidebarCollapsed && <span>{label}</span>}
      </a>
    );
  };

  private renderNavItemWithSubmenu = (
    iconName: string,
    label: string,
    view: ViewName,
    submenuItems?: Array<{ label: string; view: ViewName }>
  ): JSX.Element => {
    const { isSidebarCollapsed, currentView, onNavigate } = this.props;
    const active = currentView === view || (submenuItems && submenuItems.some(item => currentView === item.view));
    const isExpanded = this.state.expandedMenus.indexOf(label) > -1;
    const itemClass = `${styles.navItem} ${active ? styles.active : ''}`;

    const content = (
      <div key={label}>
        <a
          className={itemClass}
          onClick={() => {
            if (submenuItems && !isSidebarCollapsed) {
              this.toggleSubmenu(label);
            } else if (submenuItems && isSidebarCollapsed) {
              // When collapsed, navigate directly
              onNavigate(view);
            } else {
              onNavigate(view);
            }
          }}
        >
          <Icon iconName={iconName} />
          {!isSidebarCollapsed && (
            <>
              <span>{label}</span>
              {submenuItems && !isSidebarCollapsed && (
                <Icon
                  iconName={isExpanded ? 'ChevronDown' : 'ChevronRight'}
                  style={{ marginLeft: 'auto', fontSize: '12px' }}
                />
              )}
            </>
          )}
        </a>
        {submenuItems && isExpanded && !isSidebarCollapsed && (
          <div className={styles.submenu}>
            {submenuItems.map(item => this.renderSubmenuItem(item.label, item.view))}
          </div>
        )}
      </div>
    );

    return isSidebarCollapsed && submenuItems
      ? <TooltipHost content={label} key={label}>{content}</TooltipHost>
      : content;
  };

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
            {this.renderNavItem('Info', 'Learn More', 'about')}
          </div>
          <div className={styles.navGroup}>
            {!isSidebarCollapsed && <div className={styles.navGroupTitle}>RESULTS</div>}
            {this.renderNavItem('Home', 'Company', 'company')}
            {this.renderNavItemWithSubmenu('Info', 'ISS', 'iss', [
              { label: 'IS', view: 'is' },
              { label: 'SS', view: 'ss' }
            ])}
            {this.renderNavItem('Info', 'DTS', 'dts')}
          </div>

          {isTeamLeader && (
            <>
              <div className={styles.navGroup}>
                {!isSidebarCollapsed && <div className={styles.navGroupTitle}>Comment & Action Plan</div>}
                {this.renderNavItem('Edit', 'Action Plan', 'actionplan')}
                {this.renderNavItem('Feedback', 'Customer Feedback', 'customerfeedback')}
              </div>
              <div className={styles.navGroup}>
                {!isSidebarCollapsed && <div className={styles.navGroupTitle}>Administration</div>}
                {this.renderNavItem('ViewDashboard', 'Dashboard', 'dashboard')}
                {this.renderNavItem('Admin', 'Role Management', 'admin')}
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
