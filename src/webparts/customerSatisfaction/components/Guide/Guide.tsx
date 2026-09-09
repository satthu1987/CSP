import * as React from 'react';
import styles from './Guide.module.scss';
import { Icon } from '@fluentui/react';

interface IGuideSection {
  id: string;
  title: string;
  content: JSX.Element;
}

interface IGuideState {
  expandedIds: { [key: string]: boolean };
}

export default class Guide extends React.Component<{}, IGuideState> {
  private sections: IGuideSection[];

  constructor(props: {}) {
    super(props);
    this.state = {
      expandedIds: { gettingStarted: true }
    };
    this.sections = [
      {
        id: 'gettingStarted',
        title: 'Getting Started',
        content: (
          <div>
            <p>
              The Customer Satisfaction Program (CSP) application measures customer satisfaction and tracks
              improvement actions for the services provided by Enterprise Services Vietnam (ESVN).
            </p>
            <p>Use the menu on the left to move between the available areas of the application:</p>
            <ul>
              <li><strong>Home</strong> – the program overview page.</li>
              <li><strong>Learn More</strong> – background information about the program.</li>
              <li><strong>User Guide</strong> – this page.</li>
              <li><strong>Survey Results</strong> – satisfaction results by area and service.</li>
              <li><strong>Comment & Action Plan</strong> – manage action plans (team leaders, managers and admins).</li>
              <li><strong>Administration</strong> – role and division/service management (admins only).</li>
            </ul>
            <p>
              The menu can be collapsed by clicking the collapse button to gain more screen space; hover a
              collapsed menu item to see its name.
            </p>
          </div>
        )
      },
      {
        id: 'surveyResults',
        title: 'Viewing Survey Results',
        content: (
          <div>
            <p>
              Open <strong>Survey Results</strong> in the menu and choose the area you want to review:
            </p>
            <ul>
              <li><strong>ESVN</strong> – company-wide results.</li>
              <li><strong>Internal & Sales Support</strong> – Internal Support or Sales Support.</li>
              <li><strong>Digital Technology Support</strong> – Component Manufacturing, Engineering Technology, Enterprise Applications, ES Technology or Residential.</li>
            </ul>
            <p>
              Each result page shows the <strong>Overall Result</strong> chart for the selected year. Use the
              <strong> Year</strong> filter at the top to view data from previous years.
            </p>
            <p>
              The <strong>Action Plans</strong> section below the chart lists the plans recorded for that area.
              Click the view icon on a row to open the plan details in a read-only popup.
            </p>
          </div>
        )
      },
      {
        id: 'actionPlans',
        title: 'Working with Action Plans',
        content: (
          <div>
            <p>
              The <strong>Action Plan</strong> page is available to team leaders, managers and admins. It lists
              every action plan you are responsible for.
            </p>
            <ul>
              <li>Use the <strong>Service</strong>, <strong>Year</strong>, <strong>PIC</strong> and <strong>Status</strong> filters to narrow down the list, then <strong>Clear Filters</strong> to reset them.</li>
              <li>The grid shows Service, Original Customer Feedback, Action, PIC, Status, Result and Related Links for each plan.</li>
              <li>Click the <strong>+</strong> icon to create the action for a plan that has none, or the edit icon to update an existing plan.</li>
              <li>Use <strong>New Action</strong> in the top-right corner to create a brand new action plan.</li>
            </ul>
            <p>
              When creating a plan, select the <strong>Department</strong> first. For Digital Technology Support
              you must also select the <strong>Product Line</strong> before choosing the <strong>Service</strong>.
              Fill in the feedback, action, PIC, category, status, result and related links, then save.
            </p>
          </div>
        )
      },
      {
        id: 'administration',
        title: 'Administration (admins only)',
        content: (
          <div>
            <p>
              <strong>Role Management</strong> – manage CSP users. Assign each user a role (leader, manager or
              admin), the services they lead and the divisions they manage.
            </p>
            <p>
              <strong>Division – Service Management</strong> – maintain the list of divisions and services per
              year. Use the <strong>Year</strong> filter to switch years and the edit icon on a row to update the
              division, service, year, <strong>Service/Project Lead</strong> and <strong>Manager</strong>.
            </p>
          </div>
        )
      },
      {
        id: 'roles',
        title: 'Roles and Access',
        content: (
          <div>
            <p>The menu items you see depend on your role in the CSP program:</p>
            <ul>
              <li><strong>Visitor</strong> – Home, Learn More, User Guide and Survey Results.</li>
              <li><strong>Leader</strong> – visitor items plus the Action Plan page for the services they lead.</li>
              <li><strong>Manager</strong> – visitor items plus the Action Plan page for the services they manage.</li>
              <li><strong>Admin</strong> – all pages, including Role Management and Division – Service Management, and every action plan.</li>
            </ul>
            <p>
              If a page you expect is missing, contact a CSP admin to have your role checked in
              Role Management.
            </p>
          </div>
        )
      }
    ];
  }

  private toggleSection = (id: string): void => {
    this.setState(prev => ({
      expandedIds: { ...prev.expandedIds, [id]: !prev.expandedIds[id] }
    }));
  };

  private renderSection = (section: IGuideSection): JSX.Element => {
    const expanded = !!this.state.expandedIds[section.id];
    return (
      <div className={styles.collapsibleSection} key={section.id}>
        <button
          className={styles.sectionHeader}
          onClick={() => this.toggleSection(section.id)}
          aria-expanded={expanded}
        >
          <Icon
            iconName={expanded ? 'ChevronDown' : 'ChevronRight'}
            className={styles.chevron}
          />
          <span className={styles.sectionTitle}>{section.title}</span>
        </button>
        {expanded && (
          <div className={styles.sectionBody}>
            {section.content}
          </div>
        )}
      </div>
    );
  };

  public render(): JSX.Element {
    return (
      <main className={styles.main}>
        <section className={styles.informationSection}>
          <h2 className={styles.pageTitle}>User Guide</h2>
          <p className={styles.intro}>
            This guide explains how to use the Customer Satisfaction Program application. Click a section
            below to expand it.
          </p>
          {this.sections.map(section => this.renderSection(section))}
        </section>
      </main>
    );
  }
}
