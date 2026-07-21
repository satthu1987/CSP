import * as React from 'react';
import styles from './Results.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Icon, Spinner, SpinnerSize } from '@fluentui/react';
import { AcDataService, IAcDataItem } from '../../services/AcDataService';
import { ActionPlanService } from '../../services/ActionPlan_Service';
import { IActionplan } from '../../Models/ActionplanModel';

export interface IResultsProps {
  context: WebPartContext;
  department: string; // "IS" | "SS" | "DTS" | "Company"
}

interface IResultsState {
  acItems: IAcDataItem[];
  actionPlans: IActionplan[];
  isLoadingAc: boolean;
  isLoadingPlans: boolean;
  expandedIds: { [id: number]: boolean };
}

// Maps sidebar route department to the full department name used in the CSP list
const DEPARTMENT_MAP: { [key: string]: string } = {
  IS: 'Internal Support',
  SS: 'Sales Support',
  DTS: 'Digital Transformation Support',
};

export default class Results extends React.Component<IResultsProps, IResultsState> {
  private acDataService: AcDataService;
  private actionPlanService: ActionPlanService;

  constructor(props: IResultsProps) {
    super(props);
    this.state = {
      acItems: [],
      actionPlans: [],
      isLoadingAc: true,
      isLoadingPlans: true,
      expandedIds: {},
    };
    this.acDataService = new AcDataService(props.context);
    this.actionPlanService = new ActionPlanService(props.context, 'CSP');
  }

  public async componentDidMount(): Promise<void> {
    await this.loadData(this.props.department);
  }

  public async componentDidUpdate(prevProps: IResultsProps): Promise<void> {
    if (prevProps.department !== this.props.department) {
      this.setState({ expandedIds: {} });
      await this.loadData(this.props.department);
    }
  }

  private loadData = async (department: string): Promise<void> => {
    this.setState({ isLoadingAc: true, isLoadingPlans: true });
    const fullDepartment = DEPARTMENT_MAP[department];

    await Promise.all([
      this.acDataService.getAcDataByDepartment(department).then(acItems =>
        this.setState({ acItems, isLoadingAc: false })
      ),
      fullDepartment
        ? this.actionPlanService.getActionPlansByDepartment(fullDepartment).then(actionPlans =>
            this.setState({ actionPlans, isLoadingPlans: false })
          )
        : Promise.resolve(this.setState({ actionPlans: [], isLoadingPlans: false })),
    ]);
  };

  private toggleSection = (id: number): void => {
    this.setState(prev => ({
      expandedIds: { ...prev.expandedIds, [id]: !prev.expandedIds[id] },
    }));
  };

  private formatDate(dateStr: string): string {
    if (!dateStr) return '—';
    try {
      return new Date(dateStr).toLocaleDateString('en-GB', {
        day: '2-digit',
        month: 'short',
        year: 'numeric',
      });
    } catch {
      return dateStr;
    }
  }

  private getStatusClass(status: string): string {
    switch ((status || '').toLowerCase()) {
      case 'open': return styles.statusOpen;
      case 'in progress': return styles.statusInProgress;
      case 'closed': return styles.statusClosed;
      default: return styles.statusOpen;
    }
  }

  private renderActionPlanGrid(): JSX.Element {
    const { actionPlans, isLoadingPlans } = this.state;
    const { department } = this.props;

    if (!DEPARTMENT_MAP[department]) return <></>;

    return (
      <div className={styles.section}>
        <h2 className={styles.sectionTitle}>Action Plans</h2>

        {isLoadingPlans ? (
          <div className={styles.spinnerContainer}>
            <Spinner size={SpinnerSize.medium} label="Loading action plans..." />
          </div>
        ) : actionPlans.length === 0 ? (
          <div className={styles.emptyMessage}>No action plans found for this department.</div>
        ) : (
          <div className={styles.planGrid}>
            <div className={styles.planHeader}>
              <div className={styles.colTitle}>Title</div>
              <div className={styles.colService}>Service</div>
              <div className={styles.colPIC}>PIC</div>
              <div className={styles.colTimeline}>Timeline</div>
              <div className={styles.colStatus}>Status</div>
            </div>
            {actionPlans.map(plan => (
              <div key={plan.Id} className={styles.planRow}>
                <div className={styles.colTitle}>{plan.Title || '—'}</div>
                <div className={styles.colService}>{plan.Service || '—'}</div>
                <div className={styles.colPIC}>{plan.PIC?.Title || '—'}</div>
                <div className={styles.colTimeline}>{this.formatDate(plan.Timeline)}</div>
                <div className={styles.colStatus}>
                  <span className={`${styles.badge} ${this.getStatusClass(plan.Status)}`}>
                    {plan.Status || 'Open'}
                  </span>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    );
  }

  private renderResultsSection(): JSX.Element {
    const { acItems, isLoadingAc, expandedIds } = this.state;
    const { department } = this.props;

    return (
      <div className={styles.section}>
        <h2 className={styles.sectionTitle}>Results</h2>

        {isLoadingAc ? (
          <div className={styles.spinnerContainer}>
            <Spinner size={SpinnerSize.large} label={`Loading ${department} results...`} />
          </div>
        ) : acItems.length === 0 ? (
          <div className={styles.emptyMessage}>No result data available for {department}.</div>
        ) : (
          <div className={styles.acDataSection}>
            {acItems.map(item => {
              const expanded = !!expandedIds[item.Id];
              return (
                <div className={styles.acRow} key={item.Id}>
                  <button
                    className={styles.acHeader}
                    onClick={() => this.toggleSection(item.Id)}
                    aria-expanded={expanded}
                  >
                    <Icon
                      iconName={expanded ? 'ChevronDown' : 'ChevronRight'}
                      className={styles.acChevron}
                    />
                    <span className={styles.acTitle}>{item.Year} Result</span>
                  </button>

                  {expanded && (
                    <div className={styles.acBody}>
                      {item.DataUrl ? (
                        <img
                          src={item.DataUrl}
                          alt={item.DataAlt}
                          className={styles.acImage}
                          onError={(e) => {
                            (e.target as HTMLImageElement).style.display = 'none';
                          }}
                        />
                      ) : (
                        <div className={styles.emptyMessage}>No image available for {item.Year}.</div>
                      )}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        )}
      </div>
    );
  }

  public render(): JSX.Element {
    const { department } = this.props;
    const label = DEPARTMENT_MAP[department] || department;

    return (
      <main className={styles.mainContainer}>
        <div className={styles.breadcrumb}>
          Home › <strong>{label}</strong>
        </div>

        <div className={styles.header}>
          <h1>{label}</h1>
        </div>

        <div className={styles.content}>
          {this.renderActionPlanGrid()}
          {this.renderResultsSection()}
        </div>
      </main>
    );
  }
}
