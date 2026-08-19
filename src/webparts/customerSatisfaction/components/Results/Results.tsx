import * as React from 'react';
import styles from './Results.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { AcDataService, IAcDataItem } from '../../services/AcDataService';
import { ActionPlanService } from '../../services/ActionPlan_Service';
import { IActionplan } from '../../Models/ActionplanModel';

export interface IResultsProps {
  context: WebPartContext;
  department: string; // "IS" | "SS" | "DTS" | "Company"
  serviceFilter?: string;
  viewLabel?: string;
  hideActionPlan?: boolean; // When true, hides the Action Plans section
}

interface IResultsState {
  acItems: IAcDataItem[];
  actionPlans: IActionplan[];
  isLoadingAc: boolean;
  isLoadingPlans: boolean;
  selectedYear: string;
  availableYears: string[];
}

// Maps sidebar route department to the full department name used in the CSP list
const DEPARTMENT_MAP: { [key: string]: string } = {
  IS: 'Internal Support',
  SS: 'Sales Support',
  ISS: 'ISS',
  DTS: 'Digital Technology Support',
};

export default class Results extends React.Component<IResultsProps, IResultsState> {
  private acDataService: AcDataService;
  private actionPlanService: ActionPlanService;

  constructor(props: IResultsProps) {
    const currentYear = new Date().getFullYear().toString();
    super(props);
    this.state = {
      acItems: [],
      actionPlans: [],
      isLoadingAc: true,
      isLoadingPlans: true,
      selectedYear: currentYear,
      availableYears: this.getFixedYearOptions(currentYear),
    };
    this.acDataService = new AcDataService(props.context);
    this.actionPlanService = new ActionPlanService(props.context, 'CSP');
  }

  public async componentDidMount(): Promise<void> {
    await this.loadData(this.props.department, this.state.selectedYear);
  }

  public async componentDidUpdate(prevProps: IResultsProps): Promise<void> {
    if (
      prevProps.department !== this.props.department ||
      prevProps.serviceFilter !== this.props.serviceFilter
    ) {
      const currentYear = new Date().getFullYear().toString();
      await this.loadData(this.props.department, currentYear);
    }
  }

  private loadData = async (department: string, selectedYear: string): Promise<void> => {
    this.setState({ isLoadingAc: true, isLoadingPlans: true });
    const fullDepartment = DEPARTMENT_MAP[department];
    const currentYear = new Date().getFullYear().toString();
    const availableYears = this.getFixedYearOptions(currentYear);
    const effectiveYear = availableYears.indexOf(selectedYear) > -1 ? selectedYear : currentYear;

    try {
      const [allAcItems, allActionPlans] = await Promise.all([
        this.acDataService.getAcDataByDepartment(department),
        fullDepartment
          ? this.actionPlanService.getActionPlansByDepartment(fullDepartment)
          : Promise.resolve([]),
      ]);

      const serviceFilter = this.props.serviceFilter;
      const filteredActionPlans = serviceFilter
        ? allActionPlans.filter(item => (item.Service || '').toLowerCase() === serviceFilter.toLowerCase())
        : allActionPlans;

      this.setState({
        acItems: this.filterAcItemsByYear(allAcItems, effectiveYear),
        actionPlans: this.filterActionPlansByYear(filteredActionPlans, effectiveYear),
        selectedYear: effectiveYear,
        availableYears,
        isLoadingAc: false,
        isLoadingPlans: false,
      });
    } catch (error) {
      console.error('Error loading results data:', error);
      this.setState({
        acItems: [],
        actionPlans: [],
        availableYears,
        selectedYear: effectiveYear,
        isLoadingAc: false,
        isLoadingPlans: false,
      });
    }
  };

  private handleYearChange = async (event: React.ChangeEvent<HTMLSelectElement>): Promise<void> => {
    const selectedYear = event.target.value;
    await this.loadData(this.props.department, selectedYear);
  };

  private onYearFilterChange = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    this.handleYearChange(event).catch(error => {
      console.error('Error changing year filter:', error);
    });
  };

  private filterAcItemsByYear(items: IAcDataItem[], year: string): IAcDataItem[] {
    return items.filter(item => (item.Year || '').toString() === year);
  }

  private getActionPlanYear(plan: IActionplan): string {
    const directYear = (plan.Year || '').toString().trim();
    if (directYear) {
      return directYear;
    }

    if (plan.Timeline) {
      const timelineDate = new Date(plan.Timeline);
      if (!isNaN(timelineDate.getTime())) {
        return timelineDate.getFullYear().toString();
      }
    }

    return '';
  }

  private filterActionPlansByYear(items: IActionplan[], year: string): IActionplan[] {
    return items.filter(item => this.getActionPlanYear(item) === year);
  }

  private getFixedYearOptions(currentYear: string): string[] {
    const yearNum = parseInt(currentYear, 10);
    return [
      (yearNum - 3).toString(),
      (yearNum - 2).toString(),
      (yearNum - 1).toString(),
      yearNum.toString(),
    ];
  }

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
  private htmlToPlainText(value: string | undefined): string {
    if (!value) {
      return '';
    }

    const htmlWithLineBreaks = value.replace(/<\s*br\s*\/?>/gi, '\n');
    const temp = document.createElement('div');
    temp.innerHTML = htmlWithLineBreaks;

    const plainText = (temp.textContent || temp.innerText || '')
      .replace(/\u00a0/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();

    return plainText;
  }
    private getGridPreviewText(value: string | undefined, maxLength: number = 50): string {
    const plainText = this.htmlToPlainText(value);
    if (!plainText) {
      return '-';
    }

    return plainText.length > maxLength
      ? `${plainText.substring(0, maxLength)}...`
      : plainText;
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
              <div className={styles.colTitle}>Customer Feedback</div>
              <div className={styles.colService}>Action</div>
              <div className={styles.colPIC}>PIC</div>
              <div className={styles.colTimeline}>Timeline</div>
              <div className={styles.colStatus}>Status</div>
            </div>
            {actionPlans.map(plan => (
              <div key={plan.Id} className={styles.planRow}>
                <div className={styles.colTitle}>{this.getGridPreviewText(plan.UpdatedFeedback) || '—'}</div>
                <div className={styles.colService}>{plan.Actions || '—'}</div>
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
    const { acItems, isLoadingAc } = this.state;
    const { department } = this.props;

    return (
      <div className={styles.section}>
        <h2 className={styles.sectionTitle}>Overall Result</h2>

        {isLoadingAc ? (
          <div className={styles.spinnerContainer}>
            <Spinner size={SpinnerSize.large} label={`Loading ${department} results...`} />
          </div>
        ) : acItems.length === 0 ? (
          <div className={styles.emptyMessage}>No result data available for {department}.</div>
        ) : (
          <div className={styles.acDataSection}>
            {acItems.map(item => (
              <div className={styles.acRow} key={item.Id}>
                <div className={styles.acHeader}>
                  <span className={styles.acTitle}>{item.Year} Overall Result</span>
                </div>

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
              </div>
            ))}
          </div>
        )}
      </div>
    );
  }

  public render(): JSX.Element {
    const { department, viewLabel } = this.props;
    const { selectedYear, availableYears } = this.state;
    const label = viewLabel || DEPARTMENT_MAP[department] || department;

    return (
      <main className={styles.mainContainer}>
        <div className={styles.breadcrumb}>
          Home › <strong>{label}</strong>
        </div>

        <div className={styles.header}>
          <h1>{label}</h1>
        </div>

        <div className={styles.content}>
          <div className={styles.filterSection}>
            <div className={styles.filterGroup}>
              <label htmlFor="results-year-filter">Year</label>
              <select
                id="results-year-filter"
                className={styles.filterSelect}
                value={selectedYear}
                onChange={this.onYearFilterChange}
              >
                {availableYears.map(year => (
                  <option key={year} value={year}>
                    {year}
                  </option>
                ))}
              </select>
            </div>
          </div>

          {this.renderResultsSection()}
          {!this.props.hideActionPlan && this.renderActionPlanGrid()}
        </div>
      </main>
    );
  }
}
