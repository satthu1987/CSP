import * as React from 'react';
import styles from './Results.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Icon, Modal, Spinner, SpinnerSize } from '@fluentui/react';
import { AcDataService, IAcDataItem } from '../../services/AcDataService';
import { ActionPlanService } from '../../services/ActionPlan_Service';
import { IActionplan } from '../../Models/ActionplanModel';

export interface IResultsProps {
  context: WebPartContext;
  department: string; // "IS" | "SS" | "DTS" | "Company"
  serviceFilter?: string | string[];
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
  selectedActionPlan?: IActionplan;
  isDetailPanelOpen: boolean;
  isLoadingDetail: boolean;
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
      selectedActionPlan: undefined,
      isDetailPanelOpen: false,
      isLoadingDetail: false,
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
      JSON.stringify(prevProps.serviceFilter) !== JSON.stringify(this.props.serviceFilter)
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
      const serviceFilter = this.props.serviceFilter;
      const serviceNames = Array.isArray(serviceFilter)
        ? serviceFilter.filter(Boolean)
        : serviceFilter
          ? [serviceFilter]
          : [];
      // "Overall Result" is looked up from AC_Data by the selected service name
      // (e.g. "Application Automation"). When no specific service is selected
      // (top-level ESVN / Internal & Sales Support / Digital Technology Support
      // views), fall back to the department code itself (ESVN / ISS / DTS).
      const acDepartments = serviceNames.length > 0 ? serviceNames : [department];

      const [allAcItems, allActionPlans] = await Promise.all([
        this.acDataService.getAcDataByDepartments(acDepartments),
        fullDepartment
          ? this.actionPlanService.getActionPlansByDepartment(fullDepartment)
          : Promise.resolve([]),
      ]);

      const serviceFilterList = serviceNames.map(name => name.toLowerCase());
      console.log("Service Filter List: " + serviceFilterList);

      const filteredActionPlans = serviceFilterList.length > 0
        ? allActionPlans.filter(item => serviceFilterList.indexOf((item.Service || '').toLowerCase()) > -1)
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

  private openDetailPanel = async (plan: IActionplan): Promise<void> => {
    this.setState({ isDetailPanelOpen: true, isLoadingDetail: true, selectedActionPlan: plan });
    try {
      const fullPlan = await this.actionPlanService.getActionPlanById(plan.Id);
      this.setState({ selectedActionPlan: fullPlan || plan, isLoadingDetail: false });
    } catch (error) {
      console.error('Error loading action plan detail:', error);
      this.setState({ isLoadingDetail: false });
    }
  };

  private closeDetailPanel = (): void => {
    this.setState({
      isDetailPanelOpen: false,
      selectedActionPlan: undefined,
      isLoadingDetail: false,
    });
  };

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
              <div className={styles.colPIC}>PIC</div>
              <div className={styles.colService}>Action</div>
              <div className={styles.colTimeline}>Timeline</div>
              <div className={styles.colStatus}>Status</div>
              <div className={styles.colResult}>Results</div>
              <div className={styles.colLink}>Link</div>
              <div className={styles.colAction} />
            </div>
            {actionPlans.map(plan => (
              <div key={plan.Id} className={styles.planRow}>
                <div className={styles.colTitle}>{this.getGridPreviewText(plan.UpdatedFeedback) || '—'}</div>
                <div className={styles.colPIC}>{plan.PIC?.Title || '—'}</div>
                <div className={styles.colService}>{this.getGridPreviewText(plan.Actions) || '—'}</div>
                <div className={styles.colTimeline}>{this.formatDate(plan.Timeline)}</div>
                <div className={styles.colStatus}>
                  <span className={`${styles.badge} ${this.getStatusClass(plan.Status)}`}>
                    {plan.Status || 'Open'}
                  </span>
                </div>
                <div className={styles.colResult}>{this.getGridPreviewText(plan.Results) || '—'}</div>
                <div className={styles.colLink}>
                  {plan.RelatedLinks ? <a href={this.htmlToPlainText(plan.RelatedLinks) || '—'}>View</a> : "-" }
                </div>
                <div className={styles.colAction}>
                  <Icon
                    iconName="View"
                    className={styles.editIcon}
                    onClick={() => this.openDetailPanel(plan)}
                    title="View Detail"
                  />
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
                  {item.Data ? (
                    <div className={styles.acText}><img className={styles.acImg} src={item.Data} /></div>
                  ) : (
                    <div className={styles.emptyMessage}>No data available for {item.Year}.</div>
                  )}
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    );
  }

  private renderDetailPanel(): JSX.Element {
    const { selectedActionPlan, isDetailPanelOpen, isLoadingDetail } = this.state;
    if (!isDetailPanelOpen || !selectedActionPlan) return <></>;

    return (
      <Modal
        isOpen={isDetailPanelOpen}
        onDismiss={this.closeDetailPanel}
        isBlocking={false}
        containerClassName={styles.modalContainer}
      >
        <div className={styles.detailPanel}>
          <div className={styles.panelHeader}>
            <h2>Action Plan Details</h2>
            <button className={styles.closeBtn} onClick={this.closeDetailPanel}>
              <Icon iconName="Cancel" />
            </button>
          </div>

          {isLoadingDetail ? (
            <div className={styles.spinnerContainer}>
              <Spinner size={SpinnerSize.medium} label="Loading details..." />
            </div>
          ) : (
            <div className={styles.panelBody}>
              <div className={styles.formGroup}>
                <label>Customer Feedback</label>
                <div className={styles.readOnlyValue} dangerouslySetInnerHTML={{ __html: selectedActionPlan.UpdatedFeedback || '—' }} />
              </div>
              <div className={styles.formGroup}>
                <label>Service</label>
                <div className={styles.readOnlyText}>{selectedActionPlan.Service || '—'}</div>
              </div>
              <div className={styles.formGroup}>
                <label>PIC</label>
                <div className={styles.readOnlyText}>{selectedActionPlan.PIC?.Title || '—'}</div>
              </div>
              <div className={styles.formGroup}>
                <label>Actions</label>
                <div className={styles.readOnlyValue} dangerouslySetInnerHTML={{ __html: selectedActionPlan.Actions || '—' }} />
              </div>
              <div className={styles.formGroup}>
                <label>Timeline</label>
                <div className={styles.readOnlyText}>{this.formatDate(selectedActionPlan.Timeline)}</div>
              </div>
              <div className={styles.formGroup}>
                <label>Status</label>
                <span className={`${styles.badge} ${this.getStatusClass(selectedActionPlan.Status)}`}>
                  {selectedActionPlan.Status || 'Open'}
                </span>
              </div>
              <div className={styles.formGroup}>
                <label>Results</label>
                <div className={styles.readOnlyValue} dangerouslySetInnerHTML={{ __html: selectedActionPlan.Results || '—' }} />
              </div>
              <div className={styles.formGroup}>
                <label>Related Links</label>
                <div className={styles.readOnlyText}>
                  {selectedActionPlan.RelatedLinks ? <a href={selectedActionPlan.RelatedLinks}>{selectedActionPlan.RelatedLinks}</a> : "-" }
                </div>
              </div>
            </div>
          )}
        </div>
      </Modal>
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
              <label htmlFor="results-year-filter">Please select a year to view data</label>
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
        {this.renderDetailPanel()}
      </main>
    );
  }
}
