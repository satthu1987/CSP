import * as React from 'react';
import { Spinner, SpinnerSize, Icon } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import styles from './DTSServiceDirectory.module.scss';
import { DivisionServiceService } from '../../services/DivisionService_Service';
import { AcDataService, IAcDataItem } from '../../services/AcDataService';

export interface IDTSServiceDirectoryProps {
  context: WebPartContext;
  onShowResult: (department: string, service: string | string[], label: string) => void;
}

interface IDTSServiceItem {
  Id?: number;
  Service: string;
  PIC?: string;
  Manager?: string;
}

interface IDTSDivisionGroup {
  division: string;
  services: IDTSServiceItem[];
}

interface IDTSServiceDirectoryState {
  isLoading: boolean;
  groups: IDTSDivisionGroup[];
  expandedDivisions: { [division: string]: boolean };
  loadingServiceKey?: string;
  selectedYear: string;
  availableYears: string[];
  acItems: IAcDataItem[];
  isLoadingResult: boolean;
}

const DTS_DIVISIONS: string[] = [
  'Digital Technology Support',
  'Component Manufacturing',
  'Engineering Technology',
  'Enterprise Applications',
  'ES Technology',
  'Residential'
];

const RESULT_DEPARTMENT = 'DTS';

export default class DTSServiceDirectory extends React.Component<
  IDTSServiceDirectoryProps,
  IDTSServiceDirectoryState
> {
  private divisionServiceService: DivisionServiceService;
  private acDataService: AcDataService;

  constructor(props: IDTSServiceDirectoryProps) {
    super(props);
    const currentYear = new Date().getFullYear().toString();
    const expandedDivisions: { [division: string]: boolean } = {};
    DTS_DIVISIONS.forEach(division => { expandedDivisions[division] = true; });

    this.state = {
      isLoading: true,
      groups: [],
      expandedDivisions,
      selectedYear: currentYear,
      availableYears: [currentYear],
      acItems: [],
      isLoadingResult: true,
    };
    this.divisionServiceService = new DivisionServiceService(props.context, 'Division_Service');
    this.acDataService = new AcDataService(props.context);
  }

  public async componentDidMount(): Promise<void> {
    const selectedYear = await this.loadYears();
    await this.loadItems(selectedYear);
  }

  private loadYears = async (): Promise<string> => {
    const currentYear = new Date().getFullYear().toString();
    try {
      const years = await this.divisionServiceService.getAllYears();
      const availableYears = years.length > 0 ? years : [currentYear];
      const selectedYear = availableYears.indexOf(currentYear) > -1 ? currentYear : availableYears[0];
      this.setState({ availableYears, selectedYear });
      return selectedYear;
    } catch (error) {
      console.error('Failed to load Division_Service years:', error);
      return currentYear;
    }
  };

  private loadItems = async (year: string): Promise<void> => {
    this.setState({ isLoading: true, isLoadingResult: true });

    try {
      const acDataPromise = this.acDataService.getAcDataByDepartment(RESULT_DEPARTMENT);
      const groups: IDTSDivisionGroup[] = [];

      for (const division of DTS_DIVISIONS) {
        const services = await this.divisionServiceService.getServicesWithIdByDivision(division, year);
        if (services.length > 0) {
          groups.push({ division, services });
        }
      }

      const acItems = await acDataPromise;

      this.setState({
        isLoading: false,
        isLoadingResult: false,
        groups,
        acItems: acItems.filter(item => (item.Year || '').toString() === year),
      });
    } catch (error) {
      console.error('Failed to load Digital Technology Support directory:', error);
      this.setState({ isLoading: false, isLoadingResult: false, groups: [], acItems: [] });
    }
  };

  private onYearFilterChange = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    const selectedYear = event.target.value;
    this.setState({ selectedYear });
    this.loadItems(selectedYear).catch(error => {
      console.error('Error changing year filter:', error);
    });
  };

  private toggleDivision = (division: string): void => {
    this.setState(prev => ({
      expandedDivisions: {
        ...prev.expandedDivisions,
        [division]: !prev.expandedDivisions[division]
      }
    }));
  };

  private handleShowResult = async (division: string, item: IDTSServiceItem): Promise<void> => {
    const key = `${division}-${item.Service}`;
    this.setState({ loadingServiceKey: key });

    try {
      const serviceNames = item.Id
        ? await this.divisionServiceService.getServiceNameHistory(item.Id)
        : [];

      const services = serviceNames.length > 0 ? serviceNames : [item.Service];
      this.props.onShowResult(RESULT_DEPARTMENT, services, item.Service);
    } catch (error) {
      console.error('Failed to load service name history:', error);
      this.props.onShowResult(RESULT_DEPARTMENT, item.Service, item.Service);
    } finally {
      this.setState({ loadingServiceKey: undefined });
    }
  };

  private renderOverallResultSection = (): JSX.Element => {
    const { acItems, isLoadingResult, selectedYear } = this.state;

    return (
      <div className={styles.section}>
        <h2 className={styles.sectionTitle}>Digital Technology Support - {selectedYear} Overall Result</h2>

        {isLoadingResult ? (
          <div className={styles.spinnerContainer}>
            <Spinner size={SpinnerSize.large} label="Loading results..." />
          </div>
        ) : acItems.length === 0 ? (
          <div className={styles.emptyMessage}>No result data available for Digital Technology Support.</div>
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
  };

  public render(): JSX.Element {
    const { isLoading, groups, expandedDivisions, loadingServiceKey, selectedYear, availableYears } = this.state;

    return (
      <main className={styles.mainContainer}>
        <div className={styles.breadcrumb}>
          Home {'>'} <strong>Digital Technology Support</strong>
        </div>

        <div className={styles.header}>
          <h1>Digital Technology Support</h1>
        </div>

        <div className={styles.content}>
          <div className={styles.filterSection}>
            <div className={styles.filterGroup}>
              <label htmlFor="dts-directory-year-filter">Please select a year to view data</label>
              <select
                id="dts-directory-year-filter"
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

          {this.renderOverallResultSection()}

          <div className={styles.gridContainer}>
            {isLoading ? (
              <div className={styles.spinnerContainer}>
                <Spinner size={SpinnerSize.large} label="Loading services..." />
              </div>
            ) : groups.length === 0 ? (
              <div className={styles.emptyMessage}>No services found.</div>
            ) : (
              groups.map(group => {
                const isExpanded = !!expandedDivisions[group.division];
                return (
                  <div className={styles.divisionGroup} key={group.division}>
                    <button
                      type="button"
                      className={styles.divisionHeader}
                      onClick={() => this.toggleDivision(group.division)}
                      aria-expanded={isExpanded}
                    >
                      <Icon
                        iconName={isExpanded ? 'ChevronDown' : 'ChevronRight'}
                        className={styles.chevron}
                      />
                      <span className={styles.divisionTitle}>{group.division}</span>
                      <span className={styles.divisionCount}>{group.services.length}</span>
                    </button>

                    {isExpanded && (
                      <div className={styles.serviceTable}>
                        <div className={styles.serviceHeaderRow}>
                          <div>Service</div>
                          <div>Service/Project Lead</div>
                          {/* <div>Manager</div> */}
                          <div>Result</div>
                        </div>
                        {group.services.map((item, index) => {
                          const key = `${group.division}-${item.Service}`;
                          const isRowLoading = loadingServiceKey === key;

                          return (
                            <div className={styles.serviceRow} key={`${key}-${index}`}>
                              <div>{item.Service}</div>
                              <div>{item.PIC || '-'}</div>
                              {/* <div>{item.Manager || '-'}</div> */}
                              <div>
                                <button
                                  type="button"
                                  className={styles.actionLink}
                                  disabled={isRowLoading}
                                  onClick={() => {
                                    this.handleShowResult(group.division, item).catch(error => {
                                      console.error('Error showing result:', error);
                                    });
                                  }}
                                >
                                  {isRowLoading ? 'Loading...' : 'Show Result'}
                                </button>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    )}
                  </div>
                );
              })
            )}
          </div>
        </div>
      </main>
    );
  }
}
