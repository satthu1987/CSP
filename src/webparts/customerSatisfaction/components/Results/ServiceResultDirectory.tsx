import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import styles from './ServiceResultDirectory.module.scss';
import { DivisionServiceService } from '../../services/DivisionService_Service';
import { AcDataService, IAcDataItem } from '../../services/AcDataService';

export interface IServiceResultDirectoryProps {
  context: WebPartContext;
  title: string;
  resultDepartment: string;
  divisions: string[];
  getResultDepartmentByDivision?: (division: string) => string;
  onShowResult: (department: string, service: string | string[], label: string) => void;
}

interface IServiceResultDirectoryItem {
  Id?: number;
  Division: string;
  Service: string;
  PIC?: string;
  Manager?: string;
}

interface IServiceResultDirectoryState {
  isLoading: boolean;
  items: IServiceResultDirectoryItem[];
  loadingServiceKey?: string;
  selectedYear: string;
  availableYears: string[];
  acItems: IAcDataItem[];
  isLoadingResult: boolean;
}

export default class ServiceResultDirectory extends React.Component<
  IServiceResultDirectoryProps,
  IServiceResultDirectoryState
> {
  private divisionServiceService: DivisionServiceService;
  private acDataService: AcDataService;

  constructor(props: IServiceResultDirectoryProps) {
    super(props);
    const currentYear = new Date().getFullYear().toString();
    this.state = {
      isLoading: true,
      items: [],
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
    await this.loadItems(this.props.divisions, selectedYear);
  }

  public async componentDidUpdate(prevProps: IServiceResultDirectoryProps): Promise<void> {
    const hasDivisionChanged =
      prevProps.divisions.join('|') !== this.props.divisions.join('|');

    if (hasDivisionChanged) {
      await this.loadItems(this.props.divisions, this.state.selectedYear);
    }
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

  private loadItems = async (divisions: string[], year: string): Promise<void> => {
    this.setState({ isLoading: true, isLoadingResult: true });

    try {
      const acDataPromise = this.acDataService.getAcDataByDepartment(this.props.resultDepartment);
      const result: IServiceResultDirectoryItem[] = [];

      for (const division of divisions) {
        const services = await this.divisionServiceService.getServicesWithIdByDivision(division, year);
        services.forEach(service => {
          result.push({
            Id: service.Id,
            Division: division,
            Service: service.Service,
            PIC: service.PIC,
            Manager: service.Manager,
          });
        });
      }

      const acItems = await acDataPromise;

      this.setState({
        isLoading: false,
        isLoadingResult: false,
        items: result,
        acItems: acItems.filter(item => (item.Year || '').toString() === year),
      });
    } catch (error) {
      console.error('Failed to load Division_Service directory:', error);
      this.setState({ isLoading: false, isLoadingResult: false, items: [], acItems: [] });
    }
  };

  private onYearFilterChange = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    const selectedYear = event.target.value;
    this.setState({ selectedYear });
    this.loadItems(this.props.divisions, selectedYear).catch(error => {
      console.error('Error changing year filter:', error);
    });
  };

  private handleShowResult = async (
    rowDepartment: string,
    item: IServiceResultDirectoryItem
  ): Promise<void> => {
    const key = `${item.Division}-${item.Service}`;
    this.setState({ loadingServiceKey: key });

    try {
      const serviceNames = item.Id
        ? await this.divisionServiceService.getServiceNameHistory(item.Id)
        : [];

      const services = serviceNames.length > 0 ? serviceNames : [item.Service];
      this.props.onShowResult(rowDepartment, services, item.Service);
    } catch (error) {
      console.error('Failed to load service name history:', error);
      this.props.onShowResult(rowDepartment, item.Service, item.Service);
    } finally {
      this.setState({ loadingServiceKey: undefined });
    }
  };

  private renderOverallResultSection = (): JSX.Element => {
    const { title, resultDepartment } = this.props;
    const { acItems, isLoadingResult, selectedYear } = this.state;

    return (
      <div className={styles.section}>
        <h2 className={styles.sectionTitle}>{title} - {selectedYear} Overall Result</h2>

        {isLoadingResult ? (
          <div className={styles.spinnerContainer}>
            <Spinner size={SpinnerSize.large} label="Loading results..." />
          </div>
        ) : acItems.length === 0 ? (
          <div className={styles.emptyMessage}>No result data available for {resultDepartment}.</div>
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
    const { title, resultDepartment, getResultDepartmentByDivision } = this.props;
    const { isLoading, items, loadingServiceKey, selectedYear, availableYears } = this.state;

    return (
      <main className={styles.mainContainer}>
        <div className={styles.breadcrumb}>
          Home {'>'} <strong>{title}</strong>
        </div>

        <div className={styles.header}>
          <h1>{title}</h1>
        </div>

        <div className={styles.content}>
          <div className={styles.filterSection}>
            <div className={styles.filterGroup}>
              <label htmlFor="service-directory-year-filter">Please select a year to view data</label>
              <select
                id="service-directory-year-filter"
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
            ) : items.length === 0 ? (
              <div className={styles.emptyMessage}>No services found.</div>
            ) : (
              <>
                <div className={styles.gridHeader}>
                  <div>Service</div>
                  <div>Service/Project Lead</div>
                  {/* <div>Manager</div> */}
                  <div>Result</div>
                </div>
                {items.map((item, index) => {
                  const rowDepartment = getResultDepartmentByDivision
                    ? getResultDepartmentByDivision(item.Division)
                    : resultDepartment;
                  const key = `${item.Division}-${item.Service}`;
                  const isRowLoading = loadingServiceKey === key;

                  return (
                    <div className={styles.gridRow} key={`${key}-${index}`}>
                      <div>{item.Service}</div>
                      <div>{item.PIC || '-'}</div>
                      {/* <div>{item.Manager || '-'}</div> */}
                      <div>
                        <button
                          type="button"
                          className={styles.actionLink}
                          disabled={isRowLoading}
                          onClick={() => {
                            this.handleShowResult(rowDepartment, item).catch(error => {
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
              </>
            )}
          </div>
        </div>
      </main>
    );
  }
}
