import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import styles from './ServiceResultDirectory.module.scss';
import { DivisionServiceService } from '../../services/DivisionService_Service';

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
}

interface IServiceResultDirectoryState {
  isLoading: boolean;
  items: IServiceResultDirectoryItem[];
  loadingServiceKey?: string;
}

export default class ServiceResultDirectory extends React.Component<
  IServiceResultDirectoryProps,
  IServiceResultDirectoryState
> {
  private divisionServiceService: DivisionServiceService;

  constructor(props: IServiceResultDirectoryProps) {
    super(props);
    this.state = {
      isLoading: true,
      items: [],
    };
    this.divisionServiceService = new DivisionServiceService(props.context, 'Division_Service');
  }

  public async componentDidMount(): Promise<void> {
    await this.loadItems(this.props.divisions);
  }

  public async componentDidUpdate(prevProps: IServiceResultDirectoryProps): Promise<void> {
    const hasDivisionChanged =
      prevProps.divisions.join('|') !== this.props.divisions.join('|');

    if (hasDivisionChanged) {
      await this.loadItems(this.props.divisions);
    }
  }

  private loadItems = async (divisions: string[]): Promise<void> => {
    this.setState({ isLoading: true });

    try {
      const result: IServiceResultDirectoryItem[] = [];

      for (const division of divisions) {
        const services = await this.divisionServiceService.getServicesWithIdByDivision(division);
        services.forEach(service => {
          result.push({ Id: service.Id, Division: division, Service: service.Service, PIC: service.PIC });
        });
      }

      this.setState({ isLoading: false, items: result });
    } catch (error) {
      console.error('Failed to load Division_Service directory:', error);
      this.setState({ isLoading: false, items: [] });
    }
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

  public render(): JSX.Element {
    const { title, resultDepartment, getResultDepartmentByDivision } = this.props;
    const { isLoading, items, loadingServiceKey } = this.state;

    return (
      <main className={styles.mainContainer}>
        <div className={styles.breadcrumb}>
          Home {'>'} <strong>{title}</strong>
        </div>

        <div className={styles.header}>
          <h1>{title}</h1>
        </div>

        <div className={styles.content}>
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
                      <div>{item.PIC}</div>
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
