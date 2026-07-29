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
  onShowResult: (department: string, service: string, label: string) => void;
}

interface IServiceResultDirectoryItem {
  Division: string;
  Service: string;
PIC?: string;
}

interface IServiceResultDirectoryState {
  isLoading: boolean;
  items: IServiceResultDirectoryItem[];
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
        const services = await this.divisionServiceService.getServicesByDivision(division);
        services.forEach(service => {
          result.push({ Division: division, Service: service });
        });
      }

      this.setState({ isLoading: false, items: result });
    } catch (error) {
      console.error('Failed to load Division_Service directory:', error);
      this.setState({ isLoading: false, items: [] });
    }
  };

  public render(): JSX.Element {
    const { title, resultDepartment, getResultDepartmentByDivision, onShowResult } = this.props;
    const { isLoading, items } = this.state;

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

                  return (
                    <div className={styles.gridRow} key={`${item.Division}-${item.Service}-${index}`}>
                      <div>{item.Service}</div>
                      <div>{item.PIC}</div>
                      <div>
                        <button
                          type="button"
                          className={styles.actionLink}
                          onClick={() => onShowResult(rowDepartment, item.Service, item.Service)}
                        >
                          Show Result
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