import * as React from 'react';
import styles from './Results.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { AcDataService, IAcDataItem } from '../../services/AcDataService';

export interface IResultsProps {
  context: WebPartContext;
  department: string;
}

interface IResultsState {
  items: IAcDataItem[];
  isLoading: boolean;
}

export default class Results extends React.Component<IResultsProps, IResultsState> {
  private acDataService: AcDataService;

  constructor(props: IResultsProps) {
    super(props);
    this.state = {
      items: [],
      isLoading: true
    };
    this.acDataService = new AcDataService(props.context);
  }

  public async componentDidMount(): Promise<void> {
    await this.loadData();
  }

  public async componentDidUpdate(prevProps: IResultsProps): Promise<void> {
    if (prevProps.department !== this.props.department) {
      await this.loadData();
    }
  }

  private loadData = async (): Promise<void> => {
    this.setState({ isLoading: true });
    const items = await this.acDataService.getAcDataByDepartment(this.props.department);
    this.setState({ items, isLoading: false });
  };

  public render(): JSX.Element {
    const { department } = this.props;
    const { items, isLoading } = this.state;

    return (
      <main className={styles.mainContainer}>
        <div className={styles.breadcrumb}>Home › <strong>{department}</strong></div>

        <div className={styles.header}>
          <h1>{department}</h1>
        </div>

        <div className={styles.content}>
          {isLoading ? (
            <div className={styles.spinnerContainer}>
              <Spinner size={SpinnerSize.large} label={`Loading ${department} data...`} />
            </div>
          ) : items.length === 0 ? (
            <div className={styles.emptyMessage}>
              No data available for {department}
            </div>
          ) : (
            <div className={styles.gridContainer}>
              {items.map(item => (
                <div key={item.Id} className={styles.card}>
                  <div className={styles.cardHeader}>
                    <h3>{item.Year}</h3>
                  </div>
                  <div className={styles.cardBody}>
                    {item.DataUrl ? (
                      <img src={item.DataUrl} alt={item.DataAlt} className={styles.image} />
                    ) : (
                      <div className={styles.noImage}>No image available</div>
                    )}
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </main>
    );
  }
}
