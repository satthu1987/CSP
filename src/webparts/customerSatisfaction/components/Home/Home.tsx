import * as React from 'react';
import styles from './Home.module.scss';
import { Icon, Spinner, SpinnerSize } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { AcDataService, IAcDataItem } from '../../services/AcDataService';

export interface IHomeProps {
  context: WebPartContext;
  currentUserDisplayName: string;
}

interface IHomeState {
  acData: IAcDataItem[];
  isLoading: boolean;
  expandedIds: { [id: number]: boolean };
}

export default class Home extends React.Component<IHomeProps, IHomeState> {
  private acDataService: AcDataService;

  constructor(props: IHomeProps) {
    super(props);
    this.state = {
      acData: [],
      isLoading: true,
      expandedIds: {}
    };
    this.acDataService = new AcDataService(props.context);
  }

  public async componentDidMount(): Promise<void> {
    const allAcData = await this.acDataService.getAcData();
    // Filter to only include items with Year value
    const acData = allAcData.filter(item => item.Year && item.Year.trim() !== '');
    this.setState({ acData, isLoading: false });
  }

  private toggleSection = (id: number): void => {
    this.setState(prev => ({
      expandedIds: { ...prev.expandedIds, [id]: !prev.expandedIds[id] }
    }));
  };

  public render(): JSX.Element {
    const { acData, isLoading, expandedIds } = this.state;

    return (
      <main className={styles.mainContainer}>
        <div className={styles.breadcrumb}>Home ›</div>
        
        <div className={styles.timelineImageContainer}>
          <img 
            src={require('../../assets/csp_timeline.jpg')}
            alt="CSP Timeline"
            className={styles.timelineImage}
            onError={(e) => {
              console.error('Timeline image failed to load');
              (e.target as HTMLImageElement).style.display = 'none';
            }}
          />
        </div>

        <div className={styles.header}>
          <h1>Results Overview</h1>
        </div>

        

        <div className={styles.content}>
          {isLoading ? (
            <div className={styles.spinnerContainer}>
              <Spinner size={SpinnerSize.large} label="Loading results..." />
            </div>
          ) : acData.length === 0 ? (
            <div className={styles.emptyState}>No data found in AC_Data list.</div>
          ) : (
            <div className={styles.acDataSection}>
              {acData.map(item => {
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
                      <span className={styles.acTitle}>
                        {item.Year} Overall Result
                      </span>
                    </button>

                    {expanded && (
                      <div className={styles.acBody}>
                        {item.DataUrl ? (
                          <img src={item.DataUrl}
                            className={styles.acImage}
                            onError={(e) => {
                              console.error('Image failed to load:', item.DataUrl);
                              (e.target as HTMLImageElement).style.display = 'none';
                            }}
                          />
                        ) : (
                          <div className={styles.emptyState}>
                            No image available for {item.Year}.
                          </div>
                        )}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          )}
        </div>
      </main>
    );
  }
}
