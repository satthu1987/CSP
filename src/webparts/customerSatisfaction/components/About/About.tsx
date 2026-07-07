import * as React from 'react';
import styles from './About.module.scss';
import { Icon, Spinner, SpinnerSize } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { AcDataService, IAcDataItem } from '../../services/AcDataService';

export interface IAboutProps {
  context: WebPartContext;
}

interface IAboutState {
  acData: IAcDataItem[];
  isLoading: boolean;
  expandedIds: { [id: number]: boolean };
}

export default class About extends React.Component<IAboutProps, IAboutState> {
  private acDataService: AcDataService;

  constructor(props: IAboutProps) {
    super(props);
    this.state = {
      acData: [],
      isLoading: true,
      expandedIds: {}
    };
    this.acDataService = new AcDataService(props.context);
  }

  public async componentDidMount(): Promise<void> {
    const acData = await this.acDataService.getAcData();
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
      <div className={styles.aboutPage}>
        {/* ---------- SECTION 1: TIMELINE ---------- */}
        <section className={styles.timelineSection}>
          <div className={styles.timeline}>
            {/* Wavy connector SVG */}
            <img src={require('../../assets/csp_timeline.jpg')}
                 className={styles.acImage} />
          </div>
        </section>

        {/* ---------- SECTION 2: DIVISIONS ---------- */}
        <section className={styles.divisionsSection}>
          <div className={styles.divisionCard}>
            <div className={styles.divider} />
            <h3 className={styles.divisionTitle}>Internal &amp; Sales Support</h3>
            <button className={styles.learnMoreBtn}
              onClick={() => window.location.href = '#/about/internal-sales'}>
              Learn more
            </button>
          </div>

          <div className={styles.divisionCard}>
            <div className={styles.divider} />
            <h3 className={styles.divisionTitle}>Digital Technology Support</h3>
            <button className={styles.learnMoreBtn}
              onClick={() => window.location.href = '#/about/digital-tech'}>
              Learn more
            </button>
          </div>
        </section>

        {/* ---------- SECTION 3: AC_DATA COLLAPSIBLE ROWS ---------- */}
        <section className={styles.acDataSection}>
          {isLoading ? (
            <Spinner size={SpinnerSize.large} label="Loading results..." />
          ) : acData.length === 0 ? (
            <div className={styles.emptyState}>No data found in AC_Data list.</div>
          ) : (
            acData.map(item => {
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
            )}))}
        </section>
      </div>
    );
  }
}