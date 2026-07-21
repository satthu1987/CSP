import * as React from 'react';
import styles from './About.module.scss';
import { Icon, Spinner, SpinnerSize } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { CSPLearnMoreContentService } from '../../services/CSPLearnMoreContent_Service';
import { ICSPLearnMoreContent } from '../../Models/CSPLearnMoreContent';

export interface IAboutProps {
  context: WebPartContext;
}

interface IAboutState {
  content: ICSPLearnMoreContent[];
  isLoading: boolean;
  expandedIds: { [key: number]: boolean };
}

export default class About extends React.Component<IAboutProps, IAboutState> {
  private learnMoreService: CSPLearnMoreContentService;

  constructor(props: IAboutProps) {
    super(props);
    this.state = {
      content: [],
      isLoading: true,
      expandedIds: {}
    };
    this.learnMoreService = new CSPLearnMoreContentService(props.context);
  }

  public async componentDidMount(): Promise<void> {
    const content = await this.learnMoreService.getAllContent();
    this.setState({ content, isLoading: false });
  }

  private toggleSection = (id: number): void => {
    this.setState(prev => ({
      expandedIds: { ...prev.expandedIds, [id]: !prev.expandedIds[id] }
    }));
  };

  private renderCollapsibleSection = (item: ICSPLearnMoreContent): JSX.Element => {
    const expanded = !!this.state.expandedIds[item.Id];
    return (
      <div className={styles.collapsibleSection} key={item.Id}>
        <button
          className={styles.sectionHeader}
          onClick={() => this.toggleSection(item.Id)}
          aria-expanded={expanded}
        >
          <Icon
            iconName={expanded ? 'ChevronDown' : 'ChevronRight'}
            className={styles.chevron}
          />
          <span className={styles.sectionTitle}>{item.Title}</span>
        </button>
        {expanded && (
          <div
            className={styles.sectionBody}
            dangerouslySetInnerHTML={{ __html: item.Content }}
          />
        )}
      </div>
    );
  };

  public render(): JSX.Element {
    const { content, isLoading } = this.state;

    return (
      <main className={styles.main}>
        <section className={styles.informationSection}>
          <h4 style={{ color: '#000000' }}>Welcome to the Customer Satisfaction Program (CSP)! The program is designed to measure customer satisfaction actively and to identify opportunities for service improvement within Enterprise Services Vietnam (ESVN).
As a part of Innovation Group, we provides a variety of services in Sale Support, Internal Support, and Software Support to Simpson Strong-Tie worldwide.
</h4>
          {isLoading ? (
            <div style={{ display: 'flex', justifyContent: 'center', padding: '40px' }}>
              <Spinner size={SpinnerSize.medium} label="Loading content..." />
            </div>
          ) : content.length === 0 ? (
            <div style={{ textAlign: 'center', color: '#999', padding: '40px' }}>
              No content available.
            </div>
          ) : (
            content.map(item => this.renderCollapsibleSection(item))
          )}
        </section>
      </main>
    );
  }
}