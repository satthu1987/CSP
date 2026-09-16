import * as React from 'react';
import styles from './Home.module.scss';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { CSPLearnMoreContentService } from '../../services/CSPLearnMoreContent_Service';


export interface IHomeProps {
  context: WebPartContext;
  currentUserDisplayName: string;
}

interface IHomeState {
  howItWorksContent: string;
  isLoading: boolean;
}

export default class Home extends React.Component<IHomeProps, IHomeState> {
  private learnMoreService: CSPLearnMoreContentService;

  constructor(props: IHomeProps) {
    super(props);
    this.state = {
      howItWorksContent: '',
      isLoading: true
    };
    this.learnMoreService = new CSPLearnMoreContentService(props.context);
  }

  public async componentDidMount(): Promise<void> {
    const allContent = await this.learnMoreService.getAllContent();
    const howItWorksItem = allContent.filter(
      item => item.Title?.trim().toLowerCase() === 'how it works'
    )[0];
    this.setState({ howItWorksContent: howItWorksItem?.Content || '', isLoading: false });
  }

  public render(): JSX.Element {
    const { howItWorksContent, isLoading } = this.state;

    return (
      <main className={styles.mainContainer}>
        <div className={styles.breadcrumb}>Home ›</div>

        <div className={styles.heroImageContainer}>
          <img
            src={require('../../assets/cover.jpg')}
            alt="CSP Hero"
            className={styles.heroImage}
            onError={(e) => {
              console.error('Hero image failed to load');
              (e.target as HTMLImageElement).style.display = 'none';
            }}
          />
        </div>

        <div className={styles.welcomeSection}>
          <h3 className={styles.welcomeText}>
            Welcome to the Customer Satisfaction Program (CSP)! The program is designed to measure customer satisfaction actively and to identify opportunities for service improvement within Enterprise Services Vietnam (ESVN).
            As a part of Innovation Group, we provides a variety of services in Sale Support, Internal Support, and Software Support to Simpson Strong-Tie worldwide.
          </h3>
        </div>

        <div className={styles.howItWorksSection}>
          <h3 className={styles.sectionTitle}>How It Works</h3>
          {isLoading ? (
            <div className={styles.spinnerContainer}>
              <Spinner size={SpinnerSize.medium} label="Loading content..." />
            </div>
          ) : howItWorksContent ? (
            <div
              className={styles.sectionBody}
              dangerouslySetInnerHTML={{ __html: howItWorksContent }}
            />
          ) : (
            <div className={styles.emptyState}>
              No content available.
            </div>
          )}
        </div>
      </main>
    );
  }
}
