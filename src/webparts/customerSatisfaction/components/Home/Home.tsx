import * as React from 'react';
import styles from './Home.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IHomeProps {
  context: WebPartContext;
  currentUserDisplayName: string;
}

interface IHomeState {
  // No state needed
}

export default class Home extends React.Component<IHomeProps, IHomeState> {
  constructor(props: IHomeProps) {
    super(props);
    this.state = {};
  }

  public render(): JSX.Element {
    return (
      <main className={styles.mainContainer}>
        <div className={styles.breadcrumb}>Home ›</div>
        
        <div className={styles.heroImageContainer}>
          <img 
            src={require('../../assets/hero-image.png')}
            alt="CSP Hero"
            className={styles.heroImage}
            onError={(e) => {
              console.error('Hero image failed to load');
              (e.target as HTMLImageElement).style.display = 'none';
            }}
          />
        </div>

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
      </main>
    );
  }
}
