import * as React from 'react';
import styles from './Home.module.scss';
import { Icon, Spinner, SpinnerSize } from '@fluentui/react';

export interface IHomeProps {
  currentUserDisplayName: string;
  isTeamLeader: boolean;
  isLoading: boolean;
  sentToday: number;
  receivedToday: number;
  onActionClick: () => void;
}

export default class Home extends React.Component<IHomeProps> {
  public render(): JSX.Element {
    const { isTeamLeader, isLoading, sentToday, receivedToday, currentUserDisplayName, onActionClick } = this.props;

    return (
      <main className={styles.main}>
        <div className={styles.breadcrumb}>Home ›</div>
        <div className={styles.heroContent}>
          <div className={styles.welcomeBadge}>Nine Core</div>
          <h1 className={styles.welcomeScript}>Welcome,</h1>
          <h2 className={styles.userName}>{currentUserDisplayName}</h2>

          {isLoading ? (
            <div className={styles.loaderRow}>
              <Spinner size={SpinnerSize.small} label="Checking permissions..." />
            </div>
          ) : isTeamLeader && (
            <div className={styles.actionButtons}>
              <button className={`${styles.actionBtn} ${styles.primary}`}
                onClick={onActionClick}>
                <Icon iconName="Send" /> Send Recognition
              </button>
            </div>
          )}

          <div className={styles.statsCard}>
            Today you sent <strong>{sentToday}</strong> recognitions <br />
            and received <strong>{receivedToday}</strong> recognitions
          </div>
        </div>

        <div className={styles.heroVisual}>
          <img
            src={require('../../assets/hero-image.png')}
            alt="Customer Satisfaction Program"
            className={styles.heroImage}
          />
          <div className={styles.heroTitle}>
            <span>CUSTOMER</span>
            <span>SATISFACTION</span>
            <span>PROGRAM</span>
          </div>
        </div>
      </main>
    );
  }
}
