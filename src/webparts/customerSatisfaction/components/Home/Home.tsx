import * as React from 'react';
import styles from './Home.module.scss';

export interface IHomeProps {
  currentUserDisplayName: string;
}

export default class Home extends React.Component<IHomeProps> {
  public render(): JSX.Element {
    const { currentUserDisplayName } = this.props;

    return (
      <main className={styles.main}>
        <div className={styles.breadcrumb}>Home ›</div>
        <div className={styles.heroContent}>
          <h1 className={styles.welcomeScript}>Welcome,</h1>
          <h2 className={styles.userName}>{currentUserDisplayName}</h2>
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
