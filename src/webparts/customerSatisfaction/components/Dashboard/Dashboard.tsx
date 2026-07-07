import * as React from 'react';
import styles from './Dashboard.module.scss';
import { Spinner, SpinnerSize } from '@fluentui/react';

export interface IDashboardProps {
  currentUserDisplayName: string;
}

interface IDashboardState {
  isLoading: boolean;
}

export default class Dashboard extends React.Component<IDashboardProps, IDashboardState> {
  constructor(props: IDashboardProps) {
    super(props);
    this.state = {
      isLoading: false
    };
  }

  public render(): JSX.Element {
    const { isLoading } = this.state;

    return (
      <main className={styles.mainPlain}>
        <div className={styles.pageBreadcrumb}>
          Home › <strong>Dashboard</strong>
        </div>
        <div className={styles.container}>
          {isLoading ? (
            <div className={styles.spinnerContainer}>
              <Spinner size={SpinnerSize.large} label="Loading dashboard..." />
            </div>
          ) : (
            <>
              <h2>Dashboard</h2>
              <p>Dashboard functionality coming soon.</p>
              <p>This page is under construction.</p>
            </>
          )}
        </div>
      </main>
    );
  }
}
