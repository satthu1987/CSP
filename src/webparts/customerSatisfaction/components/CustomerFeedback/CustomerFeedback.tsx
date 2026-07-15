import * as React from 'react';
import styles from './CustomerFeedback.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Icon, Spinner, SpinnerSize } from '@fluentui/react';
import { CustomerFeedbackService } from '../../services/CustomerFeedback_Service';
import { UserRoleService } from '../../services/UserRole_Service';
import { ICustomerFeedback } from '../../Models/CustomerFeedbackModel';
import ActionPlan from '../ActionPlan/ActionPlan';

export interface ICustomerFeedbackProps {
  context: WebPartContext;
  userService: string;
}

interface ICustomerFeedbackState {
  feedbacks: ICustomerFeedback[];
  isLoading: boolean;
  userServices: string[];
  activeFeedbackId: number | undefined;
  activeFeedbackText: string | undefined;
  isActionPlanOpen: boolean;
}

export default class CustomerFeedback extends React.Component<ICustomerFeedbackProps, ICustomerFeedbackState> {
  private feedbackService: CustomerFeedbackService;
  private userRoleService: UserRoleService;

  constructor(props: ICustomerFeedbackProps) {
    super(props);
    this.state = {
      feedbacks: [],
      isLoading: true,
      userServices: [],
      activeFeedbackId: undefined,
      activeFeedbackText: undefined,
      isActionPlanOpen: false,
    };
    this.feedbackService = new CustomerFeedbackService(props.context, 'CSP_CustomerFeedback');
    this.userRoleService = new UserRoleService(props.context, 'RoleInService');
  }

  public async componentDidMount(): Promise<void> {
    await this.loadUserServices();
  }

  private loadUserServices = async (): Promise<void> => {
    try {
      const userEmail = this.props.context.pageContext.user.loginName;
      console.log('Loading user services for:', userEmail);
      
      const services = await this.userRoleService.getUserServices(userEmail);
      console.log('User services loaded:', services);
      
      this.setState({ userServices: services });
      
      if (services.length > 0) {
        await this.loadFeedbacks(services);
      } else {
        // No services found, show empty state
        this.setState({ isLoading: false, feedbacks: [] });
      }
    } catch (error) {
      console.error('Failed to load user services:', error);
      this.setState({ isLoading: false, feedbacks: [] });
    }
  };

  private loadFeedbacks = async (services?: string[]): Promise<void> => {
    this.setState({ isLoading: true });
    
    const servicesToUse = services || this.state.userServices;
    
    if (!servicesToUse || servicesToUse.length === 0) {
      console.warn('No services available to filter feedbacks');
      this.setState({ isLoading: false, feedbacks: [] });
      return;
    }
    
    const feedbacks = await this.feedbackService.getCustomerFeedbacksByServices(servicesToUse);
    this.setState({ feedbacks, isLoading: false });
  };

  private stripHtmlTags(html: string): string {
    if (!html) return '';
    
    // Create a temporary element and use textContent to extract plain text
    const div = document.createElement('div');
    div.innerHTML = html;
    let text = div.textContent || div.innerText || '';
    
    // Trim and clean up multiple spaces
    text = text.trim().replace(/\s+/g, ' ');
    
    // Limit to 150 chars for display in grid
    if (text.length > 150) {
      text = text.substring(0, 150) + '...';
    }
    
    return text;
  }

  private handleAddActionPlan = (feedbackId: number, feedbackText: string): void => {
    const cleanedText = this.stripHtmlTags(feedbackText);
    this.setState({ activeFeedbackId: feedbackId, activeFeedbackText: cleanedText, isActionPlanOpen: true });
  };

  private handleActionPlanDismiss = (): void => {
    this.setState({ activeFeedbackId: undefined, activeFeedbackText: undefined, isActionPlanOpen: false });
  };

  private handleActionPlanCreated = async (actionPlanId: number): Promise<void> => {
    const { activeFeedbackId } = this.state;
    if (activeFeedbackId !== undefined) {
      const success = await this.feedbackService.updateActionPlan(activeFeedbackId, actionPlanId);
      if (success) {
        await this.loadFeedbacks();
      } else {
        console.error('Failed to update ActionPlan for feedback:', activeFeedbackId);
      }
    }
    this.setState({ activeFeedbackId: undefined, activeFeedbackText: undefined, isActionPlanOpen: false });
  };

  private groupByService(feedbacks: ICustomerFeedback[]): Map<string, ICustomerFeedback[]> {
    const map = new Map<string, ICustomerFeedback[]>();
    feedbacks.forEach(item => {
      const key = item.Service || '(No Service)';
      if (!map.has(key)) {
        map.set(key, []);
      }
      map.get(key)!.push(item);
    });
    return map;
  }

  private renderFeedbackList(): JSX.Element {
    const { feedbacks, isLoading, userServices } = this.state;

    if (isLoading) {
      return (
        <div className={styles.spinnerContainer}>
          <Spinner size={SpinnerSize.large} label="Loading customer feedbacks..." />
        </div>
      );
    }

    if (!userServices || userServices.length === 0) {
      return (
        <div className={styles.emptyMessage}>
          You do not have any assigned services. Please contact your administrator.
        </div>
      );
    }

    if (feedbacks.length === 0) {
      return (
        <div className={styles.emptyMessage}>
          No customer feedback items found for your services.
        </div>
      );
    }

    const grouped = this.groupByService(feedbacks);
    const groups: JSX.Element[] = [];

    grouped.forEach((items, service) => {
      groups.push(
        <div key={service} className={styles.serviceGroup}>
          <div className={styles.serviceGroupHeader}>{service}</div>
          <div className={styles.feedbackRowHeader}>
            <div>Title</div>
            <div>Customer Feedback</div>
            <div>Action Plan ID</div>
            <div>Status</div>
            <div />
          </div>
          <div className={styles.feedbackList}>
            {items.map(item => (
              <div key={item.Id} className={styles.feedbackRow}>
                <div className={styles.colTitle}>{item.Title || '-'}</div>
                <div className={styles.colFeedback} title={this.stripHtmlTags(item.CustomerFeedback || '')}>
                  {this.stripHtmlTags(item.CustomerFeedback || '') || '-'}
                </div>
                <div className={styles.colActionPlan}>
                  {item.ActionPlan ? `#${item.ActionPlan}` : '-'}
                </div>
                <div className={styles.colStatus}>
                  <span className={`${styles.badge} ${item.ActionPlan ? styles.badgeLinked : styles.badgeNone}`}>
                    {item.ActionPlan ? 'Linked' : 'No Plan'}
                  </span>
                </div>
                <div className={styles.colActions}>
                  <button
                    className={styles.btnAdd}
                    title="Add Action Plan"
                    onClick={() => this.handleAddActionPlan(item.Id, item.CustomerFeedback || '')}
                  >
                    <Icon iconName="Add" />
                  </button>
                </div>
              </div>
            ))}
          </div>
        </div>
      );
    });

    return <>{groups}</>;
  }

  public render(): JSX.Element {
    const { isActionPlanOpen, activeFeedbackId, activeFeedbackText } = this.state;

    return (
      <main className={styles.mainContainer}>
        <div className={styles.breadcrumb}>Home › <strong>Customer Feedback</strong></div>

        <div className={styles.header}>
          <h1>Customer Feedback</h1>
        </div>

        <div className={styles.content}>
          {this.renderFeedbackList()}
        </div>

        {isActionPlanOpen && activeFeedbackId !== undefined && (
          <ActionPlan
            context={this.props.context}
            userService={this.props.userService}
            isInlineMode={true}
            sourceFeedbackId={activeFeedbackId}
            sourceFeedbackText={activeFeedbackText}
            onActionPlanCreated={this.handleActionPlanCreated}
            onInlineDismiss={this.handleActionPlanDismiss}
          />
        )}
      </main>
    );
  }
}
