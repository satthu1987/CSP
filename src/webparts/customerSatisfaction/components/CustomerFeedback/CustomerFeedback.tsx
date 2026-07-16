import * as React from 'react';
import styles from './CustomerFeedback.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Icon, Spinner, SpinnerSize } from '@fluentui/react';
import { CustomerFeedbackService } from '../../services/CustomerFeedback_Service';
import { ActionPlanService } from '../../services/ActionPlan_Service';
import { UserRoleService } from '../../services/UserRole_Service';
import { ICustomerFeedback } from '../../Models/CustomerFeedbackModel';
import { IActionplan } from '../../Models/ActionplanModel';
import ActionPlan from '../ActionPlan/ActionPlan';
import ReactQuill from 'react-quill';

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
  isDetailView: boolean;
  selectedFeedback: ICustomerFeedback | undefined;
  selectedActionPlan: IActionplan | undefined;
  isActionPlanLoading: boolean;
}

export default class CustomerFeedback extends React.Component<ICustomerFeedbackProps, ICustomerFeedbackState> {
  private feedbackService: CustomerFeedbackService;
  private actionPlanService: ActionPlanService;
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
      isDetailView: false,
      selectedFeedback: undefined,
      selectedActionPlan: undefined,
      isActionPlanLoading: false,
    };
    this.feedbackService = new CustomerFeedbackService(props.context, 'CSP_CustomerFeedback');
    this.actionPlanService = new ActionPlanService(props.context, 'CSP');
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

  private stripHtmlTags(html: string, maxLength?: number): string {
    if (!html) return '';
    
    // Create a temporary element and use textContent to extract plain text
    const div = document.createElement('div');
    div.innerHTML = html;
    let text = div.textContent || div.innerText || '';
    
    // Trim and clean up multiple spaces
    text = text.trim().replace(/\s+/g, ' ');
    
    // Limit text length only when maxLength is provided.
    if (maxLength !== undefined && maxLength > 0 && text.length > maxLength) {
      text = text.substring(0, maxLength) + '...';
    }
    
    return text;
  }

  private handleAddActionPlan = (feedbackId: number, feedbackText: string): void => {
    const cleanedText = this.stripHtmlTags(feedbackText);
    this.setState({ activeFeedbackId: feedbackId, activeFeedbackText: cleanedText, isActionPlanOpen: true });
  };

  private openDetailView = async (feedback: ICustomerFeedback): Promise<void> => {
    this.setState({
      isDetailView: true,
      selectedFeedback: feedback,
      selectedActionPlan: undefined,
      isActionPlanLoading: !!feedback.ActionPlan,
    });

    if (!feedback.ActionPlan) {
      return;
    }

    const actionPlan = await this.actionPlanService.getActionPlanById(feedback.ActionPlan);
    this.setState({ selectedActionPlan: actionPlan, isActionPlanLoading: false });
  };

  private backToFeedbackList = (): void => {
    this.setState({
      isDetailView: false,
      selectedFeedback: undefined,
      selectedActionPlan: undefined,
      isActionPlanLoading: false,
    });
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
            <div>Updated Feedback</div>
            <div>Action Plan ID</div>
            <div>Status</div>
            <div />
          </div>
          <div className={styles.feedbackList}>
            {items.map(item => (
              <div key={item.Id} className={styles.feedbackRow} onClick={() => this.openDetailView(item)}>
                <div className={styles.colTitle}>
                    {item.Title || '-'}
                </div>
                <div className={styles.colFeedback} title={this.stripHtmlTags(item.UpdatedFeedback || '')}>{this.stripHtmlTags(item.UpdatedFeedback || '', 150) || '-'}
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
                  {!item.ActionPlan && (
                    <button
                      className={styles.btnAdd}
                      title="Add Action Plan"
                      onClick={(event) => {
                        event.stopPropagation();
                        this.handleAddActionPlan(item.Id, item.UpdatedFeedback || '');
                      }}
                    >
                      <Icon iconName="Add" />
                    </button>
                  )}
                </div>
              </div>
            ))}
          </div>
        </div>
      );
    });

    return <>{groups}</>;
  }

  private renderActionPlanDetailSection(): JSX.Element {
    const { selectedFeedback, selectedActionPlan, isActionPlanLoading } = this.state;

    if (!selectedFeedback) {
      return <></>;
    }

    return (
      <section className={styles.detailSection}>
        <h2>Linked Action Plan</h2>

        {isActionPlanLoading && (
          <div className={styles.spinnerContainer}>
            <Spinner size={SpinnerSize.medium} label="Loading action plan..." />
          </div>
        )}

        {!isActionPlanLoading && !selectedFeedback.ActionPlan && (
          <div className={styles.emptyMessage}>No Action Plan linked to this feedback.</div>
        )}

        {!isActionPlanLoading && selectedFeedback.ActionPlan && !selectedActionPlan && (
          <div className={styles.emptyMessage}>Linked Action Plan could not be loaded.</div>
        )}

        {!isActionPlanLoading && selectedActionPlan && (
          <div className={styles.linkedPlanCard}>
            <div className={styles.linkedPlanHeader}>
              <div className={styles.linkedPlanTitle}>{selectedActionPlan.Title || '-'}</div>
              <div className={styles.linkedPlanMeta}>#{selectedActionPlan.Id}</div>
            </div>
            <div className={styles.linkedPlanGrid}>
              <div>
                <span>Status:</span> {selectedActionPlan.Status || '-'}
              </div>
              <div>
                <span>Service:</span> {selectedActionPlan.Service || '-'}
              </div>
              <div>
                <span>PIC:</span> {selectedActionPlan.PIC?.Title || '-'}
              </div>
              <div>
                <span>Timeline:</span> {selectedActionPlan.Timeline || '-'}
              </div>
            </div>
          </div>
        )}
      </section>
    );
  }

  private renderDetailView(): JSX.Element {
    const { selectedFeedback } = this.state;

    if (!selectedFeedback) {
      return <></>;
    }

    return (
      <>
        <div className={styles.detailHeaderActions}>
          <button className={styles.btnBack} onClick={this.backToFeedbackList}>
            <Icon iconName="Back" />
            <span>Back</span>
          </button>
        </div>

        <section className={styles.detailSection}>
          <h2>Feedback Information</h2>
          <div className={styles.detailRow}>
            <label>Title</label>
            <div>{selectedFeedback.Title || '-'}</div>
          </div>
          <div className={styles.detailRow}>
            <label>Original Feedback</label>
            <div className={styles.feedbackContent}>
              <ReactQuill
                theme="snow"
                value={Array.isArray(selectedFeedback.CustomerFeedback) ? selectedFeedback.CustomerFeedback.join('\n') : (selectedFeedback.CustomerFeedback || '')}
                placeholder="Enter feedback with full format..."
              />
            </div>
          </div>
          <div className={styles.detailRow}>
            <label>Updated Feedback</label>
            <div className={styles.feedbackContent}>
              <ReactQuill
                theme="snow"
                value={Array.isArray(selectedFeedback.UpdatedFeedback) ? selectedFeedback.UpdatedFeedback.join('\n') : (selectedFeedback.CustomerFeedback || '')}
                placeholder="Enter feedback with full format..."
              />
            </div>
          </div>
        </section>

        {this.renderActionPlanDetailSection()}
      </>
    );
  }

  public render(): JSX.Element {
    const { isActionPlanOpen, activeFeedbackId, activeFeedbackText, isDetailView } = this.state;

    return (
      <main className={styles.mainContainer}>
        <div className={styles.breadcrumb}>
          Home › <strong>Customer Feedback</strong>
          {isDetailView && <span> › <strong>Detail Feedback</strong></span>}
        </div>

        <div className={styles.header}>
          <h1>{isDetailView ? 'Detail Feedback' : 'Customer Feedback'}</h1>
        </div>

        <div className={styles.content}>
          {isDetailView ? this.renderDetailView() : this.renderFeedbackList()}
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
