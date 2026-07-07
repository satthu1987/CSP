import * as React from 'react';
import styles from './ActionPlan.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Icon, Spinner, SpinnerSize } from '@fluentui/react';
import { ActionPlanService } from '../../services/ActionPlan_Service';
import { IActionplan } from '../../Models/ActionplanModel';

export interface IActionPlanProps {
  context: WebPartContext;
  userService: string;
}

interface IActionPlanState {
  actionPlans: IActionplan[];
  isLoading: boolean;
  selectedActionPlan: IActionplan | null;
  isDetailPanelOpen: boolean;
  isNewMode: boolean;
  formData: Partial<IActionplan>;
}

export default class ActionPlan extends React.Component<IActionPlanProps, IActionPlanState> {
  private actionPlanService: ActionPlanService;

  constructor(props: IActionPlanProps) {
    super(props);
    this.state = {
      actionPlans: [],
      isLoading: true,
      selectedActionPlan: null,
      isDetailPanelOpen: false,
      isNewMode: false,
      formData: {}
    };
    this.actionPlanService = new ActionPlanService(props.context, 'ActionPlan');
  }

  public async componentDidMount(): Promise<void> {
    await this.loadActionPlans();
  }

  private loadActionPlans = async (): Promise<void> => {
    this.setState({ isLoading: true });
    const actionPlans = await this.actionPlanService.getActionPlansByService(this.props.userService);
    this.setState({ actionPlans, isLoading: false });
  };

  private openDetailPanel = (actionPlan: IActionplan): void => {
    this.setState({
      selectedActionPlan: actionPlan,
      isDetailPanelOpen: true,
      isNewMode: false,
      formData: { ...actionPlan }
    });
  };

  private openNewActionPanel = (): void => {
    this.setState({
      selectedActionPlan: null,
      isDetailPanelOpen: true,
      isNewMode: true,
      formData: { Service: this.props.userService }
    });
  };

  private closeDetailPanel = (): void => {
    this.setState({
      isDetailPanelOpen: false,
      selectedActionPlan: null,
      isNewMode: false,
      formData: {}
    });
  };

  private updateFormField = (field: keyof IActionplan, value: any): void => {
    this.setState(prev => ({
      formData: {
        ...prev.formData,
        [field]: value
      }
    }));
  };

  private handleSave = async (): Promise<void> => {
    const { isNewMode, formData } = this.state;

    if (isNewMode) {
      const result = await this.actionPlanService.createActionPlan(formData);
      if (result) {
        await this.loadActionPlans();
        this.closeDetailPanel();
      }
    } else if (this.state.selectedActionPlan) {
      const success = await this.actionPlanService.updateActionPlan(
        this.state.selectedActionPlan.Id,
        formData
      );
      if (success) {
        await this.loadActionPlans();
        this.closeDetailPanel();
      }
    }
  };

  private renderGrid(): JSX.Element {
    const { actionPlans, isLoading } = this.state;

    if (isLoading) {
      return (
        <div className={styles.spinnerContainer}>
          <Spinner size={SpinnerSize.large} label="Loading action plans..." />
        </div>
      );
    }

    if (actionPlans.length === 0) {
      return (
        <div className={styles.emptyState}>
          No action plans found for your service.
        </div>
      );
    }

    return (
      <div className={styles.gridContainer}>
        <div className={styles.gridHeader}>
          <div className={styles.colTitle}>Title</div>
          <div className={styles.colStatus}>Status</div>
          <div className={styles.colTimeline}>Timeline</div>
          <div className={styles.colCategory}>Category</div>
        </div>
        {actionPlans.map(plan => (
          <div
            key={plan.Id}
            className={styles.gridRow}
            onClick={() => this.openDetailPanel(plan)}
          >
            <div className={styles.colTitle}>{plan.Title}</div>
            <div className={styles.colStatus}>
              <span className={`${styles.badge} ${styles[`status-${plan.Status?.toLowerCase()}` as keyof typeof styles]}`}>
                {plan.Status}
              </span>
            </div>
            <div className={styles.colTimeline}>
              {plan.Timeline ? new Date(plan.Timeline).toLocaleDateString() : '-'}
            </div>
            <div className={styles.colCategory}>{plan.Category || '-'}</div>
          </div>
        ))}
      </div>
    );
  }

  private renderDetailPanel(): JSX.Element {
    const { formData, isNewMode } = this.state;

    return (
      <div className={styles.detailPanel}>
        <div className={styles.panelHeader}>
          <h2>{isNewMode ? 'New Action Plan' : 'Action Plan Details'}</h2>
          <button className={styles.closeBtn} onClick={this.closeDetailPanel}>
            <Icon iconName="Cancel" />
          </button>
        </div>

        <div className={styles.panelBody}>
          <div className={styles.formGroup}>
            <label>Title</label>
            <input
              type="text"
              value={formData.Title || ''}
              onChange={(e) => this.updateFormField('Title', e.target.value)}
              placeholder="Enter title"
            />
          </div>

          <div className={styles.formGroup}>
            <label>Service</label>
            <input
              type="text"
              value={formData.Service || ''}
              readOnly
            />
          </div>

          <div className={styles.formGroup}>
            <label>Status</label>
            <select
              value={formData.Status || ''}
              onChange={(e) => this.updateFormField('Status', e.target.value)}
            >
              <option value="">Select Status</option>
              <option value="Open">Open</option>
              <option value="In Progress">In Progress</option>
              <option value="Closed">Closed</option>
            </select>
          </div>

          <div className={styles.formGroup}>
            <label>Category</label>
            <input
              type="text"
              value={formData.Category || ''}
              onChange={(e) => this.updateFormField('Category', e.target.value)}
              placeholder="Enter category"
            />
          </div>

          <div className={styles.formGroup}>
            <label>Product Line</label>
            <input
              type="text"
              value={formData.ProductLine || ''}
              onChange={(e) => this.updateFormField('ProductLine', e.target.value)}
              placeholder="Enter product line"
            />
          </div>

          <div className={styles.formGroup}>
            <label>Timeline</label>
            <input
              type="date"
              value={formData.Timeline ? formData.Timeline.split('T')[0] : ''}
              onChange={(e) => this.updateFormField('Timeline', e.target.value)}
            />
          </div>

          <div className={styles.formGroup}>
            <label>Year</label>
            <input
              type="text"
              value={formData.Year || ''}
              onChange={(e) => this.updateFormField('Year', e.target.value)}
              placeholder="Enter year"
            />
          </div>

          <div className={styles.formGroup}>
            <label>Department</label>
            <input
              type="text"
              value={formData.Department || ''}
              onChange={(e) => this.updateFormField('Department', e.target.value)}
              placeholder="Enter department"
            />
          </div>

          <div className={styles.formGroup}>
            <label>Division</label>
            <input
              type="text"
              value={formData.Division || ''}
              onChange={(e) => this.updateFormField('Division', e.target.value)}
              placeholder="Enter division"
            />
          </div>

          <div className={styles.formGroup}>
            <label>Customer Feedback</label>
            <textarea
              value={formData.CustomerFeedback?.join('\n') || ''}
              onChange={(e) => this.updateFormField('CustomerFeedback', e.target.value.split('\n'))}
              placeholder="Enter feedback (one per line)"
              rows={4}
            />
          </div>

          <div className={styles.formGroup}>
            <label>Actions</label>
            <textarea
              value={formData.Actions?.join('\n') || ''}
              onChange={(e) => this.updateFormField('Actions', e.target.value.split('\n'))}
              placeholder="Enter actions (one per line)"
              rows={4}
            />
          </div>

          <div className={styles.formGroup}>
            <label>Results</label>
            <textarea
              value={formData.Results?.join('\n') || ''}
              onChange={(e) => this.updateFormField('Results', e.target.value.split('\n'))}
              placeholder="Enter results (one per line)"
              rows={4}
            />
          </div>

          <div className={styles.formGroup}>
            <label>Related Links</label>
            <input
              type="text"
              value={formData.RelatedLinks || ''}
              onChange={(e) => this.updateFormField('RelatedLinks', e.target.value)}
              placeholder="Enter related links"
            />
          </div>
        </div>

        <div className={styles.panelFooter}>
          <button className={styles.btnCancel} onClick={this.closeDetailPanel}>
            Cancel
          </button>
          <button className={styles.btnSave} onClick={this.handleSave}>
            {isNewMode ? 'Create' : 'Update'}
          </button>
        </div>
      </div>
    );
  }

  public render(): JSX.Element {
    const { isDetailPanelOpen } = this.state;

    return (
      <main className={styles.mainContainer}>
        <div className={styles.breadcrumb}>Home › <strong>Action Plan</strong></div>

        <div className={styles.header}>
          <h1>Action Plans</h1>
          <button className={styles.btnNewAction} onClick={this.openNewActionPanel}>
            <Icon iconName="Add" /> New Action
          </button>
        </div>

        <div className={`${styles.content} ${isDetailPanelOpen ? styles.contentWithPanel : ''}`}>
          {this.renderGrid()}
          {isDetailPanelOpen && this.renderDetailPanel()}
        </div>
      </main>
    );
  }
}
