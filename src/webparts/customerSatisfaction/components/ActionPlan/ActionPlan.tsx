import * as React from 'react';
import styles from './ActionPlan.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Icon, IPersonaProps, Modal, Spinner, SpinnerSize } from '@fluentui/react';
import { ActionPlanService, IActionPlanUpsert } from '../../services/ActionPlan_Service';
import { DivisionServiceService } from '../../services/DivisionService_Service';
import { IActionplan } from '../../Models/ActionplanModel';
import ReactQuill from 'react-quill';
import 'react-quill/dist/quill.snow.css'; 
import { IPeoplePickerContext, PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

type IActionPlanFormData = IActionPlanUpsert;

export interface IActionPlanProps {
  context: WebPartContext;
  userService: string;
  /** When true, renders only the modal in new-plan mode (no full page) */
  isInlineMode?: boolean;
  /** ID of the CSP_CustomerFeedback item that triggered this form */
  sourceFeedbackId?: number;
  /** Text content of the Customer Feedback that triggered this form */
  sourceFeedbackText?: string;
  /** Called with the new ActionPlan ID after successful creation in inline mode */
  onActionPlanCreated?: (actionPlanId: number) => void;
  /** Called when the user dismisses the modal in inline mode */
  onInlineDismiss?: () => void;
}

interface IActionPlanState {
  actionPlans: IActionplan[];
  isLoading: boolean;
  selectedActionPlan?: IActionplan;
  isDetailPanelOpen: boolean;
  isNewMode: boolean;
  formData: IActionPlanFormData;
  choiceOptions: { [key: string]: string[] };
  departmentServices: string[];
  isDepartmentServicesLoading: boolean;
}

export default class ActionPlan extends React.Component<IActionPlanProps, IActionPlanState> {
  private actionPlanService: ActionPlanService;
  private divisionServiceService: DivisionServiceService;
  private readonly digitalTechnologySupportDivision = 'Digital Transformation Support';

  constructor(props: IActionPlanProps) {
    super(props);
    this.state = {
      actionPlans: [],
      isLoading: true,
      selectedActionPlan: undefined,
      isDetailPanelOpen: props.isInlineMode === true,
      isNewMode: props.isInlineMode === true,
      formData: {
        Service: props.userService,
        CustomerFeedback: props.sourceFeedbackText || undefined
      },
      choiceOptions: {},
      departmentServices: [],
      isDepartmentServicesLoading: false
    };
    this.actionPlanService = new ActionPlanService(props.context, 'CSP');
    this.divisionServiceService = new DivisionServiceService(props.context, 'Division_Service');
  }

  public async componentDidMount(): Promise<void> {
    await Promise.all([
      this.loadActionPlans(),
      this.loadChoiceOptions()
    ]);
  }

  private loadActionPlans = async (): Promise<void> => {
    this.setState({ isLoading: true });
    const actionPlans = await this.actionPlanService.getActionPlansByService(this.props.userService);
    this.setState({ actionPlans, isLoading: false });
  };

  private loadChoiceOptions = async (): Promise<void> => {
    const choiceOptions = await this.actionPlanService.getAllChoiceOptions();
    this.setState({ choiceOptions });
  };

   private openDetailPanel = (actionPlan: IActionplan): void => {
     this.setState({
       selectedActionPlan: actionPlan,
       isDetailPanelOpen: true,
       isNewMode: false,
       formData: { ...actionPlan, PICId: actionPlan.PICId }
     });

     this.loadDepartmentServices(actionPlan.Department).catch(error => console.error('Failed to load Department services:', error));
   };

  private openNewActionPanel = (): void => {
    this.setState({
      selectedActionPlan: undefined,
      isDetailPanelOpen: true,
      isNewMode: true,
      formData: { Service: this.props.userService },
      departmentServices: [],
      isDepartmentServicesLoading: false
    });
  };

  private closeDetailPanel = (): void => {
    if (this.props.isInlineMode && this.props.onInlineDismiss) {
      this.props.onInlineDismiss();
      return;
    }
    this.setState({
      isDetailPanelOpen: false,
      selectedActionPlan: undefined,
      isNewMode: false,
      formData: {},
      departmentServices: [],
      isDepartmentServicesLoading: false
    });
  };

  private updateFormField = (field: keyof IActionPlanFormData, value: IActionPlanFormData[keyof IActionPlanFormData]): void => {
    this.setState(prev => ({
      formData: {
        ...prev.formData,
        [field]: value
      }
    }));
  };

  private isDigitalTechnologySupport(department: string | undefined): boolean {
    return (department || '').trim().toLowerCase() === this.digitalTechnologySupportDivision.toLowerCase();
  }

  private loadDepartmentServices = async (department: string | undefined): Promise<void> => {
    if (!department || this.isDigitalTechnologySupport(department)) {
      this.setState({ departmentServices: [], isDepartmentServicesLoading: false });
      return;
    }

    this.setState({ isDepartmentServicesLoading: true });
    const departmentServices = await this.divisionServiceService.getServicesByDivision(department);
    this.setState({ departmentServices, isDepartmentServicesLoading: false });
  };

  private handleDepartmentChange = (department: string): void => {
    const isDigitalTechnologySupport = this.isDigitalTechnologySupport(department);

    this.setState(prev => ({
      formData: {
        ...prev.formData,
        Department: department,
        Service: undefined,
        ProductLine: isDigitalTechnologySupport ? prev.formData.ProductLine : undefined
      },
      departmentServices: isDigitalTechnologySupport ? [] : prev.departmentServices,
      isDepartmentServicesLoading: isDigitalTechnologySupport ? false : prev.isDepartmentServicesLoading
    }));

    this.loadDepartmentServices(department).catch(error => console.error('Failed to load Department services:', error));
  };

  private handleSave = async (): Promise<void> => {
    const { isNewMode, formData } = this.state;
    
    // Ensure PICId is included
    const payload: IActionPlanUpsert = {
      ...formData,
      PICId: formData.PICId
    };

    console.log('Saving ActionPlan:', { isNewMode, payload });

    if (isNewMode) {
      const result = await this.actionPlanService.createActionPlan(payload);
      console.log('Create ActionPlan result:', result);
      if (result) {
        if (this.props.isInlineMode && this.props.onActionPlanCreated) {
          this.props.onActionPlanCreated(result.Id);
        } else {
          await this.loadActionPlans();
          this.closeDetailPanel();
        }
      }
    } else if (this.state.selectedActionPlan) {
      const success = await this.actionPlanService.updateActionPlan(
        this.state.selectedActionPlan.Id,
        payload
      );
      if (success) {
        await this.loadActionPlans();
        this.closeDetailPanel();
      }
    }
  };

  private getStatusClassName(status: string | undefined): string {
    switch (status?.toLowerCase()) {
      case 'open':
        return styles.statusOpen;
      case 'in progress':
        return styles.statusInProgress;
      case 'closed':
        return styles.statusClosed;
      default:
        return '';
    }
  }

  private renderGrid(): JSX.Element {
    const { actionPlans, isLoading } = this.state;

    return (
      <div className={styles.gridContainer}>
        {isLoading ? (
          <div className={styles.spinnerContainer}>
            <Spinner size={SpinnerSize.large} label="Loading action plans..." />
          </div>
        ) : (
          <>
            <div className={styles.gridHeader}>
              <div className={styles.colTitle}>Title</div>
              <div className={styles.colService}>Service</div>
              <div className={styles.colPIC}>PIC</div>
              <div className={styles.colStatus}>Status</div>
            </div>
            {actionPlans.length === 0 ? (
              <div className={styles.emptyMessage}>
                There is no items in Action List
              </div>
            ) : (
              actionPlans.map(plan => (
                <div
                  key={plan.Id}
                  className={styles.gridRow}
                  onClick={() => this.openDetailPanel(plan)}
                >
                  <div className={styles.colTitle}>{plan.Title}</div>
                  <div className={styles.colService}>{plan.Service || '-'}</div>
                  <div className={styles.colPIC}>
                    {plan.PIC?.Title || '-'}
                  </div>
                  <div className={styles.colStatus}>
                    <span className={`${styles.badge} ${this.getStatusClassName(plan.Status)}`}>
                      {plan.Status || '-'}
                    </span>
                  </div>
                </div>
              ))
            )}
          </>
        )}
      </div>
    );
  }

   private renderDetailPanel(): JSX.Element {
     const { formData, isNewMode, choiceOptions, isDetailPanelOpen, departmentServices, isDepartmentServicesLoading } = this.state;
     const isDigitalTechnologySupport = this.isDigitalTechnologySupport(formData.Department);
     const peoplePickerContext: IPeoplePickerContext = {
       absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
       msGraphClientFactory: this.props.context.msGraphClientFactory,
       spHttpClient: this.props.context.spHttpClient,
     };
     

     return (
       <Modal
        isOpen={isDetailPanelOpen}
        onDismiss={this.closeDetailPanel}
        isBlocking={false}
        containerClassName={styles.modalContainer}
      >
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

          {this.props.sourceFeedbackId !== undefined && (
            <div className={styles.formGroup}>
              <label>Customer Feedback ID</label>
              <input
                type="text"
                value={`#${this.props.sourceFeedbackId}`}
                readOnly
                style={{ background: '#f5f5f5', color: '#555', cursor: 'default' }}
              />
            </div>
          )}

          <div className={styles.formGroup}>
            <label>Department</label>
            <select
              value={formData.Department || ''}
              onChange={(e) => this.handleDepartmentChange(e.target.value)}
            >
              <option value="">Select Department</option>
              {Array.isArray(choiceOptions.Department) && choiceOptions.Department.length > 0 ? (
                choiceOptions.Department.map(choice => (
                  <option key={choice} value={choice}>{choice}</option>
                ))
              ) : (
                <option>Loading options...</option>
              )}
            </select>
          </div>

          <div className={styles.formGroup}>
            <label>Division</label>
            <select
              value={formData.Division || ''}
              onChange={(e) => this.updateFormField('Division', e.target.value)}
            >
              <option value="">Select Division</option>
              {Array.isArray(choiceOptions.Division) && choiceOptions.Division.length > 0 ? (
                choiceOptions.Division.map(choice => (
                  <option key={choice} value={choice}>{choice}</option>
                ))
              ) : (
                <option>Loading options...</option>
              )}
            </select>
          </div>

          {!isDigitalTechnologySupport && (
            <div className={styles.formGroup}>
              <label>Service</label>
              <select
                value={formData.Service || ''}
                onChange={(e) => this.updateFormField('Service', e.target.value)}
                disabled={isDepartmentServicesLoading || !formData.Department}
              >
                <option value="">Select Service</option>
                {departmentServices.length > 0 ? (
                  departmentServices.map(service => (
                    <option key={service} value={service}>{service}</option>
                  ))
                ) : (
                  <option disabled>{isDepartmentServicesLoading ? 'Loading options...' : 'No service available'}</option>
                )}
              </select>
            </div>
          )}

          {isDigitalTechnologySupport && (
            <div className={styles.formGroup}>
              <label>Product Line</label>
              <select
                value={formData.ProductLine || ''}
                onChange={(e) => this.updateFormField('ProductLine', e.target.value)}
              >
                <option value="">Select Product Line</option>
                {Array.isArray(choiceOptions.ProductLine) && choiceOptions.ProductLine.length > 0 ? (
                  choiceOptions.ProductLine.map(choice => (
                    <option key={choice} value={choice}>{choice}</option>
                  ))
                ) : (
                  <option>Loading options...</option>
                )}
              </select>
            </div>
          )}

          <div className={styles.formGroup}>
            <label>Timeline</label>
            <input
              type="date"
              value={formData.Timeline ? formData.Timeline.split('T')[0] : ''}
              onChange={(e) => this.updateFormField('Timeline', e.target.value)}
            />
          </div>

          <div className={styles.formGroup}>
            <label>Customer Feedback</label>
            <ReactQuill
              theme="snow"
              value={Array.isArray(formData.CustomerFeedback) ? formData.CustomerFeedback.join('\n') : (formData.CustomerFeedback || '')}
              onChange={(content) => this.updateFormField('CustomerFeedback', content)}
              placeholder="Enter feedback with full format..."
            />
          </div>

          <div className={styles.formGroup}>
            <label>Action</label>
            <ReactQuill
              theme="snow"
              value={Array.isArray(formData.Actions) ? formData.Actions.join('\n') : (formData.Actions || '')}
              onChange={(content) => this.updateFormField('Actions', content)}
              placeholder="Enter actions with full format..."
            />
          </div>

          <div className={styles.formGroup}>
            <label>PIC</label>
            <PeoplePicker
              context={peoplePickerContext}
              personSelectionLimit={1}
              groupName=""
              ensureUser={true}
              principalTypes={[PrincipalType.User]}
              defaultSelectedUsers={formData.PICId ? [formData.PIC?.EMail || ''] : []}
              onChange={(items: IPersonaProps[]) => {
                if (items.length > 0) {
                  // Extract user ID from various possible locations
                  let personId: number | undefined = undefined;
                  
                  // Try to get ID from items[0].id
                  if (items[0].id) {
                    const idValue = typeof items[0].id === 'string' ? Number(items[0].id) : items[0].id;
                    if (typeof idValue === 'number' && !isNaN(idValue)) {
                      personId = idValue;
                    }
                  }
                  
                  // Log for debugging
                  console.log('PIC Selected:', {
                    text: items[0].text,
                    email: items[0].secondaryText,
                    id: items[0].id,
                    extractedPersonId: personId,
                    fullItem: items[0]
                  });
                  
                  this.updateFormField('PICId', personId);
                  this.updateFormField('PIC', {
                    Title: items[0].text || '',
                    EMail: items[0].secondaryText || ''
                  });
                  return;
                }

                this.updateFormField('PICId', undefined);
                this.updateFormField('PIC', undefined);
              }}
            />
          </div>

          <div className={styles.formGroup}>
            <label>Category</label>
            <select
              value={formData.Category || ''}
              onChange={(e) => this.updateFormField('Category', e.target.value)}
            >
              <option value="">Select Category</option>
              {Array.isArray(choiceOptions.Category) && choiceOptions.Category.length > 0 ? (
                choiceOptions.Category.map(choice => (
                  <option key={choice} value={choice}>{choice}</option>
                ))
              ) : (
                <option>Loading options...</option>
              )}
            </select>
          </div>

          <div className={styles.formGroup}>
            <label>Status</label>
            <select
              value={formData.Status || ''}
              onChange={(e) => this.updateFormField('Status', e.target.value)}
            >
              <option value="">Select Status</option>
              {Array.isArray(choiceOptions.Status) && choiceOptions.Status.length > 0 ? (
                choiceOptions.Status.map(choice => (
                  <option key={choice} value={choice}>{choice}</option>
                ))
              ) : (
                <option>Loading options...</option>
              )}
            </select>
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
            <label>Result</label>
            <ReactQuill
              theme="snow"
              value={Array.isArray(formData.Results) ? formData.Results.join('\n') : (formData.Results || '')}
              onChange={(content) => this.updateFormField('Results', content)}
              placeholder="Enter results with full format..."
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
      </Modal>
    );
  }

  public render(): JSX.Element {
    if (this.props.isInlineMode) {
      return this.renderDetailPanel();
    }

    return (
      <main className={styles.mainContainer}>
        <div className={styles.breadcrumb}>Home › <strong>Action Plan</strong></div>

        <div className={styles.header}>
          <h1>Action Plans</h1>
          <button className={styles.btnNewAction} onClick={this.openNewActionPanel}>
            <Icon iconName="Add" /> New Action
          </button>
        </div>

        <div className={styles.content}>
          {this.renderGrid()}
        </div>
        {this.renderDetailPanel()}
      </main>
    );
  }
}
