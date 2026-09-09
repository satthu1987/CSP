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
  /**
   * Controls how action plans are loaded:
   * - 'admin': loads all action plans (no filter)
  * - 'leader': loads plans by services where current user is PIC in Division_Service list
  * - 'manager': loads plans by services where current user is Manager in Division_Service list
   * Defaults to 'leader' behaviour if omitted.
   */
  filterMode?: 'admin' | 'leader' | 'manager';
  /** When true, renders only the modal in new-plan mode (no full page) */
  isInlineMode?: boolean;
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
  filterService: string;
  filterPIC: string;
  filterStatus: string;
  filterYear: string;
  userServices: string[];
  isLoadingUserServices: boolean;
}

export default class ActionPlan extends React.Component<IActionPlanProps, IActionPlanState> {
  private actionPlanService: ActionPlanService;
  private divisionServiceService: DivisionServiceService;
  private readonly digitalTechnologySupportDivision = 'Digital Technology Support';

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
      },
      choiceOptions: {},
      departmentServices: [],
      isDepartmentServicesLoading: false,
      filterService: '',
      filterPIC: '',
      filterStatus: '',
      filterYear: '',
      userServices: [],
      isLoadingUserServices: true
    };
    this.actionPlanService = new ActionPlanService(props.context, 'CSP');
    this.divisionServiceService = new DivisionServiceService(props.context, 'Division_Service');
  }

  public async componentDidMount(): Promise<void> {
    await this.loadUserServices();
    await Promise.all([
      this.loadActionPlans(),
      this.loadChoiceOptions()
    ]);
  }

  private getCurrentUserEmail(): string {
    const user = this.props.context.pageContext.user;
    const rawEmail = (user && user.email ? user.email : '').trim();
    if (rawEmail) {
      return rawEmail;
    }

    const loginName = (user && user.loginName ? user.loginName : '').trim();
    const claimParts = loginName.split('|');
    return claimParts.length > 0 ? (claimParts[claimParts.length - 1] || '').trim() : loginName;
  }

  private loadUserServices = async (): Promise<void> => {
    this.setState({ isLoadingUserServices: true });
    try {
      const userEmail = this.getCurrentUserEmail();
      const { filterMode } = this.props;
      let userServices: string[] = [];

      if (filterMode === 'manager') {
        userServices = await this.divisionServiceService.getServicesByManager(userEmail);
      } else if (filterMode === 'leader') {
        userServices = await this.divisionServiceService.getServicesByPIC(userEmail);
      }

      this.setState({ userServices, isLoadingUserServices: false });
    } catch (error) {
      console.error('Error loading user services:', error);
      this.setState({ userServices: [], isLoadingUserServices: false });
    }
  };

  private loadActionPlans = async (): Promise<void> => {
    this.setState({ isLoading: true });
    try {
      const { filterMode } = this.props;
      const { userServices } = this.state;
      let actionPlans: IActionplan[] = [];

      if (filterMode === 'admin') {
        // Admin sees all action plans
        actionPlans = await this.actionPlanService.getAllActionPlans();
      } else if (userServices.length > 1) {
        actionPlans = await this.actionPlanService.getActionPlansByServices(userServices);
      } else if (userServices.length === 1) {
        actionPlans = await this.actionPlanService.getActionPlansByService(userServices[0]);
      } else {
        actionPlans = [];
      }

      this.setState({ actionPlans, isLoading: false });
    } catch (error) {
      console.error('Error loading action plans:', error);
      this.setState({ actionPlans: [], isLoading: false });
    }
  };

  private loadChoiceOptions = async (): Promise<void> => {
    const choiceOptions = await this.actionPlanService.getAllChoiceOptions();
    this.setState({ choiceOptions });
  };

   private openDetailPanel = (actionPlan: IActionplan): void => {
    const isDigitalTechnologySupport = this.isDigitalTechnologySupport(actionPlan.Department);
     this.setState({
       selectedActionPlan: actionPlan,
       isDetailPanelOpen: true,
       isNewMode: false,
       formData: {
         ...actionPlan,
         PICId: actionPlan.PICId,
         Category: Array.isArray(actionPlan.Category)
           ? actionPlan.Category
           : (actionPlan.Category ? [actionPlan.Category] : [])
       }
     });

     const serviceLookupDivision = isDigitalTechnologySupport
       ? (actionPlan.ProductLine || actionPlan.Department)
       : actionPlan.Department;
     this.loadDepartmentServices(serviceLookupDivision).catch(error => console.error('Failed to load Department services:', error));
   };

  private openNewActionPanel = async (): Promise<void> => {
    // Load division based on service first
    const division = await this.divisionServiceService.getDivisionByService(this.props.userService);
    const isDigitalTechnologySupport = this.isDigitalTechnologySupport(division);
    
    // Set state with both Service and Department values
    this.setState({
      selectedActionPlan: undefined,
      isDetailPanelOpen: true,
      isNewMode: true,
      formData: {
        Service: isDigitalTechnologySupport ? undefined : this.props.userService,
        ProductLine: undefined,
        Department: division || undefined
      },
      departmentServices: [],
      isDepartmentServicesLoading: false
    });

    // Load service options immediately only for non-DTS departments.
    if (division && !isDigitalTechnologySupport) {
      this.loadDepartmentServices(division).catch(error => console.error('Failed to load Department services:', error));
    }
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
    if (!department) {
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
      departmentServices: [],
      isDepartmentServicesLoading: false
    }));

    if (!department || isDigitalTechnologySupport) {
      return;
    }

    this.loadDepartmentServices(department).catch(error => console.error('Failed to load Department services:', error));
  };

  private handleProductLineChange = (productLine: string): void => {
    this.setState(prev => ({
      formData: {
        ...prev.formData,
        ProductLine: productLine,
        Service: undefined,
      },
      departmentServices: [],
      isDepartmentServicesLoading: false
    }));

    if (!productLine) {
      return;
    }

    this.loadDepartmentServices(productLine).catch(error => console.error('Failed to load Product Line services:', error));
  };

  private handleSave = async (): Promise<void> => {
    const { isNewMode, formData } = this.state;
    const selectedCategories = this.getSelectedCategories();
    
    // Ensure PICId is included
    const payload: IActionPlanUpsert = {
      ...formData,
      PICId: formData.PICId,
      Category: selectedCategories
    };

    if (selectedCategories.length === 0) {
      delete payload.Category;
    }

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

  private getFilteredActionPlans(): IActionplan[] {
    const { actionPlans, filterService, filterPIC, filterStatus, filterYear } = this.state;

    return actionPlans.filter(plan => {
      const serviceMatch = filterService === '' || (plan.Service || '').toLowerCase() === filterService.toLowerCase();
      const picMatch = (plan.PIC?.Title || '').toLowerCase().indexOf(filterPIC.toLowerCase()) > -1;
      const statusMatch = filterStatus === '' || (plan.Status || '').toLowerCase() === filterStatus.toLowerCase();
      const yearMatch = filterYear === '' || (plan.Year || '') === filterYear;

      return serviceMatch && picMatch && statusMatch && yearMatch;
    });
  }

  private getServiceFilterOptions(): string[] {
    const { actionPlans, choiceOptions, userServices } = this.state;
    const serviceMap: { [key: string]: boolean } = {};
    const options: string[] = [];

    actionPlans.forEach(plan => {
      const service = (plan.Service || '').trim();
      if (service && !serviceMap[service]) {
        serviceMap[service] = true;
        options.push(service);
      }
    });

    if (options.length === 0) {
      userServices.forEach(service => {
        const value = (service || '').trim();
        if (value && !serviceMap[value]) {
          serviceMap[value] = true;
          options.push(value);
        }
      });
    }

    if (options.length === 0 && Array.isArray(choiceOptions.Service)) {
      choiceOptions.Service.forEach(service => {
        const value = (service || '').trim();
        if (value && !serviceMap[value]) {
          serviceMap[value] = true;
          options.push(value);
        }
      });
    }

    options.sort();
    return options;
  }

  private getSelectedCategories(): string[] {
    const categoryValue = this.state.formData.Category;

    if (Array.isArray(categoryValue)) {
      return categoryValue;
    }

    if (typeof categoryValue === 'string' && categoryValue.trim()) {
      return [categoryValue.trim()];
    }

    return [];
  }

  private toggleCategoryChoice(choice: string): void {
    const selectedCategories = this.getSelectedCategories().slice();
    const selectedIndex = selectedCategories.indexOf(choice);

    if (selectedIndex >= 0) {
      selectedCategories.splice(selectedIndex, 1);
    } else {
      selectedCategories.push(choice);
    }

    this.updateFormField('Category', selectedCategories);
  }

  private getYearOptions(): string[] {
    const currentYear = new Date().getFullYear();
    const years: string[] = [];
    for (let i = -3; i <= 0; i++) {
      years.push((currentYear + i).toString());
    }
    return years;
  }

  private htmlToPlainText(value: string | undefined): string {
    if (!value) {
      return '';
    }

    const htmlWithLineBreaks = value.replace(/<\s*br\s*\/?>/gi, '\n');
    const temp = document.createElement('div');
    temp.innerHTML = htmlWithLineBreaks;

    const plainText = (temp.textContent || temp.innerText || '')
      .replace(/\u00a0/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();

    return plainText;
  }

  private getGridPreviewText(value: string | undefined, maxLength: number = 50): string {
    const plainText = this.htmlToPlainText(value);
    if (!plainText) {
      return '-';
    }

    return plainText.length > maxLength
      ? `${plainText.substring(0, maxLength)}...`
      : plainText;
  }

  private renderGrid(): JSX.Element {
    const { actionPlans, isLoading, filterService, filterPIC, filterStatus, filterYear, choiceOptions } = this.state;
    const filteredPlans = this.getFilteredActionPlans();
    const statusOptions = Array.isArray(choiceOptions.Status) ? choiceOptions.Status : [];
    const serviceOptions = this.getServiceFilterOptions();
    
    // Build unique PIC list without Array.from for ES5 compatibility
    const picSet: { [key: string]: boolean } = {};
    const uniquePICs: string[] = [];
    actionPlans.forEach(plan => {
      const picTitle = plan.PIC?.Title;
      if (picTitle && !picSet[picTitle]) {
        picSet[picTitle] = true;
        uniquePICs.push(picTitle);
      }
    });
    uniquePICs.sort();

    return (
      <div className={styles.gridContainer}>
        {isLoading ? (
          <div className={styles.spinnerContainer}>
            <Spinner size={SpinnerSize.large} label="Loading action plans..." />
          </div>
        ) : (
          <>
            <div className={styles.filterSection}>
              <div className={styles.filterGroup}>
                <label>Service</label>
                <select
                  value={filterService}
                  onChange={(e) => this.setState({ filterService: e.target.value })}
                  className={styles.filterSelect}
                >
                  <option value="">All Services</option>
                  {serviceOptions.length > 0 ? (
                    serviceOptions.map(service => (
                      <option key={service} value={service}>{service}</option>
                    ))
                  ) : (
                    <option disabled>No services available</option>
                  )}
                </select>
              </div>

              <div className={styles.filterGroup}>
                <label>Year</label>
                <select
                  value={filterYear}
                  onChange={(e) => this.setState({ filterYear: e.target.value })}
                  className={styles.filterSelect}
                >
                  <option value="">All Years</option>
                  {this.getYearOptions().map(year => (
                    <option key={year} value={year}>{year}</option>
                  ))}
                </select>
              </div>

              <div className={styles.filterGroup}>
                <label>PIC</label>
                <select
                  value={filterPIC}
                  onChange={(e) => this.setState({ filterPIC: e.target.value })}
                  className={styles.filterSelect}
                >
                  <option value="">All PICs</option>
                  {uniquePICs.map((pic: string) => (
                    <option key={pic} value={pic}>{pic}</option>
                  ))}
                </select>
              </div>

              <div className={styles.filterGroup}>
                <label>Status</label>
                <select
                  value={filterStatus}
                  onChange={(e) => this.setState({ filterStatus: e.target.value })}
                  className={styles.filterSelect}
                >
                  <option value="">All Status</option>
                  {statusOptions.map(status => (
                    <option key={status} value={status}>{status}</option>
                  ))}
                </select>
              </div>

              <button
                className={styles.clearFilterBtn}
                onClick={() => this.setState({ filterService: '', filterPIC: '', filterStatus: '', filterYear: '' })}
              >
                Clear Filters
              </button>
            </div>

            <div className={styles.gridHeader}>
              <div className={styles.colService}>Service</div>
              <div className={styles.colTitle}>Original Customer Feedback</div>
              <div className={styles.colTitle}>Updated Feedback</div>
              <div className={styles.colActions}>Action</div>
              <div className={styles.colPIC}>PIC</div>
              <div className={styles.colStatus}>Status</div>
              <div className={styles.colResult}>Action Output</div>
              <div className={styles.colRelatedLinks}>Related Links</div>
              <div className={styles.colAction} />
            </div>
            {filteredPlans.length === 0 ? (
              <div className={styles.emptyMessage}>
                {actionPlans.length === 0 ? 'There is no items in Action List' : 'No results matching the filters'}
              </div>
            ) : (
              filteredPlans.map(plan => (
                <div
                  key={plan.Id}
                  className={styles.gridRow}
                >
                  <div className={styles.colService}>{plan.Service || '-'}</div>
                  <div className={styles.colTitle}>
                    {this.getGridPreviewText(plan.CustomerFeedback)}
                  </div>
                  <div className={styles.colTitle}>
                    {this.getGridPreviewText(plan.UpdatedFeedback)}
                  </div>
                  <div className={styles.colActions}>
                    {this.getGridPreviewText(Array.isArray(plan.Actions) ? plan.Actions.join('\n') : plan.Actions)}
                  </div>
                  <div className={styles.colPIC}>{plan.PIC?.Title || '-'}</div>
                  <div className={styles.colStatus}>
                    <span className={`${styles.badge} ${this.getStatusClassName(plan.Status)}`}>
                      {plan.Status || '-'}
                    </span>
                  </div>
                  <div className={styles.colResult}>
                    {this.getGridPreviewText(Array.isArray(plan.Results) ? plan.Results.join('\n') : (plan.Results as unknown as string))}
                  </div>
                  <div className={styles.colRelatedLinks}>
                    {plan.RelatedLinks ? <a href={this.getGridPreviewText(plan.RelatedLinks) || '—'}>View</a> : "-" }
                  </div>
                  <div className={styles.colAction}>
                    {!plan.Actions || plan.Actions.length === 0 ? (
                      <Icon
                        iconName="Add"
                        className={styles.addIcon}
                        onClick={() => this.openDetailPanel(plan)}
                        title="Add Action Plan"
                      />
                    ) : (
                      <Icon
                        iconName="Edit"
                        className={styles.editIcon}
                        onClick={() => this.openDetailPanel(plan)}
                        title="Edit Action Plan"
                      />
                    )}
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
          {/* <div className={styles.formGroup}>
            <label>Title</label>
            <input
              type="text"
              value={formData.Title || ''}
              onChange={(e) => this.updateFormField('Title', e.target.value)}
              placeholder="Enter title"
            />
          </div> */}


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

          {isDigitalTechnologySupport && (
            <div className={styles.formGroup}>
              <label>Product Line</label>
              <select
                value={formData.ProductLine || ''}
                onChange={(e) => this.handleProductLineChange(e.target.value)}
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
            <label>Service</label>
            <select
              value={formData.Service || ''}
              onChange={(e) => this.updateFormField('Service', e.target.value)}
              disabled={
                isDepartmentServicesLoading ||
                !formData.Department ||
                (isDigitalTechnologySupport && !formData.ProductLine)
              }
            >
              <option value="">Select Service</option>
              {departmentServices.length > 0 ? (
                departmentServices.map(service => (
                  <option key={service} value={service}>{service}</option>
                ))
              ) : (
                <option disabled>
                  {isDepartmentServicesLoading
                    ? 'Loading options...'
                    : (isDigitalTechnologySupport && !formData.ProductLine
                      ? 'Please select Product Line first'
                      : 'No service available')}
                </option>
              )}
            </select>
          </div>

          <div className={styles.formGroup}>
            <label>Timeline</label>
            <input
              type="date"
              value={formData.Timeline ? formData.Timeline.split('T')[0] : ''}
              onChange={(e) => this.updateFormField('Timeline', e.target.value)}
            />
          </div>

        {isNewMode ?
          
          <div className={styles.formGroup}>
            <label>Updated Feedback</label>
            <ReactQuill
              theme="snow"
              value={formData.UpdatedFeedback || ''}
              onChange={(content) => this.updateFormField('UpdatedFeedback', content)}
              placeholder="Enter updated feedback..."
            />
          </div>
:
          (<><div className={styles.formGroup}>
            <label>Customer Feedback</label>
            <ReactQuill
              theme="snow"
              value={formData.CustomerFeedback || ''}
              onChange={(content) => this.updateFormField('CustomerFeedback', content)}
              placeholder="Enter customer feedback..."
            />
          </div>
          <div className={styles.formGroup}>
            <label>Updated Feedback</label>
            <ReactQuill
              theme="snow"
              value={formData.UpdatedFeedback || ''}
              onChange={(content) => this.updateFormField('UpdatedFeedback', content)}
              placeholder="Enter updated feedback..."
            />
          </div></>)
   }
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
            <div className={styles.multiChoiceBox}>
              {Array.isArray(choiceOptions.Category) && choiceOptions.Category.length > 0 ? (
                choiceOptions.Category.map(choice => (
                  <label key={choice} className={styles.multiChoiceItem}>
                    <input
                      type="checkbox"
                      checked={this.getSelectedCategories().indexOf(choice) >= 0}
                      onChange={() => this.toggleCategoryChoice(choice)}
                    />
                    <span>{choice}</span>
                  </label>
                ))
              ) : (
                <div className={styles.fieldHint}>Loading options...</div>
              )}
            </div>
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
            <label>Action Output</label>
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
          <button className={styles.btnNewAction} onClick={() => this.openNewActionPanel()}>
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
