import * as React from 'react';
import styles from './DivisionServiceManagement.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Icon, IPersonaProps, Modal, Spinner, SpinnerSize } from '@fluentui/react';
import { IPeoplePickerContext, PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { DivisionServiceService } from '../../services/DivisionService_Service';

export interface IDivisionServiceManagementProps {
  context: WebPartContext;
}

interface IDivisionServiceItem {
  Id: number;
  Division: string;
  Service: string;
  Year: string;
  PICId?: number;
  PICTitle?: string;
  PICEmail?: string;
  ManagerId?: number;
  ManagerTitle?: string;
  ManagerEmail?: string;
}

interface IFormData {
  division: string;
  service: string;
  year: string;
  picId?: number;
  picDisplayName?: string;
  picEmail?: string;
  managerId?: number;
  managerDisplayName?: string;
  managerEmail?: string;
}

interface IDivisionServiceManagementState {
  items: IDivisionServiceItem[];
  isLoading: boolean;
  selectedYear: string;
  availableYears: string[];
  isPanelOpen: boolean;
  selectedItem?: IDivisionServiceItem;
  formData: IFormData;
  isSaving: boolean;
  errorMsg: string;
}

export default class DivisionServiceManagement extends React.Component<
  IDivisionServiceManagementProps,
  IDivisionServiceManagementState
> {
  private divisionServiceService: DivisionServiceService;

  constructor(props: IDivisionServiceManagementProps) {
    super(props);
    const currentYear = new Date().getFullYear().toString();
    this.state = {
      items: [],
      isLoading: true,
      selectedYear: currentYear,
      availableYears: [currentYear],
      isPanelOpen: false,
      selectedItem: undefined,
      formData: { division: '', service: '', year: currentYear },
      isSaving: false,
      errorMsg: '',
    };
    this.divisionServiceService = new DivisionServiceService(props.context, 'Division_Service');
  }

  public async componentDidMount(): Promise<void> {
    const selectedYear = await this.loadYears();
    await this.loadItems(selectedYear);
  }

  private loadYears = async (): Promise<string> => {
    const currentYear = new Date().getFullYear().toString();
    try {
      const years = await this.divisionServiceService.getAllYears();
      const availableYears = years.length > 0 ? years : [currentYear];
      const selectedYear = availableYears.indexOf(currentYear) > -1 ? currentYear : availableYears[0];
      this.setState({ availableYears, selectedYear });
      return selectedYear;
    } catch (error) {
      console.error('Failed to load Division_Service years:', error);
      return currentYear;
    }
  };

  private loadItems = async (year: string): Promise<void> => {
    this.setState({ isLoading: true });
    try {
      const items = await this.divisionServiceService.getAllItemsWithDetails(year);
      this.setState({ items, isLoading: false });
    } catch (error) {
      console.error('Failed to load Division_Service items:', error);
      this.setState({ items: [], isLoading: false });
    }
  };

  private onYearFilterChange = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    const selectedYear = event.target.value;
    this.setState({ selectedYear });
    this.loadItems(selectedYear).catch(error => {
      console.error('Error changing year filter:', error);
    });
  };

  private openEditPanel = (item: IDivisionServiceItem): void => {
    this.setState({
      isPanelOpen: true,
      selectedItem: item,
      formData: {
        division: item.Division,
        service: item.Service,
        year: item.Year,
        picId: item.PICId,
        picDisplayName: item.PICTitle,
        picEmail: item.PICEmail,
        managerId: item.ManagerId,
        managerDisplayName: item.ManagerTitle,
        managerEmail: item.ManagerEmail,
      },
      errorMsg: '',
    });
  };

  private closePanel = (): void => {
    this.setState({ isPanelOpen: false, selectedItem: undefined, errorMsg: '' });
  };

  private handleSave = async (): Promise<void> => {
    const { formData, selectedItem } = this.state;

    if (!selectedItem) {
      return;
    }
    if (!formData.division.trim() || !formData.service.trim() || !formData.year.trim()) {
      this.setState({ errorMsg: 'Division, Service and Year are required.' });
      return;
    }

    this.setState({ isSaving: true, errorMsg: '' });

    const success = await this.divisionServiceService.updateItemDetails(selectedItem.Id, {
      division: formData.division.trim(),
      service: formData.service.trim(),
      year: formData.year.trim(),
      picId: formData.picId,
      managerId: formData.managerId,
    });

    if (success) {
      this.setState({ isSaving: false, isPanelOpen: false, selectedItem: undefined });
      await this.loadYears();
      await this.loadItems(this.state.selectedYear);
    } else {
      this.setState({ isSaving: false, errorMsg: 'Failed to save. Please try again.' });
    }
  };

  private renderGrid(): JSX.Element {
    const { items, isLoading } = this.state;

    if (isLoading) {
      return (
        <div className={styles.spinnerContainer}>
          <Spinner size={SpinnerSize.large} label="Loading services..." />
        </div>
      );
    }

    if (items.length === 0) {
      return <div className={styles.emptyState}>No Division/Service records found for the selected year.</div>;
    }

    return (
      <div className={styles.gridContainer}>
        <div className={styles.gridHeader}>
          <div className={styles.colNo}>#</div>
          <div className={styles.colDivision}>Division</div>
          <div className={styles.colService}>Service</div>
          <div className={styles.colYear}>Year</div>
          <div className={styles.colPIC}>Service/Project Lead</div>
          <div className={styles.colPIC}>Manager</div>
          <div className={styles.colAction}>Action</div>
        </div>
        {items.map((item, index) => (
          <div key={item.Id} className={styles.gridRow}>
            <div className={styles.colNo}>{index + 1}</div>
            <div className={styles.colDivision}>{item.Division || '—'}</div>
            <div className={styles.colService}>{item.Service || '—'}</div>
            <div className={styles.colYear}>{item.Year || '—'}</div>
            <div className={styles.colPIC}>{item.PICTitle || '—'}</div>
            <div className={styles.colPIC}>{item.ManagerTitle || '—'}</div>
            <div className={styles.colAction}>
              <button
                className={styles.editIcon}
                title="Edit"
                onClick={() => this.openEditPanel(item)}
              >
                <Icon iconName="Edit" />
              </button>
            </div>
          </div>
        ))}
      </div>
    );
  }

  private renderPanel(): JSX.Element {
    const { isPanelOpen, formData, isSaving, errorMsg } = this.state;

    const peoplePickerContext: IPeoplePickerContext = {
      absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
      msGraphClientFactory: this.props.context.msGraphClientFactory,
      spHttpClient: this.props.context.spHttpClient,
    };

    const defaultUsers = formData.picEmail ? [formData.picEmail] : [];
    const defaultManagerUsers = formData.managerEmail ? [formData.managerEmail] : [];

    return (
      <Modal
        isOpen={isPanelOpen}
        onDismiss={this.closePanel}
        isBlocking={false}
        containerClassName={styles.modalContainer}
      >
        <div className={styles.detailPanel}>
          <div className={styles.panelHeader}>
            <h2>Edit Division / Service</h2>
            <button className={styles.closeBtn} onClick={this.closePanel}>
              <Icon iconName="Cancel" />
            </button>
          </div>

          <div className={styles.panelBody}>
            <div className={styles.formGroup}>
              <label>Division</label>
              <input
                type="text"
                value={formData.division}
                onChange={(e) => this.setState({ formData: { ...formData, division: e.target.value } })}
                disabled={isSaving}
              />
            </div>

            <div className={styles.formGroup}>
              <label>Service</label>
              <input
                type="text"
                value={formData.service}
                onChange={(e) => this.setState({ formData: { ...formData, service: e.target.value } })}
                disabled={isSaving}
              />
            </div>

            <div className={styles.formGroup}>
              <label>Year</label>
              <input
                type="text"
                value={formData.year}
                onChange={(e) => this.setState({ formData: { ...formData, year: e.target.value } })}
                disabled={isSaving}
              />
            </div>

            <div className={styles.formGroup}>
              <label>Service/Project Lead</label>
              <PeoplePicker
                key={String(this.state.selectedItem?.Id)}
                context={peoplePickerContext}
                personSelectionLimit={1}
                groupName=""
                ensureUser={true}
                principalTypes={[PrincipalType.User]}
                defaultSelectedUsers={defaultUsers}
                onChange={(items: IPersonaProps[]) => {
                  if (items.length > 0) {
                    const rawId = items[0].id;
                    const picId = rawId ? Number(rawId) : undefined;
                    this.setState({
                      formData: {
                        ...formData,
                        picId: picId && !isNaN(picId) ? picId : undefined,
                        picDisplayName: items[0].text || '',
                        picEmail: items[0].secondaryText || '',
                      },
                    });
                  } else {
                    this.setState({
                      formData: { ...formData, picId: undefined, picDisplayName: '', picEmail: '' },
                    });
                  }
                }}
              />
            </div>

            <div className={styles.formGroup}>
              <label>Manager</label>
              <PeoplePicker
                key={String(this.state.selectedItem?.Id)}
                context={peoplePickerContext}
                personSelectionLimit={1}
                groupName=""
                ensureUser={true}
                principalTypes={[PrincipalType.User]}
                defaultSelectedUsers={defaultManagerUsers}
                onChange={(items: IPersonaProps[]) => {
                  if (items.length > 0) {
                    const rawId = items[0].id;
                    const managerId = rawId ? Number(rawId) : undefined;
                    this.setState({
                      formData: {
                        ...formData,
                        managerId: managerId && !isNaN(managerId) ? managerId : undefined,
                        managerDisplayName: items[0].text || '',
                        managerEmail: items[0].secondaryText || '',
                      },
                    });
                  } else {
                    this.setState({
                      formData: { ...formData, managerId: undefined, managerDisplayName: '', managerEmail: '' },
                    });
                  }
                }}
              />
            </div>

            {errorMsg && <div className={styles.errorMsg}>{errorMsg}</div>}
          </div>

          <div className={styles.panelFooter}>
            <button className={styles.btnCancel} onClick={this.closePanel} disabled={isSaving}>
              Cancel
            </button>
            <button className={styles.btnSave} onClick={this.handleSave} disabled={isSaving}>
              {isSaving ? 'Saving...' : 'Save'}
            </button>
          </div>
        </div>
      </Modal>
    );
  }

  public render(): JSX.Element {
    const { selectedYear, availableYears } = this.state;

    return (
      <main className={styles.mainContainer}>
        <div className={styles.breadcrumb}>
          Home › <strong>Division – Service Management</strong>
        </div>

        <div className={styles.header}>
          <h1>Division – Service Management</h1>
        </div>

        <div className={styles.content}>
          <div className={styles.filterSection}>
            <div className={styles.filterGroup}>
              <label htmlFor="division-service-year-filter">Year</label>
              <select
                id="division-service-year-filter"
                className={styles.filterSelect}
                value={selectedYear}
                onChange={this.onYearFilterChange}
              >
                {availableYears.map(year => (
                  <option key={year} value={year}>
                    {year}
                  </option>
                ))}
              </select>
            </div>
          </div>

          {this.renderGrid()}
        </div>

        {this.renderPanel()}
      </main>
    );
  }
}
