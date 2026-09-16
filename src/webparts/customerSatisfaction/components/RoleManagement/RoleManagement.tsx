import * as React from 'react';
import styles from './RoleManagement.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Icon, IPersonaProps, Modal, Spinner, SpinnerSize } from '@fluentui/react';
import { IPeoplePickerContext, PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { CSPUserRole_Service } from '../../services/CSPUserRole_Service';
import { ICSPUserRole } from '../../Models/CSPUserRole';
import { DivisionServiceService } from '../../services/DivisionService_Service';

export interface IRoleManagementProps {
  context: WebPartContext;
}

interface IFormData {
  picId?: number;
  picDisplayName?: string;
  picEmail?: string;
  role: string; // visitor, leader, admin
  selectedAssignments: string[];
}

interface IRoleManagementState {
  roles: ICSPUserRole[];
  roleScopes: { [roleId: number]: string[] };
  isLoading: boolean;
  isPanelOpen: boolean;
  isNewMode: boolean;
  selectedRole?: ICSPUserRole;
  formData: IFormData;
  isSaving: boolean;
  isLoadingAssignments: boolean;
  assignmentOptions: string[];
  errorMsg: string;
}

export default class RoleManagement extends React.Component<IRoleManagementProps, IRoleManagementState> {
  private roleService: CSPUserRole_Service;
  private divisionService: DivisionServiceService;

  constructor(props: IRoleManagementProps) {
    super(props);
    this.state = {
      roles: [],
      roleScopes: {},
      isLoading: true,
      isPanelOpen: false,
      isNewMode: false,
      selectedRole: undefined,
      formData: { role: 'visitor', selectedAssignments: [] },
      isSaving: false,
      isLoadingAssignments: false,
      assignmentOptions: [],
      errorMsg: '',
    };
    this.roleService = new CSPUserRole_Service(props.context);
    this.divisionService = new DivisionServiceService(props.context, 'Division_Service');
  }

  public async componentDidMount(): Promise<void> {
    await this.loadRoles();
  }

  private loadRoles = async (): Promise<void> => {
    this.setState({ isLoading: true });
    try {
      const roles = await this.roleService.getAllRoles();
      const roleScopes = await this.loadRoleScopes(roles);
      this.setState({ roles, roleScopes, isLoading: false });
    } catch (error) {
      console.error('Error loading roles:', error);
      this.setState({ isLoading: false });
    }
  };

  private loadRoleScopes = async (roles: ICSPUserRole[]): Promise<{ [roleId: number]: string[] }> => {
    const scopeEntries = await Promise.all(roles.map(async role => {
      const roleName = (role.Role || '').trim().toLowerCase();
      const email = role.PIC?.EMail;

      if (!email || !this.requiresAssignments(roleName)) {
        return [role.Id, []] as [number, string[]];
      }

      const scope = roleName === 'manager'
        ? await this.divisionService.getDivisionsByManager(email)
        : await this.divisionService.getServicesByPIC(email);

      return [role.Id, scope] as [number, string[]];
    }));

    return scopeEntries.reduce((accumulator, entry) => {
      accumulator[entry[0]] = entry[1];
      return accumulator;
    }, {} as { [roleId: number]: string[] });
  };

  private renderRoleScope(role: ICSPUserRole): JSX.Element {
    const roleName = (role.Role || '').trim().toLowerCase();
    const scope = this.state.roleScopes[role.Id] || [];

    if (roleName !== 'Leader' && roleName !== 'Manager') {
      return <span className={styles.scopeEmpty}>—</span>;
    }

    if (scope.length === 0) {
      return <span className={styles.scopeEmpty}>No {roleName === 'Manager' ? 'divisions' : 'services'} assigned</span>;
    }

    return (
      <div className={styles.scopeList}>
        {scope.map(item => (
          <span key={item} className={styles.scopeTag}>{item}</span>
        ))}
      </div>
    );
  }

  private openNewPanel = (): void => {
    this.setState({
      isPanelOpen: true,
      isNewMode: true,
      selectedRole: undefined,
      formData: { role: 'visitor', selectedAssignments: [] },
      assignmentOptions: [],
      isLoadingAssignments: false,
      errorMsg: '',
    });
  };

  private openEditPanel = async (role: ICSPUserRole): Promise<void> => {
    this.setState({
      isPanelOpen: true,
      isNewMode: false,
      selectedRole: role,
      formData: {
        role: role.Role || 'visitor',
        picId: role.PIC?.Id,
        picDisplayName: role.PIC?.Title,
        picEmail: role.PIC?.EMail,
        selectedAssignments: [],
      },
      assignmentOptions: [],
      isLoadingAssignments: false,
      errorMsg: '',
    });

    if (role.PIC?.EMail) {
      await this.loadAssignmentOptions(role.Role || 'visitor', role.PIC.EMail);
    }
  };

  private closePanel = (): void => {
    this.setState({
      isPanelOpen: false,
      selectedRole: undefined,
      isLoadingAssignments: false,
      assignmentOptions: [],
      errorMsg: '',
    });
  };

  private requiresAssignments(role: string): boolean {
    return role === 'Manager' || role === 'Leader';
  }

  private getAssignmentLabel(role: string): string {
    return role === 'Manager' ? 'Divisions' : 'Services';
  }

  private getAssignmentPlaceholder(role: string): string {
    return role === 'Manager'
      ? 'Select divisions this manager will manage.'
      : 'Select services this leader will lead.';
  }

  private loadAssignmentOptions = async (role: string, userEmail?: string): Promise<void> => {
    if (!this.requiresAssignments(role)) {
      this.setState({ assignmentOptions: [], isLoadingAssignments: false });
      return;
    }

    this.setState({ isLoadingAssignments: true, assignmentOptions: [] });

    try {
      const options = role === 'Manager'
        ? await this.divisionService.getAllDivisions()
        : await this.divisionService.getAllServices();

      let selectedAssignments: string[] = [];
      if (userEmail) {
        selectedAssignments = role === 'Manager'
          ? await this.divisionService.getDivisionsByManager(userEmail)
          : await this.divisionService.getServicesByPIC(userEmail);
      }

      this.setState(prevState => ({
        assignmentOptions: options,
        isLoadingAssignments: false,
        formData: {
          ...prevState.formData,
          selectedAssignments,
        },
      }));
    } catch (error) {
      console.error('Failed to load role assignment options:', error);
      this.setState({ isLoadingAssignments: false, assignmentOptions: [] });
    }
  };

  private handleRoleChange = async (role: string): Promise<void> => {
    this.setState(prevState => ({
      formData: {
        ...prevState.formData,
        role,
        selectedAssignments: [],
      },
      errorMsg: '',
    }));

    await this.loadAssignmentOptions(role, this.state.formData.picEmail);
  };

  private toggleAssignment = (value: string): void => {
    this.setState(prevState => {
      const selectedAssignments = prevState.formData.selectedAssignments.slice();
      const index = selectedAssignments.indexOf(value);

      if (index >= 0) {
        selectedAssignments.splice(index, 1);
      } else {
        selectedAssignments.push(value);
      }

      return {
        formData: {
          ...prevState.formData,
          selectedAssignments,
        },
      };
    });
  };

  private handleSave = async (): Promise<void> => {
    const { formData, isNewMode, selectedRole } = this.state;

    if (!formData.picId) {
      this.setState({ errorMsg: 'PIC is required.' });
      return;
    }
    if (!formData.role.trim()) {
      this.setState({ errorMsg: 'Role is required.' });
      return;
    }
    if (this.requiresAssignments(formData.role) && formData.selectedAssignments.length === 0) {
      this.setState({
        errorMsg: `Please select at least one ${this.getAssignmentLabel(formData.role).toLowerCase().slice(0, -1)}.`,
      });
      return;
    }

    this.setState({ isSaving: true, errorMsg: '' });

    let success = false;
    if (isNewMode) {
      success = await this.roleService.createRole(formData.picId, formData.role.trim());
    } else if (selectedRole) {
      success = await this.roleService.updateRole(selectedRole.Id, formData.picId, formData.role.trim());
    }

    if (success && formData.picId) {
      if (formData.role === 'Manager') {
        success = await this.divisionService.assignManagerToDivisions(formData.picId, formData.selectedAssignments);
      } else if (formData.role === 'Leader') {
        success = await this.divisionService.assignLeaderToServices(formData.picId, formData.selectedAssignments);
      }
    }

    if (success) {
      this.setState({
        isSaving: false,
        isPanelOpen: false,
        assignmentOptions: [],
        isLoadingAssignments: false,
      });
      await this.loadRoles();
    } else {
      this.setState({ isSaving: false, errorMsg: 'Failed to save. Please try again.' });
    }
  };

  private handleDelete = async (id: number, picId?: number): Promise<void> => {
    if (!window.confirm('Are you sure you want to delete this role assignment?')) {
      return;
    }

    try {
      const success = await this.roleService.deleteRole(id);
      if (success) {
        if (picId) {
          await this.divisionService.clearAssignmentsForUser(picId);
        }
        await this.loadRoles();
      } else {
        alert('Failed to delete role assignment.');
      }
    } catch (error) {
      console.error('Error deleting role:', error);
      alert('Error deleting role assignment.');
    }
  };

  private renderGrid(): JSX.Element {
    const { roles, isLoading } = this.state;

    if (isLoading) {
      return (
        <div className={styles.spinnerContainer}>
          <Spinner size={SpinnerSize.large} label="Loading roles..." />
        </div>
      );
    }

    if (roles.length === 0) {
      return <div className={styles.emptyState}>No roles found. Click &quot;New Role&quot; to create one.</div>;
    }

    return (
      <div className={styles.gridContainer}>
        <div className={styles.gridHeader}>
          <div className={styles.colNo}>#</div>
          <div className={styles.colPIC}>PIC</div>
          <div className={styles.colRole}>Role</div>
          <div className={styles.colScope}>Scope</div>
          <div className={styles.colAction}>Action</div>
        </div>
        {roles.map((role, index) => (
          <div key={role.Id} className={styles.gridRow}>
            <div className={styles.colNo}>{index + 1}</div>
            <div className={styles.colPIC}>{role.PIC?.Title || '—'}</div>
            <div className={styles.colRole}>{role.Role || '—'}</div>
            <div className={styles.colScope}>{this.renderRoleScope(role)}</div>
            <div className={styles.colAction}>
              <button
                className={styles.editIcon}
                title="Edit"
                onClick={() => this.openEditPanel(role)}
              >
                <Icon iconName="Edit" />
              </button>
              <button
                className={styles.editIcon}
                title="Delete"
                onClick={() => this.handleDelete(role.Id, role.PIC?.Id)}
              >
                <Icon iconName="Delete" />
              </button>
            </div>
          </div>
        ))}
      </div>
    );
  }

  private renderPanel(): JSX.Element {
    const { isPanelOpen, isNewMode, formData, isSaving, errorMsg, assignmentOptions, isLoadingAssignments } = this.state;

    const peoplePickerContext: IPeoplePickerContext = {
      absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
      msGraphClientFactory: this.props.context.msGraphClientFactory,
      spHttpClient: this.props.context.spHttpClient,
    };

    const defaultUsers = formData.picEmail ? [formData.picEmail] : [];

    return (
      <Modal
        isOpen={isPanelOpen}
        onDismiss={this.closePanel}
        isBlocking={false}
        containerClassName={styles.modalContainer}
      >
        <div className={styles.detailPanel}>
          <div className={styles.panelHeader}>
            <h2>{isNewMode ? 'Assign User Role' : 'Edit User Role'}</h2>
            <button className={styles.closeBtn} onClick={this.closePanel}>
              <Icon iconName="Cancel" />
            </button>
          </div>

          <div className={styles.panelBody}>
            <div className={styles.formGroup}>
              <label>Role</label>
              <select
                value={formData.role}
                onChange={(e) => {
                  this.handleRoleChange(e.target.value).catch(error => {
                    console.error('Failed to change role selection:', error);
                  });
                }}
              >
                <option value="Visitor">Visitor</option>
                <option value="Leader">Leader</option>
                <option value="Manager">Manager</option>
                <option value="Admin">Admin</option>
              </select>
            </div>

            {this.requiresAssignments(formData.role) && (
              <div className={styles.formGroup}>
                <label>{this.getAssignmentLabel(formData.role)}</label>
                <div className={styles.assignmentBox}>
                  {isLoadingAssignments ? (
                    <Spinner size={SpinnerSize.small} label={`Loading ${this.getAssignmentLabel(formData.role).toLowerCase()}...`} />
                  ) : assignmentOptions.length === 0 ? (
                    <div className={styles.assignmentEmpty}>{this.getAssignmentPlaceholder(formData.role)}</div>
                  ) : (
                    assignmentOptions.map(option => (
                      <label key={option} className={styles.assignmentItem}>
                        <input
                          type="checkbox"
                          checked={formData.selectedAssignments.indexOf(option) >= 0}
                          onChange={() => this.toggleAssignment(option)}
                          disabled={isSaving}
                        />
                        <span>{option}</span>
                      </label>
                    ))
                  )}
                </div>
              </div>
            )}

            <div className={styles.formGroup}>
              <label>PIC</label>
              <PeoplePicker
                key={isNewMode ? 'new' : String(this.state.selectedRole?.Id)}
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

                    if (formData.role === 'Leader' || formData.role === 'Manager') {
                      this.loadAssignmentOptions(formData.role, items[0].secondaryText || '').catch(error => {
                        console.error('Failed to load assignment options for selected user:', error);
                      });
                    }
                  } else {
                    this.setState({
                      formData: { ...formData, picId: undefined, picDisplayName: '', picEmail: '', selectedAssignments: [] },
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
    return (
      <main className={styles.mainContainer}>
        <div className={styles.breadcrumb}>
          Home › <strong>Role Management</strong>
        </div>

        <div className={styles.header}>
          <h1>Role Management</h1>
          <button className={styles.btnNewRole} onClick={this.openNewPanel}>
            <Icon iconName="Add" />
            Assign User Role
          </button>
        </div>

        <div className={styles.content}>
          {this.renderGrid()}
        </div>

        {this.renderPanel()}
      </main>
    );
  }
}
