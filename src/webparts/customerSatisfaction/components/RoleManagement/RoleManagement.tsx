import * as React from 'react';
import styles from './RoleManagement.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Icon, IPersonaProps, Modal, Spinner, SpinnerSize } from '@fluentui/react';
import { IPeoplePickerContext, PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { UserRoleService } from '../../services/UserRole_Service';
import { IUserRole } from '../../Models/UserRole';

export interface IRoleManagementProps {
  context: WebPartContext;
}

interface IFormData {
  title: string;
  picId?: number;
  picDisplayName?: string;
  picEmail?: string;
}

interface IRoleManagementState {
  roles: IUserRole[];
  isLoading: boolean;
  isPanelOpen: boolean;
  isNewMode: boolean;
  selectedRole?: IUserRole;
  formData: IFormData;
  isSaving: boolean;
  errorMsg: string;
}

export default class RoleManagement extends React.Component<IRoleManagementProps, IRoleManagementState> {
  private roleService: UserRoleService;

  constructor(props: IRoleManagementProps) {
    super(props);
    this.state = {
      roles: [],
      isLoading: true,
      isPanelOpen: false,
      isNewMode: false,
      selectedRole: undefined,
      formData: { title: '' },
      isSaving: false,
      errorMsg: '',
    };
    this.roleService = new UserRoleService(props.context, 'RoleInService');
  }

  public async componentDidMount(): Promise<void> {
    await this.loadRoles();
  }

  private loadRoles = async (): Promise<void> => {
    this.setState({ isLoading: true });
    try {
      const roles = await this.roleService.getAllRoles();
      this.setState({ roles, isLoading: false });
    } catch (error) {
      console.error('Error loading roles:', error);
      this.setState({ isLoading: false });
    }
  };

  private openNewPanel = (): void => {
    this.setState({
      isPanelOpen: true,
      isNewMode: true,
      selectedRole: undefined,
      formData: { title: '' },
      errorMsg: '',
    });
  };

  private openEditPanel = (role: IUserRole): void => {
    this.setState({
      isPanelOpen: true,
      isNewMode: false,
      selectedRole: role,
      formData: {
        title: role.Title,
        picDisplayName: role.PIC?.Title,
        picEmail: role.PIC?.EMail,
      },
      errorMsg: '',
    });
  };

  private closePanel = (): void => {
    this.setState({ isPanelOpen: false, selectedRole: undefined, errorMsg: '' });
  };

  private handleSave = async (): Promise<void> => {
    const { formData, isNewMode, selectedRole } = this.state;

    if (!formData.title.trim()) {
      this.setState({ errorMsg: 'Role name is required.' });
      return;
    }
    if (!formData.picId) {
      this.setState({ errorMsg: 'PIC is required.' });
      return;
    }

    this.setState({ isSaving: true, errorMsg: '' });

    let success = false;
    if (isNewMode) {
      success = await this.roleService.createRole(formData.title.trim(), formData.picId);
    } else if (selectedRole) {
      success = await this.roleService.updateRole(selectedRole.Id, formData.title.trim(), formData.picId);
    }

    if (success) {
      this.setState({ isSaving: false, isPanelOpen: false });
      await this.loadRoles();
    } else {
      this.setState({ isSaving: false, errorMsg: 'Failed to save. Please try again.' });
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
          <div className={styles.colRole}>Role</div>
          <div className={styles.colPIC}>PIC</div>
          <div className={styles.colAction}>Action</div>
        </div>
        {roles.map((role, index) => (
          <div key={role.Id} className={styles.gridRow}>
            <div className={styles.colNo}>{index + 1}</div>
            <div className={styles.colRole}>{role.Title}</div>
            <div className={styles.colPIC}>{role.PIC?.Title || '—'}</div>
            <div className={styles.colAction}>
              <button
                className={styles.editIcon}
                title="Edit"
                onClick={() => this.openEditPanel(role)}
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
    const { isPanelOpen, isNewMode, formData, isSaving, errorMsg } = this.state;

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
            <h2>{isNewMode ? 'New Role' : 'Edit Role'}</h2>
            <button className={styles.closeBtn} onClick={this.closePanel}>
              <Icon iconName="Cancel" />
            </button>
          </div>

          <div className={styles.panelBody}>
            <div className={styles.formGroup}>
              <label>Role</label>
              <input
                type="text"
                value={formData.title}
                onChange={(e) => this.setState({ formData: { ...formData, title: e.target.value } })}
                placeholder="Enter role name"
              />
            </div>

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
                  } else {
                    this.setState({
                      formData: { ...formData, picId: undefined, picDisplayName: '', picEmail: '' },
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
            New Role
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
