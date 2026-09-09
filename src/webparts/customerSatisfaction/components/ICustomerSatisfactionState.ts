export type ViewName =
  'home'
  | 'about'
  | 'actionplan'
  | 'dashboard'
  | 'admin'
  | 'divisionServiceManagement'
  | 'esvn'
  | 'iss'
  | 'is'
  | 'ss'
  | 'dts'
  | 'serviceResult'
  | 'dtsComponentManufacturing'
  | 'dtsEngineeringTechnology'
  | 'dtsEnterpriseApplications'
  | 'dtsESTechnology'
  | 'dtsResidential'
  | 'guide';

export interface ICustomerSatisfactionState {
  isTeamLeader: boolean;
  userRole: 'visitor' | 'leader' | 'manager' | 'admin'; // visitor, leader, manager, or admin
  isLoading: boolean;
  sentToday: number;
  receivedToday: number;
  isSidebarCollapsed: boolean;
  currentView: ViewName;
  userService: string;
  selectedResultDepartment?: string;
  selectedResultService?: string | string[];
  selectedResultLabel?: string;
}
