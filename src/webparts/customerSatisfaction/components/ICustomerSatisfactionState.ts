export type ViewName =
  'home'
  | 'about'
  | 'actionplan'
  | 'dashboard'
  | 'admin'
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
  | 'dtsResidential';

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
  selectedResultService?: string;
  selectedResultLabel?: string;
}