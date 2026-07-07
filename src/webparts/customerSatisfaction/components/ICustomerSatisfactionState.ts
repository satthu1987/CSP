export type ViewName = 'home' | 'about' | 'actionplan' | 'dashboard' | 'admin';

export interface ICustomerSatisfactionState {
  isTeamLeader: boolean;
  isLoading: boolean;
  sentToday: number;
  receivedToday: number;
  isSidebarCollapsed: boolean;
  currentView: ViewName;
  userService: string;
}