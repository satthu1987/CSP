import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ICustomerSatisfactionProps {
  context: WebPartContext;
  currentUserEmail: string;
  currentUserLoginName: string;
  currentUserDisplayName: string;
  teamLeaderListName: string;
}