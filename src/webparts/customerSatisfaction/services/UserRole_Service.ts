import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IUserRole } from '../Models/UserRole';

export class UserRoleService {
  private context: WebPartContext;
  private listName: string;

  constructor(context: WebPartContext, listName: string) {
    this.context = context;
    this.listName = listName;
  }

  /**
   * Gets the user role information from the UserRole SharePoint list.
   * Returns the user role if found, undefined if not found.
   */
  public async getUserRole(userEmail: string): Promise<IUserRole | undefined> {
    try {
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items` +
        `?$select=Id,Title,PIC/EMail,PIC/Title&$expand=PIC` +
        `&$filter=PIC/EMail eq '${userEmail}'`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        return undefined;
      }

      const data = await response.json();
      if (Array.isArray(data.value) && data.value.length > 0) {
        return data.value[0] as IUserRole;
      }

      return undefined;
    } catch (error) {
      console.error('UserRoleService error:', error);
      return undefined;
    }
  }

  /**
   * Gets all services for a user from the RoleInService list.
   * Returns an array of service titles (strings) that the user has roles in.
   */
  public async getUserServices(userEmail: string): Promise<string[]> {
    try {
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items` +
        `?$select=Title,PIC/EMail&$expand=PIC` +
        `&$filter=PIC/EMail eq '${userEmail}'`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        console.error('Failed to fetch user services');
        return [];
      }

      const data = await response.json();
      if (Array.isArray(data.value)) {
        return data.value.map((role: IUserRole) => role.Title).filter((title: string) => title);
      }

      return [];
    } catch (error) {
      console.error('UserRoleService getUserServices error:', error);
      return [];
    }
  }

  /**
   * Checks if the current user exists in the UserRole SharePoint list.
   */
  public async isUserInRole(userEmail: string): Promise<boolean> {
    const userRole = await this.getUserRole(userEmail);
    return userRole !== undefined;
  }
}
