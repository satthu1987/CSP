import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export class TeamLeaderService {
  private context: WebPartContext;
  private listName: string;

  constructor(context: WebPartContext, listName: string) {
    this.context = context;
    this.listName = listName;
  }

  /**
   * Checks if the current user exists in the TeamLeader SharePoint list.
   * Assumes the list has a "User" (Person) column or "Email" text column.
   */
  public async isCurrentUserTeamLeader(userEmail: string): Promise<boolean> {
    try {
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items` +
        `?$select=Id,User/EMail,User/Title&$expand=User` +
        `&$filter=User/EMail eq '${userEmail}'`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        // Fallback: try a plain Email column
        return await this.fallbackEmailColumnCheck(userEmail);
      }

      const data = await response.json();
      return Array.isArray(data.value) && data.value.length > 0;
    } catch (error) {
      console.error('TeamLeaderService error:', error);
      return false;
    }
  }

  private async fallbackEmailColumnCheck(userEmail: string): Promise<boolean> {
    try {
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items` +
        `?$select=Id,Email&$filter=Email eq '${userEmail}'`;

      const response = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) return false;

      const data = await response.json();
      return Array.isArray(data.value) && data.value.length > 0;
    } catch {
      return false;
    }
  }
}