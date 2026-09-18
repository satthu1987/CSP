import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ICSPUserRole } from '../Models/CSPUserRole';

export class CSPUserRole_Service {
  private context: WebPartContext;
  private listName: string = 'CSP_UserRole';

  constructor(context: WebPartContext) {
    this.context = context;
  }

  /**
   * Gets the user role from CSP_UserRole list.
   * Returns the user's role record if found, undefined if not found.
   */
  public async getUserRole(userEmail: string): Promise<ICSPUserRole | undefined> {
    try {
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items` +
        `?$select=Id,Title,PIC/Id,PIC/EMail,PIC/Title,Role&$expand=PIC` +
        `&$filter=PIC/EMail eq '${userEmail}'`;

      console.log('CSPUserRole_Service getUserRole - fetching for email:', userEmail);

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        console.error('CSPUserRole_Service getUserRole - response not ok:', response.status);
        return undefined;
      }

      const data = await response.json();
      console.log('CSPUserRole_Service getUserRole - response data:', data);

      if (Array.isArray(data.value) && data.value.length > 0) {
        const userRole = data.value[0] as ICSPUserRole;
        console.log('CSPUserRole_Service getUserRole - found user role:', userRole);
        return userRole;
      }

      console.log('CSPUserRole_Service getUserRole - no user role found for email:', userEmail);
      return undefined;
    } catch (error) {
      console.error('CSPUserRole_Service getUserRole error:', error);
      return undefined;
    }
  }

  /**
   * Gets the user's role type (visitor, leader, manager, admin).
   * Returns the role string if found, 'visitor' as default.
   */
  public async getUserRoleType(userEmail: string): Promise<'visitor' | 'leader' | 'manager' | 'admin'> {
    console.log('CSPUserRole_Service getUserRoleType - checking role for email:', userEmail);

    const userRole = await this.getUserRole(userEmail);
    console.log('CSPUserRole_Service getUserRoleType - userRole object:', userRole);

    if (!userRole) {
      console.log('CSPUserRole_Service getUserRoleType - no user role found, returning visitor');
      return 'visitor';
    }

    const role = (userRole.Role || '').trim().toLowerCase();
    console.log('CSPUserRole_Service getUserRoleType - role value (trimmed, lowercase):', role);

    if (role === 'leader' || role === 'admin' || role === 'manager') {
      console.log('CSPUserRole_Service getUserRoleType - returning role:', role);
      return role as 'leader' | 'manager' | 'admin';
    }

    console.log('CSPUserRole_Service getUserRoleType - role not recognized, returning visitor');
    return 'visitor';
  }

  /**
   * Gets all items from CSP_UserRole list.
   */
  public async getAllRoles(): Promise<ICSPUserRole[]> {
    try {
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items` +
        `?$select=Id,Title,PIC/Id,PIC/EMail,PIC/Title,Role&$expand=PIC&$orderby=Id asc`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        console.error('Failed to fetch all roles');
        return [];
      }

      const data = await response.json();
      return Array.isArray(data.value) ? (data.value as ICSPUserRole[]) : [];
    } catch (error) {
      console.error('CSPUserRole_Service getAllRoles error:', error);
      return [];
    }
  }

  private async getUserLoginNameById(userId: number): Promise<string | undefined> {
    try {
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/getuserbyid(${userId})?$select=LoginName`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        console.error('getUserLoginNameById failed for user', userId, response.status);
        return undefined;
      }

      const data = await response.json();
      return data.LoginName as string;
    } catch (error) {
      console.error('CSPUserRole_Service getUserLoginNameById error:', error);
      return undefined;
    }
  }

  /**
   * Sets the PIC multi-user field on a CSP_UserRole item using ValidateUpdateListItem,
   * which accepts field values as an opaque JSON-stringified array of
   * `{Key: <login name>}` objects instead of a typed OData multi-value array. This avoids
   * the "StartArray/StartObject/PrimitiveValue expected" errors seen when sending
   * `PICId` as part of a plain item POST/MERGE body.
   */
  private async setPicField(itemId: number, picId: number): Promise<boolean> {
    try {
      const loginName = await this.getUserLoginNameById(picId);
      if (!loginName) {
        console.error('setPicField: could not resolve login name for user', picId);
        return false;
      }

      const fieldValue = JSON.stringify([{ Key: loginName }]);

      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items(${itemId})/validateupdatelistitem`;

      const body = {
        formValues: [
          {
            FieldName: 'PIC',
            FieldValue: fieldValue,
          },
        ],
        bNewDocumentUpdate: false,
      };

      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          body: JSON.stringify(body),
          headers: {
            Accept: 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata',
          },
        }
      );

      if (!response.ok) {
        console.error('setPicField request failed:', response.status);
        return false;
      }

      const data = await response.json();
      const results: Array<{ ErrorMessage?: string; HasException?: boolean }> = Array.isArray(data)
        ? data
        : Array.isArray(data.value)
        ? data.value
        : [];
      const fieldError = results.filter(r => r.HasException || r.ErrorMessage);
      if (fieldError.length > 0) {
        console.error('setPicField field-level error:', fieldError);
        return false;
      }

      return true;
    } catch (error) {
      console.error('CSPUserRole_Service setPicField error:', error);
      return false;
    }
  }

  /**
   * Creates a new role item in CSP_UserRole list.
   */
  public async createRole(picId: number, role: string): Promise<boolean> {
    try {
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items`;

      const body = JSON.stringify({
        Role: role,
      });

      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          body,
          headers: {
            Accept: 'application/json',
            'Content-Type': 'application/json',
          },
        }
      );

      if (!response.ok) {
        console.error('CSPUserRole_Service createRole - item creation failed:', response.status);
        return false;
      }

      const data = await response.json();
      const newItemId = data.Id as number;
      if (!newItemId) {
        console.error('CSPUserRole_Service createRole - no Id in created item response:', data);
        return false;
      }

      return this.setPicField(newItemId, picId);
    } catch (error) {
      console.error('CSPUserRole_Service createRole error:', error);
      return false;
    }
  }

  /**
   * Updates an existing role item in CSP_UserRole list.
   */
  public async updateRole(id: number, picId: number, role: string): Promise<boolean> {
    try {
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items(${id})`;

      const body = JSON.stringify({
        Role: role,
      });

      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          body,
          headers: {
            Accept: 'application/json',
            'Content-Type': 'application/json',
            'X-HTTP-Method': 'MERGE',
            'If-Match': '*',
          },
        }
      );

      if (!response.ok) {
        console.error('CSPUserRole_Service updateRole - item update failed:', response.status);
        return false;
      }

      return this.setPicField(id, picId);
    } catch (error) {
      console.error('CSPUserRole_Service updateRole error:', error);
      return false;
    }
  }

  /**
   * Deletes a role item from CSP_UserRole list.
   */
  public async deleteRole(id: number): Promise<boolean> {
    try {
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items(${id})`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'X-HTTP-Method': 'DELETE',
            'If-Match': '*',
          },
        }
      );

      return response.ok;
    } catch (error) {
      console.error('CSPUserRole_Service deleteRole error:', error);
      return false;
    }
  }
}
