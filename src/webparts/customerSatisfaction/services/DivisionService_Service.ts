import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IDivisionService } from '../Models/DivisionServiceModel';

export class DivisionServiceService {
  private context: WebPartContext;
  private listName: string;

  constructor(context: WebPartContext, listName: string) {
    this.context = context;
    this.listName = listName;
  }

  private escapeODataValue(value: string): string {
    return value.replace(/'/g, "''");
  }

  /**
   * Updates a list item using the default JSON Light (minimal metadata, OData v4) format that
   * SPHttpClient.configurations.v1 sends by default. In this format, multi-value fields
   * (e.g. multi-user PIC/Manager columns) must be sent as a PLAIN ARRAY of ids — wrapping
   * them in `{ results: [...] }` (an odata=verbose construct) causes SharePoint to reject the
   * request with "StartObject node found... PrimitiveValue node was expected."
   */
  private async updateListItem(itemId: number, payload: { [key: string]: unknown }): Promise<boolean> {
    const endpoint =
      `${this.context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists/getbytitle('${this.listName}')/items(${itemId})`;

    const response: SPHttpClientResponse = await this.context.spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        body: JSON.stringify(payload),
        headers: {
          Accept: 'application/json',
          'Content-Type': 'application/json',
          'X-HTTP-Method': 'MERGE',
          'If-Match': '*',
        },
      }
    );

    return response.ok;
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
      console.error('DivisionServiceService getUserLoginNameById error:', error);
      return undefined;
    }
  }

  /**
   * Updates a multi-value Person/Group field (e.g. PIC, Manager) using SharePoint's
   * ValidateUpdateListItem REST method. Unlike a plain item MERGE/PATCH, this method
   * accepts field values as an opaque JSON-stringified array of `{Key: <login name>}`
   * objects, which avoids the typed OData multi-value array format
   * (`[...]` vs `{results:[...]}`) that has been rejected inconsistently by this
   * tenant with "StartArray/StartObject/PrimitiveValue expected" errors.
   */
  private async updateMultiUserField(
    itemId: number,
    fieldInternalName: string,
    userIds: number[]
  ): Promise<boolean> {
    try {
      const loginNames = (
        await Promise.all(userIds.map(id => this.getUserLoginNameById(id)))
      ).filter((loginName): loginName is string => !!loginName);

      const fieldValue = JSON.stringify(loginNames.map(loginName => ({ Key: loginName })));

      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items(${itemId})/validateupdatelistitem`;

      const body = {
        formValues: [
          {
            FieldName: fieldInternalName,
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
        console.error('updateMultiUserField request failed:', fieldInternalName, response.status);
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
        console.error('updateMultiUserField field-level error:', fieldInternalName, fieldError);
        return false;
      }

      return true;
    } catch (error) {
      console.error('DivisionServiceService updateMultiUserField error:', error);
      return false;
    }
  }

  public async getAllDivisions(): Promise<string[]> {
    try {
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items` +
        `?$select=Division&$orderby=Division asc`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        console.error('Failed to fetch divisions:', response.status);
        return [];
      }

      const data = await response.json();
      const rows = Array.isArray(data.value) ? (data.value as Partial<IDivisionService>[]) : [];
      const uniqueDivisions: string[] = [];

      rows.forEach(row => {
        const division = (row.Division || '').trim();
        if (division && uniqueDivisions.indexOf(division) === -1) {
          uniqueDivisions.push(division);
        }
      });

      return uniqueDivisions;
    } catch (error) {
      console.error('DivisionServiceService getAllDivisions error:', error);
      return [];
    }
  }

  public async getAllServices(): Promise<string[]> {
    try {
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items` +
        `?$select=Service&$orderby=Service asc`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        console.error('Failed to fetch services:', response.status);
        return [];
      }

      const data = await response.json();
      const rows = Array.isArray(data.value) ? (data.value as Partial<IDivisionService>[]) : [];
      const uniqueServices: string[] = [];

      rows.forEach(row => {
        const service = (row.Service || '').trim();
        if (service && uniqueServices.indexOf(service) === -1) {
          uniqueServices.push(service);
        }
      });

      return uniqueServices;
    } catch (error) {
      console.error('DivisionServiceService getAllServices error:', error);
      return [];
    }
  }

  public async getServicesByDivision(division: string): Promise<string[]> {
    if (!division) {
      return [];
    }

    try {
      const escapedDivision = this.escapeODataValue(division);
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items` +
        `?$select=Service,PICId,PIC/EMail,PIC/Title&$filter=Division eq '${escapedDivision}'&$orderby=Service asc&$expand=PIC`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        console.error('Failed to fetch services by division:', response.status);
        return [];
      }

      const data = await response.json();
      const rows = Array.isArray(data.value) ? (data.value as Partial<IDivisionService>[]) : [];
      const uniqueServices: string[] = [];

      rows.forEach(row => {
        const service = (row.Service || '').trim();
        if (service && uniqueServices.indexOf(service) === -1) {
          uniqueServices.push(service);
        }
      });

      return uniqueServices;
    } catch (error) {
      console.error('DivisionServiceService getServicesByDivision error:', error);
      return [];
    }
  }

  /**
   * Gets services (with Id, PIC and Manager) for a given division. The Id is needed to
   * subsequently look up the Service field's version history. PIC and Manager are
   * multi-user fields, so all assigned users are resolved and joined for display.
   */
  public async getServicesWithIdByDivision(division: string, year?: string): Promise<Array<{ Id: number; Service: string; PIC?: string; Manager?: string }>> {
    if (!division) {
      return [];
    }

    try {
      const escapedDivision = this.escapeODataValue(division);
      let filter = `Division eq '${escapedDivision}'`;
      if (year) {
        filter += ` and Year eq '${this.escapeODataValue(year)}'`;
      }
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items` +
        `?$select=Id,Service,PICId,PIC/Title,PIC/EMail,ManagerId,Manager/Title,Manager/EMail` +
        `&$filter=${filter}&$orderby=Service asc&$expand=PIC,Manager`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        console.error('Failed to fetch services with id by division:', response.status);
        return [];
      }

      const data = await response.json();
      const rows = Array.isArray(data.value)
        ? (data.value as Array<{
            Id: number;
            Service?: string;
            PICId?: number | number[];
            PIC?: { Title?: string; EMail?: string } | Array<{ Title?: string; EMail?: string }>;
            ManagerId?: number | number[];
            Manager?: { Title?: string; EMail?: string } | Array<{ Title?: string; EMail?: string }>;
          }>)
        : [];

      const results = await Promise.all(
        rows
          .filter(row => (row.Service || '').trim())
          .map(async row => {
            const picUsers = await this.resolveAllUserDetails(row.PIC, row.PICId);
            const managerUsers = await this.resolveAllUserDetails(row.Manager, row.ManagerId);

            return {
              Id: row.Id,
              Service: (row.Service || '').trim(),
              PIC: picUsers.map(u => u.Title).filter(Boolean).join(', ') || undefined,
              Manager: managerUsers.map(u => u.Title).filter(Boolean).join(', ') || undefined,
            };
          })
      );

      return results;
    } catch (error) {
      console.error('DivisionServiceService getServicesWithIdByDivision error:', error);
      return [];
    }
  }

  /** Normalizes a person field's id value to an array of user ids (handles both single and multi-value fields). */
  private allUserIds(value: number | number[] | undefined): number[] {
    if (Array.isArray(value)) {
      return value;
    }
    return value !== undefined && value !== null ? [value] : [];
  }

  /** Normalizes an expanded person field value to an array of user objects (handles both single and multi-value fields). */
  private allExpandedUsers(
    value: { Title?: string; EMail?: string } | Array<{ Title?: string; EMail?: string }> | undefined
  ): Array<{ Title?: string; EMail?: string }> {
    if (Array.isArray(value)) {
      return value;
    }
    return value ? [value] : [];
  }

  /**
   * Resolves ALL users assigned to a (possibly multi-value) person field, matching each
   * id to its expanded Title/EMail when available and falling back to a lookup by id
   * when the expansion is missing or incomplete.
   */
  private async resolveAllUserDetails(
    expanded: { Title?: string; EMail?: string } | Array<{ Title?: string; EMail?: string }> | undefined,
    ids: number | number[] | undefined
  ): Promise<Array<{ Id: number; Title: string; EMail: string }>> {
    const userIds = this.allUserIds(ids);
    const expandedUsers = this.allExpandedUsers(expanded);

    return Promise.all(
      userIds.map(async (id, index) => {
        let title = (expandedUsers[index]?.Title || '').trim();
        let email = (expandedUsers[index]?.EMail || '').trim();

        if (!title || !email) {
          const details = await this.getUserDetailsById(id);
          title = title || (details?.Title || '');
          email = email || (details?.EMail || '');
        }

        return { Id: id, Title: title, EMail: email };
      })
    );
  }

  /** Fallback lookup: resolves a user's Title and EMail by SharePoint user Id. */
  private async getUserDetailsById(userId: number): Promise<{ Title?: string; EMail?: string } | undefined> {
    try {
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/getuserbyid(${userId})?$select=Id,Title,Email`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        console.error('Failed to resolve user by id:', response.status);
        return undefined;
      }

      const user = await response.json() as { Title?: string; Email?: string };
      return { Title: (user.Title || '').trim(), EMail: (user.Email || '').trim() };
    } catch (error) {
      console.error('DivisionServiceService getUserDetailsById error:', error);
      return undefined;
    }
  }

  /** Gets the distinct set of Year values present in the Division_Service list. */
  public async getAllYears(): Promise<string[]> {
    try {
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items` +
        `?$select=Year&$orderby=Year desc`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        console.error('Failed to fetch years:', response.status);
        return [];
      }

      const data = await response.json();
      const rows = Array.isArray(data.value) ? (data.value as Array<{ Year?: string }>) : [];
      const uniqueYears: string[] = [];

      rows.forEach(row => {
        const year = (row.Year || '').toString().trim();
        if (year && uniqueYears.indexOf(year) === -1) {
          uniqueYears.push(year);
        }
      });

      return uniqueYears.sort((a, b) => Number(b) - Number(a));
    } catch (error) {
      console.error('DivisionServiceService getAllYears error:', error);
      return [];
    }
  }

  /** Gets all Division_Service items (optionally filtered by Year) for management purposes. PIC and Manager are multi-user fields, so all assigned users are returned. */
  public async getAllItemsWithDetails(year?: string): Promise<Array<{
    Id: number;
    Division: string;
    Service: string;
    Year: string;
    PICUsers: Array<{ Id: number; Title: string; EMail: string }>;
    ManagerUsers: Array<{ Id: number; Title: string; EMail: string }>;
  }>> {
    try {
      const filter = year ? `?$filter=Year eq '${this.escapeODataValue(year)}'&` : '?';
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items` +
        `${filter}$select=Id,Division,Service,Year,PICId,PIC/Title,PIC/EMail,ManagerId,Manager/Title,Manager/EMail&$expand=PIC,Manager&$orderby=Division asc,Service asc`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        console.error('Failed to fetch Division_Service items:', response.status);
        return [];
      }

      const data = await response.json();
      const rows = Array.isArray(data.value)
        ? (data.value as Array<{
            Id: number;
            Division?: string;
            Service?: string;
            Year?: string;
            PICId?: number | number[];
            PIC?: { Title?: string; EMail?: string } | Array<{ Title?: string; EMail?: string }>;
            ManagerId?: number | number[];
            Manager?: { Title?: string; EMail?: string } | Array<{ Title?: string; EMail?: string }>;
          }>)
        : [];

      const results = await Promise.all(rows.map(async row => {
        const picUsers = await this.resolveAllUserDetails(row.PIC, row.PICId);
        const managerUsers = await this.resolveAllUserDetails(row.Manager, row.ManagerId);

        return {
          Id: row.Id,
          Division: row.Division || '',
          Service: row.Service || '',
          Year: row.Year || '',
          PICUsers: picUsers,
          ManagerUsers: managerUsers,
        };
      }));

      return results;
    } catch (error) {
      console.error('DivisionServiceService getAllItemsWithDetails error:', error);
      return [];
    }
  }

  /** Updates Division, Service, Year and/or PIC/Manager (multi-user) for a Division_Service item. */
  public async updateItemDetails(
    itemId: number,
    payload: { division?: string; service?: string; year?: string; picIds?: number[]; managerIds?: number[] }
  ): Promise<boolean> {
    const updatePayload: { [key: string]: unknown } = {};
    if (payload.division !== undefined) {
      updatePayload.Division = payload.division;
    }
    if (payload.service !== undefined) {
      updatePayload.Service = payload.service;
    }
    if (payload.year !== undefined) {
      updatePayload.Year = payload.year;
    }

    if (Object.keys(updatePayload).length > 0) {
      const updated = await this.updateListItem(itemId, updatePayload);
      if (!updated) {
        return false;
      }
    }

    if (payload.picIds !== undefined) {
      const updated = await this.updateMultiUserField(itemId, 'PIC', payload.picIds);
      if (!updated) {
        return false;
      }
    }

    if (payload.managerIds !== undefined) {
      const updated = await this.updateMultiUserField(itemId, 'Manager', payload.managerIds);
      if (!updated) {
        return false;
      }
    }

    return true;
  }

  /**
   * Gets the distinct set of Service names (including the current value) this list
   * item has ever had, based on the list's version history. Used so that when a
   * Service is renamed, results/action plans filtered by the old and new names can
   * both be surfaced.
   */
  public async getServiceNameHistory(itemId: number): Promise<string[]> {
    if (!itemId) {
      return [];
    }

    try {
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items(${itemId})/versions` +
        `?$select=Service`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        console.error('Failed to fetch service version history:', response.status);
        return [];
      }

      const data = await response.json();
      const rows = Array.isArray(data.value) ? (data.value as Array<{ Service?: string }>) : [];
      const names: string[] = [];

      rows.forEach(row => {
        const name = (row.Service || '').trim();
        if (name && names.indexOf(name) === -1) {
          names.push(name);
        }
      });

      return names;
    } catch (error) {
      console.error('DivisionServiceService getServiceNameHistory error:', error);
      return [];
    }
  }

  /**
   * Gets all service names from Division_Service where the user is the PIC.
   */
  public async getServicesByPIC(userEmail: string): Promise<string[]> {
    if (!userEmail) {
      return [];
    }

    try {
      const escaped = this.escapeODataValue(userEmail);
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items` +
        `?$select=Service,PIC/EMail&$expand=PIC` +
        `&$filter=PIC/EMail eq '${escaped}'&$orderby=Service asc`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        console.error('Failed to fetch services by PIC:', response.status);
        return [];
      }

      const data = await response.json();
      const rows = Array.isArray(data.value) ? data.value : [];
      const uniqueServices: string[] = [];

      rows.forEach((row: { Service?: string }) => {
        const service = (row.Service || '').trim();
        if (service && uniqueServices.indexOf(service) === -1) {
          uniqueServices.push(service);
        }
      });

      return uniqueServices;
    } catch (error) {
      console.error('DivisionServiceService getServicesByPIC error:', error);
      return [];
    }
  }

  /**
   * Gets all service names from Division_Service where the user is the Manager.
   */
  public async getServicesByManager(userEmail: string): Promise<string[]> {
    if (!userEmail) {
      return [];
    }

    try {
      const escaped = this.escapeODataValue(userEmail);
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items` +
        `?$select=Service,Manager/EMail&$expand=Manager` +
        `&$filter=Manager/EMail eq '${escaped}'&$orderby=Service asc`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        console.error('Failed to fetch services by Manager:', response.status);
        return [];
      }

      const data = await response.json();
      const rows = Array.isArray(data.value) ? data.value : [];
      const uniqueServices: string[] = [];

      rows.forEach((row: { Service?: string }) => {
        const service = (row.Service || '').trim();
        if (service && uniqueServices.indexOf(service) === -1) {
          uniqueServices.push(service);
        }
      });

      return uniqueServices;
    } catch (error) {
      console.error('DivisionServiceService getServicesByManager error:', error);
      return [];
    }
  }

  public async getDivisionsByManager(userEmail: string): Promise<string[]> {
    if (!userEmail) {
      return [];
    }

    try {
      const escaped = this.escapeODataValue(userEmail);
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items` +
        `?$select=Division,Manager/EMail&$expand=Manager` +
        `&$filter=Manager/EMail eq '${escaped}'&$orderby=Division asc`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        console.error('Failed to fetch divisions by Manager:', response.status);
        return [];
      }

      const data = await response.json();
      const rows = Array.isArray(data.value) ? (data.value as Partial<IDivisionService>[]) : [];
      const uniqueDivisions: string[] = [];

      rows.forEach(row => {
        const division = (row.Division || '').trim();
        if (division && uniqueDivisions.indexOf(division) === -1) {
          uniqueDivisions.push(division);
        }
      });

      return uniqueDivisions;
    } catch (error) {
      console.error('DivisionServiceService getDivisionsByManager error:', error);
      return [];
    }
  }

  public async getDivisionByService(service: string): Promise<string | undefined> {
    if (!service) {
      return undefined;
    }

    try {
      const escapedService = this.escapeODataValue(service);
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items` +
        `?$select=Division&$filter=Service eq '${escapedService}'&$top=1`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        console.error('Failed to fetch division by service:', response.status);
        return undefined;
      }

      const data = await response.json();
      const rows = Array.isArray(data.value) ? (data.value as Partial<IDivisionService>[]) : [];
      if (rows.length > 0) {
        return rows[0].Division;
      }

      return undefined;
    } catch (error) {
      console.error('DivisionServiceService getDivisionByService error:', error);
      return undefined;
    }
  }

  public async clearAssignmentsForUser(userId: number): Promise<boolean> {
    try {
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items` +
        `?$select=Id,PICId,ManagerId`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        console.error('Failed to fetch assignments for clearing:', response.status);
        return false;
      }

      const data = await response.json();
      const rows = Array.isArray(data.value)
        ? (data.value as Array<{ Id: number; PICId?: number | number[]; ManagerId?: number | number[] }>)
        : [];

      for (const row of rows) {
        const picIds = this.allUserIds(row.PICId);
        const managerIds = this.allUserIds(row.ManagerId);

        if (picIds.indexOf(userId) !== -1) {
          const updated = await this.updateMultiUserField(
            row.Id,
            'PIC',
            picIds.filter(id => id !== userId)
          );
          if (!updated) {
            return false;
          }
        }

        if (managerIds.indexOf(userId) !== -1) {
          const updated = await this.updateMultiUserField(
            row.Id,
            'Manager',
            managerIds.filter(id => id !== userId)
          );
          if (!updated) {
            return false;
          }
        }
      }

      return true;
    } catch (error) {
      console.error('DivisionServiceService clearAssignmentsForUser error:', error);
      return false;
    }
  }

  /**
   * Syncs the given user's Manager (multi-user field) assignment across ALL Division_Service
   * items: the user is added to the Manager column of items whose Division is in
   * `selectedDivisions` (if not already present), and removed from the Manager column of
   * items whose Division is NOT in `selectedDivisions` (if currently present). Other users
   * already in PIC/Manager on those items are always preserved.
   */
  public async assignManagerToDivisions(userId: number, selectedDivisions: string[]): Promise<boolean> {
    try {
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items` +
        `?$select=Id,Division,ManagerId`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        console.error('Failed to fetch items for manager assignment sync:', response.status);
        return false;
      }

      const data = await response.json();
      const rows = Array.isArray(data.value)
        ? (data.value as Array<{ Id: number; Division?: string; ManagerId?: number | number[] }>)
        : [];

      for (const row of rows) {
        const existingIds = this.allUserIds(row.ManagerId);
        const hasUser = existingIds.indexOf(userId) !== -1;
        const isSelected = selectedDivisions.indexOf(row.Division || '') !== -1;

        let newIds: number[] | undefined;
        if (isSelected && !hasUser) {
          newIds = [...existingIds, userId];
        } else if (!isSelected && hasUser) {
          newIds = existingIds.filter(id => id !== userId);
        }

        if (newIds) {
          const updated = await this.updateMultiUserField(row.Id, 'Manager', newIds);
          if (!updated) {
            return false;
          }
        }
      }

      return true;
    } catch (error) {
      console.error('DivisionServiceService assignManagerToDivisions error:', error);
      return false;
    }
  }

  /**
   * Syncs the given user's PIC (multi-user field) assignment across ALL Division_Service
   * items: the user is added to the PIC column of items whose Service is in
   * `selectedServices` (if not already present), and removed from the PIC column of items
   * whose Service is NOT in `selectedServices` (if currently present). Other users already
   * in PIC/Manager on those items are always preserved.
   */
  public async assignLeaderToServices(userId: number, selectedServices: string[]): Promise<boolean> {
    try {
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items` +
        `?$select=Id,Service,PICId`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        console.error('Failed to fetch items for leader assignment sync:', response.status);
        return false;
      }

      const data = await response.json();
      const rows = Array.isArray(data.value)
        ? (data.value as Array<{ Id: number; Service?: string; PICId?: number | number[] }>)
        : [];

      for (const row of rows) {
        const existingIds = this.allUserIds(row.PICId);
        const hasUser = existingIds.indexOf(userId) !== -1;
        const isSelected = selectedServices.indexOf(row.Service || '') !== -1;

        let newIds: number[] | undefined;
        if (isSelected && !hasUser) {
          newIds = [...existingIds, userId];
        } else if (!isSelected && hasUser) {
          newIds = existingIds.filter(id => id !== userId);
        }

        if (newIds) {
          const updated = await this.updateMultiUserField(row.Id, 'PIC', newIds);
          if (!updated) {
            return false;
          }
        }
      }

      return true;
    } catch (error) {
      console.error('DivisionServiceService assignLeaderToServices error:', error);
      return false;
    }
  }
}
