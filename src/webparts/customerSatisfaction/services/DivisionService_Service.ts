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

  private async updateListItem(itemId: number, payload: Partial<{ PICId: number | null; ManagerId: number | null; Division: string; Service: string; Year: string }>): Promise<boolean> {
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
   * Gets services (with Id and PIC) for a given division. The Id is needed to
   * subsequently look up the Service field's version history.
   */
  public async getServicesWithIdByDivision(division: string, year?: string): Promise<Array<{ Id: number; Service: string; PIC?: string }>> {
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
        `?$select=Id,Service,PICId,PIC/Title&$filter=${filter}&$orderby=Service asc&$expand=PIC`;

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
        ? (data.value as Array<{ Id: number; Service?: string; PICId?: number; PIC?: { Title?: string } | Array<{ Title?: string }> }>)
        : [];

      const results = await Promise.all(
        rows
          .filter(row => (row.Service || '').trim())
          .map(async row => {
            const picValue = Array.isArray(row.PIC) ? row.PIC[0] : row.PIC;
            let pic = picValue?.Title;

            if (!pic && row.PICId) {
              pic = await this.getUserTitleById(row.PICId);
            }

            return {
              Id: row.Id,
              Service: (row.Service || '').trim(),
              PIC: pic,
            };
          })
      );

      return results;
    } catch (error) {
      console.error('DivisionServiceService getServicesWithIdByDivision error:', error);
      return [];
    }
  }

  /** Fallback lookup: resolves a user's display name by SharePoint user Id when $expand doesn't yield one. */
  private async getUserTitleById(userId: number): Promise<string | undefined> {
    try {
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/getuserbyid(${userId})`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        return undefined;
      }

      const data = await response.json();
      return data?.Title;
    } catch (error) {
      console.error('DivisionServiceService getUserTitleById error:', error);
      return undefined;
    }
  }

  /**
   * Expanded person fields can come back as an object, an array (multi-person
   * field) or be missing even though the id is set; this normalizes the value
   * to a single user object.
   */
  private normalizeExpandedUser(
    value: { Title?: string; EMail?: string } | Array<{ Title?: string; EMail?: string }> | undefined
  ): { Title?: string; EMail?: string } | undefined {
    if (Array.isArray(value)) {
      return value.length > 0 ? value[0] : undefined;
    }
    return value;
  }

  /** Returns the first user id when the person field id comes back as an array (multi-person field). */
  private firstUserId(value: number | number[] | undefined): number | undefined {
    if (Array.isArray(value)) {
      return value.length > 0 ? value[0] : undefined;
    }
    return value;
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

  /**
   * Resolves a person field's Title/EMail from the expanded value, falling back
   * to a lookup by id when the expansion is missing or incomplete.
   */
  private async resolveUserDetails(
    expanded: { Title?: string; EMail?: string } | Array<{ Title?: string; EMail?: string }> | undefined,
    userId: number | number[] | undefined
  ): Promise<{ Title?: string; EMail?: string } | undefined> {
    const user = this.normalizeExpandedUser(expanded);
    const fallbackId = this.firstUserId(userId);

    let title = (user?.Title || '').trim();
    let email = (user?.EMail || '').trim();

    if ((!title || !email) && fallbackId !== undefined) {
      const details = await this.getUserDetailsById(fallbackId);
      title = title || (details?.Title || '');
      email = email || (details?.EMail || '');
    }

    return title || email ? { Title: title, EMail: email } : undefined;
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

  /** Gets all Division_Service items (optionally filtered by Year) for management purposes. */
  public async getAllItemsWithDetails(year?: string): Promise<Array<{
    Id: number;
    Division: string;
    Service: string;
    Year: string;
    PICId?: number;
    PICTitle?: string;
    PICEmail?: string;
    ManagerId?: number;
    ManagerTitle?: string;
    ManagerEmail?: string;
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
        const pic = await this.resolveUserDetails(row.PIC, row.PICId);
        const manager = await this.resolveUserDetails(row.Manager, row.ManagerId);

        return {
          Id: row.Id,
          Division: row.Division || '',
          Service: row.Service || '',
          Year: row.Year || '',
          PICId: this.firstUserId(row.PICId),
          PICTitle: pic?.Title,
          PICEmail: pic?.EMail,
          ManagerId: this.firstUserId(row.ManagerId),
          ManagerTitle: manager?.Title,
          ManagerEmail: manager?.EMail,
        };
      }));

      return results;
    } catch (error) {
      console.error('DivisionServiceService getAllItemsWithDetails error:', error);
      return [];
    }
  }

  /** Updates Division, Service, Year and/or PIC for a Division_Service item. */
  public async updateItemDetails(
    itemId: number,
    payload: { division?: string; service?: string; year?: string; picId?: number; managerId?: number }
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
    if (payload.picId !== undefined) {
      updatePayload.PICId = payload.picId;
    }
    if (payload.managerId !== undefined) {
      updatePayload.ManagerId = payload.managerId;
    }
    return this.updateListItem(itemId, updatePayload as Partial<{ Division: string; Service: string; Year: string; PICId: number; ManagerId: number }>);
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
        `?$select=Id,PICId,ManagerId&$filter=PICId eq ${userId} or ManagerId eq ${userId}`;

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
        ? (data.value as Array<{ Id: number; PICId?: number; ManagerId?: number }>)
        : [];

      for (const row of rows) {
        const payload: { PICId?: null; ManagerId?: null } = {};

        if (row.PICId === userId) {
          payload.PICId = null;
        }
        if (row.ManagerId === userId) {
          payload.ManagerId = null;
        }

        if (Object.keys(payload).length > 0) {
          const updated = await this.updateListItem(row.Id, payload);
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

  public async assignManagerToDivisions(userId: number, divisions: string[]): Promise<boolean> {
    try {
      for (const division of divisions) {
        const escapedDivision = this.escapeODataValue(division);
        const endpoint =
          `${this.context.pageContext.web.absoluteUrl}` +
          `/_api/web/lists/getbytitle('${this.listName}')/items` +
          `?$select=Id&$filter=Division eq '${escapedDivision}'`;

        const response: SPHttpClientResponse = await this.context.spHttpClient.get(
          endpoint,
          SPHttpClient.configurations.v1
        );

        if (!response.ok) {
          console.error('Failed to fetch items for division assignment:', response.status, division);
          return false;
        }

        const data = await response.json();
        const rows = Array.isArray(data.value) ? (data.value as Array<{ Id: number }>) : [];

        for (const row of rows) {
          const updated = await this.updateListItem(row.Id, { ManagerId: userId });
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

  public async assignLeaderToServices(userId: number, services: string[]): Promise<boolean> {
    try {
      for (const service of services) {
        const escapedService = this.escapeODataValue(service);
        const endpoint =
          `${this.context.pageContext.web.absoluteUrl}` +
          `/_api/web/lists/getbytitle('${this.listName}')/items` +
          `?$select=Id&$filter=Service eq '${escapedService}'`;

        const response: SPHttpClientResponse = await this.context.spHttpClient.get(
          endpoint,
          SPHttpClient.configurations.v1
        );

        if (!response.ok) {
          console.error('Failed to fetch items for service assignment:', response.status, service);
          return false;
        }

        const data = await response.json();
        const rows = Array.isArray(data.value) ? (data.value as Array<{ Id: number }>) : [];

        for (const row of rows) {
          const updated = await this.updateListItem(row.Id, { PICId: userId });
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
