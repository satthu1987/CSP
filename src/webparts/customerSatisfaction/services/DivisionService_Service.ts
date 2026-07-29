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

  public async getServicesByDivision(division: string): Promise<string[]> {
    if (!division) {
      return [];
    }

    try {
      const escapedDivision = division.replace(/'/g, "''");
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
   * Gets all service names from Division_Service where the user is the PIC.
   */
  public async getServicesByPIC(userEmail: string): Promise<string[]> {
    if (!userEmail) {
      return [];
    }

    try {
      const escaped = userEmail.replace(/'/g, "''");
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
      const escaped = userEmail.replace(/'/g, "''");
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

  public async getDivisionByService(service: string): Promise<string | undefined> {
    if (!service) {
      return undefined;
    }

    try {
      const escapedService = service.replace(/'/g, "''");
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
}
