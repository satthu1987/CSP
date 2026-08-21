import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IAcDataItem {
  Id: number;
  Year: string;
  Department: string;
  Data: string;
}

export class AcDataService {
  private context: WebPartContext;
  private listName: string = 'AC_Data';

  constructor(context: WebPartContext) {
    this.context = context;
  }

  /** Fetch Year + Data (text column) from AC_Data list */
  public async getAcData(): Promise<IAcDataItem[]> {
    const webUrl = this.context.pageContext.web.absoluteUrl;

    const endpoint =
      `${webUrl}/_api/web/lists/getbytitle('${this.listName}')/items` +
      `?$select=Id,Year,Department,Data&$orderby=Year desc`;

    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      console.error('Failed to fetch AC_Data list', await response.text());
      return [];
    }

    const data = await response.json();
    const items = data.value || [];

    const results: IAcDataItem[] = items.map((item: any) => ({
      Id: item.Id,
      Year: item.Year || '',
      Department: item.Department || '',
      Data: item.Data || ''
    }));

    return results;
  }

  /** Fetch AC_Data items filtered by Department */
  public async getAcDataByDepartment(department: string): Promise<IAcDataItem[]> {
    const webUrl = this.context.pageContext.web.absoluteUrl;

    const endpoint =
      `${webUrl}/_api/web/lists/getbytitle('${this.listName}')/items` +
      `?$select=Id,Year,Department,Data` +
      `&$filter=Department eq '${department.replace(/'/g, "''")}'` +
      `&$orderby=Year desc`;

    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      console.error('Failed to fetch AC_Data by Department:', department);
      return [];
    }

    const data = await response.json();
    const items = data.value || [];

    const results: IAcDataItem[] = items.map((item: any) => ({
      Id: item.Id,
      Year: item.Year || '',
      Department: item.Department || '',
      Data: item.Data || ''
    }));

    return results;
  }
}
