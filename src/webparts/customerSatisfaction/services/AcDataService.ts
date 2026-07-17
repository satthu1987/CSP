import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IAcDataItem {
  Id: number;
  Year: string;
  Department: string;
  DataUrl: string;
  DataAlt: string;
}

export class AcDataService {
  private context: WebPartContext;
  private listName: string = 'AC_Data';
  private listIdCache: string = '';

  constructor(context: WebPartContext) {
    this.context = context;
  }

  /** Fetch Year + Data (image column) from AC_Data list */
  public async getAcData(): Promise<IAcDataItem[]> {
    const webUrl = this.context.pageContext.web.absoluteUrl;

    // 1. Get list ID once (needed to build image URL fallback)
    const listId = await this.getListId();

    // 2. Fetch items
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

    // 3. For each item, also fetch FieldValuesAsHtml to get the rendered <img>
    const results: IAcDataItem[] = await Promise.all(
      items.map(async (item: any) => {
        console.log('Processing item:', item.Data);
        let imageUrl = this.buildImageUrlFromJson(item.Data, webUrl, listId, item.Id);

        // Fallback: try FieldValuesAsHtml if URL is missing
        if (!imageUrl) {
          imageUrl = await this.getImageFromHtmlEndpoint(item.Id);
        }

        return {
          Id: item.Id,
          Year: item.Year || '',
          Department: item.Department || '',
          DataUrl: imageUrl,
          DataAlt: `${item.Year} Data`
        };
      })
    );

    return results;
  }

  /** Fetch AC_Data items filtered by Department */
  public async getAcDataByDepartment(department: string): Promise<IAcDataItem[]> {
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const listId = await this.getListId();

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

    const results: IAcDataItem[] = await Promise.all(
      items.map(async (item: any) => {
        let imageUrl = this.buildImageUrlFromJson(item.Data, webUrl, listId, item.Id);

        if (!imageUrl) {
          imageUrl = await this.getImageFromHtmlEndpoint(item.Id);
        }

        return {
          Id: item.Id,
          Year: item.Year || '',
          Department: item.Department || '',
          DataUrl: imageUrl,
          DataAlt: `${item.Year} Data`
        };
      })
    );

    return results;
  }

  /** Get and cache the list GUID */
  private async getListId(): Promise<string> {
    if (this.listIdCache) return this.listIdCache;

    const endpoint =
      `${this.context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists/getbytitle('${this.listName}')?$select=Id`;

    const response = await this.context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    );
    if (!response.ok) return '';

    const data = await response.json();
    this.listIdCache = data.Id || '';
    return this.listIdCache;
  }

  /**
   * Modern Image column stores JSON like:
   *   {"fileName":"Reserved_ImageAttachment_[4]_[Data]...jpg","originalImageName":"2025 data"}
   * Files live at: {webUrl}/SiteAssets/Lists/{listId}/{fileName}
   */
  private buildImageUrlFromJson(raw: string, webUrl: string, listId: string,itemid:string): string {
    if (!raw) return '';

    try {
      const parsed = JSON.parse(raw);

      // Case 1: Already has serverRelativeUrl (older / migrated lists)
      if (parsed.serverRelativeUrl) {
        return this.toAbsolute(parsed.serverRelativeUrl);
      }
      if (parsed.serverUrl && parsed.serverRelativeUrl) {
        return parsed.serverUrl + parsed.serverRelativeUrl;
      }

      // Case 2: Modern image column → build from fileName + list ID
      if (parsed.fileName && listId) {
       // const cleanListId = listId.replace(/[{}]/g, '');
        return `${webUrl}/Lists/${this.listName}/Attachments/${itemid}/${encodeURIComponent(parsed.fileName)}`;
      }

      return '';
    } catch {
      // If it's already a plain URL string
      if (raw.indexOf('http') === 0 || raw.indexOf('/') === 0) {
        return this.toAbsolute(raw);
      }
      return '';
    }
  }

  /**
   * Reliable fallback: SharePoint's FieldValuesAsHtml endpoint
   * returns the rendered HTML for the Image column with the correct <img src="...">.
   */
  private async getImageFromHtmlEndpoint(itemId: number): Promise<string> {
    try {
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items(${itemId})/FieldValuesAsHtml`;

      const response = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return '';

      const data = await response.json();
      const html: string = data.Data || '';

      // Extract src="..." from the returned <img ... />
      const match = html.match(/src=["']([^"']+)["']/i);
      if (match && match[1]) {
        // Decode HTML entities (&amp; etc.) and resolve to absolute
        const decoded = match[1].replace(/&amp;/g, '&');
        return this.toAbsolute(decoded);
      }
      return '';
    } catch (err) {
      console.warn('FieldValuesAsHtml fallback failed for item', itemId, err);
      return '';
    }
  }

  /** Convert server-relative URL → absolute */
  private toAbsolute(url: string): string {
    if (!url) return '';
    if (url.indexOf('http') === 0) return url;
    return `${window.location.origin}${url}`;
  }
}