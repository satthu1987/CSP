import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ICSPLearnMoreContent } from '../Models/CSPLearnMoreContent';

export class CSPLearnMoreContentService {
  private context: WebPartContext;
  private listName: string = 'CSP_LearnMoreContent';

  constructor(context: WebPartContext) {
    this.context = context;
  }

  /**
   * Gets all items from the CSP_LearnMoreContent list.
   */
  public async getAllContent(): Promise<ICSPLearnMoreContent[]> {
    try {
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items` +
        `?$select=Id,Title,Content` +
        `&$orderby=Created asc`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        console.error('Failed to fetch CSP_LearnMoreContent list', response.status);
        return [];
      }

      const data = await response.json();
      return Array.isArray(data.value) ? (data.value as ICSPLearnMoreContent[]) : [];
    } catch (error) {
      console.error('CSPLearnMoreContentService getAllContent error:', error);
      return [];
    }
  }
}
