import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IActionplan } from '../Models/ActionplanModel';

export class ActionPlanService {
  private context: WebPartContext;
  private listName: string;

  constructor(context: WebPartContext, listName: string) {
    this.context = context;
    this.listName = listName;
  }

  /**
   * Gets all action plans filtered by service.
   */
  public async getActionPlansByService(service: string): Promise<IActionplan[]> {
    try {
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items` +
        `?$select=Id,Title,Service,CustomerFeedback,Actions,PIC/EMail,PIC/Title,Timeline,Status,Results,RelatedLinks,Year,Category,ProductLine,Department,Division` +
        `&$expand=PIC` +
        `&$filter=Service eq '${service}'` +
        `&$orderby=Timeline desc`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        console.error('Failed to fetch action plans');
        return [];
      }

      const data = await response.json();
      return Array.isArray(data.value) ? (data.value as IActionplan[]) : [];
    } catch (error) {
      console.error('ActionPlanService getActionPlansByService error:', error);
      return [];
    }
  }

  /**
   * Gets a single action plan by ID.
   */
  public async getActionPlanById(id: number): Promise<IActionplan | undefined> {
    try {
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items(${id})` +
        `?$select=Id,Title,Service,CustomerFeedback,Actions,PIC/EMail,PIC/Title,Timeline,Status,Results,RelatedLinks,Year,Category,ProductLine,Department,Division` +
        `&$expand=PIC`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        return undefined;
      }

      return (await response.json()) as IActionplan;
    } catch (error) {
      console.error('ActionPlanService getActionPlanById error:', error);
      return undefined;
    }
  }

  /**
   * Creates a new action plan item.
   */
  public async createActionPlan(actionplan: Partial<IActionplan>): Promise<IActionplan | undefined> {
    try {
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items`;

      const body = JSON.stringify(actionplan);

      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          body,
          headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/json',
          },
        }
      );

      if (!response.ok) {
        console.error('Failed to create action plan');
        return undefined;
      }

      return (await response.json()) as IActionplan;
    } catch (error) {
      console.error('ActionPlanService createActionPlan error:', error);
      return undefined;
    }
  }

  /**
   * Updates an existing action plan item.
   */
  public async updateActionPlan(id: number, actionplan: Partial<IActionplan>): Promise<boolean> {
    try {
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items(${id})`;

      const body = JSON.stringify(actionplan);

      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          body,
          headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/json',
            'X-HTTP-Method': 'MERGE',
            'If-Match': '*',
          },
        }
      );

      return response.ok;
    } catch (error) {
      console.error('ActionPlanService updateActionPlan error:', error);
      return false;
    }
  }
}
