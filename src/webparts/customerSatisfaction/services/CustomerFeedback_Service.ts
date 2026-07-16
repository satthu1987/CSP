import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ICustomerFeedback } from '../Models/CustomerFeedbackModel';

export class CustomerFeedbackService {
  private context: WebPartContext;
  private listName: string;

  constructor(context: WebPartContext, listName: string) {
    this.context = context;
    this.listName = listName;
  }

  public async getAllCustomerFeedbacks(): Promise<ICustomerFeedback[]> {
    try {
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items` +
        `?$select=Id,Title,CustomerFeedback,UpdatedFeedback,Service,ActionPlan` +
        `&$orderby=Service asc,Title asc`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        console.error('Failed to fetch customer feedbacks');
        return [];
      }

      const data = await response.json();
      return Array.isArray(data.value) ? (data.value as ICustomerFeedback[]) : [];
    } catch (error) {
      console.error('CustomerFeedbackService getAllCustomerFeedbacks error:', error);
      return [];
    }
  }

  /**
   * Gets customer feedbacks filtered by services.
   * @param services Array of service names to filter by
   */
  public async getCustomerFeedbacksByServices(services: string[]): Promise<ICustomerFeedback[]> {
    try {
      if (!services || services.length === 0) {
        console.warn('No services provided for filtering customer feedbacks');
        return [];
      }

      // Build filter: Service eq 'Service1' or Service eq 'Service2' ...
      const filters = services
        .filter(s => s && s.trim())
        .map(s => `Service eq '${s.replace(/'/g, "''")}'`)
        .join(' or ');

      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items` +
        `?$select=Id,Title,CustomerFeedback,UpdatedFeedback,Service,ActionPlan` +
        `&$filter=${filters}` +
        `&$orderby=Service asc,Title asc`;

      console.log('Fetching customer feedbacks with filter:', filters);

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        console.error('Failed to fetch customer feedbacks by services', { status: response.status });
        return [];
      }

      const data = await response.json();
      const feedbacks = Array.isArray(data.value) ? (data.value as ICustomerFeedback[]) : [];
      console.log('Fetched customer feedbacks:', { count: feedbacks.length, services });
      return feedbacks;
    } catch (error) {
      console.error('CustomerFeedbackService getCustomerFeedbacksByServices error:', error);
      return [];
    }
  }

  public async updateActionPlan(feedbackId: number, actionPlanId: number): Promise<boolean> {
    try {
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items(${feedbackId})`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          body: JSON.stringify({ ActionPlan: actionPlanId.toString() }),
          headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/json',
            'X-HTTP-Method': 'MERGE',
            'If-Match': '*',
          },
        }
      );

      if (!response.ok) {
        console.error('Failed to update ActionPlan for feedback', { feedbackId, actionPlanId, status: response.status });
        return false;
      }

      console.log('Successfully updated ActionPlan for feedback:', { feedbackId, actionPlanId });
      return true;
    } catch (error) {
      console.error('CustomerFeedbackService updateActionPlan error:', error);
      return false;
    }
  }
}
