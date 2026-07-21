import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IActionplan } from '../Models/ActionplanModel';

export interface IUserDirectoryEntry {
  id: number;
  loginName: string;
  displayName: string;
  email: string;
}

export interface IActionPlanUpsert extends Partial<Omit<IActionplan, 'UpdatedFeedback' | 'Actions' | 'Results' | 'PICId'>> {
  UpdatedFeedback?: string;
  Actions?: string[] | string;
  Results?: string[] | string;
  PICId?: number;
}

interface IPeoplePickerEntity {
  Key?: string;
  DisplayText?: string;
  Description?: string;
  EntityData?: {
    Email?: string;
  };
}

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
        `?$select=Id,Title,Service,UpdatedFeedback,Actions,PICId,PIC/EMail,PIC/Title,Timeline,Status,Results,RelatedLinks,Year,Category,ProductLine,Department,Division` +
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
   * Gets all action plans filtered by multiple services.
   */
  public async getActionPlansByServices(services: string[]): Promise<IActionplan[]> {
    try {
      if (!services || services.length === 0) {
        return [];
      }

      // Build filter for multiple services
      const filters = services
        .filter(s => s && s.trim())
        .map(s => `Service eq '${s.replace(/'/g, "''")}'`)
        .join(' or ');

      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items` +
        `?$select=Id,Title,Service,UpdatedFeedback,Actions,PICId,PIC/EMail,PIC/Title,Timeline,Status,Results,RelatedLinks,Year,Category,ProductLine,Department,Division` +
        `&$expand=PIC` +
        `&$filter=${filters}` +
        `&$orderby=Timeline desc`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        console.error('Failed to fetch action plans by services');
        return [];
      }

      const data = await response.json();
      return Array.isArray(data.value) ? (data.value as IActionplan[]) : [];
    } catch (error) {
      console.error('ActionPlanService getActionPlansByServices error:', error);
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
        `?$select=Id,Title,Service,CustomerFeedback,UpdatedFeedback,Actions,PICId,PIC/EMail,PIC/Title,Timeline,Status,Results,RelatedLinks,Year,Category,ProductLine,Department,Division` +
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
  public async createActionPlan(actionplan: IActionPlanUpsert): Promise<IActionplan | undefined> {
    try {
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items`;

      const payload = this.buildActionPlanPayload(actionplan);
      const body = JSON.stringify(payload);

      console.log('Creating ActionPlan with payload:', payload);

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
        console.error('Failed to create action plan', { status: response.status, statusText: response.statusText });
        return undefined;
      }

      const result = (await response.json()) as IActionplan;
      console.log('ActionPlan created successfully:', result);
      return result;
    } catch (error) {
      console.error('ActionPlanService createActionPlan error:', error);
      return undefined;
    }
  }

   /**
    * Updates an existing action plan item.
    */
   public async updateActionPlan(id: number, actionplan: IActionPlanUpsert): Promise<boolean> {
     try {
       const endpoint =
         `${this.context.pageContext.web.absoluteUrl}` +
         `/_api/web/lists/getbytitle('${this.listName}')/items(${id})`;

       const body = JSON.stringify(this.buildActionPlanPayload(actionplan));

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

   /**
    * Gets choice options for a specific field.
    */
   public async getFieldChoices(fieldName: string): Promise<string[]> {
     try {
       const endpoint =
         `${this.context.pageContext.web.absoluteUrl}` +
         `/_api/web/lists/getbytitle('${this.listName}')/fields/getbytitle('${fieldName}')` +
         `?$select=Choices`;

       const response: SPHttpClientResponse = await this.context.spHttpClient.get(
         endpoint,
         SPHttpClient.configurations.v1
       );

       if (!response.ok) {
         console.warn(`Failed to fetch choices for field ${fieldName}:`, response.status);
         return [];
       }

       const data = await response.json();
       console.log(`Choices for ${fieldName}:`, data.Choices);
       return Array.isArray(data.Choices) ? data.Choices : [];
     } catch (error) {
       console.error(`ActionPlanService getFieldChoices error for ${fieldName}:`, error);
       return [];
     }
   }

   /**
    * Gets all choice field options for dropdown lists.
    */
   public async getAllChoiceOptions(): Promise<{ [key: string]: string[] }> {
     const fields = ['Service', 'Status', 'Category', 'ProductLine', 'Department', 'Division'];
     const result: { [key: string]: string[] } = {};

     for (const field of fields) {
       result[field] = await this.getFieldChoices(field);
     }

     return result;
   }

   public async searchUsers(query: string): Promise<IUserDirectoryEntry[]> {
     try {
       const endpoint =
         `${this.context.pageContext.web.absoluteUrl}` +
         `/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;

       const response: SPHttpClientResponse = await this.context.spHttpClient.post(
         endpoint,
         SPHttpClient.configurations.v1,
         {
           body: JSON.stringify({
             queryParams: {
               AllowEmailAddresses: true,
               AllowMultipleEntities: false,
               AllUrlZones: false,
               MaximumEntitySuggestions: 12,
               PrincipalSource: 15,
               PrincipalType: 1,
               QueryString: query,
             },
           }),
           headers: {
             Accept: 'application/json;odata=verbose',
             'Content-Type': 'application/json;odata=verbose',
           },
         }
       );

       if (!response.ok) {
         console.warn('Failed to search users:', response.status);
         return [];
       }

       const data = await response.json();
       const rawResults = data?.d?.ClientPeoplePickerSearchUser;
       const parsedResults = typeof rawResults === 'string' ? (JSON.parse(rawResults) as IPeoplePickerEntity[]) : [];
       const users = await Promise.all(
         parsedResults.map(async (user) => {
           const loginName = user.Key || '';
           const displayName = user.DisplayText || '';
           const email = user.EntityData?.Email || user.Description || '';
           const ensuredUserId = await this.ensureUser(loginName);

           if (!ensuredUserId) {
             return undefined;
           }

           return {
             id: ensuredUserId,
             loginName,
             displayName,
             email,
           } as IUserDirectoryEntry;
         })
       );

       return users.filter((user): user is IUserDirectoryEntry => Boolean(user));
     } catch (error) {
       console.error('ActionPlanService searchUsers error:', error);
       return [];
     }
   }

  /**
   * Gets all action plans filtered by department.
   */
  public async getActionPlansByDepartment(department: string): Promise<IActionplan[]> {
    try {
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items` +
        `?$select=Id,Title,Service,PICId,PIC/EMail,PIC/Title,Timeline,Status,Department` +
        `&$expand=PIC` +
        `&$filter=Department eq '${department.replace(/'/g, "''")}'` +
        `&$orderby=Timeline desc`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        console.error('Failed to fetch action plans by department');
        return [];
      }

      const data = await response.json();
      return Array.isArray(data.value) ? (data.value as IActionplan[]) : [];
    } catch (error) {
      console.error('ActionPlanService getActionPlansByDepartment error:', error);
      return [];
    }
  }

   private buildActionPlanPayload(actionplan: IActionPlanUpsert): IActionPlanUpsert {
     const payload: IActionPlanUpsert = { ...actionplan };

     // Remove complex objects that shouldn't be sent to SharePoint REST API
     delete payload.PIC;
     
     // Ensure PICId is explicitly set
     if (actionplan.PICId) {
       payload.PICId = actionplan.PICId;
     }

     console.log('buildActionPlanPayload - Input:', actionplan, 'Output:', payload);
     return payload;
   }

   private async ensureUser(loginName: string): Promise<number | undefined> {
     if (!loginName) {
       return undefined;
     }

     try {
       const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/ensureuser`;
       const response: SPHttpClientResponse = await this.context.spHttpClient.post(
         endpoint,
         SPHttpClient.configurations.v1,
         {
           body: JSON.stringify({ logonName: loginName }),
           headers: {
             Accept: 'application/json;odata=nometadata',
             'Content-Type': 'application/json;odata=nometadata',
           },
         }
       );

       if (!response.ok) {
         console.warn('Failed to ensure user:', loginName, response.status);
         return undefined;
       }

       const user = await response.json();
       return user?.Id as number | undefined;
     } catch (error) {
       console.error('ActionPlanService ensureUser error:', error);
       return undefined;
     }
   }
}
