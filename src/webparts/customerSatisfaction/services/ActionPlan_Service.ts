import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IActionplan } from '../Models/ActionplanModel';

export interface IUserDirectoryEntry {
  id: number;
  loginName: string;
  displayName: string;
  email: string;
}

type IMultiChoicePayload = { results: string[] };

export interface IActionPlanUpsert extends Partial<Omit<IActionplan, 'UpdatedFeedback' | 'Actions' | 'Results' | 'PICId' | 'Category'>> {
  UpdatedFeedback?: string;
  Actions?: string;
  Results?: string;
  Category?: string[] | string | IMultiChoicePayload;
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

  private normalizeCategoryValues(category: IActionPlanUpsert['Category']): string[] {
    if (!category) {
      return [];
    }

    if (Array.isArray(category)) {
      return category
        .map(value => (value || '').trim())
        .filter(value => Boolean(value));
    }

    if (typeof category === 'string') {
      const trimmed = category.trim();
      return trimmed ? [trimmed] : [];
    }

    if (Array.isArray(category.results)) {
      return category.results
        .map(value => (value || '').trim())
        .filter(value => Boolean(value));
    }

    return [];
  }

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
        `?$select=Id,Title,Service,CustomerFeedback,UpdatedFeedback,Actions,PICId,PIC/EMail,PIC/Title,Timeline,Status,Results,RelatedLinks,Year,Category,ProductLine,Department,Division` +
        `&$expand=PIC` +
        `&$filter=Service eq '${service}'` +
        `&$orderby=Timeline desc`;

      return await this.fetchActionPlans(endpoint);
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
        `?$select=Id,Title,Service,CustomerFeedback,UpdatedFeedback,Actions,PICId,PIC/EMail,PIC/Title,Timeline,Status,Results,RelatedLinks,Year,Category,ProductLine,Department,Division` +
        `&$expand=PIC` +
        `&$filter=${filters}` +
        `&$orderby=Timeline desc`;

      return await this.fetchActionPlans(endpoint);
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

      const rows = await this.fetchActionPlans(endpoint);
      return rows.length > 0 ? rows[0] : undefined;
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
   * Fetches action plan items and normalizes the PIC value. The expanded PIC
   * can come back as an object, an array (multi-person field) or be missing
   * even though PICId is set; in the last case the user is resolved by id.
   */
  private async fetchActionPlans(endpoint: string): Promise<IActionplan[]> {
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Action plan list request failed with status ${response.status}`);
    }

    const data = await response.json();
    const rows = (Array.isArray(data?.value)
      ? data.value
      : (Array.isArray(data?.d?.results) ? data.d.results : [])) as IActionplan[];

    await this.resolveMissingPIC(rows);
    return rows;
  }

  private normalizePICValue(value: unknown): { Title: string; EMail: string } | undefined {
    if (Array.isArray(value)) {
      return this.normalizePICValue(value[0]);
    }

    if (value && typeof value === 'object') {
      const pic = value as { Title?: unknown; EMail?: unknown };
      const title = typeof pic.Title === 'string' ? pic.Title.trim() : '';
      const email = typeof pic.EMail === 'string' ? pic.EMail.trim() : '';
      if (title || email) {
        return { Title: title, EMail: email };
      }
    }

    return undefined;
  }

  private async resolveMissingPIC(plans: IActionplan[]): Promise<void> {
    const missingUserIds: number[] = [];
    plans.forEach(plan => {
      plan.PIC = this.normalizePICValue(plan.PIC);
      if (!plan.PIC && typeof plan.PICId === 'number' && missingUserIds.indexOf(plan.PICId) === -1) {
        missingUserIds.push(plan.PICId);
      }
    });

    if (missingUserIds.length === 0) {
      return;
    }

    const resolvedUsers = await Promise.all(missingUserIds.map(async userId => {
      try {
        const endpoint =
          `${this.context.pageContext.web.absoluteUrl}` +
          `/_api/web/getuserbyid(${userId})?$select=Id,Title,Email`;
        const response: SPHttpClientResponse = await this.context.spHttpClient.get(
          endpoint,
          SPHttpClient.configurations.v1
        );

        if (!response.ok) {
          console.warn(`Failed to resolve PIC user ${userId}:`, response.status);
          return undefined;
        }

        const user = await response.json() as { Title?: string; Email?: string };
        const title = (user.Title || '').trim();
        const email = (user.Email || '').trim();
        return (title || email) ? { Title: title, EMail: email } : undefined;
      } catch (error) {
        console.error(`ActionPlanService resolveMissingPIC error for user id ${userId}:`, error);
        return undefined;
      }
    }));

    const userMap: { [id: number]: { Title: string; EMail: string } } = {};
    resolvedUsers.forEach((user, index) => {
      if (user) {
        userMap[missingUserIds[index]] = user;
      }
    });

    plans.forEach(plan => {
      if (!plan.PIC && typeof plan.PICId === 'number' && userMap[plan.PICId]) {
        plan.PIC = userMap[plan.PICId];
      }
    });
  }

  /**
   * Gets all action plans (no filter - for admin use).
   */
  public async getAllActionPlans(): Promise<IActionplan[]> {
    try {
      const endpoint =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${this.listName}')/items` +
        `?$select=Id,Title,Service,CustomerFeedback,UpdatedFeedback,Actions,PICId,PIC/EMail,PIC/Title,Timeline,Status,Results,RelatedLinks,Year,Category,ProductLine,Department,Division` +
        `&$expand=PIC` +
        `&$orderby=Timeline desc`;

      return await this.fetchActionPlans(endpoint);
    } catch (error) {
      console.error('ActionPlanService getAllActionPlans error:', error);
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
        `?$select=Id,Title,CustomerFeedback,UpdatedFeedback,Service,PICId,PIC/EMail,PIC/Title,Timeline,Status,Actions,Results,RelatedLinks,Department,Year` +
        `&$expand=PIC` +
        `&$filter=Department eq '${department.replace(/'/g, "''")}'` +
        `&$orderby=Timeline desc`;

      return await this.fetchActionPlans(endpoint);
    } catch (error) {
      console.error('ActionPlanService getActionPlansByDepartment error:', error);
      return [];
    }
  }

   private buildActionPlanPayload(actionplan: IActionPlanUpsert): IActionPlanUpsert {
     // Build payload from an explicit allow-list to avoid sending SharePoint annotation fields.
     const payload: IActionPlanUpsert = {};

     if (actionplan.Title !== undefined) {
       payload.Title = actionplan.Title;
     }
     if (actionplan.Service !== undefined) {
       payload.Service = actionplan.Service;
     }
     if (actionplan.CustomerFeedback !== undefined) {
       payload.CustomerFeedback = actionplan.CustomerFeedback;
     }
     if (actionplan.UpdatedFeedback !== undefined) {
       payload.UpdatedFeedback = actionplan.UpdatedFeedback;
     }
     if (actionplan.Actions !== undefined) {
       payload.Actions = Array.isArray(actionplan.Actions) ? actionplan.Actions.join('\n') : actionplan.Actions;
     }
     if (actionplan.Timeline !== undefined) {
       payload.Timeline = actionplan.Timeline;
     }
     if (actionplan.Status !== undefined) {
       payload.Status = actionplan.Status;
     }
     if (actionplan.Results !== undefined) {
       payload.Results = Array.isArray(actionplan.Results) ? actionplan.Results.join('\n') : actionplan.Results;
     }
     if (actionplan.RelatedLinks !== undefined) {
       payload.RelatedLinks = actionplan.RelatedLinks;
     }
     if (actionplan.Year !== undefined) {
       payload.Year = actionplan.Year;
     }
     if (actionplan.Category !== undefined) {
       const categoryValues = this.normalizeCategoryValues(actionplan.Category);
       if (categoryValues.length > 0) {
         // Multi-choice field expects a JSON array in this API mode.
         payload.Category = categoryValues;
       }
     }
     if (actionplan.ProductLine !== undefined) {
       payload.ProductLine = actionplan.ProductLine;
     }
     if (actionplan.Department !== undefined) {
       payload.Department = actionplan.Department;
     }
     if (actionplan.Division !== undefined) {
       payload.Division = actionplan.Division;
     }

     // For Person field updates, only send lookup ID.
     if (typeof actionplan.PICId === 'number') {
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
