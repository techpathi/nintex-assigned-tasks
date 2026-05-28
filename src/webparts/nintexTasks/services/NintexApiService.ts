import { HttpClient, HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';
import { INintexTask, ITaskListResponse, TaskStatus } from '../models/INintexTask';
import { getNintexTaskUrl, IRawNintexTask, transformToTaskFormUrl } from '../utils/NintexTaskUtils';

export interface INintexUser {
  id: string;
  email: string;
  firstName?: string;
  lastName?: string;
}

export interface IListTasksOptions {
  status?: TaskStatus;
  assignee?: string;
  from?: string;
  to?: string;
}

export class NintexApiService {
  private httpClient: HttpClient;
  private tenantUrl: string;
  private dashboardUrl: string;
  private accessToken: string;
  private currentUserEmail: string;

  constructor(httpClient: HttpClient, tenantUrl: string, dashboardUrl: string, accessToken: string, currentUserEmail: string) {
    this.httpClient = httpClient;
    this.tenantUrl = tenantUrl.replace(/\/+$/, '');
    this.dashboardUrl = dashboardUrl.replace(/\/+$/, '');
    this.accessToken = accessToken;
    this.currentUserEmail = currentUserEmail;
  }

  /**
   * List tasks with optional filters and pagination
   */
  public async listTasks(options: IListTasksOptions = {}): Promise<ITaskListResponse> {
    let url = `${this.tenantUrl}/workflows/v2/tasks?workflowMode=standard&$top=25`;

    if (options.status) {
      url += `&status=${options.status}`;
    }

    options.assignee = 'dev2@techpathi.site';

    if (options.assignee) {
      url += `&assignee=${encodeURIComponent(options.assignee.toLowerCase())}`;
    }

    if (options.from) {
      url += `&from=${encodeURIComponent(options.from)}`;
    }

    if (options.to) {
      url += `&to=${encodeURIComponent(options.to)}`;
    }

    const requestHeaders: Headers = new Headers();
    requestHeaders.append('Authorization', `Bearer ${this.accessToken}`);
    requestHeaders.append('Accept', 'application/json');

    const httpClientOptions: IHttpClientOptions = {
      headers: requestHeaders
    };

    const response: HttpClientResponse = await this.httpClient.get(
      url,
      HttpClient.configurations.v1,
      httpClientOptions
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to list tasks: ${response.status} ${response.statusText} - ${errorText}`);
    }

    const data = await response.json();

    const tasks: INintexTask[] = (data.tasks || data || []).map((task: Record<string, unknown>) => {
      const assignments = (task.taskAssignments as Record<string, unknown>[]) || [];
      const primaryAssignment = assignments[0] || {};
      const urlsObj = (task.urls || primaryAssignment.urls) as Record<string, unknown> | undefined;

      return {
        id: (primaryAssignment.id as string) || (task.id as string) || '',
        name: task.name as string || task.subject as string || 'Untitled Task',
        status: primaryAssignment.status as string || task.status as string || '',
        workflowName: task.workflowName as string || (task.workflow as Record<string, unknown>)?.name as string || '',
        assignee: task.assignee as string || primaryAssignment.assignee as string || (task.assignees as string[])?.join(', ') || '',
        createdDate: task.createdDate as string || task.dateCreated as string || '',
        dueDate: task.dueDate as string || '',
        completedDate: task.completedDate as string || primaryAssignment.completedDate as string || task.dateCompleted as string || undefined,
        completedBy: task.completedBy as string || primaryAssignment.completedBy as string || (primaryAssignment.outcome ? primaryAssignment.assignee as string : undefined) || '-',
        description: task.description as string || task.message as string || '',
        outcome: task.outcome as string || primaryAssignment.outcome as string || undefined,
        taskType: task.taskType as string || task.type as string || undefined,
        workflowInstanceId: task.workflowInstanceId as string || undefined,
        formUrl: transformToTaskFormUrl(task.formUrl as string || (urlsObj ? urlsObj.formUrl as string : undefined) || undefined),
        taskUrl: transformToTaskFormUrl(getNintexTaskUrl(task as IRawNintexTask, this.currentUserEmail, this.dashboardUrl) || (urlsObj ? (urlsObj.taskUrl as string || urlsObj.formUrl as string) : undefined) || task.formUrl as string || undefined),
        initiator: task.initiator as string || undefined,
        taskId: task.id as string,
        assignmentId: primaryAssignment.id as string || task.id as string
      };
    });

    return {
      tasks,
      totalCount: data.totalCount || undefined
    };
  }

  /**
   * Get a single task by ID
   */
  public async getTask(taskId: string): Promise<INintexTask> {
    const url = `${this.tenantUrl}/workflows/v2/tasks/${taskId}`;

    const requestHeaders: Headers = new Headers();
    requestHeaders.append('Authorization', `Bearer ${this.accessToken}`);
    requestHeaders.append('Accept', 'application/json');

    const httpClientOptions: IHttpClientOptions = {
      headers: requestHeaders
    };

    const response: HttpClientResponse = await this.httpClient.get(
      url,
      HttpClient.configurations.v1,
      httpClientOptions
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to get task: ${response.status} ${response.statusText} - ${errorText}`);
    }

    const task: Record<string, unknown> = await response.json();

    const assignments = (task.taskAssignments as Record<string, unknown>[]) || [];
    const primaryAssignment = assignments[0] || {};
    const urlsObj = (task.urls || primaryAssignment.urls) as Record<string, unknown> | undefined;

    return {
      id: (primaryAssignment.id as string) || (task.id as string) || '',
      name: task.name as string || task.subject as string || 'Untitled Task',
      status: primaryAssignment.status as string || task.status as string || '',
      workflowName: task.workflowName as string || (task.workflow as Record<string, unknown>)?.name as string || '',
      assignee: task.assignee as string || primaryAssignment.assignee as string || (task.assignees as string[])?.join(', ') || '',
      createdDate: task.createdDate as string || task.dateCreated as string || '',
      dueDate: task.dueDate as string || '',
      completedDate: task.completedDate as string || primaryAssignment.completedDate as string || task.dateCompleted as string || undefined,
      completedBy: task.completedBy as string || primaryAssignment.completedBy as string || (primaryAssignment.outcome ? primaryAssignment.assignee as string : undefined) || '-',
      description: task.description as string || task.message as string || '',
      outcome: task.outcome as string || primaryAssignment.outcome as string || undefined,
      taskType: task.taskType as string || task.type as string || undefined,
      workflowInstanceId: task.workflowInstanceId as string || undefined,
      formUrl: transformToTaskFormUrl(task.formUrl as string || (urlsObj ? urlsObj.formUrl as string : undefined) || undefined),
      taskUrl: transformToTaskFormUrl(getNintexTaskUrl(task as IRawNintexTask, this.currentUserEmail, this.dashboardUrl) || (urlsObj ? (urlsObj.taskUrl as string || urlsObj.formUrl as string) : undefined) || task.formUrl as string || undefined),
      initiator: task.initiator as string || undefined,
      taskId: task.id as string,
      assignmentId: primaryAssignment.id as string || task.id as string
    };
  }

  /**
   * Delegate a task assignment to other users
   */
  public async delegateTaskAssignment(taskId: string, assignmentId: string, delegateeEmails: string[], message: string, token: string): Promise<boolean> {
    const endpoint = `${this.tenantUrl}/workflows/v2/tasks/${taskId}/assignments/${assignmentId}/delegate`;
    
    const options: IHttpClientOptions = {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        assignees: delegateeEmails,
        message: message || undefined
      }),
      method: 'PUT'
    };

    const response: HttpClientResponse = await this.httpClient.fetch(endpoint, HttpClient.configurations.v1, options);
    
    if (!response.ok) {
      const errorMsg = await response.text();
      throw new Error(`Failed to delegate task ${taskId} assignment ${assignmentId}: ${errorMsg}`);
    }

    return true;
  }

  /**
   * Search Nintex users for delegation
   */
  public async searchNintexUsers(filterText: string, token: string): Promise<INintexUser[]> {
    if (!filterText) return [];
    const endpoint = `${this.tenantUrl}/tenants/v1/users?filter=${encodeURIComponent(filterText)}&limit=20`;
    
    const options: IHttpClientOptions = {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/json'
      }
    };

    const response: HttpClientResponse = await this.httpClient.get(endpoint, HttpClient.configurations.v1, options);
    if (!response.ok) {
      console.error("Failed to fetch Nintex users", await response.text());
      return [];
    }

    const data = await response.json();
    return Array.isArray(data) ? data : data.users || [];
  }
}
