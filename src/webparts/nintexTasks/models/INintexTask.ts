export interface INintexTask {
  id: string;
  name: string;
  status: string;
  workflowName: string;
  assignee: string;
  createdDate: string;
  dueDate: string;
  completedDate?: string;
  completedBy?: string;
  description: string;
  outcome?: string;
  taskType?: string;
  workflowInstanceId?: string;
  formUrl?: string;
  taskUrl?: string;
  initiator?: string;
}

export interface ITaskListResponse {
  tasks: INintexTask[];
  totalCount?: number;
}

export type TaskStatus = 'active' | 'expired' | 'complete' | 'overridden' | 'terminated' | 'all';
