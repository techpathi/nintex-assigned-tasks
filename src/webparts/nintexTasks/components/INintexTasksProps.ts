import { HttpClient, SPHttpClient } from '@microsoft/sp-http';
import { TaskStatus } from '../models/INintexTask';

export interface INintexTasksProps {
  tenantUrl: string;
  tokenListUrl: string;
  tokenTitleFilter?: string;
  tokenColumnName?: string;
  dashboardUrl: string;
  defaultStatusFilter: TaskStatus;
  httpClient: HttpClient;
  spHttpClient: SPHttpClient;
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  userDisplayName: string;
  userEmail: string;
}
