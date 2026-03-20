export interface INintexTaskAssignment {
  id: string;
  assignee: string;
  [key: string]: unknown;
}

export interface IRawNintexTask {
  id: string;
  workflowInstanceId: string;
  taskAssignments?: INintexTaskAssignment[];
  urls?: {
    formUrl?: string;
    taskUrl?: string;
  };
  [key: string]: unknown;
}

/**
 * Constructs a deep link to the Nintex Workflow Cloud task dashboard.
 * 
 * @param task The raw task object from the Nintex tasks API response
 * @param currentUserEmail The current logged-in user's email for finding their specific assignment
 * @param tenantHostname The NWC tenant hostname (e.g., "mydev26-507348.workflowcloud.com")
 * @returns A fully formed URL string, or undefined if required info is missing
 */
export function getNintexTaskUrl(
  task: IRawNintexTask, 
  currentUserEmail: string, 
  tenantHostname: string
): string | undefined {
  if (!task || !task.id || !task.workflowInstanceId) {
    return undefined;
  }

  const assignments = task.taskAssignments || [];
  
  if (assignments.length === 0) {
    return undefined; // Tasks without assignments cannot be routed this way
  }

  // 1. Try to find the exact assignment for the current user (case-insensitive)
  let matchedAssignment = assignments.find(
    a => a.assignee && a.assignee.toLowerCase() === currentUserEmail.toLowerCase()
  );

  // 2. Fallback to the first assignment if the current user isn't found
  if (!matchedAssignment) {
    matchedAssignment = assignments[0];
  }

  if (!matchedAssignment || !matchedAssignment.id) {
    return undefined;
  }

  // Ensure tenant hostname is clean of protocols/slashes if user passed full URL
  const cleanTenant = tenantHostname.replace(/^https?:\/\//, '').replace(/\/+$/, '');

  // Construct the expected NWC Dashboard URL
  return `https://${cleanTenant}/dashboard/my/tasks?activity-feed-route=instance/${encodeURIComponent(task.workflowInstanceId)}/task/${encodeURIComponent(task.id)}/assignment/${encodeURIComponent(matchedAssignment.id)}`;
}
