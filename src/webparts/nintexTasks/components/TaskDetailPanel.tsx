import * as React from 'react';
import {
  Panel,
  PanelType,
  Stack,
  Text,
  Label,
  Icon,
  Separator
} from '@fluentui/react';
import { INintexTask } from '../models/INintexTask';
import styles from './NintexTasks.module.scss';

export interface ITaskDetailPanelProps {
  task: INintexTask | undefined;
  isOpen: boolean;
  onDismiss: () => void;
}

export default class TaskDetailPanel extends React.Component<ITaskDetailPanelProps> {

  private _formatDate(dateStr: string | undefined): string {
    if (!dateStr) return '—';
    try {
      const date = new Date(dateStr);
      return date.toLocaleDateString('en-US', {
        year: 'numeric',
        month: 'short',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
      });
    } catch {
      return dateStr;
    }
  }

  private _getStatusIcon(status: string): string {
    switch (status?.toLowerCase()) {
      case 'active': return 'Clock';
      case 'completed': return 'CheckMark';
      case 'expired': return 'Warning';
      case 'overridden': return 'Forward';
      default: return 'Info';
    }
  }

  public render(): React.ReactElement<ITaskDetailPanelProps> {
    const { task, isOpen, onDismiss } = this.props;

    if (!task) return <></>;

    return (
      <Panel
        isOpen={isOpen}
        onDismiss={onDismiss}
        type={PanelType.medium}
        headerText="Task Details"
        closeButtonAriaLabel="Close"
        className={styles.detailPanel}
      >
        <Stack tokens={{ childrenGap: 16 }} className={styles.detailContent}>
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
            <Icon
              iconName={this._getStatusIcon(task.status)}
              className={`${styles.statusIcon} ${(styles as Record<string, string>)[`status${task.status?.charAt(0).toUpperCase()}${task.status?.slice(1).toLowerCase()}`] || ''}`}
            />
            <Text variant="xLarge" className={styles.detailTitle}>{task.name}</Text>
          </Stack>

          <Separator />

          <Stack tokens={{ childrenGap: 12 }}>
            <Stack className={styles.detailField}>
              <Label className={styles.detailLabel}>Status</Label>
              <Text className={styles.detailValue}>{task.status || '—'}</Text>
            </Stack>

            <Stack className={styles.detailField}>
              <Label className={styles.detailLabel}>Workflow</Label>
              <Text className={styles.detailValue}>{task.workflowName || '—'}</Text>
            </Stack>

            <Stack className={styles.detailField}>
              <Label className={styles.detailLabel}>Assignee</Label>
              <Text className={styles.detailValue}>{task.assignee || '—'}</Text>
            </Stack>

            {task.initiator && (
              <Stack className={styles.detailField}>
                <Label className={styles.detailLabel}>Initiated By</Label>
                <Text className={styles.detailValue}>{task.initiator}</Text>
              </Stack>
            )}

            <Stack className={styles.detailField}>
              <Label className={styles.detailLabel}>Created Date</Label>
              <Text className={styles.detailValue}>{this._formatDate(task.createdDate)}</Text>
            </Stack>

            <Stack className={styles.detailField}>
              <Label className={styles.detailLabel}>Due Date</Label>
              <Text className={styles.detailValue}>{this._formatDate(task.dueDate)}</Text>
            </Stack>

            {task.completedDate && (
              <Stack className={styles.detailField}>
                <Label className={styles.detailLabel}>Completed Date</Label>
                <Text className={styles.detailValue}>{this._formatDate(task.completedDate)}</Text>
              </Stack>
            )}

            {task.outcome && (
              <Stack className={styles.detailField}>
                <Label className={styles.detailLabel}>Outcome</Label>
                <Text className={styles.detailValue}>{task.outcome}</Text>
              </Stack>
            )}

            {task.taskType && (
              <Stack className={styles.detailField}>
                <Label className={styles.detailLabel}>Task Type</Label>
                <Text className={styles.detailValue}>{task.taskType}</Text>
              </Stack>
            )}

            {task.description && (
              <Stack className={styles.detailField}>
                <Label className={styles.detailLabel}>Description</Label>
                <div className={styles.detailValue} dangerouslySetInnerHTML={{ __html: task.description }} />
              </Stack>
            )}

            {task.formUrl && (
              <Stack className={styles.detailField}>
                <Label className={styles.detailLabel}>Task Form</Label>
                <a href={task.formUrl} target="_blank" rel="noreferrer" className={styles.formLink}>
                  Open Task Form
                </a>
              </Stack>
            )}
          </Stack>
        </Stack>
      </Panel>
    );
  }
}
