import * as React from 'react';
import styles from './NintexTasks.module.scss';
import type { INintexTasksProps } from './INintexTasksProps';
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  CheckboxVisibility,
  IColumn,
  IDetailsRowProps,
  DetailsRow
} from '@fluentui/react/lib/DetailsList';
import {
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  SearchBox,
  Dropdown,
  IDropdownOption,
  Stack,
  Text,
  Icon,
  Callout,
  DirectionalHint,
  DatePicker
} from '@fluentui/react';
import { INintexTask, TaskStatus } from '../models/INintexTask';
import { NintexApiService } from '../services/NintexApiService';
import { TokenService } from '../services/TokenService';
import TaskDetailPanel from './TaskDetailPanel';


interface INintexTasksState {
  tasks: INintexTask[];
  filteredTasks: INintexTask[];
  isLoading: boolean;
  error: string | undefined;
  searchText: string;
  statusFilter: TaskStatus;
  selectedTask: INintexTask | undefined;
  isPanelOpen: boolean;
  sortColumn: string;
  sortDescending: boolean;
  // Date Filter
  isDateFilterOpen: boolean;
  dateFilterTab: 'Preset' | 'Custom';
  dateFilterPreset: '30' | '60' | '90' | 'custom' | 'all';
  dateType: 'assigned' | 'completed';
  customDateStart?: Date;
  customDateEnd?: Date;
}

export default class NintexTasks extends React.Component<INintexTasksProps, INintexTasksState> {
  private _columns: IColumn[];
  private _dateFilterButton = React.createRef<HTMLDivElement>();

  constructor(props: INintexTasksProps) {
    super(props);

    this.state = {
      tasks: [],
      filteredTasks: [],
      isLoading: false,
      error: undefined,
      searchText: '',
      statusFilter: props.defaultStatusFilter || 'active',
      selectedTask: undefined,
      isPanelOpen: false,
      sortColumn: 'createdDate',
      sortDescending: true,
      isDateFilterOpen: false,
      dateFilterTab: 'Preset',
      dateFilterPreset: '30',
      dateType: 'assigned',
      customDateStart: undefined,
      customDateEnd: undefined
    };

    this._columns = this._buildColumns();
  }

  private _buildColumns(): IColumn[] {
    return [
      {
        key: 'name',
        name: 'Task',
        fieldName: 'name',
        minWidth: 150,
        maxWidth: 280,
        isResizable: true,
        isSorted: this.state?.sortColumn === 'name',
        isSortedDescending: this.state?.sortColumn === 'name' ? this.state?.sortDescending : false,
        onColumnClick: this._onColumnClick.bind(this),
        onRender: (item: INintexTask) => {
          if (item.taskUrl) {
            return (
              <div data-selection-disabled={true}>
                <a
                  href={item.taskUrl}
                  target="_blank"
                  data-interception="off"
                  rel="noopener noreferrer"
                  className={styles.taskNameLink}
                  title={item.name}
                  onClick={(e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    window.open(item.taskUrl, '_blank', 'noopener,noreferrer');
                  }}
                >
                  {item.name}
                </a>
              </div>
            );
          }
          return <Text className={styles.taskName} title={item.name}>{item.name}</Text>;
        }
      },
      {
        key: 'workflowName',
        name: 'Workflow',
        fieldName: 'workflowName',
        minWidth: 120,
        maxWidth: 220,
        isResizable: true,
        isSorted: this.state?.sortColumn === 'workflowName',
        isSortedDescending: this.state?.sortColumn === 'workflowName' ? this.state?.sortDescending : false,
        onColumnClick: this._onColumnClick.bind(this),
        onRender: (item: INintexTask) => (
          <Text className={styles.cellText} title={item.workflowName}>{item.workflowName || '—'}</Text>
        )
      },
      {
        key: 'status',
        name: 'Status',
        fieldName: 'status',
        minWidth: 80,
        maxWidth: 120,
        isResizable: true,
        isSorted: this.state?.sortColumn === 'status',
        isSortedDescending: this.state?.sortColumn === 'status' ? this.state?.sortDescending : false,
        onColumnClick: this._onColumnClick.bind(this),
        onRender: (item: INintexTask) => {
          const statusClass = `status${item.status?.charAt(0).toUpperCase()}${item.status?.slice(1).toLowerCase()}Outline`;
          return (
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
              <span className={`${styles.statusBadgeOutline} ${(styles as Record<string, string>)[statusClass] || ''}`}>
                {item.status || '—'}
              </span>
            </Stack>
          );
        }
      },
      {
        key: 'createdDate',
        name: 'Assigned',
        fieldName: 'createdDate',
        minWidth: 120,
        maxWidth: 180,
        isResizable: true,
        isSorted: this.state?.sortColumn === 'createdDate',
        isSortedDescending: this.state?.sortColumn === 'createdDate' ? this.state?.sortDescending : false,
        onColumnClick: this._onColumnClick.bind(this),
        onRender: (item: INintexTask) => (
          <Text className={styles.cellText}>{this._formatDateTime(item.createdDate)}</Text>
        )
      },
      {
        key: 'completedDate',
        name: 'Completed',
        fieldName: 'completedDate',
        minWidth: 120,
        maxWidth: 180,
        isResizable: true,
        isSorted: this.state?.sortColumn === 'completedDate',
        isSortedDescending: this.state?.sortColumn === 'completedDate' ? this.state?.sortDescending : false,
        onColumnClick: this._onColumnClick.bind(this),
        onRender: (item: INintexTask) => (
          <Text className={styles.cellText}>{this._formatDateTime(item.completedDate)}</Text>
        )
      },
      {
        key: 'completedBy',
        name: 'Completed by',
        fieldName: 'completedBy',
        minWidth: 120,
        maxWidth: 200,
        isResizable: true,
        isSorted: this.state?.sortColumn === 'completedBy',
        isSortedDescending: this.state?.sortColumn === 'completedBy' ? this.state?.sortDescending : false,
        onColumnClick: this._onColumnClick.bind(this),
        onRender: (item: INintexTask) => (
          <Text className={styles.cellText}>{item.completedBy || '-'}</Text>
        )
      }
    ];
  }

  private _formatDateTime(dateStr: string | undefined): string {
    if (!dateStr || dateStr === '-') return '-';
    try {
      const date = new Date(dateStr);
      if (isNaN(date.getTime())) return dateStr;
      return date.toLocaleString('en-US', {
        year: 'numeric',
        month: 'short',
        day: 'numeric',
        hour: 'numeric',
        minute: '2-digit'
      });
    } catch {
      return dateStr;
    }
  }



  public componentDidMount(): void {
    this._loadTasks().catch(() => { /* handled in _loadTasks */ });
  }

  public componentDidUpdate(prevProps: INintexTasksProps): void {
    if (
      prevProps.tenantUrl !== this.props.tenantUrl ||
      prevProps.tokenListUrl !== this.props.tokenListUrl ||
      prevProps.tokenTitleFilter !== this.props.tokenTitleFilter ||
      prevProps.tokenColumnName !== this.props.tokenColumnName ||
      prevProps.dashboardUrl !== this.props.dashboardUrl
    ) {
      this._loadTasks().catch(() => { /* handled in _loadTasks */ });
    }
  }

  private async _loadTasks(): Promise<void> {
    const { tenantUrl, tokenListUrl, dashboardUrl, httpClient, spHttpClient, userEmail } = this.props;

    if (!tenantUrl || !tokenListUrl || !dashboardUrl) {
      this.setState({
        error: 'Please configure the Nintex API Base URL, Dashboard URL, and Token List URL in the web part properties.',
        isLoading: false,
        tasks: [],
        filteredTasks: []
      });
      return;
    }

    this.setState({ isLoading: true, error: undefined });

    try {
      const tokenService = new TokenService(
        spHttpClient,
        tokenListUrl,
        this.props.tokenTitleFilter,
        this.props.tokenColumnName
      );
      const token = await tokenService.getToken();

      const service = new NintexApiService(httpClient, tenantUrl, dashboardUrl, token, userEmail);

      let fromStr: string | undefined;
      let toStr: string | undefined;

      if (this.state.dateFilterPreset !== 'all') {
        const now = new Date();
        if (this.state.dateFilterPreset === '30') {
          fromStr = new Date(now.setDate(now.getDate() - 30)).toISOString();
        } else if (this.state.dateFilterPreset === '60') {
          fromStr = new Date(now.setDate(now.getDate() - 60)).toISOString();
        } else if (this.state.dateFilterPreset === '90') {
          fromStr = new Date(now.setDate(now.getDate() - 90)).toISOString();
        } else if (this.state.dateFilterPreset === 'custom') {
          if (this.state.customDateStart) fromStr = this.state.customDateStart.toISOString();
          if (this.state.customDateEnd) {
            const endDate = new Date(this.state.customDateEnd);
            endDate.setHours(23, 59, 59, 999); // Set to end of the day
            toStr = endDate.toISOString();
          }
        }
      }

      const response = await service.listTasks({
        status: this.state.statusFilter,
        assignee: userEmail,
        from: fromStr,
        to: toStr
      });

      this.setState({
        tasks: response.tasks,
        filteredTasks: this._applyFilters(response.tasks, this.state.searchText),
        isLoading: false,
        error: undefined
      });
    } catch (err) {
      try {
        const tokenService = new TokenService(
          spHttpClient,
          tokenListUrl,
          this.props.tokenTitleFilter,
          this.props.tokenColumnName
        );
        tokenService.clearCache();
      } catch { /* best effort */ }

      this.setState({
        error: `Failed to load tasks: ${err instanceof Error ? err.message : String(err)}`,
        isLoading: false,
        tasks: [],
        filteredTasks: []
      });
    }
  }

  private _applyFilters(tasks: INintexTask[], searchText: string): INintexTask[] {
    let filtered = [...tasks];

    if (searchText) {
      const lowerSearch = searchText.toLowerCase();
      filtered = filtered.filter(task =>
        (task.name || '').toLowerCase().indexOf(lowerSearch) >= 0 ||
        (task.workflowName || '').toLowerCase().indexOf(lowerSearch) >= 0 ||
        (task.assignee || '').toLowerCase().indexOf(lowerSearch) >= 0 ||
        (task.description || '').toLowerCase().indexOf(lowerSearch) >= 0
      );
    }

    if (this.state.dateFilterPreset) {
      const now = new Date();
      let cutoffDate: Date | undefined;

      if (this.state.dateFilterPreset === '30') {
        cutoffDate = new Date(now.setDate(now.getDate() - 30));
      } else if (this.state.dateFilterPreset === '60') {
        cutoffDate = new Date(now.setDate(now.getDate() - 60));
      } else if (this.state.dateFilterPreset === '90') {
        cutoffDate = new Date(now.setDate(now.getDate() - 90));
      }

      const getDateForTask = (task: INintexTask): string | undefined => {
        return this.state.dateType === 'completed' ? task.completedDate : task.createdDate;
      };

      if (cutoffDate) {
        filtered = filtered.filter(task => {
          const taskDate = getDateForTask(task);
          if (!taskDate) return false;
          return new Date(taskDate) >= cutoffDate!;
        });
      } else if (this.state.dateFilterPreset === 'custom') {
        filtered = filtered.filter(task => {
          const taskDate = getDateForTask(task);
          if (!taskDate) return true;
          const tDate = new Date(taskDate).getTime();
          const start = this.state.customDateStart ? this.state.customDateStart.getTime() : 0;
          // Set end of day for the end date
          let end = Number.MAX_SAFE_INTEGER;
          if (this.state.customDateEnd) {
            const endD = new Date(this.state.customDateEnd);
            endD.setHours(23, 59, 59, 999);
            end = endD.getTime();
          }
          return tDate >= start && tDate <= end;
        });
      }
    }

    filtered = this._sortTasks(filtered, this.state.sortColumn, this.state.sortDescending);
    return filtered;
  }

  private _sortTasks(tasks: INintexTask[], column: string, descending: boolean): INintexTask[] {
    return tasks.sort((a, b) => {
      const aVal = ((a as unknown) as Record<string, string>)[column] || '';
      const bVal = ((b as unknown) as Record<string, string>)[column] || '';

      if (column === 'createdDate' || column === 'dueDate') {
        const aDate = aVal ? new Date(aVal).getTime() : 0;
        const bDate = bVal ? new Date(bVal).getTime() : 0;
        return descending ? bDate - aDate : aDate - bDate;
      }

      const comparison = aVal.toString().localeCompare(bVal.toString(), undefined, { sensitivity: 'base' });
      return descending ? -comparison : comparison;
    });
  }

  private _onColumnClick = (_ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const newSortDescending = this.state.sortColumn === column.fieldName ? !this.state.sortDescending : false;

    this.setState({
      sortColumn: column.fieldName || 'name',
      sortDescending: newSortDescending
    }, () => {
      this._columns = this._buildColumns();
      const filtered = this._applyFilters(this.state.tasks, this.state.searchText);
      this.setState({ filteredTasks: filtered });
    });
  }

  private _onSearchChange = (_ev?: React.ChangeEvent<HTMLInputElement>, newValue?: string): void => {
    const searchText = newValue || '';
    this.setState({
      searchText,
      filteredTasks: this._applyFilters(this.state.tasks, searchText)
    });
  }

  private _onDateTypeChange = (_ev: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (option) {
      this.setState({
        dateType: option.key as 'assigned' | 'completed'
      }, this._applyFilterState);
    }
  }

  private _onStatusFilterChange = (_ev: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (option) {
      const newStatus = option.key as TaskStatus;
      let newDateType: 'assigned' | 'completed' = 'assigned';
      
      if (newStatus === 'complete' || newStatus === 'expired' || newStatus === 'terminated') {
        newDateType = 'completed';
      }

      this.setState({
        statusFilter: newStatus,
        dateType: newDateType
      }, () => {
        this._loadTasks().catch(() => { /* handled in _loadTasks */ });
      });
    }
  }

  private _toggleDateFilter = (): void => {
    this.setState({ isDateFilterOpen: !this.state.isDateFilterOpen });
  }

  private _onApplyPreset = (key: '30' | '60' | '90'): void => {
    this.setState({
      dateFilterPreset: key,
      isDateFilterOpen: false
    }, () => {
      this._loadTasks().catch(() => {});
    });
  }

  private _applyFilterState = (): void => {
    this._loadTasks().catch(() => {});
  }

  private _getDateFilterDisplayText(): string {
    if (this.state.dateFilterPreset === 'custom') {
      if (this.state.customDateStart && this.state.customDateEnd) {
        return `${this.state.customDateStart.toLocaleDateString()} - ${this.state.customDateEnd.toLocaleDateString()}`;
      } else if (this.state.customDateStart) {
        return `Since ${this.state.customDateStart.toLocaleDateString()}`;
      } else if (this.state.customDateEnd) {
        return `Until ${this.state.customDateEnd.toLocaleDateString()}`;
      }
      return 'Custom Range';
    }

    switch (this.state.dateFilterPreset) {
      case '30': return 'In the past 30 days';
      case '60': return 'In the past 60 days';
      case '90': return 'In the past 90 days';
      default: return 'All Tasks';
    }
  }

  private _onRenderRow = (props?: IDetailsRowProps): JSX.Element | null => {
    if (!props) return null;
    return (
      <div 
        onClick={() => this._onRowClick(props.item as INintexTask)}
        style={{ cursor: 'pointer' }}
      >
        <DetailsRow {...props} />
      </div>
    );
  };

  private _onRowClick = (item: INintexTask): void => {
    this.setState({
      selectedTask: item,
      isPanelOpen: true
    });
  }

  private _onPanelDismiss = (): void => {
    this.setState({
      isPanelOpen: false,
      selectedTask: undefined
    });
  }

  private _onRefresh = (): void => {
    this._loadTasks().catch(() => { /* handled in _loadTasks */ });
  }

  private _statusOptions: IDropdownOption[] = [
    { key: 'all', text: 'All' },
    { key: 'active', text: 'Active' },
    { key: 'complete', text: 'Complete' },
    { key: 'expired', text: 'Expired' },
    { key: 'terminated', text: 'Terminated' }
  ];

  public render(): React.ReactElement<INintexTasksProps> {
    const {
      isLoading,
      error,
      filteredTasks,
      statusFilter,
      selectedTask,
      isPanelOpen,
      searchText
    } = this.state;

    const { hasTeamsContext } = this.props;

    return (
      <section className={`${styles.nintexTasks} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.container}>
          <div className={styles.headerContainer}>
            <h2 className={styles.headerTitle}>Assigned Tasks</h2>
          </div>

          <div className={styles.toolbarContainer}>
            {/* Toolbar Filters (Left) */}
            <Stack horizontal wrap tokens={{ childrenGap: 12 }} verticalAlign="center" className={styles.toolbarFilters}>
              <Dropdown
                selectedKey={statusFilter}
                options={this._statusOptions}
                onChange={this._onStatusFilterChange}
                className={styles.statusDropdown}
                onRenderTitle={(options) => 
                  <span><span style={{ color: '#8a8886' }}>Status: </span>{options && options[0].text}</span>
                }
              />
              <Dropdown
                selectedKey={this.state.dateType}
                options={statusFilter === 'active' ? 
                  [{ key: 'assigned', text: 'Assigned' }] : 
                  [{ key: 'assigned', text: 'Assigned' }, { key: 'completed', text: 'Completed' }]
                }
                onChange={this._onDateTypeChange}
                className={styles.dateDropdown}
                onRenderTitle={(options) => 
                  <span><span style={{ color: '#8a8886' }}>Date: </span>{options && options[0].text}</span>
                }
              />
              <div 
                className={styles.dateRangePickerContainer}
                onClick={this._toggleDateFilter}
                ref={this._dateFilterButton}
                style={{ cursor: 'pointer' }}
              >
                <span className={styles.dateRangeText}>{this._getDateFilterDisplayText()}</span>
                <Icon iconName="Calendar" className={styles.calendarIcon} />
              </div>
            </Stack>
            
            <div className={styles.spacer} />

            {/* Toolbar Search (Right) */}
            <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="center" className={styles.searchStack}>
              <span className={styles.searchLabel}>Search task name:</span>
              <SearchBox
                placeholder=""
                value={searchText}
                onChange={this._onSearchChange}
                className={styles.searchBox}
                iconProps={{ iconName: 'Search' }}
              />
              <button className={styles.refreshButton} onClick={this._onRefresh} title="Refresh">
                <Icon iconName="Refresh" />
              </button>
            </Stack>
          </div>

          {/* Error */}
          {error !== undefined && (
            <MessageBar
              messageBarType={MessageBarType.error}
              isMultiline={false}
              onDismiss={() => this.setState({ error: undefined })}
              dismissButtonAriaLabel="Close"
              className={styles.errorBar}
            >
              {error}
            </MessageBar>
          )}

          {/* Loading */}
          {isLoading && (
            <div className={styles.loadingContainer}>
              <Spinner size={SpinnerSize.large} label="Loading tasks..." />
            </div>
          )}

          {/* Empty State */}
          {!isLoading && !error && filteredTasks.length === 0 && (
            <div className={styles.emptyState}>
              <Icon iconName="TaskSolid" className={styles.emptyIcon} />
              <Text variant="large">No tasks found</Text>
              <Text variant="small" className={styles.emptySubtext}>
                {searchText ? 'Try adjusting your search or filters' : 'There are no tasks matching the current filter'}
              </Text>
            </div>
          )}

          {/* Task List */}
          {!isLoading && filteredTasks.length > 0 && (
            <div className={styles.listContainer}>
              <DetailsList
                items={filteredTasks}
                columns={this._columns}
                layoutMode={DetailsListLayoutMode.justified}
                selectionMode={SelectionMode.none}
                checkboxVisibility={CheckboxVisibility.hidden}
                onRenderRow={this._onRenderRow}
                onItemInvoked={this._onRowClick}
                className={styles.taskList}
                isHeaderVisible={true}
              />
            </div>
          )}

          {/* Detail Panel */}
          <TaskDetailPanel
            task={selectedTask}
            isOpen={isPanelOpen}
            onDismiss={this._onPanelDismiss}
          />

          {/* Date Range Callout */}
          {this.state.isDateFilterOpen && (
            <Callout
              target={this._dateFilterButton.current}
              onDismiss={this._toggleDateFilter}
              setInitialFocus
              directionalHint={DirectionalHint.bottomLeftEdge}
              className={styles.dateFilterCallout}
            >
              <div className={styles.segmentedControl}>
                <button 
                  className={`${styles.segmentButton} ${this.state.dateFilterTab === 'Preset' ? styles.segmentActive : ''}`}
                  onClick={() => this.setState({ dateFilterTab: 'Preset' })}
                >
                  Preset
                </button>
                <button 
                  className={`${styles.segmentButton} ${this.state.dateFilterTab === 'Custom' ? styles.segmentActive : ''}`}
                  onClick={() => this.setState({ dateFilterTab: 'Custom' })}
                >
                  Custom
                </button>
              </div>

              {this.state.dateFilterTab === 'Preset' && (
                <div className={styles.presetList}>
                  {[
                    { key: '30', text: 'In the past 30 days' },
                    { key: '60', text: 'In the past 60 days' },
                    { key: '90', text: 'In the past 90 days' }
                  ].map(opt => (
                    <div 
                      key={opt.key}
                      className={`${styles.presetOption} ${this.state.dateFilterPreset === opt.key ? styles.presetSelected : ''}`}
                      onClick={() => this._onApplyPreset(opt.key as '30' | '60' | '90')}
                    >
                      {opt.text}
                    </div>
                  ))}
                </div>
              )}

              {this.state.dateFilterTab === 'Custom' && (
                <div className={styles.customDateContainer}>
                  <DatePicker 
                    label="Start Date" 
                    value={this.state.customDateStart}
                    maxDate={new Date()}
                    onSelectDate={(d) => {
                      this.setState({ 
                        customDateStart: d || undefined,
                        dateFilterPreset: 'custom' 
                      }, this._applyFilterState);
                    }}
                  />
                  <DatePicker 
                    label="End Date" 
                    value={this.state.customDateEnd}
                    maxDate={new Date()}
                    onSelectDate={(d) => {
                      this.setState({ 
                        customDateEnd: d || undefined,
                        dateFilterPreset: 'custom'
                      }, this._applyFilterState);
                    }}
                  />
                </div>
              )}
            </Callout>
          )}
        </div>
      </section>
    );
  }
}
