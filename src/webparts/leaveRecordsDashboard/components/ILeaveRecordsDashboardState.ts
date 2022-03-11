export type LoadingStatus = 'loading' | 'loaded' | 'loadNoData' | 'loadError'
export type ContentView = 'calendar' | 'list'

// Data Model
export interface ILeaveTypeData {
  typeID: string;
  typeNameEn: string;
  typeNameZh: string;
}

export interface IUserAnnualLeaveData {
  annualLeaveTotal: number;
  annualLeaveTaken: number;
  sickLeaveTaken: number;
}

export interface IUserLeaveData {
  leaveTypeID: string;
  leaveDateFrom: string;
  leaveDateTo: string;
  daysCount: number;
}

// Data Object Model
export interface ILeaveType {
  loadingStatus: LoadingStatus;
  data: ILeaveTypeData[];
}

export interface IUserAnnualLeave {
  loadingStatus: LoadingStatus;
  data: IUserAnnualLeaveData;
}

export interface IUserLeave {
  loadingStatus: LoadingStatus;
  data: IUserLeaveData[];
}

export interface ILeaveRecordsDashboardState {
  isExpanded: boolean;
  contentView: ContentView;
  userCardNo: string;
  leaveTypes: ILeaveType;
  userAnnualLeave: IUserAnnualLeave;
  userLeaves: IUserLeave;
  filterYear: Date;
  filterLeaveType: string;
}
