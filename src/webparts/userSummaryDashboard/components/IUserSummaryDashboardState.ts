export type LoadingStatus = 'loading' | 'loaded' | 'loadNoData' | 'loadError'

export interface IUserAnnualLeaveData {
  annualTaken: number;
  annualRemaining: number;
}

export interface IUserAttendanceData {
  logDatetime: string;
  logLocation: string;
}

export interface IUserAnnualLeave {
  loadingStatus: LoadingStatus;
  data: IUserAnnualLeaveData;
}

export interface IUserAttendance {
  loadingStatus: LoadingStatus;
  data: IUserAttendanceData[];
}

export interface IUserSummaryDashboardState {
  userCardNo: string;
  userAnnualLeave: IUserAnnualLeave;
  userAttendance: IUserAttendance;
}
