export type LoadingStatus = 'loading' | 'loaded' | 'loadNoData' | 'loadError'

export interface IUserAttendanceData {
  logDate: string;
  logTime: string;
  logLocation: string;
}

export interface IUserAttendance {
  loadingStatus: LoadingStatus;
  data: IUserAttendanceData[];
}

export interface IAttendanceRecordsDashboardState {
  userCardNo: string;
  userAttendance: IUserAttendance;
  filterDateFrom: Date;
  filterDateTo: Date
}
