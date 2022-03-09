export type LoadingStatus = 'loading' | 'loaded' | 'loadNoData' | 'loadError'

export interface IUserAttendanceData {
  logDateTime: string;
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
  filterLocation: string;
}
