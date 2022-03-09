import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IUserSummaryDashboardProps {
  description: string;
  leaveLink: string;
  attendanceLink: string;
  context: WebPartContext;
}
