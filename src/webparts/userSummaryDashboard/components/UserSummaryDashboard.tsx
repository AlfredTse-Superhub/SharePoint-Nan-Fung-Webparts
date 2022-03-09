import * as React from 'react';
import { IUserSummaryDashboardProps } from './IUserSummaryDashboardProps';
import { IUserAttendanceData, IUserSummaryDashboardState } from './IUserSummaryDashboardState';
import { escape } from '@microsoft/sp-lodash-subset';
import axios from 'Axios';
import { Grid } from '@material-ui/core';
import { DateRange, Error } from '@material-ui/icons';
import { Tab, Tabs, TabList, TabPanel } from 'react-tabs';
import * as moment from 'moment';

import 'react-tabs/style/react-tabs.css';
import styles from './UserSummaryDashboard.module.scss';



export default class UserSummaryDashboard extends React.Component<IUserSummaryDashboardProps, IUserSummaryDashboardState> {
  private _userEmail: string = this.props.context.pageContext.legacyPageContext.userEmail;
  private _absoluteUrl: string = this.props.context.pageContext.web.absoluteUrl;

  constructor(props) {
    super(props);

    this.state = {
      userCardNo: '',
      userAnnualLeave: {
        loadingStatus: 'loading',
        data: {
          annualTaken: 0, 
          annualRemaining: 0
        }
      },
      userAttendance: {
        loadingStatus: 'loading',
        data: []
      },
    };
  }

  public async componentDidMount() {
    await this.getData();
  }

  private async getData() {
    await this.getUserInfo();
    this.getUserAnnualLeave();
    this.getUserAttendance();
  }

  private async getUserInfo() {
    try {
      const response = await axios.get(
        this._absoluteUrl + `/_api/web/lists/getbytitle('employee')/items?` +
          `$select=CARD_NO &` +
          `$filter=EMAIL eq '${this._userEmail}'`
      );
      this.setState({
        userCardNo: response.data.value[0]['CARD_NO']
      })

    } catch (error) {
      console.error(error);
    }
  }

  private async getUserAnnualLeave() {
    try {
      const response = await axios.get(
        this._absoluteUrl + `/_api/web/lists/getbytitle('LEAVEANU')/items?` +
          `$filter=CARD_NO eq '${this.state.userCardNo}'`
      );
      if (response.data.value.length > 0) {
        this.setState({
          userAnnualLeave: {
            loadingStatus: 'loaded',
            data: {
              annualTaken: response.data.value[0]['ANNUAL'],
              annualRemaining: response.data.value[0]['ANU_TOTAL'] - response.data.value[0]['ANNUAL']
            }
          },
        });
      } else {
        this.setState({userAnnualLeave: {loadingStatus: 'loadNoData' , data: null}});
      }

    } catch (error) {
      console.error(error);
      this.setState({userAnnualLeave: {loadingStatus: 'loadError' , data: null}});
    }
  }

  private async getUserAttendance() {
    try {
      const response = await axios.get(
        this._absoluteUrl + `/_api/web/lists/getbytitle('attendance')/items?` +
          `$filter=H_CARD_NO eq '${this.state.userCardNo}' &` +
          `$orderby=H_LOG_DATETIME desc &` +
          `$top=5`
      );
      if (response.data.value.length > 0) {
        let attendanceData: IUserAttendanceData[] = [];
        response.data.value.map((item: IUserAttendanceData) => {
          attendanceData.push({
            logDatetime: item['H_LOG_DATETIME'],
            logLocation: item['H_LOCATION_ID']
          });
        });
        this.setState({
          userAttendance: {
            loadingStatus: 'loaded',
            data: attendanceData
          }
        })
      } else {
        this.setState({userAttendance: {loadingStatus: 'loadNoData' , data: null}});
      }

    } catch (error) {
      console.error(error);
      this.setState({userAttendance: {loadingStatus: 'loadError' , data: null}});
    }
  }

  private formatDateTime(dateTime: string): string {
    return moment(dateTime).format('MMMM Do, h:mm a');
  }

  public render(): React.ReactElement<IUserSummaryDashboardProps> {
    const { userAnnualLeave, userAttendance } = this.state;

    return (
      <section className={styles.userSummaryDashboard}>
        <div className={styles.webpartContent}>
          <Tabs>
            <TabList>
              <Tab style={{padding: '6px 7px'}}>Leave</Tab>
              <Tab style={{padding: '6px 7px'}}>Lunch</Tab>
              <Tab style={{padding: '6px 7px'}}>Attendance</Tab>
              <Tab style={{padding: '6px 7px'}}>Snack</Tab>
            </TabList>

            <TabPanel>
              {userAnnualLeave.loadingStatus === 'loading' &&
                <div>Loading...</div>
              }
              {userAnnualLeave.loadingStatus === 'loadNoData' &&
                <div>-No Data-</div>
              }
              {userAnnualLeave.loadingStatus === 'loadError' &&
                <div>
                  <Error style={{color: 'slategrey'}}/>
                  <div>Oops, Something went wrong</div>
                </div>
              }
              {userAnnualLeave.loadingStatus === 'loaded' &&
                <a className={styles.link} href={this.props.leaveLink}>
                  <div>
                    <Grid container>
                      <Grid item sm={6} md={6}>已取年假</Grid>
                      <Grid item sm={6} md={6}>{userAnnualLeave.data.annualTaken}天</Grid>
                    </Grid>
                    <div className={styles.divider}></div>
                    <Grid container>
                      <Grid item sm={6} md={6}>年假剩餘</Grid>
                      <Grid item sm={6} md={6}>{userAnnualLeave.data.annualRemaining}天</Grid>
                    </Grid>
                  </div>
                </a>
              }
            </TabPanel>
            <TabPanel>
              <h2>Lunch</h2>
            </TabPanel>

            <TabPanel>
              <Grid container>
                <Grid item sm={3} md={3}>
                  <a href={this.props.attendanceLink}>
                    <DateRange id={styles.largeIcon} />
                  </a>
                </Grid>
                <Grid item sm={9} md={9}>
                  {userAttendance.loadingStatus === 'loading' &&
                    <div className={styles.attendanceListEmpty}>Loading...</div>
                  }
                  {userAttendance.loadingStatus === 'loadNoData' &&
                    <div className={styles.attendanceListEmpty}>-No Data-</div>
                  }
                  {userAttendance.loadingStatus === 'loadError' &&
                    <div className={styles.attendanceListEmpty}>
                      <Error style={{color: 'slategrey'}}/>
                      <div>Oops, Something went wrong</div>
                    </div>
                  }
                  {userAttendance.loadingStatus === 'loaded' &&
                    <div>
                      { userAttendance.data.map((item: IUserAttendanceData) => {
                          return (
                            <div className={styles.attendanceRow}>
                              {this.formatDateTime(item.logDatetime)} - {item.logLocation}
                            </div>
                          );
                      })}
                    </div>
                  }
                </Grid>
              </Grid>
            </TabPanel>

            <TabPanel>
              <h2>Snack</h2>
            </TabPanel>
          </Tabs>
        </div>
      </section>
    );
  }
}
