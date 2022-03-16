import * as React from 'react';
import { IUserSummaryDashboardProps } from './IUserSummaryDashboardProps';
import { IUserAttendanceData, IUserSummaryDashboardState } from './IUserSummaryDashboardState';
import axios from 'Axios';
import { DateRange, Error } from '@material-ui/icons';
import { Tab, Tabs, TabList, TabPanel } from 'react-tabs';
import * as moment from 'moment';
import classnames from 'classnames';

import 'react-tabs/style/react-tabs.css';
import 'bootstrap-css-only/css/bootstrap.min.css';
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
        this._absoluteUrl + `/_api/web/lists/getbytitle('Employee Card Record')/items?` +
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
        <div style={{height: '170px'}}>
          <div className='row'>
            <div className={classnames('col-12', styles.noPadding)}>
              <Tabs className={styles.tabs}>
                <TabList>
                  <Tab>Leave</Tab>
                  <Tab>Lunch</Tab>
                  <Tab>Attendance</Tab>
                  <Tab>Snack</Tab>
                </TabList>

                <TabPanel>
                  <div className='col-12'>
                    {userAnnualLeave.loadingStatus === 'loading' &&
                      <div className={styles.colCenter}>Loading...</div>
                    }
                    {userAnnualLeave.loadingStatus === 'loadNoData' &&
                      <div className={styles.colCenter}>-No Data-</div>
                    }
                    {userAnnualLeave.loadingStatus === 'loadError' &&
                      <div className={styles.colCenter}>
                        <Error style={{color: 'slategrey'}}/>
                        <div>Oops, Something went wrong</div>
                      </div>
                    }
                    {userAnnualLeave.loadingStatus === 'loaded' &&
                      <a className={styles.link} href={this.props.leaveLink}>
                        <div className='row'>
                          <div className='col-6'>已取年假</div>
                          <div className={classnames('col-6', styles.colCenter)}>
                            {userAnnualLeave.data.annualTaken}天
                          </div>
                          <div className='col-12'>
                            <div className={styles.divider} />
                          </div>
                        </div>

                        <div className='row'>
                          <div className='col-6'>年假剩餘</div>
                          <div className={classnames('col-6', styles.colCenter)}>
                            {userAnnualLeave.data.annualRemaining}天
                          </div>
                        </div>
                      </a>
                    }
                  </div>
                </TabPanel>
                <TabPanel>
                  <div className={classnames('col-12', styles.colCenter)}>
                    <h2>Lunch</h2>
                  </div>
                </TabPanel>

                <TabPanel>
                  <div className={classnames('col-12', styles.colCenter)}>
                    {userAttendance.loadingStatus === 'loading' &&
                      <div>Loading...</div>
                    }
                    {userAttendance.loadingStatus === 'loadNoData' &&
                      <div>-No Data-</div>
                    }
                    {userAttendance.loadingStatus === 'loadError' &&
                      <div>
                        <Error style={{color: 'slategrey'}}/> Something went wrong
                      </div>
                    }
                    {userAttendance.loadingStatus === 'loaded' &&
                      <div className='row'>
                        <div className={styles.colCenter}>
                          <a href={this.props.attendanceLink}>
                            <DateRange id={styles.largeIcon} />
                          </a>
                        </div>
                        <div style={{marginLeft: '5px'}}>
                          { userAttendance.data.map((item: IUserAttendanceData) => {
                              return (
                                <div className={styles.attendanceRow}>
                                  {this.formatDateTime(item.logDatetime)} - {item.logLocation}
                                </div>
                              );
                          })}
                        </div>
                      </div>
                    }
                  </div>
                </TabPanel>

                <TabPanel>
                  <div className={classnames('col-12', styles.colCenter)}>
                    <h2>Snack</h2>
                  </div>
                </TabPanel>
              </Tabs>
            </div>
          </div>
        </div>
      </section>
    );
  }
}
