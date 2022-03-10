import * as React from 'react';
import { IAttendanceRecordsDashboardProps } from './IAttendanceRecordsDashboardProps';
import { IUserAttendanceData, IAttendanceRecordsDashboardState } from './IAttendanceRecordsDashboardState';
import axios from 'Axios';
import { Clear, Error, Search } from '@material-ui/icons';
import { Grid } from '@material-ui/core';
import { Dropdown, IDropdownOption, IconButton } from 'office-ui-fabric-react';
import DatePicker from "react-datepicker";
import { MDBDataTable } from 'mdbreact';
import { isNull } from 'lodash';
import * as moment from 'moment';

import 'bootstrap-css-only/css/bootstrap.min.css';
import 'mdbreact/dist/css/mdb.css';
import 'react-datepicker/dist/react-datepicker.css';
import styles from './AttendanceRecordsDashboard.module.scss';



export default class AttendanceRecordsDashboard extends React.Component<IAttendanceRecordsDashboardProps, IAttendanceRecordsDashboardState> {
  private _userEmail: string = this.props.context.pageContext.user.email;
  private _absoluteUrl: string = this.props.context.pageContext.web.absoluteUrl;
  private _locationsOptions: IDropdownOption[] = [
    {key: 'default', text: '所有地點'},
    {key: '20main', text: '20 Main'},
    {key: '5Fmain', text: '5/F Main'}
  ]

  constructor(props: IAttendanceRecordsDashboardProps) {
    super(props);

    this.state = {
      userCardNo: '',
      userAttendance: {
        loadingStatus: 'loading',
        data: []
      },
      filterDateFrom: null,
      filterDateTo: null,
      filterLocation: '所有地點',
    };
  }

  public async componentDidMount(): Promise<void> {
    this.getData();
  }

  private async getData(): Promise<void> {
    await this.getUserInfo();
    this.getUserAttendance();
  }

  private async getUserInfo(): Promise<void> {
    try {
      const response = await axios.get(
        this._absoluteUrl + `/_api/web/lists/getbytitle('employee')/items?` +
          `$select=CARD_NO &` +
          `$filter=EMAIL eq '${this._userEmail}'`
      );
      this.setState({userCardNo: response.data.value[0]['CARD_NO']});

    } catch (error) {
      console.error(error);
    }
  }

  private async getUserAttendance(
    dateFrom: Date = null,
    dateTo: Date = null,
    location: string = '所有地點'
  ): Promise<void> {
    try {
      let dateFromCondition = ``;
      let dateToCondition = ``;
      let locationCondition = ``;
      if (!isNull(dateFrom)) { dateFromCondition = `and H_LOG_DATETIME ge datetime'${dateFrom.toISOString()}'`; }
      if (!isNull(dateTo)) { dateToCondition = `and H_LOG_DATETIME le datetime'${dateTo.toISOString()}'`; }
      if (location != '所有地點') { locationCondition = `and H_LOCATION_ID eq '${location}'`; }
        
      const response = await axios.get(
        this._absoluteUrl + `/_api/web/lists/getbytitle('attendance')/items?` +
          `$orderby=H_LOG_DATETIME desc &` +
          `$filter=H_CARD_NO eq '${this.state.userCardNo}'` +
          dateFromCondition +
          dateToCondition +
          locationCondition
      );
      if (response.data.value.length > 0) {
        let attendanceData: IUserAttendanceData[] = [];
        response.data.value.map((item: IUserAttendanceData) => {
          attendanceData.push({
            logDate: this.formatDate(item['H_LOG_DATETIME']),
            logTime: this.formatTime(item['H_LOG_DATETIME']),
            logLocation: item['H_LOCATION_ID']
          });
        });
        this.setState({
          userAttendance: {
            loadingStatus: 'loaded',
            data: attendanceData
          },
        });
      } else {
        this.setState({userAttendance: {loadingStatus: 'loadNoData', data: null}});
      }

    } catch (error) {
      console.error(error);
      this.setState({userAttendance: {loadingStatus: 'loadError', data: null}});
    }
  }

  private applyFilter(): void {
    this.getUserAttendance(this.state.filterDateFrom, this.state.filterDateTo, this.state.filterLocation);
  }

  private resetFilter(): void {
    this.setState({
      filterDateFrom: null,
      filterDateTo: null,
      filterLocation: '所有地點'
    })
  }

  private getDropdownSelectedKey(): number | string {
    return (this._locationsOptions.filter(item => item.text === this.state.filterLocation)[0].key);
  }

  private formatDate(dateTime: string): string {
    return moment(dateTime).format('yyyy-MM-DD');
  }

  private formatTime(dateTime: string): string {
    return moment(dateTime).format('h:mm a');
  }
  
  public render(): React.ReactElement<IAttendanceRecordsDashboardProps> {
    const { userAttendance, filterDateFrom, filterDateTo, filterLocation } = this.state;
    const attendanceTableColumns = [
      {
        label: '日期',
        field: 'logDate',
        sort: 'asc',
      },
      {
        label: '時間',
        field: 'logTime',
        sort: 'asc',
      },
      {
        label: '地點',
        field: 'logLocation',
        sort: 'asc',
      },
    ];

    return (
      <section className={styles.attendanceRecordsDashboard}>
        {/* Header section */}
        <div className={styles.attendanceTitleBar}>
          考勤查詢
        </div>
        <div className={styles.divider} />

        {/* Filter section */}
        <div className={styles.filterSection}>
          <Grid container>
            <Grid item xs={12} sm={3} md={3}>開始日期</Grid>
            <Grid item xs={12} sm={3} md={3}>結束日期</Grid>
            <Grid item xs={12} sm={3} md={3}>地點 </Grid>
          </Grid>
          <div className={styles.box5px}/>
          <Grid container>
            <Grid item xs={12} sm={3} md={3}>
              <DatePicker
                placeholderText='Select date from'
                className={styles.datePicker}
                value={isNull(filterDateFrom)
                        ? null 
                        : this.formatDate(filterDateFrom.toISOString())}
                onChange={(date) => {
                  this.setState({
                    filterDateFrom: date
                  });
                }}
              />
            </Grid>
            <Grid item xs={12} sm={3} md={3}>
              <DatePicker
                placeholderText='Select date to'
                className={styles.datePicker}
                value={isNull(filterDateTo) 
                        ? null 
                        : moment(filterDateTo).format('yyyy-MM-DD')}
                onChange={(date) => {
                  this.setState({
                    filterDateTo: date
                  })
                }}
              />
            </Grid>
            <Grid item xs={12} sm={3} md={3}>
              <Dropdown
                className={styles.dropdown}
                placeholder='Select Location'
                options={this._locationsOptions}
                selectedKey={this.getDropdownSelectedKey()}
                onChange={(onChange, option) => {
                  this.setState({
                    filterLocation: option.text
                  })
                }}
              />
            </Grid>
            <Grid item xs={12} sm={3} md={3}>
              <IconButton className={styles.iconButton} onClick={() => this.applyFilter()}><Search /></IconButton>
              <IconButton className={styles.iconButton} onClick={() => this.resetFilter()}><Clear /></IconButton>
            </Grid>
          </Grid>
        </div>

        {/* Attendance list section */}
        <div className={styles.attendanceList}>
          {userAttendance.loadingStatus === 'loading' &&
            <div className={styles.attendanceListEmpty}>
              Loading...
            </div>
          }
          {userAttendance.loadingStatus === 'loadNoData' &&
            <div className={styles.attendanceListEmpty}>
              -No Data-
            </div>
          }
          {userAttendance.loadingStatus === 'loadError' &&
            <div className={styles.attendanceListEmpty}>
              <Error style={{color: 'slategrey'}}/>
              <div>Oops, Something went wrong</div>
            </div>
          }
          {userAttendance.loadingStatus === 'loaded' &&
            <div>
              <MDBDataTable
                striped
                bordered
                small
                noBottomColumns
                sortable={false}
                className={styles.attendanceTable}
                data={{
                  columns: attendanceTableColumns,
                  rows: userAttendance.data
                }}
              />
            </div>
          }
        </div>
      </section>
    );
  }
}
