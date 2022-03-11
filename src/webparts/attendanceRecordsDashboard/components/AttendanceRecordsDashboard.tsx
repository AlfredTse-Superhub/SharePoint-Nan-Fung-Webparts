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

import 'react-datepicker/dist/react-datepicker.css';
import styles from './AttendanceRecordsDashboard.module.scss';
import 'bootstrap-css-only/css/bootstrap.min.css';
import 'mdbreact/dist/css/mdb.css';



export default class AttendanceRecordsDashboard extends React.Component<IAttendanceRecordsDashboardProps, IAttendanceRecordsDashboardState> {
  private _userEmail: string = this.props.context.pageContext.user.email;
  private _absoluteUrl: string = this.props.context.pageContext.web.absoluteUrl;

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
    dateTo: Date = null
  ): Promise<void> {
    try {
      let dateFromCondition = ``;
      let dateToCondition = ``;
      let today = new Date();

      if (isNull(dateFrom))
        dateFrom = new Date(today.getFullYear(), today.getMonth(), 1);
      if (isNull(dateTo))
        dateTo = new Date(today.getFullYear(), today.getMonth() + 1, 1);

      dateFromCondition = `and H_LOG_DATETIME ge datetime'${dateFrom.toISOString()}'`;
      dateToCondition = `and H_LOG_DATETIME le datetime'${dateTo.toISOString()}'`;
        
      const response = await axios.get(
        this._absoluteUrl + `/_api/web/lists/getbytitle('attendance')/items?` +
          `$orderby=H_LOG_DATETIME desc &` +
          `$filter=H_CARD_NO eq '${this.state.userCardNo}'` +
          dateFromCondition +
          dateToCondition
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
        this.setState({userAttendance: {loadingStatus: 'loadNoData', data: []}});
      }

    } catch (error) {
      console.error(error);
      this.setState({userAttendance: {loadingStatus: 'loadError', data: null}});
    }
  }

  private applyFilter(): void {
    this.getUserAttendance(this.state.filterDateFrom, this.state.filterDateTo);
  }

  private resetFilter(): void {
    this.setState({
      filterDateFrom: null,
      filterDateTo: null
    })
  }

  private formatDate(dateTime: string): string {
    return moment(dateTime).format('yyyy-MM-DD');
  }

  private formatTime(dateTime: string): string {
    return moment(dateTime).format('h:mm a');
  }
  
  public render(): React.ReactElement<IAttendanceRecordsDashboardProps> {
    const { userAttendance, filterDateFrom, filterDateTo } = this.state;
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
          <Grid container spacing={1}>
            <Grid item xs={12} sm={3} md={3}>開始日期</Grid>
            <Grid item xs={12} sm={3} md={3}>結束日期</Grid>
          </Grid>
          <Grid container spacing={1}>
            <Grid item xs={12} sm={3} md={3}>
              <DatePicker
                placeholderText='--選擇日期--'
                className='form-control form-control-sm'
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
                placeholderText='--選擇日期--'
                className='form-control form-control-sm'
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
            <Grid item xs={12} sm={2} md={2}>
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
          {userAttendance.loadingStatus === 'loadError' &&
            <div className={styles.attendanceListEmpty}>
              <Error style={{color: 'slategrey'}}/>
              <div>Oops, Something went wrong</div>
            </div>
          }
          {(userAttendance.loadingStatus === 'loaded' || userAttendance.loadingStatus === 'loadNoData') &&
            <div>
              <MDBDataTable
                className={styles.attendanceTable}
                striped
                bordered
                small
                noBottomColumns
                sortable={false}
                entriesLabel='顯示項目'
                entriesOptions={[10, 20]}
                searchLabel='搜尋'
                paginationLabel={['上貢', '下頁']}
                infoLabel={['顯示第', '至' ,'項，共' ,'項記錄']}
                noRecordsFoundLabel='沒有記錄'
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
