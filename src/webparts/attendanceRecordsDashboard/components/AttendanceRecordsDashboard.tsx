import * as React from 'react';
import { IAttendanceRecordsDashboardProps } from './IAttendanceRecordsDashboardProps';
import { IUserAttendanceData, IAttendanceRecordsDashboardState } from './IAttendanceRecordsDashboardState';
import axios from 'Axios';
import { Clear, Error, Search } from '@material-ui/icons';
import { IconButton } from 'office-ui-fabric-react';
import DatePicker from "react-datepicker";
import { MDBDataTable } from 'mdbreact';
import { isNull } from 'lodash';
import * as moment from 'moment';
import classnames from 'classnames';

import 'react-datepicker/dist/react-datepicker.css';
import 'bootstrap-css-only/css/bootstrap.min.css';
import 'mdbreact/dist/css/mdb.css';
import styles from './AttendanceRecordsDashboard.module.scss';



export default class AttendanceRecordsDashboard extends React.Component<IAttendanceRecordsDashboardProps, IAttendanceRecordsDashboardState> {
  private _userEmail: string = this.props.context.pageContext.user.email;
  private _absoluteUrl: string = this.props.context.pageContext.web.absoluteUrl;
  private _defaultDateFrom: Date = new Date(new Date().getFullYear(), new Date().getMonth(), 1);
  private _defaultDateTo: Date = new Date(new Date().getFullYear(), new Date().getMonth() + 1, 1);

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
        this._absoluteUrl + `/_api/web/lists/getbytitle('Employee Card Record')/items?` +
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
        this._absoluteUrl + `/_api/web/lists/getbytitle('Attendance')/items?` +
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
        label: '??????',
        field: 'logDate',
        sort: 'asc',
      },
      {
        label: '??????',
        field: 'logTime',
        sort: 'asc',
      },
      {
        label: '??????',
        field: 'logLocation',
        sort: 'asc',
      },
    ];

    return (
      <section className={styles.attendanceRecordsDashboard}>
        <div className='container' style={{width: '100vw'}}>
          {/* Header section */}
          <div className='row'>
            <div className={classnames('col-12', styles.noPadding)} style={{margin: '10px 0px'}}>
              <div className={styles.divider} />
            </div>
          </div>
          
          {/* Filter section */}
          <div className='row'>
            <div className={classnames('col-sm-4 col-md-3', styles.noPadding)}>
              <div className='col-12'>
                ????????????
              </div>
              <div className='col-12'>
                <DatePicker
                  placeholderText='--????????????--'
                  className='form-control form-control-sm'
                  value={
                    isNull(filterDateFrom)
                      // ? null
                      ? this.formatDate(this._defaultDateFrom.toISOString())
                      : this.formatDate(filterDateFrom.toISOString())
                  }
                  onChange={(date) => {
                    this.setState({
                      filterDateFrom: date
                    });
                  }}
                />
              </div>
            </div>
            <div className={classnames('col-sm-4 col-md-3', styles.noPadding)}>
              <div className='col-12'>
                ????????????
              </div>
              <div className='col-12'>
                <DatePicker
                  placeholderText='--????????????--'
                  className='form-control form-control-sm'
                  value={
                    isNull(filterDateTo)
                      ? this.formatDate(this._defaultDateTo.toISOString())
                      : moment(filterDateTo).format('yyyy-MM-DD')
                  }
                  onChange={(date) => {
                    this.setState({
                      filterDateTo: date
                    })
                  }}
                />
              </div>
            </div>
            <div className={classnames('col-sm-4 col-md-6', styles.filterButton)}>
              <IconButton className={styles.iconButton} onClick={() => this.applyFilter()}>
                <Search />
              </IconButton>
              <IconButton className={styles.iconButton} onClick={() => this.resetFilter()}>
                <Clear />
              </IconButton>
            </div>
          </div>

          {/* Attendance list section */}
          <div className={classnames("col-12", styles.attendanceList, styles.noPadding)}>
            {userAttendance.loadingStatus === 'loading' &&
              <div className={styles.attendanceListEmpty}>
                Loading...
              </div>
            }
            {userAttendance.loadingStatus === 'loadError' &&
              <div className={styles.attendanceListEmpty}>
                <Error style={{color: 'slategrey'}}/>
                Something went wrong
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
                  entriesLabel='????????????'
                  entriesOptions={[10, 20]}
                  searchLabel='??????'
                  paginationLabel={['??????', '??????']}
                  infoLabel={['?????????', '???' ,'?????????' ,'?????????']}
                  noRecordsFoundLabel='????????????'
                  data={{
                    columns: attendanceTableColumns,
                    rows: userAttendance.data
                  }}
                />
              </div>
            }
          </div>
        </div>
      </section>
    );
  }
}
