import * as React from 'react';
import { ILeaveRecordsDashboardProps } from './ILeaveRecordsDashboardProps';
import { ILeaveTypeData, IUserAnnualLeaveData, IUserLeaveData, ILeaveRecordsDashboardState } from './ILeaveRecordsDashboardState';
// import { escape } from '@microsoft/sp-lodash-subset';
import axios from 'Axios';
import { Grid } from '@material-ui/core';
import { ExpandMore, ExpandLess, DateRange, List, Error, Search, Clear } from '@material-ui/icons';
import { Dropdown, IDropdownOption, IconButton, Toggle } from 'office-ui-fabric-react';
// import { ToggleButton, ToggleButtonGroup } from '@material-ui/lab';
import FullCalendar, { EventContentArg } from '@fullcalendar/react';
import dayGridPlugin from '@fullcalendar/daygrid';
import DatePicker from 'react-datepicker';
import { MDBDataTable,  } from 'mdbreact';
import { isNull } from 'lodash';
import * as moment from 'moment';

import 'bootstrap-css-only/css/bootstrap.min.css';
import 'mdbreact/dist/css/mdb.css';
import 'react-datepicker/dist/react-datepicker.css';
import styles from './LeaveRecordsDashboard.module.scss';



export default class LeaveRecordsDashboard extends React.Component<ILeaveRecordsDashboardProps, ILeaveRecordsDashboardState> {
  private _userEmail: string = this.props.context.pageContext.legacyPageContext.userEmail;
  private _absoluteUrl: string = this.props.context.pageContext.web.absoluteUrl;
  private _selectedYear = new Date();
  // private _maxRecords: number = 20;
  private _leaveTypesOptions: IDropdownOption[] = [];
  private _userCalendarEvents: Array<any> = [];
  private _calendarRef = React.createRef<FullCalendar>();

  constructor(props) {
    super(props);

    this.state = {
      showDetails: false,
      contentView: 'calendar',
      userCardNo: '',
      leaveTypes: {
        loadingStatus: 'loading',
        data: []
      },
      userAnnualLeave: {
        loadingStatus: 'loading',
        data: null
      },
      userLeaves: {
        loadingStatus: 'loading',
        data: []
      },
      filterYear: null,
      filterLeaveType: '所有',
    };
  }

  public async componentDidMount(): Promise<void> {
    this.getData();
  }

  private async getData(): Promise<void> {
    this.getLeaveTypes();
    await this.getUserInfo();
    this.getUserAnnualLeave();
    this.getUserLeaves();
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

  private async getLeaveTypes(): Promise<void> {
    try {
      let leaveTypesData: ILeaveTypeData[] = [];
      this._leaveTypesOptions = [
        {key: 'default', text: '所有'},
      ]
      const response = await axios.get(
        this._absoluteUrl + `/_api/web/lists/getbytitle('Leave Type')/items?`
      );
      if (response.data.value.length > 0) {
        response.data.value.map((item: ILeaveTypeData) => {
          leaveTypesData.push({
            typeID: item['leave_type'],
            typeNameEn: item['Title_en'],
            typeNameZh: item['Title_zh']
          });
          this._leaveTypesOptions.push({
            key: item['leave_type'],
            text: item['Title_zh']
          })
        })
        this.setState({leaveTypes: {loadingStatus: 'loaded', data: leaveTypesData}});

      } else {
        this.setState({leaveTypes: {loadingStatus: 'loadNoData', data: null}});
      }

    } catch (error) {
      console.error(error);
      this.setState({leaveTypes: {loadingStatus: 'loadError', data: null}});
    }
  }

  private async getUserAnnualLeave(year: Date = new Date()): Promise<void> {
    try {
      const response = await axios.get(
        this._absoluteUrl + `/_api/web/lists/getbytitle('LEAVEANU')/items?` +
          `$filter=CARD_NO eq '${this.state.userCardNo}'` +
          `and YEAR eq '${year.getFullYear().toString()}'`
      );
      if (response.data.value.length > 0) {
        this.setState({
          userAnnualLeave: {
            loadingStatus: 'loaded',
            data: {
              annualLeaveTotal: response.data.value[0]['ANU_TOTAL'],
              annualLeaveTaken: response.data.value[0]['ANNUAL'],
              sickLeaveTaken: response.data.value[0]['SICK']
            }
          }
        })
      } else {
        this.setState({userAnnualLeave: {loadingStatus: 'loadNoData', data: null}});
      }

    } catch (error) {
      console.error(error);
      this.setState({userAnnualLeave: {loadingStatus: 'loadError', data: null}});
    }
  }

  private async getUserLeaves(filterYear: Date = null, leaveType: string = '所有'): Promise<void> {
    try {
      let dateCondition = ``;
      let leaveTypeCondition = ``;
      
      if (leaveType != '所有') { leaveTypeCondition = ` and leave_type eq '${this.getLeaveTypeID(leaveType)}'`; }
      if (isNull(filterYear)) { filterYear = new Date(new Date().getFullYear(), 0, 1);}

      let nextYear = new Date(filterYear.toISOString());
      nextYear.setFullYear(nextYear.getFullYear() + 1);
      dateCondition = ` and leave_date_to ge datetime'${filterYear.toISOString()}' and leave_date_to lt datetime'${nextYear.toISOString()}'`;

      const response = await axios.get(
        this._absoluteUrl + `/_api/web/lists/getbytitle('LEAVE')/items?` +
          `$orderby=leave_date_to desc &` +
          `$filter=card_no eq '${this.state.userCardNo}'` +
          dateCondition +
          leaveTypeCondition
      );
      this._userCalendarEvents = [];

      if (response.data.value.length > 0) {
        let userLeavesData: IUserLeaveData[] = [];
        response.data.value.map((item: IUserLeaveData) => {
          userLeavesData.push({
            leaveTypeID: this.getLeaveTypeTitle(item['leave_type']),
            leaveDateFrom: this.formatDate(item['leave_date_from']),
            leaveDateTo: this.formatDate(item['leave_date_to']),
            daysCount: item['total_day']
          });
          this._userCalendarEvents.push({
            title: this.getLeaveTypeTitle(item['leave_type']),
            start: item['leave_date_from'],
            end: item['leave_date_to']
          })
        })
        this.setState({userLeaves: {loadingStatus: 'loaded', data: userLeavesData}});

      } else {
        this.setState({userLeaves: {loadingStatus: 'loadNoData', data: null}});
      }

    } catch (error) {
      console.error(error);
      this.setState({userLeaves: {loadingStatus: 'loadError', data: null}});
    }
  }

  private async applyFilter(): Promise<void> {
    if (isNull(this.state.filterYear)) {
      this._selectedYear = new Date();
      this.getUserAnnualLeave();
    } else {
      this._selectedYear = this.state.filterYear;
      this.getUserAnnualLeave(this.state.filterYear);
    }
    await this.getUserLeaves(this.state.filterYear, this.state.filterLeaveType);
    this._calendarRef.current.getApi().changeView('dayGridMonth', this.formatDate(this.state.filterYear.toISOString()));
  }


  private resetFilter(): void {
    this.setState({
      filterYear: null,
      filterLeaveType: '所有'
    })
  }

  private getLeaveTypeTitle(leaveTypeID: string): string {
    return (this.state.leaveTypes.data.filter(item => item.typeID === leaveTypeID)[0].typeNameZh);
  }

  private getLeaveTypeID(leaveTypeTitle: string): string {
    return (this.state.leaveTypes.data.filter(item => item.typeNameZh === leaveTypeTitle)[0].typeID);
  }

  private getDropdownSelectedKey(): number | string {
    return (this._leaveTypesOptions.filter(item => item.text === this.state.filterLeaveType)[0].key);
  }

  private formatDate(dateTime: string): string {
    return moment(dateTime).format('yyyy-MM-DD');
  }

  private formatTime(dateTime: string): string {
    return moment(dateTime).format('h:mm a');
  }

  private renderEventContent(eventInfo: EventContentArg): JSX.Element {
    return (
      <>
        <div className={styles.calendarEvents}>
          <i>{eventInfo.event.title}</i>
        </div>
      </>
    );
  }
  
  public render(): React.ReactElement<ILeaveRecordsDashboardProps> {
    const { showDetails, contentView, userAnnualLeave, userLeaves, filterYear, filterLeaveType } = this.state;
    const leaveTableColumns = [
      {
        label: '類別',
        field: 'leaveTypeID',
        sort: 'asc',
        width: 150
      },
      {
        label: '由日期',
        field: 'leaveDateFrom',
        sort: 'asc',
        width: 270
      },
      {
        label: '至日期',
        field: 'leaveDateTo',
        sort: 'asc',
        width: 270
      },
      {
        label: '日數',
        field: 'daysCount',
        sort: 'asc',
        width: 200
      },
    ];

    return (
      <section className={styles.leaveRecordsDashboard}>
        {/* Top Bar section */}
        <div className={styles.leaveOverview}>
          {/* Title section */}
          <div className={styles.leaveTitleBar}>
            <Grid container>
              <Grid item sm={8} md={9}>
                <div className={styles.leaveTitleBarItem} onClick={()=> this.setState({showDetails: !showDetails})}>
                  {this._selectedYear.getFullYear().toString()}假期查詢
                  {!showDetails && <ExpandMore id={styles.expandIcon} />}
                  {showDetails && <ExpandLess id={styles.expandIcon} />}
                </div>
              </Grid>
              <Grid item sm={4} md={3}>
                {/* <ToggleButtonGroup
                  color="primary"
                  exclusive
                  size='small'
                  onChange={(onClick, value) => {this.setState({contentView: value})}}
                >
                  <ToggleButton selected={contentView === 'calendar'} value='calendar'>Calendar</ToggleButton>
                  <ToggleButton selected={contentView === 'list'} value='list'>List</ToggleButton>
                </ToggleButtonGroup> */}
                <div style={{whiteSpace: 'normal'}} className={styles.leaveTitleBarItem}>
                  <Toggle
                    className={styles.toggle}
                    // style={{margin: '0'}}
                    defaultChecked
                    onText="Show List"
                    offText="Show Calendar"
                    onChange={(onChange, isChecked) => {
                      if(isChecked) 
                        this.setState({contentView: 'calendar'})
                      else
                        this.setState({contentView: 'list'})
                    }}
                  />
                </div>
              </Grid>
            </Grid>
          </div>

          {/* Leave Overiew section */}
          <div className={styles.divider} />
          <div >
            <Grid container>
              <Grid item sm={3} md={3}>年假</Grid>
              <Grid item sm={3} md={3}>總數 <span className={styles.leaveTotal}>{isNull(userAnnualLeave.data) ? 0 : userAnnualLeave.data.annualLeaveTotal}</span></Grid>
              <Grid item sm={3} md={3}>已取 <span className={styles.leaveTaken}>{isNull(userAnnualLeave.data) ? 0 : userAnnualLeave.data.annualLeaveTaken}</span></Grid>
              <Grid item sm={3} md={3}>餘額 <span className={styles.leaveRemaining}>{isNull(userAnnualLeave.data) ? 0 : (userAnnualLeave.data.annualLeaveTotal - userAnnualLeave.data.annualLeaveTaken)}</span></Grid>
            </Grid>
            <div className={styles.divider} />
            <Grid container>
              <Grid item sm={6} md={6}>病假</Grid>
              <Grid item sm={6} md={6}>已取 <span className={styles.leaveTaken}>{isNull(userAnnualLeave.data) ? 0 : userAnnualLeave.data.sickLeaveTaken}</span></Grid>
            </Grid>
          </div>
        </div>

        {/* User Leave Calendar View section */}
        {showDetails &&
          <div className={styles.leaveContentView}>
            {/* Filter section */}
            <div className={styles.filterSection}>
              <Grid container>
                <Grid item xs={12} sm={4} md={3}>選擇年份</Grid>
                <Grid item xs={12} sm={4} md={3}>類別</Grid>
              </Grid>
              <div className={styles.box5px}/>
              <Grid container>
                <Grid item xs={12} sm={4} md={3}>
                  <DatePicker
                    placeholderText='Pick a year'
                    className={styles.datePicker}
                    value={isNull(filterYear)
                            ? null 
                            : moment(filterYear).format('yyyy')}
                    dateFormat='yyyy'
                    showYearPicker
                    onChange={(date) => {
                      this.setState({filterYear: date});
                    }}
                  />
                {/* <DatePicker
                  
                  firstDayOfWeek={DayOfWeek.Sunday}
                  placeholder="Select a date..."
                  // strings={}
                /> */}
                </Grid>
                <Grid item xs= {12} sm={4} md={3}>
                  <Dropdown
                    className={styles.dropdown}
                    placeholder='Select Leave Type'
                    options={this._leaveTypesOptions}
                    selectedKey={this.getDropdownSelectedKey()}
                    onChange={(onChange, option) => {
                      this.setState({
                        filterLeaveType: option.text
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
            
            {/* Calendar section */}
            { contentView === 'calendar' &&
              <div style={{width: '100%'}}>
                <FullCalendar
                  ref={this._calendarRef}
                  plugins={[ dayGridPlugin ]}
                  initialView="dayGridMonth"
                  events={this._userCalendarEvents}
                  eventContent={this.renderEventContent}
                  navLinkDayClick={() => {console.log("clicked")}}
                />
              </div>
            }

            {/* List section */}
            { contentView === 'list' &&
              <div>
                {userLeaves.loadingStatus === 'loading' &&
                  <div className={styles.leaveListEmpty}>
                    Loading...
                  </div>
                }
                {userLeaves.loadingStatus === 'loadNoData' &&
                  <div className={styles.leaveListEmpty}>
                    -No Data-
                  </div>
                }
                {userLeaves.loadingStatus === 'loadError' &&
                  <div className={styles.leaveListEmpty}>
                    <Error style={{color: 'slategrey'}}/>
                    <div>Oops, Something went wrong</div>
                  </div>
                }
                {userLeaves.loadingStatus === 'loaded' &&
                  <div>
                    <MDBDataTable
                      striped
                      bordered
                      small
                      noBottomColumns
                      sortable={false}
                      className={styles.leaveTable}
                      data={{
                        columns: leaveTableColumns,
                        rows: userLeaves.data
                      }}
                    />
                  </div>
                }
              </div>
            }
          </div>
        }

        {/* Footer section */}
        {showDetails &&
          <div className={styles.footer}>
            備註︰
            <br />年假必需每年放清，不可累積，如有特殊情況，由公司酌情處理。
          </div>
        }
      </section>
    );
  }
}
