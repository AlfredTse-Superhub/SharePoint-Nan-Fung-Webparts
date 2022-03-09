import * as React from 'react';
import { ILeaveRecordsDashboardProps } from './ILeaveRecordsDashboardProps';
import { ContentView, ILeaveTypeData, IUserAnnualLeaveData, IUserLeaveData, ILeaveRecordsDashboardState } from './ILeaveRecordsDashboardState';
import { escape } from '@microsoft/sp-lodash-subset';
import axios from 'Axios';
import { Grid, Button } from '@material-ui/core';
import { ExpandMore, ExpandLess, Error } from '@material-ui/icons';
import { Dropdown,IDropdownOption } from 'office-ui-fabric-react';
import { ToggleButton, ToggleButtonGroup } from '@material-ui/lab';
import FullCalendar, { EventContentArg } from '@fullcalendar/react'
import dayGridPlugin from '@fullcalendar/daygrid'
import DatePicker from 'react-datepicker';
import { isNull } from 'lodash';
import * as moment from 'moment';

import styles from './LeaveRecordsDashboard.module.scss';


export default class LeaveRecordsDashboard extends React.Component<ILeaveRecordsDashboardProps, ILeaveRecordsDashboardState> {
  private _userEmail: string = this.props.context.pageContext.legacyPageContext.userEmail;
  private _absoluteUrl: string = this.props.context.pageContext.web.absoluteUrl;
  private _currentYear: string = new Date().getFullYear().toString();
  private _maxRecords: number = 10;
  private _leaveTypesOptions: IDropdownOption[] = [];
  private _userCalendarEvents: Array<any> = [];
  private _calendarRef = React.createRef<FullCalendar>();
  private _monthPicker = React.createRef<DatePicker>();

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
        data: {
          annualLeaveTotal: 0,
          annualLeaveTaken: 0,
          sickLeaveTaken: 0
        }
      },
      userLeaves: {
        loadingStatus: 'loading',
        data: []
      },
      filterMonth: null,
      filterDateFrom: null,
      filterDateTo: null,
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

  private async getUserAnnualLeave(): Promise<void> {
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

  private async getUserLeaves(
    month: Date = null,
    dateFrom: Date = null, 
    dateTo: Date = null, 
    leaveType: string = '所有'
  ): Promise<void> {
    try {
      let dateFromCondition = ``;
      let dateToCondition = ``;
      let leaveTypeCondition = ``;

      if (!isNull(dateFrom)) { dateFromCondition = `and leave_date_from ge datetime'${dateFrom.toISOString()}'`; }
      if (!isNull(dateTo)) { dateToCondition = `and leave_date_to le datetime'${dateTo.toISOString()}'`; }
      if (!isNull(month)) {
        let nextMonth = new Date(month.toISOString());
        nextMonth.setMonth(nextMonth.getMonth() + 1);
        dateFromCondition = `and leave_date_from ge datetime'${month.toISOString()}'`;
        dateToCondition = `and leave_date_to lt datetime'${nextMonth.toISOString()}'`;
      }
      if (leaveType != '所有') { leaveTypeCondition = `and leave_type eq '${this.getLeaveTypeID(leaveType)}'`; }

      const response = await axios.get(
        this._absoluteUrl + `/_api/web/lists/getbytitle('LEAVE')/items?` +
          `$top=${this._maxRecords} &` +
          `$orderby=leave_date_to desc &` +
          `$filter=card_no eq '${this.state.userCardNo}'` +
          dateFromCondition +
          dateToCondition +
          leaveTypeCondition
      );
      this._userCalendarEvents = [];

      if (response.data.value.length > 0) {
        let userLeavesData: IUserLeaveData[] = [];
        response.data.value.map((item: IUserLeaveData) => {
          userLeavesData.push({
            leaveTypeID: item['leave_type'],
            leaveDateFrom: item['leave_date_from'],
            leaveDateTo: item['leave_date_to'],
            daysCount: item['total_day']
          });
          this._userCalendarEvents.push({
            title: this.getLeaveTypeTitle(item['leave_type']),
            start: item['leave_date_from'],
            end: item['leave_date_to']
          })
        })
        this.setState({userLeaves: {loadingStatus: 'loaded',data: userLeavesData}});

      } else {
        this.setState({userLeaves: {loadingStatus: 'loadNoData', data: null}});
      }

    } catch (error) {
      console.error(error);
      this.setState({userLeaves: {loadingStatus: 'loadError', data: null}});
    }
  }

  // private changeMonth(date: Date): void {
  //   this.setState({filterMonth: date});
  //   // this._calendarRef.current.getApi().changeView('dayGridMonth', this.formatDate(date.toISOString()));
  // }

  private async applyFilter(filterType: ContentView): Promise<void> {
    if (filterType === 'calendar')
      await this.getUserLeaves(this.state.filterMonth, null ,null, this.state.filterLeaveType);
    if (filterType === 'list')
      await this.getUserLeaves(null, this.state.filterDateFrom, this.state.filterDateTo, this.state.filterLeaveType);
    // change calendar view
    this._calendarRef.current.getApi().changeView('dayGridMonth', this.formatDate(this.state.filterMonth.toISOString()));
  }

  private resetFilter(): void {
    this.setState({
      filterMonth: null,
      filterDateFrom: null,
      filterDateTo: null,
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

  // private handleEventClick = (clickInfo) => {
  //   if (confirm(`Are you sure you want to delete the event '${clickInfo.event.title}'`)) {
  //     // clickInfo.event.remove()
  //   }
  // }

  
  public render(): React.ReactElement<ILeaveRecordsDashboardProps> {
    const { showDetails, contentView, userAnnualLeave, userLeaves, filterMonth, filterDateFrom, filterDateTo, filterLeaveType } = this.state;

    return (
      <section className={styles.leaveRecordsDashboard}>
        {/* Top Bar section */}
        <div className={styles.leaveOverview}>
          {/* Title section */}
          <Grid container>
            <Grid item sm={10} md={10}>
              <div className={styles.leaveTitleBar} onClick={()=> this.setState({showDetails: !showDetails})}>
                {this._currentYear}假期查詢
                {!showDetails && <ExpandMore id={styles.mediumIcon} />}
                {showDetails && <ExpandLess id={styles.mediumIcon} />}
              </div>
            </Grid>
            <Grid item sm={2} md={2}>
              <ToggleButtonGroup
                color="primary"
                exclusive
                size='small'
                onChange={(onClick, value) => {this.setState({contentView: value})}}
              >
                <ToggleButton selected={contentView === 'calendar'} value='calendar'>Calendar</ToggleButton>
                <ToggleButton selected={contentView === 'list'} value='list'>List</ToggleButton>
              </ToggleButtonGroup>
            </Grid>
          </Grid>

          {/* Leave Overiew section */}
          <div className={styles.divider} />
          <div>
            <Grid container>
              <Grid item sm={3} md={3}>年假</Grid>
              <Grid item sm={3} md={3}>總數 <span className={styles.leaveTotal}>{userAnnualLeave.data.annualLeaveTotal}</span></Grid>
              <Grid item sm={3} md={3}>已取 <span className={styles.leaveTaken}>{userAnnualLeave.data.annualLeaveTaken}</span></Grid>
              <Grid item sm={3} md={3}>餘額 <span className={styles.leaveRemaining}>{userAnnualLeave.data.annualLeaveTotal - userAnnualLeave.data.annualLeaveTaken}</span></Grid>
            </Grid>
            <div className={styles.divider} />
            <Grid container>
              <Grid item sm={6} md={6}>病假</Grid>
              <Grid item sm={6} md={6}>已取 <span className={styles.leaveTaken}>{userAnnualLeave.data.sickLeaveTaken}</span></Grid>
            </Grid>
          </div>
        </div>

        {/* User Leave Calendar View section */}
        {(showDetails && contentView === 'calendar') &&
          <div className={styles.leaveContentView}>
            {/* Filter section */}
            <div className={styles.filterSection}>
              <Grid container>
                <Grid item sm={3} md={3}>選擇月份</Grid>
                <Grid item sm={3} md={3}>類別</Grid>
              </Grid>
              <div className={styles.box5px}/>
              <Grid container>
                <Grid item sm={3} md={3}>
                  <DatePicker
                    ref={this._monthPicker}
                    placeholderText='Pick a month'
                    className={styles.datePicker}
                    value={isNull(filterMonth)
                            ? null 
                            : moment(filterMonth).format('MM/yyyy')}
                    dateFormat='MM/yyyy'
                    showMonthYearPicker={true}
                    onChange={(date) => {
                      this.setState({filterMonth: date});
                    }}
                  />
                </Grid>
                <Grid item sm={3} md={3}>
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
              </Grid>
              <Grid container>
                <Grid item sm={12} md={12}>
                  <Button style={{color: 'darkslategrey'}} onClick={() => this.applyFilter('calendar')}>Filter</Button>
                  <Button style={{color: 'darkslategrey'}} onClick={() => this.resetFilter()}>Reset</Button>
                </Grid>
              </Grid>
            </div>
            
            {/* Calendar section */}
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

        {/* User Leave List View section */}
        {(showDetails && contentView==='list') &&
          <div className={styles.leaveContentView}>
            {/* Filter section */}
            <div className={styles.filterSection}>
              <Grid container>
                <Grid item sm={3} md={3}>由日期</Grid>
                <Grid item sm={3} md={3}>至日期</Grid>
                <Grid item sm={3} md={3}>類別 </Grid>
              </Grid>
              <div className={styles.box5px}/>
              <Grid container>
                <Grid item sm={3} md={3}>
                  <DatePicker
                    placeholderText='Select date from'
                    className={styles.datePicker}
                    value={isNull(filterDateFrom)
                            ? null 
                            : moment(filterDateFrom).format('yyyy-MM-DD')}
                    onChange={(date) => {
                      this.setState({
                        filterDateFrom: date
                      });
                    }}
                  />
                </Grid>
                <Grid item sm={3} md={3}>
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
                <Grid item sm={3} md={3}>
                  <Dropdown
                    className={styles.dropdown}
                    placeholder='Select Leave Type'
                    options={this._leaveTypesOptions}
                    selectedKey={this.getDropdownSelectedKey()}
                    onChange={(onChange, option) => {
                      this.setState({
                        filterLeaveType: option.text
                      })
                      console.log("Leave Type: " + filterLeaveType)
                    }}
                  />
                </Grid>
              </Grid>
              <Grid container>
                <Grid item sm={12} md={12}>
                  <Button style={{color: 'darkslategrey'}} onClick={() => this.applyFilter('list')}>Filter</Button>
                  <Button style={{color: 'darkslategrey'}} onClick={() => this.resetFilter()}>Reset</Button>
                </Grid>
              </Grid>
            </div>
            
            {/* List section */}
            <div className={styles.leaveListHeader}>
              <Grid container>
                <Grid item sm={3} md={3}>類別</Grid>
                <Grid item sm={8} md={8}>日期</Grid>
                <Grid item sm={1} md={1}>日數</Grid>
              </Grid>
            </div>

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
                { userLeaves.data.map((item) => {
                    return (
                      <div className={styles.leaveListRow}>
                        <Grid container>
                          <Grid item sm={3} md={3}>{this.getLeaveTypeTitle(item.leaveTypeID)}</Grid>
                          <Grid item sm={8} md={8}>
                            {item.leaveDateFrom == item.leaveDateTo 
                              ? this.formatDate(item.leaveDateTo) 
                              : this.formatDate(item.leaveDateFrom) + ' - ' + this.formatDate(item.leaveDateTo)}
                          </Grid>
                          <Grid item sm={1} md={1}>{item.daysCount}</Grid>
                        </Grid>
                      </div>
                    );
                  })
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
