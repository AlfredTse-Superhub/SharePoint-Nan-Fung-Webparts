import * as React from 'react';
import { ILeaveRecordsDashboardProps } from './ILeaveRecordsDashboardProps';
import { ILeaveTypeData, IUserLeaveData, ILeaveRecordsDashboardState } from './ILeaveRecordsDashboardState';
import axios from 'Axios';
import { ExpandMore, ExpandLess, Error, Search, Clear } from '@material-ui/icons';
import { Dropdown, IDropdownOption, IconButton, Toggle } from 'office-ui-fabric-react';
import FullCalendar, { EventContentArg } from '@fullcalendar/react';
import dayGridPlugin from '@fullcalendar/daygrid';
import DatePicker from 'react-datepicker';
import { MDBDataTable } from 'mdbreact';
import { isNull } from 'lodash';
import * as moment from 'moment';
import classnames from 'classnames';

import 'react-datepicker/dist/react-datepicker.css';
import 'bootstrap-css-only/css/bootstrap.min.css';
import 'mdbreact/dist/css/mdb.css';
import styles from './LeaveRecordsDashboard.module.scss';



export default class LeaveRecordsDashboard extends React.Component<ILeaveRecordsDashboardProps, ILeaveRecordsDashboardState> {
  private _userEmail: string = this.props.context.pageContext.legacyPageContext.userEmail;
  private _absoluteUrl: string = this.props.context.pageContext.web.absoluteUrl;
  private _selectedYear = new Date();
  private _leaveTypesOptions: IDropdownOption[] = [{key: 'default', text: '所有'}];
  private _userCalendarEvents: Array<any> = [];
  private _calendarRef = React.createRef<FullCalendar>();

  constructor(props) {
    super(props);

    this.state = {
      isExpanded: true,
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
        this._absoluteUrl + `/_api/web/lists/getbytitle('Employee Card Record')/items?` +
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
        this.setState({leaveTypes: {loadingStatus: 'loadNoData', data: []}});
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
        this.setState({userLeaves: {loadingStatus: 'loadNoData', data: []}});
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
    if (this.state.contentView === 'calendar') {
      this._calendarRef.current.getApi().changeView('dayGridMonth', this.formatDate(this._selectedYear.toISOString()));
    }
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
    const { isExpanded, contentView, userAnnualLeave, userLeaves, filterYear, filterLeaveType } = this.state;
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
        <div className='container' style={{width: '100vw'}}>
          {/* Top Bar section */}
          <div className={classnames('row', styles.titleBar)}>
            <div className={classnames('col-9', styles.noPadding)}>
              <div className={styles.titleBarItem} onClick={()=> this.setState({isExpanded: !isExpanded})}>
                {this._selectedYear.getFullYear().toString()}假期查詢
                {!isExpanded && <ExpandMore id={styles.expandIcon} />}
                {isExpanded && <ExpandLess id={styles.expandIcon} />}
              </div>
            </div>
            <div className={classnames('col-3', styles.colCenter)}>
              <div className={styles.titleBarItem}>
                <Toggle
                  className={styles.toggle}
                  defaultChecked
                  onText="切換至列表"
                  offText="切換至月歷"
                  onChange={(onChange, isChecked) => {
                    if(isChecked) 
                      this.setState({contentView: 'calendar'})
                    else
                      this.setState({contentView: 'list'})
                  }}
                />
              </div>
            </div>
          </div>

          <div className='row'>
            <div className={classnames('col-12', styles.noPadding)}><div className={styles.divider} /></div>
            <div className={classnames('col-3', styles.noPadding)}>年假</div>
            <div className={classnames('col-3', styles.noPadding)}>總數 <span className={styles.decoratedTextBlue}>{isNull(userAnnualLeave.data) ? 0 : userAnnualLeave.data.annualLeaveTotal}</span></div>
            <div className={classnames('col-3', styles.noPadding)}>已取 <span className={styles.decoratedTextGray}>{isNull(userAnnualLeave.data) ? 0 : userAnnualLeave.data.annualLeaveTaken}</span></div>
            <div className={classnames('col-3', styles.noPadding)}>餘額 <span className={styles.decoratedTextGreen}>{isNull(userAnnualLeave.data) ? 0 : (userAnnualLeave.data.annualLeaveTotal - userAnnualLeave.data.annualLeaveTaken)}</span></div>
          </div>

          <div className='row'>
            <div className={classnames('col-12', styles.noPadding)}><div className={styles.divider} /></div>
            <div className={classnames('col-6', styles.noPadding)}>病假</div>
            <div className={classnames('col-6', styles.noPadding)}>已取 <span className={styles.decoratedTextGray}>{isNull(userAnnualLeave.data) ? 0 : userAnnualLeave.data.sickLeaveTaken}</span></div>
          </div>

          {/* User Leave Content section */}
          {isExpanded &&
            <div style={{marginTop: '50px'}}>
              <div className='row'>
                {/* Filter section */}
                <div className={classnames('col-sm-4 col-md-3', styles.noPadding)}>
                  <div className='col-12'>
                    年份
                  </div>
                  <div className='col-12'>
                    <DatePicker
                      placeholderText='--選擇年份--'
                      className='form-control form-control-sm'
                      popperClassName={styles.datePickerPopper}
                      value={isNull(filterYear)
                              ? null 
                              : moment(filterYear).format('yyyy')}
                      dateFormat='yyyy'
                      showYearPicker
                      onChange={(date) => {
                        this.setState({filterYear: date});
                      }}
                    />
                  </div>
                </div>
                <div className={classnames('col-sm-4 col-md-3', styles.noPadding)}>
                  <div className='col-12'>
                    類別
                  </div>
                  <div className='col-12'>
                    <Dropdown
                      className={styles.dropdown}
                      placeholder='--假期類別--'
                      options={this._leaveTypesOptions}
                      selectedKey={this.getDropdownSelectedKey()}
                      onChange={(onChange, option) => {
                        this.setState({
                          filterLeaveType: option.text
                        })
                      }}
                    />
                  </div>
                </div>
                <div className={classnames('col-sm-4 col-md-6', styles.filterButton)}>
                  <IconButton className={styles.iconButton} onClick={() => this.applyFilter()}><Search /></IconButton>
                  <IconButton className={styles.iconButton} onClick={() => this.resetFilter()}><Clear /></IconButton>
                </div>
              </div>
              
              <div style={{margin: '30px 0px'}}>
                {/* Calendar section */}
                { contentView === 'calendar' &&
                  <div className={classnames('col-12', styles.noPadding)}>
                    <FullCalendar
                      ref={this._calendarRef}
                      plugins={[ dayGridPlugin ]}
                      initialView="dayGridMonth"
                      events={this._userCalendarEvents}
                      eventContent={this.renderEventContent}
                    />
                  </div>
                }

                {/* List section */}
                { contentView === 'list' &&
                  <div className={classnames('col-12', styles.noPadding)}>
                    {userLeaves.loadingStatus === 'loading' &&
                      <div className={styles.leaveListEmpty}>
                        Loading...
                      </div>
                    }
                    {userLeaves.loadingStatus === 'loadError' &&
                      <div className={styles.leaveListEmpty}>
                        <Error style={{color: 'slategrey'}}/>
                        <div>Oops, Something went wrong</div>
                      </div>
                    }
                    {(userLeaves.loadingStatus === 'loaded' || userLeaves.loadingStatus === 'loadNoData')  &&
                      <div>
                        <MDBDataTable
                          className={styles.leaveTable}
                          striped
                          bordered
                          small
                          noBottomColumns
                          sortable={false}
                          entriesOptions={[10, 20]}
                          entriesLabel='顯示項目'
                          searchLabel='搜尋'
                          paginationLabel={['上貢', '下頁']}
                          infoLabel={['顯示第', '至' ,'項，共' ,'項記錄']}
                          noRecordsFoundLabel='沒有記錄'
                          data={{
                            columns: leaveTableColumns,
                            rows: userLeaves.data
                          }}
                        />
                      </div>
                    }
                  </div>
                }

                {/* Footer section */}
                <div className={classnames('col-12', styles.footer)}>
                  備註︰
                  <br />年假必需每年放清，不可累積，如有特殊情況，由公司酌情處理。
                </div>
              </div>
            </div>
          }
        </div>
      </section>
    );
  }
}
