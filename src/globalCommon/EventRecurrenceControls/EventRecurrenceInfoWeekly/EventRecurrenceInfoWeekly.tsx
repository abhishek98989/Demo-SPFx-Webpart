import * as React from 'react'; 
import styles from './EventRecurrenceInfoWeekly.module.scss';
import strings from "../constants/strings";
import { IEventRecurrenceInfoWeeklyProps } from './IEventRecurrenceInfoWeeklyProps';
import { IEventRecurrenceInfoWeeklyState } from './IEventRecurrenceInfoWeeklyState';
//import { escape } from '@microsoft/sp-lodash-subset';
import * as moment from 'moment';
//import { parseString } from "xml2js";
import {
  ChoiceGroup,
  IChoiceGroupOption,
  Label,
  MaskedTextField,
  Checkbox,
} from 'office-ui-fabric-react';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { toLocaleShortDateString } from '../utils/dateUtils';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/regional-settings/web";


//import spservices from '../../services/spservices';

const DayPickerStrings: IDatePickerStrings = {
  months: [strings.January, strings.February, strings.March, strings.April, strings.May, strings.June, strings.July, strings.August, strings.September, strings.October, strings.November, strings.December],

  shortMonths: [strings.Jan, strings.Feb, strings.Mar, strings.Apr, strings.May, strings.Jun, strings.Jul, strings.Aug, strings.Sep, strings.Oct, strings.Nov, strings.Dez],

  days: [strings.Sunday, strings.Monday, strings.Tuesday, strings.Wednesday, strings.Thursday, strings.Friday, strings.Saturday],

  shortDays: [strings.ShortDay_Saunday, strings.ShortDay_M, strings.ShortDay_T, strings.ShortDay_W, strings.ShortDay_Tursday, strings.ShortDay_Friday, strings.ShortDay_S],

  goToToday: strings.GoToDay,
  prevMonthAriaLabel: strings.PrevMonth,
  nextMonthAriaLabel: strings.NextMonth,
  prevYearAriaLabel: strings.PrevYear,
  nextYearAriaLabel: strings.NextYear,
  closeButtonAriaLabel: strings.CloseDate
};

/**
 *
 *
 * @export
 * @class EventRecurrenceInfoDaily
 * @extends {React.Component<IEventRecurrenceInfoWeeklyProps, IEventRecurrenceInfoWeeklyState>}
 */
export class EventRecurrenceInfoWeekly extends React.Component<IEventRecurrenceInfoWeeklyProps, IEventRecurrenceInfoWeeklyState> {
  //private spService: spservices = null;
  public constructor(props: IEventRecurrenceInfoWeeklyProps) {
    super(props);
    this.onPaternChange = this.onPaternChange.bind(this);


    this.state = {
      selectedKey: 'daily',
      selectPatern: 'every',
      startDate: this.props.startDate ?? moment().toDate(),
      endDate: this.props?.DueDate ??moment().endOf('month').toDate(), 
      // endDate:props?.DueDate!=undefined?props?.DueDate: "",
      numberOcurrences: '10',
      numberOfWeeks: '1',
      disableNumberOfWeeks: false,
      disableNumberOcurrences: true,
      selectdateRangeOption: 'noDate',
      disableEndDate: true,
      weeklySunday:  moment().weekday() === 0 ? true: false,
      weeklyMonday: moment().weekday() === 1 ? true: false,
      weekklyTuesday: moment().weekday() === 2 ? true: false,
      weekklyWednesday: moment().weekday() === 3 ? true: false,
      weekklyThursday: moment().weekday() === 4 ? true: false,
      weeklyFriday: moment().weekday() === 5 ? true: false,
      weeklySaturday: moment().weekday() === 6 ? true: false,
      isLoading: false,
      errorMessageNumberOfWeeks: '',
      PatternType:"Weekly",
      selectedPeople : this.props.userName,
      selectedPeopleEmail : this.props.PeopleEmail,
      selectedPeopleId : this.props.userId,
      UserDataIndex : this.props.UserDataIndex
    };

    //
    this.onNumberOfWeeksChange = this.onNumberOfWeeksChange.bind(this);
    this.onNumberOfOcurrencesChange = this.onNumberOfOcurrencesChange.bind(this);
    this.onDataRangeOptionChange = this.onDataRangeOptionChange.bind(this);
    this.onEndDateChange = this.onEndDateChange.bind(this);
    this.onStartDateChange = this.onStartDateChange.bind(this);
    this.onApplyRecurrence = this.onApplyRecurrence.bind(this);
    this.onCheckboxSundayChange = this.onCheckboxSundayChange.bind(this);
    this.onCheckboxMondayChange = this.onCheckboxMondayChange.bind(this);
    this.onCheckboxTuesdayChange = this.onCheckboxTuesdayChange.bind(this);
    this.onCheckboxWednesdayChange = this.onCheckboxWednesdayChange.bind(this);
    this.onCheckboxThursdayChange = this.onCheckboxThursdayChange.bind(this);
    this.onCheckboxFridayChange = this.onCheckboxFridayChange.bind(this);
    this.onCheckboxSaturdayChange = this.onCheckboxSaturdayChange.bind(this);

    //this.spService = new spservices(this.props.context);
  }

  /**
   *
   *
   * @private
   * @param {Date} date
   * @memberof EventRecurrenceInfoDaily
   */
   private onStartDateChange(date: Date) {
    //Put the applyRecurrence() function in the callback of the setState() method to make sure that applyRecurrence() applied after the state change is complete.
    this.setState({ startDate: date }, () => {
      this.applyRecurrence();
    });
  }

  /**
   *
   *
   * @private
   * @param {Date} date
   * @memberof EventRecurrenceInfoDaily
   */
   private onEndDateChange(date: Date) {
    //Put the applyRecurrence() function in the callback of the setState() method to make sure that applyRecurrence() applied after the state change is complete.
    this.setState({ endDate: date}, () => {
        this.applyRecurrence();
      }
    );
  }
  /**
   *
   *
   * @private
   * @param {React.SyntheticEvent<HTMLElement>} ev
   * @param {string} value
   * @memberof EventRecurrenceInfoDaily
   */
  private onNumberOfWeeksChange(ev: React.SyntheticEvent<HTMLElement>, value: string) {
    ev.preventDefault();
    setTimeout(() => {
      let errorMessage:string ='';
      if (Number(value.trim()) == 0 || Number(value.trim()) > 255) {
        value = '1  ';
        errorMessage = 'Allowed values 1 to 255';
      }
      this.setState({  numberOfWeeks: value , errorMessageNumberOfWeeks: errorMessage  });
      this.applyRecurrence();
    }, 2000);


  }


  /**
   *
   *
   * @private
   * @param {React.SyntheticEvent<HTMLElement>} ev
   * @param {string} value
   * @memberof EventRecurrenceInfoDaily
   */
  private onNumberOfOcurrencesChange(ev: React.SyntheticEvent<HTMLElement>, value: string) {
    ev.preventDefault();
    setTimeout(() => {
      this.setState({ numberOcurrences: value.trim().length > 0 ? value : "10 " });
      this.applyRecurrence();
    }, 2000);

  }

  /**
   *
   *
   * @private
   * @param {React.SyntheticEvent<HTMLElement>} ev
   * @param {IChoiceGroupOption} option
   * @memberof EventRecurrenceInfoDaily
   */
   private onDataRangeOptionChange(
    ev: React.SyntheticEvent<HTMLElement>,
    option: IChoiceGroupOption
  ): void {
    ev.preventDefault();
    //Put the applyRecurrence() function in the callback of the setState() method to make sure that applyRecurrence() applied after the state change is complete.
    this.setState(
      {
        selectdateRangeOption: option.key,
        disableNumberOcurrences: option.key == "endAfter" ? false : true,
        disableEndDate: option.key == "endDate" ? false : true,
      },
      () => {
        this.applyRecurrence();
      }
    );
  }

  private onPaternChange(
    ev: React.SyntheticEvent<HTMLElement>,
    option: IChoiceGroupOption
  ): void {
    ev.preventDefault();
    //Put the applyRecurrence() function in the callback of the setState() method to make sure that applyRecurrence() applied after the state change is complete.
    this.setState(
      {
        selectPatern: option.key,
        disableNumberOfWeeks: option.key == "every" ? false : true,
      },
      () => {
        this.applyRecurrence();
      }
    );
  }

  public async componentWillMount() {
    //  await this.load();
    await this.load();
  }


  public async  componentDidUpdate(prevProps: IEventRecurrenceInfoWeeklyProps, prevState: IEventRecurrenceInfoWeeklyState) {

  }

  private async load() {
    let patern: any = {};
    let dateRange: { repeatForever?: string, repeatInstances?: string, windowEnd?: Date } = {};
    let weeklyPatern: { weekFrequency?: string, su?: boolean, mo?: boolean, tu?: boolean, we?: boolean, th?: boolean, fr?: boolean, sa?: boolean } = {};


    if (this.props.recurrenceData) {

      const XMLParser: any = require('xml2js');

      XMLParser.parseString(this.props.recurrenceData, { explicitArray: false }, (error:any, result:any) => {

        if (result.recurrence.rule.repeat) {
          patern = result.recurrence.rule.repeat;
        }

        //
        if (result.recurrence.rule.repeatForever) {
          dateRange = { repeatForever: result.recurrence.rule.repeatForever };
        }
        if (result.recurrence.rule.repeatInstances) {
          dateRange = { repeatInstances: result.recurrence.rule.repeatInstances };
        }
        if (result.recurrence.rule.windowEnd) {
          dateRange = { windowEnd: result.recurrence.rule.windowEnd };
        }

      });
      // daily Patern
      if (patern.weekly) {

        weeklyPatern = patern.weekly.$.weekFrequency ? { weekFrequency: patern.weekly.$.weekFrequency } : { weekFrequency: 1 };
        const weeklysu = patern.weekly.$.su ? true : false;
        const weeklymo = patern.weekly.$.mo ? true : false;
        const weeklytu = patern.weekly.$.tu ? true : false;
        const weeklywe = patern.weekly.$.we ? true : false;
        const weeklyth = patern.weekly.$.th ? true : false;
        const weeklyfr = patern.weekly.$.fr ? true : false;
        const weeklysa = patern.weekly.$.sa ? true : false;
        weeklyPatern = { su: weeklysu, mo: weeklymo, tu: weeklytu, we: weeklywe, th: weeklyth, fr: weeklyfr, sa: weeklysa };

      }

      let selectDateRangeOption: string = 'noDate';
      if (dateRange.repeatForever) {
        selectDateRangeOption = 'noDate';
      } else if (dateRange.repeatInstances) {
        selectDateRangeOption = 'endAfter';
      } else if (dateRange.windowEnd) {
        selectDateRangeOption = 'endDate';
      }


      console.log(selectDateRangeOption, new Date(moment(dateRange.windowEnd).format('YYYY/MM/DD')));
      // weekday patern
      this.setState({
        weeklySunday: weeklyPatern.su ?? false,
        weeklyMonday: weeklyPatern.mo?? false,
        weekklyTuesday: weeklyPatern.tu?? false,
        weekklyWednesday: weeklyPatern.we?? false,
        weekklyThursday: weeklyPatern.th?? false,
        weeklyFriday: weeklyPatern.fr?? false,
        weeklySaturday: weeklyPatern.sa?? false,
        selectPatern: weeklyPatern.weekFrequency?? '',
        numberOfWeeks: weeklyPatern.weekFrequency ? weeklyPatern.weekFrequency : '1',
        selectdateRangeOption: selectDateRangeOption,
        numberOcurrences: dateRange.repeatInstances ? dateRange.repeatInstances : '1',
        disableNumberOcurrences: dateRange.repeatInstances ? false : true,
        endDate: dateRange.windowEnd ? new Date(moment(dateRange.windowEnd).format('YYYY/MM/DD')) : this.state.endDate,
        disableEndDate: dateRange.windowEnd ? false : true,
        isLoading: false,
      },
      async () => await this.applyRecurrence()
      );

    }
    else {
      await this.applyRecurrence();
    }
  }


  private async onApplyRecurrence(ev: React.MouseEvent<HTMLButtonElement>) {
    await this.applyRecurrence();
  }
  /**
   *
   *
   * @private
   * @param {React.MouseEvent<HTMLButtonElement>} ev
   * @memberof EventRecurrenceInfoDaily
   */
  private async applyRecurrence() {
    const endDate = await this.getUtcTime(this.state.endDate);
    let selectDateRangeOption;
    switch (this.state.selectdateRangeOption) {
      case 'noDate':
        selectDateRangeOption = `<repeatForever>FALSE</repeatForever>`;
        break;
      case 'endAfter':
        selectDateRangeOption = `<repeatInstances>${this.state.numberOcurrences.trim()}</repeatInstances>`;
        break;
      case 'endDate':
        selectDateRangeOption = `<windowEnd>${endDate}</windowEnd>`;
        break;
      default:
        break;
    }

    // test weekDays
    let weekdays: string = '';
    if (this.state.weeklySunday) {
      weekdays = 'su="TRUE" ';
    }
    if (this.state.weeklyMonday) {
      weekdays = `${weekdays} mo="TRUE"`;
    }
    if (this.state.weekklyTuesday) {
      weekdays = `${weekdays} tu="TRUE"`;
    }
    if (this.state.weekklyWednesday) {
      weekdays = `${weekdays} we="TRUE"`;
    }
    if (this.state.weekklyThursday) {
      weekdays = `${weekdays} th="TRUE"`;
    }
    if (this.state.weeklyFriday) {
      weekdays = `${weekdays} fr="TRUE"`;
    }
    if (this.state.weeklySaturday) {
      weekdays = `${weekdays} sa="TRUE"`;
    }
    const recurrenceXML = `<recurrence><rule><firstDayOfWeek>su</firstDayOfWeek><repeat>` +
      `<weekly ${weekdays} weekFrequency="${this.state.numberOfWeeks.trim()}" /></repeat>${selectDateRangeOption}</rule></recurrence>`;
    console.log(recurrenceXML);
    //change
    this.props.returnRecurrenceData(this.state.startDate,this.state.endDate, recurrenceXML,this.state?.selectedPeople,this.state?.selectedPeopleId,this.state?.selectedPeopleEmail,this.state.PatternType,this.state.selectdateRangeOption,this.state?.UserDataIndex);
  }
  /**                                                                                 
   *
   * @private
   * @returns {Promise<string>}
   * @memberof spservices
   */
public async getUtcTime(date: string | Date): Promise<string> {
    try {
      // Initialize PnPjs SPFI with your site URL and properly set up the observer
      const sp = spfi(this.props.siteUrl).using(SPFx(this.props.context));

      // Ensure date is a Date object
      const dateObj = typeof date === "string" ? new Date(date) : date;

      // Convert local time to UTC using the web's regional settings
      const utcTime = await sp.web.regionalSettings.timeZone.localTimeToUTC(dateObj);

      // Return as ISO string, or format as needed
      return new Date(utcTime).toISOString();
    } catch (error) {
      console.error("Error converting time to UTC:", error);
      return Promise.reject(error);
    }
  }

  private async onCheckboxSundayChange(ev: React.FormEvent<HTMLElement>, isChecked: boolean) {
    this.setState({ weeklySunday: isChecked });
    await this.applyRecurrence();
  }
  private async onCheckboxMondayChange(ev: React.FormEvent<HTMLElement>, isChecked: boolean) {
    this.setState({ weeklyMonday: isChecked });
    await this.applyRecurrence();
  }
  private async onCheckboxTuesdayChange(ev: React.FormEvent<HTMLElement>, isChecked: boolean) {
    this.setState({ weekklyTuesday: isChecked });
    await this.applyRecurrence();
  }
  private async onCheckboxWednesdayChange(ev: React.FormEvent<HTMLElement>, isChecked: boolean) {
    this.setState({ weekklyWednesday: isChecked });
    await this.applyRecurrence();
  }
  private async onCheckboxThursdayChange(ev: React.FormEvent<HTMLElement>, isChecked: boolean) {
    this.setState({ weekklyThursday: isChecked });
    await this.applyRecurrence();
  }
  private async onCheckboxFridayChange(ev: React.FormEvent<HTMLElement>, isChecked: boolean) {
    this.setState({ weeklyFriday: isChecked });
    await this.applyRecurrence();
  }
  private async onCheckboxSaturdayChange(ev: React.FormEvent<HTMLElement>, isChecked: boolean) {
    this.setState({ weeklySaturday: isChecked });
    await this.applyRecurrence();
  }
  /**
   *
   *
   * @returns {React.ReactElement<IEventRecurrenceInfoWeeklyProps>}
   * @memberof EventRecurrenceInfoWeekly
   */
  public render(): React.ReactElement<IEventRecurrenceInfoWeeklyProps> {
    return (
      <div >
        {
          <div>
            <div style={{ display: 'inline-block', float: 'right', paddingTop: '10px', height: '40px' }}>

            </div>
            <div style={{ width: '100%', paddingTop: '10px' }}>
              <Label className='border-bottom mb-1'>{strings.PaternLabel}</Label>
              <div className='alignCenter'>
                <Label styles={{ root: { display: 'inline-block', verticalAlign: 'top', width: '40px' } }}>{strings.every}</Label>
                <MaskedTextField
                  styles={{ root: { display: 'inline-block', verticalAlign: 'top', width: '100px', paddingLeft: '5px' } }}
                  mask="999"
                  maskChar=' '
                  errorMessage={this.state.errorMessageNumberOfWeeks}
                  value={this.state.numberOfWeeks}
                  onChange={this.onNumberOfWeeksChange} />
                <Label styles={{ root: { display: 'inline-block', verticalAlign: 'top', width: '80px', paddingLeft: '10px' } }}>{strings.WeeksOnLabel}</Label>

              </div>
              <div className='everyweek mt-2'>
                <Checkbox label="Sunday" className={styles.ckeckBoxInline} checked={this.state.weeklySunday} onChange={this.onCheckboxSundayChange} />
                <Checkbox label="Monday" className={styles.ckeckBoxInline} checked={this.state.weeklyMonday} onChange={this.onCheckboxMondayChange} />
                <Checkbox label="Tuesday" className={styles.ckeckBoxInline} checked={this.state.weekklyTuesday} onChange={this.onCheckboxTuesdayChange} />
                <Checkbox label="Wednesday" className={styles.ckeckBoxInline} checked={this.state.weekklyWednesday} onChange={this.onCheckboxWednesdayChange} />
              </div>
              <div className='everyweek mt-2'>
                <Checkbox label="Thursday" className={styles.ckeckBoxInline} checked={this.state.weekklyThursday} onChange={this.onCheckboxThursdayChange} />
                <Checkbox label="Friday" className={styles.ckeckBoxInline} checked={this.state.weeklyFriday} onChange={this.onCheckboxFridayChange} />
                <Checkbox label="Saturday" className={styles.ckeckBoxInline} checked={this.state.weeklySaturday} onChange={this.onCheckboxSaturdayChange} />
              </div>
            </div>

            <div className='mt-4'>
              <Label className='border-bottom'>{strings.dateRangeLabel}</Label>
              <div className='d-flex gap-5'>
              <div className={'styles.dateRange col-4'}>

                <DatePicker
                  firstDayOfWeek={DayOfWeek.Sunday}
                  strings={DayPickerStrings} 
                  placeholder={strings.StartDatePlaceHolder}
                  ariaLabel={strings.StartDatePlaceHolder}
                  label={strings.StartDateLabel}
                  value={this.state.startDate}
                  onSelectDate={this.onStartDateChange}
                  formatDate={toLocaleShortDateString}    
                />

              </div>
              <div className='col'>
                <ChoiceGroup
                  selectedKey={this.state.selectdateRangeOption}
                  onChange={this.onDataRangeOptionChange}
                  options={[
                    {
                      key: 'noDate',
                      text: strings.noEndDate,
                    },
                    {
                      key: 'endDate',
                      text: strings.EndByLabel,
                      onRenderField: (props, render) => {
                        return (
                          <div className='alignCenter'>
                            {render!(props)}
                            <DatePicker
                              firstDayOfWeek={DayOfWeek.Sunday}
                              strings={DayPickerStrings}
                              placeholder={strings.StartDatePlaceHolder}
                              ariaLabel={strings.StartDatePlaceHolder}
                              style={{ display: 'inline-block', verticalAlign: 'top', paddingLeft: '22px', width:'250px'}}
                              onSelectDate={this.onEndDateChange}
                              formatDate={toLocaleShortDateString}
                              value={this.state.endDate}
                              disabled={this.state.disableEndDate}
                            />
                          </div>
                        );
                      }
                    },
                    {
                      key: 'endAfter',
                      text: strings.EndAfterLabel,
                      onRenderField: (props, render) => {
                        return (
                          <div className='alignCenter'>
                            {render!(props)}
                            <MaskedTextField
                              styles={{ root: { display: 'inline-block', verticalAlign: 'top', width: '100px', paddingLeft: '10px' } }}
                              mask="999"
                              maskChar=' '
                              value={this.state.numberOcurrences}
                              disabled={this.state.disableNumberOcurrences}
                              onChange={this.onNumberOfOcurrencesChange} />
                            <Label styles={{ root: { display: 'inline-block', verticalAlign: 'top', paddingLeft: '10px' } }}>{strings.OcurrencesLabel}</Label>
                          </div>
                        );
                      }
                    },
                  ]}
                  required={true}
                />
              </div>
              </div>
            </div>
          </div>
        }
      </div>
    );
  }
}