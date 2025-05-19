import * as React from 'react';
//import styles from './EventRecurrenceInfoDaily.module.scss';
import strings from '../constants/strings';
import { IEventRecurrenceInfoDailyProps } from './IEventRecurrenceInfoDailyProps';
import { IEventRecurrenceInfoDailyState } from './IEventRecurrenceInfoDailyState';
//import { escape } from '@microsoft/sp-lodash-subset';
import * as moment from 'moment';
//import { parseString } from "xml2js";
import {
  ChoiceGroup,
  IChoiceGroupOption,
  Label,
  MaskedTextField,
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
  shortDays: [strings.ShortDay_S, strings.ShortDay_M, strings.ShortDay_T, strings.ShortDay_W, strings.ShortDay_T, strings.ShortDay_Friday, strings.ShortDay_S],
  goToToday: strings.GoToDay,
  prevMonthAriaLabel: strings.PrevMonth,
  nextMonthAriaLabel: strings.NextMonth,
  prevYearAriaLabel: strings.PrevYear,
  nextYearAriaLabel: strings.NextYear,
  closeButtonAriaLabel: strings.CloseDate,
  isRequiredErrorMessage: strings.IsRequired,
  invalidInputErrorMessage: strings.InvalidDateFormat,
};

/**
 *
 *
 * @export
 * @class EventRecurrenceInfoDaily
 * @extends {React.Component<IEventRecurrenceInfoDailyProps, IEventRecurrenceInfoDailyState>}
 */
export class EventRecurrenceInfoDaily extends React.Component<IEventRecurrenceInfoDailyProps, IEventRecurrenceInfoDailyState> {
  //private spService: spservices = null;
  public constructor(props: IEventRecurrenceInfoDailyProps) {
    super(props);


    this.onPatternChange = this.onPatternChange.bind(this);
    this.state = {
      selectedKey: 'daily',
      selectPatern: 'every',
      startDate: this.props.startDate ?? moment().toDate(),
      endDate: this.props?.DueDate ?? moment().endOf('month').toDate(),
      // endDate:props?.DueDate!=undefined?props?.DueDate:"",
      numberOcurrences: '1',
      numberOfDays: '1',
      disableNumberOfDays: false,
      disableNumberOcurrences: true,
      selectdateRangeOption: 'noDate',
      disableEndDate: true,
      selectedRecurrenceRule: 'daily',
      isLoading: false,
      errorMessageNumberOcurrences: '',
      errorMessageNumberOfDays: '',
      PatternType: "Daily",
      selectedPeople: this.props.userName,
      selectedPeopleEmail: this.props.PeopleEmail,
      selectedPeopleId: this.props.userId,
      UserDataIndex: this.props?.UserDataIndex
    };

    //
    this.onNumberOfDaysChange = this.onNumberOfDaysChange.bind(this);
    this.onNumberOfOcurrencesChange = this.onNumberOfOcurrencesChange.bind(this);
    this.onDataRangeOptionChange = this.onDataRangeOptionChange.bind(this);
    this.onEndDateChange = this.onEndDateChange.bind(this);
    this.onStartDateChange = this.onStartDateChange.bind(this);
    this.onApplyRecurrence = this.onApplyRecurrence.bind(this);

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
    this.setState({ endDate: date }, () => {
      this.applyRecurrence();
    });
  }

  /**
   *
   *
   * @private
   * @param {React.SyntheticEvent<HTMLElement>} ev
   * @param {string} value
   * @memberof EventRecurrenceInfoDaily
   */
  private onNumberOfDaysChange(ev: React.SyntheticEvent<HTMLElement>, value: string) {
    ev.preventDefault();
    let errorMessage = '';
    setTimeout(() => {

      if (Number(value.trim()) == 0 || Number(value.trim()) > 255) {
        value = '1  ';
        errorMessage = 'Allowed values 1 to 255';
      }
      this.setState({ numberOfDays: value, errorMessageNumberOfDays: errorMessage });
      this.applyRecurrence();
    }, 2500);

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
    let errorMessage = '';
    setTimeout(() => {

      if (Number(value.trim()) == 0 || Number(value.trim()) > 999) {
        value = '1  ';
        errorMessage = 'Allowed values 1 to 999';
      }
      this.setState({ numberOcurrences: value, errorMessageNumberOcurrences: errorMessage });
      this.applyRecurrence();
    }, 2500);

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
        // endDate: this.getNextDateAfter30DAys(this.state.startDate)

      },
      () => {

        this.applyRecurrence();
      }
    );
  }


  private onPatternChange(
    ev: React.SyntheticEvent<HTMLElement>,
    option: IChoiceGroupOption
  ): void {
    ev.preventDefault();
    //Put the applyRecurrence() function in the callback of the setState() method to make sure that applyRecurrence() applied after the state change is complete.
    this.setState(
      {
        selectPatern: option.key,
        disableNumberOfDays: option.key == "every" ? false : true,
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


  public async componentDidUpdate(prevProps: IEventRecurrenceInfoDailyProps, prevState: IEventRecurrenceInfoDailyState) {

  }

  private async load() {
    let patern: any = {};
    let dateRange: { repeatForever?: string, repeatInstances?: string, windowEnd?: Date } = {};
    let dailyPatern: { dayFrequency?: string, weekDay?: string } = {};
    let recurrenceRule: string = '';

    if (this.props.recurrenceData) {

      const XMLParser: any = require('xml2js');

      XMLParser.parseString(this.props.recurrenceData, { explicitArray: false }, (error: any, result: any) => {

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
      if (patern.daily) {
        recurrenceRule = 'daily';
        if (patern.daily.$.dayFrequency) {
          dailyPatern = { dayFrequency: patern.daily.$.dayFrequency };
        }
        if (patern.daily.$.weekday) {
          dailyPatern = { weekDay: 'weekDay' };
        }
      }

      let selectDateRangeOption: string = 'noDate';
      if (dateRange.repeatForever) {
        selectDateRangeOption = 'noDate';
      } else if (dateRange.repeatInstances) {
        selectDateRangeOption = 'endAfter';
      } else if (dateRange.windowEnd) {
        selectDateRangeOption = 'endDate';
      }

      // weekday patern
      this.setState({
        selectedRecurrenceRule: recurrenceRule,
        selectPatern: dailyPatern.dayFrequency ? 'every' : 'everweekday',
        numberOfDays: dailyPatern.dayFrequency ? dailyPatern.dayFrequency : '1',
        disableNumberOfDays: dailyPatern.dayFrequency ? false : true,
        selectdateRangeOption: selectDateRangeOption,
        numberOcurrences: dateRange.repeatInstances ? dateRange.repeatInstances : '10',
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
        selectDateRangeOption = `<repeatInstances>${this.state.numberOcurrences}</repeatInstances>`;
        break;
      case 'endDate':
        selectDateRangeOption = `<windowEnd>${endDate}</windowEnd>`;
        break;
      default:
        break;
    }
    const recurrenceXML = `<recurrence><rule><firstDayOfWeek>su</firstDayOfWeek><repeat>` +
      (this.state.selectPatern === 'every'
        ? `<daily dayFrequency="${this.state.numberOfDays.trim()}" />`
        : `<daily weekday="TRUE" />`) +
      `</repeat>${selectDateRangeOption}</rule></recurrence>`;
    //  console.log(recurrenceXML);
    //endDate change
    this.props.returnRecurrenceData(this.state.startDate, this.state.endDate, recurrenceXML, this.state?.selectedPeople, this.state?.selectedPeopleId, this.state?.selectedPeopleEmail, this.state.PatternType, this.state.selectdateRangeOption, this.state?.UserDataIndex);
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

  /**
   *
   *
   * @returns {React.ReactElement<IEventRecurrenceInfoDailyProps>}
   * @memberof EventRecurrenceInfoDaily
   */
  public render(): React.ReactElement<IEventRecurrenceInfoDailyProps> {
    return (
      <div >
        {
          <div>
            <div style={{ display: 'inline-block', float: 'right', paddingTop: '10px', height: '40px' }}>

            </div>
            <div style={{ width: '100%', paddingTop: '10px' }}>
              <Label className='Pattern border-bottom'>{strings.patternLabel}</Label>
              <div className='alignCenter'>
                <ChoiceGroup
                  selectedKey={this.state.selectPatern}
                  options={[
                    {
                      key: 'every',
                      text: strings.every,
                      ariaLabel: 'every',

                      onRenderField: (props, render) => {
                        return (
                          <div className='alignCenter'>
                            {render!(props)}
                            <MaskedTextField
                              styles={{ root: { display: 'inline-block', verticalAlign: 'top', width: '100px', paddingLeft: '10px' } }}
                              mask="999"
                              maskChar=' '
                              disabled={this.state.disableNumberOfDays}
                              value={this.state.numberOfDays}
                              errorMessage={this.state.errorMessageNumberOfDays}
                              onChange={this.onNumberOfDaysChange} />
                            <Label styles={{ root: { display: 'inline-block', verticalAlign: 'top', width: '60px', paddingLeft: '10px' } }}>{strings.days}</Label>
                          </div>
                        );
                      }
                    },
                    {
                      key: 'everweekday',
                      text: strings.everyweekdays,
                    }
                  ]}
                  onChange={this.onPatternChange}
                  required={true}
                />
              </div>
            </div>

            <div style={{ paddingTop: '22px' }}>
              <Label className='border-bottom'>{strings.dateRangeLabel}</Label>
              <div className='d-flex gap-5'>
                <div className='col-4'>

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
                                ariaLabel="Select a date"
                                style={{ display: 'inline-block', verticalAlign: 'top', paddingLeft: '22px', width: '250px' }}
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
                                errorMessage={this.state.errorMessageNumberOcurrences}
                                onChange={this.onNumberOfOcurrencesChange} />
                              <Label styles={{ root: { display: 'inline-block', verticalAlign: 'top', paddingLeft: '10px' } }}>{strings.occurrencesLabel}</Label>
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