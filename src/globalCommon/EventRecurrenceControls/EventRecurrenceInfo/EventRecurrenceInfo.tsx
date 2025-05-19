// Modified EventRecurrenceInfo.tsx with improved recurrence pattern parsing
import * as React from 'react';
import styles from './EventRecurrenceInfo.module.scss';
import strings from '../constants/strings';
import { IEventRecurrenceInfoProps } from './IEventRecurrenceInfoProps';
import { IEventRecurrenceInfoState } from './IEventRecurrenceInfoState';
import * as moment from 'moment';
import {
  ChoiceGroup,
  IChoiceGroupOption,
} from 'office-ui-fabric-react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker"
import { SPHttpClient } from "@microsoft/sp-http";

import { EventRecurrenceInfoDaily } from './../EventRecurrenceInfoDaily/EventRecurrenceInfoDaily';
import { EventRecurrenceInfoWeekly } from './../EventRecurrenceInfoWeekly/EventRecurrenceInfoWeekly';
import { EventRecurrenceInfoMonthly } from './../EventRecurrenceInfoMonthly/EventRecurrenceInfoMonthly';
import { EventRecurrenceInfoYearly } from './../EventRecurrenceInfoYearly/EventRecurrenceInfoYearly';
import { RecurrenceDataParser, IRecurrenceSettings } from './RecurrenceDataParser';

let selectedPeople: any
let userData: any = [];
let chnagededUserIds :any = [];

export class EventRecurrenceInfo extends React.Component<IEventRecurrenceInfoProps, IEventRecurrenceInfoState> {
  private parsedRecurrenceSettings: IRecurrenceSettings | null = null;
  
  public constructor(props: IEventRecurrenceInfoProps) {
    super(props);

    this._onRecurrenceFrequenceChange = this._onRecurrenceFrequenceChange.bind(this);

    this.state = {
      selectedKey: props?.selectedKey,
      selectPatern: 'every',
      startDate: this.props?.startDate != undefined ? this.props.startDate : moment().toDate(),
      endDate: props?.DueDate != undefined ? props?.DueDate : moment().endOf('month').toDate(),
      numberOcurrences: '1',
      numberOfDays: '1',
      disableNumberOfDays: false,
      disableNumberOcurrences: true,
      selectdateRangeOption: 'noDate',
      disableEndDate: true,
      selectedRecurrenceRule: props?.selectedRecurrenceRule,
      removeRecurrence: props?.removeRecurrence,
      PeopleId: props?.recurrenceData?.userId ? props?.recurrenceData?.userId : "",
      PeopleName: props?.recurrenceData?.userTitle ? props?.recurrenceData?.userTitle : "",
      PeopleEmail: props?.recurrenceData?.userEmail ? props?.recurrenceData?.userEmail : " ",
      defaultSelectedUsers: this.props?.removeRecurrence ? [props?.recurrenceData?.userEmail] : [],
      UserDataIndex: this.props.dataIndex,
      // Add parsed recurrence pattern to state
      parsedRecurrenceData: null
    };
  }

  private _onRecurrenceFrequenceChange(ev: React.SyntheticEvent<HTMLElement>, option: IChoiceGroupOption): void {
    this.setState({
      selectedRecurrenceRule: option.key
    });
  }

  /**
   * Parse recurrence XML data into structured object
   */
  private parseRecurrenceData(recurrenceData: string): void {
    if (!recurrenceData) return;
    
    try {
      this.parsedRecurrenceSettings = RecurrenceDataParser.parseRecurrenceData(recurrenceData);
      this.setState({ 
        parsedRecurrenceData: this.parsedRecurrenceSettings,
        selectedRecurrenceRule: this.parsedRecurrenceSettings?.recurrenceType || ''
      });
      
      console.log('Parsed recurrence settings:', this.parsedRecurrenceSettings);
    } catch (error) {
      console.error('Error parsing recurrence data:', error);
    }
  }

  /**
   * Component mount lifecycle method
   */
  componentDidMount() {
    // Parse recurrence data if available
    if (this.props.recurrenceData) {
      this.parseRecurrenceData(this.props.recurrenceData);
    } else {
      this.handaleselectedRecurrenceRule();
    }
  }

  handaleselectedRecurrenceRule() {
    if (!this.props?.removeRecurrence) {
      this.setState({ selectedRecurrenceRule: '' });
    } else {
      if (this.props?.recurrenceData?.recurrenceData) {
        if (this.props?.recurrenceData?.recurrenceData.indexOf('<daily') != -1) {
          this.setState({ selectedRecurrenceRule: 'daily' });
        }
        if (this.props?.recurrenceData?.recurrenceData.indexOf('<weekly') != -1) {
          this.setState({ selectedRecurrenceRule: 'weekly' });
        }
        if (this.props?.recurrenceData?.recurrenceData.indexOf('<monthly') != -1) {
          this.setState({ selectedRecurrenceRule: 'monthly' });
        }
        if (this.props?.recurrenceData?.recurrenceData.indexOf('<monthlyByDay') != -1) {
          this.setState({ selectedRecurrenceRule: 'monthly' });
        }
        if (this.props?.recurrenceData?.recurrenceData.indexOf('<yearly') != -1) {
          this.setState({ selectedRecurrenceRule: 'yearly' });
        }
      }
    }
  }

  people = async (people: any) => {
    let userId: any[] = [];
    let userTitle: any[] = [];
    let userMail: any[] = [];

    if (people?.length > 0) {
      people.forEach((item: any) => {
        if (item?.id !== undefined) {
          userMail.push(item.id.split("|")[2]);
        }
      });

      if (userMail?.length > 0) {
        let userInfo = await this.getUserInfo(userMail);
        userData = userInfo
        if (userInfo && userInfo.length > 0) {
          userInfo.forEach((item: any) => {
            if (item?.Title !== undefined) {
              userTitle.push(item.Title);
              userId.push(item.Id);
            }
          });

          let isValidUser = this.props.recurrenceDataInfo.some((userData:any)=> userData.userId == userId?.[0])
          if(isValidUser) {
            alert("Taggeduser")
          } else {
            this.setState({
              PeopleId: userId?.[0],
              PeopleName: userTitle?.[0],
              PeopleEmail: userMail?.[0]
            })
          }
        }
      }
    } else {
      let userInfo = await this.getUserInfo(
        this.props.context._pageContext._legacyPageContext.userPrincipalName);
      if (userInfo && userInfo.length > 0) {
        userInfo.forEach((item: any) => {
          userTitle.push(item.Title);
          userId.push(item.Id);
        });

        let isValidUser = this.props.recurrenceDataInfo.some((userData:any)=> userData.userId == userId?.[0])
        if(isValidUser) {
          alert("Taggeduser")
        } else {
          this.setState({
            PeopleId: userId,
            PeopleName: userTitle,
            PeopleEmail: userMail
          })
        }
      } else {
        chnagededUserIds.push(this.state.PeopleId)
        this.props.setDeletedUserID(chnagededUserIds)
        this.setState({
          PeopleId: "",
          PeopleName: "",
          PeopleEmail: ""
        })
      }
    }
  };

  getUserInfo = async (userMails: string[]): Promise<any[]> => {
    const userInfoArray: any[] = [];
    const userEndPoint: string = `${this.props.context.pageContext.web.absoluteUrl}/_api/Web/EnsureUser`;
    try {
      const requests = userMails.map(async (userMail) => {
        const userData: string = JSON.stringify({ logonName: userMail });
        const userReqData = { body: userData };
        const resUserInfo: any = await this.props.context.spHttpClient.post(userEndPoint, SPHttpClient.configurations.v1, userReqData);
        if (!resUserInfo.ok) {
          throw new Error(`Failed to fetch user info for ${userMail}`);
        }
        const userInfo: any = await resUserInfo.json();
        return userInfo;
      });
      const userInfos = await Promise.all(requests);
      userInfoArray.push(...userInfos);
    } catch (error) {
      console.error('Error fetching user info:', error);
    }
    return userInfoArray;
  };

  /**
   * Render the component
   */
  public render(): React.ReactElement<IEventRecurrenceInfoProps> {
    return (
      <div className={styles.divWrraper}>
        <div className='d-flex justify-content-between'>
          <ChoiceGroup
            label={strings.recurrenceInformationLabel}
            selectedKey={this.props?.removeRecurrence ? this.state.selectedRecurrenceRule : ""}
            disabled={this.props?.removeRecurrence ? false : true}
            options={[
              {
                key: 'daily',
                iconProps: { iconName: 'CalendarDay' },
                text: strings.dailyLabel
              },
              {
                key: 'weekly',
                iconProps: { iconName: 'CalendarWeek' },
                text: strings.weeklyLabel
              },
              {
                key: 'monthly',
                iconProps: { iconName: 'Calendar' },
                text: strings.monthlyLabel,
              },
              {
                key: 'yearly',
                iconProps: { iconName: 'Calendar' },
                text: strings.yearlyLabel,
              }
            ]}
            onChange={this._onRecurrenceFrequenceChange}
          />
            
          {this.props?.recurrenceData?.userId && this.props?.useFor == "EditTaskPopup" && 
            <div className='deleteUserData mt-4'>
              <a className='hreflink siteColor me-1'>
                <span className='alignIcon svg__iconbox hreflink mini svg__icon--trash' title='Delete' onClick={() => this.props.deleteUserRecurence(this.state.UserDataIndex)}></span>
              </a>
            </div>
          }  
        </div>

        {this?.props?.removeRecurrence && this.props?.useFor == "EditTaskPopup" &&
          <div className='PeoplePicker'>
            <PeoplePicker
              context={this.props.context}
              principalTypes={[PrincipalType.User]}
              personSelectionLimit={1}
              titleText="Select People"
              resolveDelay={1000}
              onChange={this.people}
              showtooltip={true}
              required={true}
              defaultSelectedUsers={[this.props?.recurrenceData?.userEmail]}
            />
          </div>
        }

        {/* Daily recurrence component */}
        {this.renderRecurrenceComponent('daily', EventRecurrenceInfoDaily)}

        {/* Weekly recurrence component */}
        {this.renderRecurrenceComponent('weekly', EventRecurrenceInfoWeekly)}

        {/* Monthly recurrence component */}
        {this.renderRecurrenceComponent('monthly', EventRecurrenceInfoMonthly)}

        {/* Yearly recurrence component */}
        {this.renderRecurrenceComponent('yearly', EventRecurrenceInfoYearly)}
      </div>
    );
  }

  /**
   * Helper to render the appropriate recurrence component based on type
   */
  private renderRecurrenceComponent(type: string, Component: any) {
    const shouldRender = this.state.selectedRecurrenceRule === type && this.props?.removeRecurrence;
    
    if (!shouldRender) return null;
    
    // Determine if we need to check for people selection based on the useFor prop
    const needsPeopleCheck = this.props.useFor === "EditTaskPopup";
    const hasPeople = !needsPeopleCheck || this.state.PeopleName !== "";
    
    if (needsPeopleCheck && !hasPeople) return null;
    
    const commonProps = {
      display: true,
      recurrenceData: this.props.recurrenceData?.recurrenceData,
      startDate: needsPeopleCheck ? this.state?.startDate : this.props.startDate,
      context: this.props.context,
      siteUrl: this.props.siteUrl,
      endDate: this?.props?.DueDate,
      DueDate: this?.props?.DueDate,
      returnRecurrenceData: this.props.returnRecurrenceInfo || (() => {}),
      // If parsed data is available, pass it to the component
      parsedRecurrenceSettings: this.parsedRecurrenceSettings
    };
    
    // Add people-specific props if needed
    const peopleProps = needsPeopleCheck ? {
      userName: this.state.PeopleName,
      userId: this.state.PeopleId,
      PeopleEmail: this.state.PeopleEmail,
      UserDataIndex: this.state.UserDataIndex
    } : {};
    
    return <Component {...commonProps} {...peopleProps} />;
  }
}