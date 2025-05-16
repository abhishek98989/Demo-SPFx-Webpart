import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface IEventRecurrenceInfoYearlyProps {
    display: boolean;
    recurrenceData: string;
    startDate: Date;
    context: WebPartContext;
    siteUrl: string;
    DueDate?:any;
    endDate?:any;
    returnRecurrenceData: (startDate: Date,endDate:Date, recurrenceData: string,userTitel: string, userId: number, PeopleEmail:string,PatternType:string,selectdateRangeOption:any, UserDataIndex : any) => void|any;
    deleteUserRecurence?:any
  useFor?:any
  recurrenceDataInfo?:any
  setAddForOther?:any
  userName?:any
  PeopleEmail?:any
  userId?:any
  UserDataIndex? : any
}