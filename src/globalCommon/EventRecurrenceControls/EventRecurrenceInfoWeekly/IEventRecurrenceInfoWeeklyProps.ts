import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface  IEventRecurrenceInfoWeeklyProps {
  display:boolean;
  recurrenceData: string;
  startDate:Date;
  context: WebPartContext;
  siteUrl:string;
  DueDate?:any;
  endDate?:any;
  returnRecurrenceData: (startDate:Date,endDat:Date,erecurrenceData:string,userTitel: string, userId: number, PeopleEmail:string,PatternType:string,selectdateRangeOption:any,UserDataIndex:AnalyserOptions) => void;
  deleteUserRecurence?:any
  useFor?:any
  recurrenceDataInfo?:any
  setAddForOther?:any
  userName?:any
  PeopleEmail?:any
  userId?:any
  UserDataIndex? : any
}