import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface  IEventRecurrenceInfoMonthlyProps {
  display:boolean;
  recurrenceData: string;
  startDate:Date;
  context: WebPartContext;
  siteUrl:string;
  DueDate?:any;
  endDate?:any;
  returnRecurrenceData: (startDate:Date,endDate:Date,recurrenceData:string,userTitel: string, userId: number, PeopleEmail:string,PatternType:string,selectdateRangeOption:any,UserDataIndex:any) => void;
  userId?:any
  userName?:string
  PeopleEmail?:string
   deleteUserRecurence?:any
  useFor?:any
  recurrenceDataInfo?:any
  setAddForOther?:any
  UserDataIndex? : any
 
}