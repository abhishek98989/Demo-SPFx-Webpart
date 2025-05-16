import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface  IEventRecurrenceInfoDailyProps {
  display:boolean;
  recurrenceData: string;
  startDate:Date;
  context: WebPartContext;
  siteUrl:string;
  DueDate?:any;
  endDate?:any
  returnRecurrenceData: (startDate:Date,endDate:Date,recurrenceData:string,userTitel: string, userId: number,userEmail:string,PatternType:string,selectdateRangeOption:any,UserDataIndex:any) => void;
  userId?:any
  userName?:string,
  PeopleEmail?:string
  UserDataIndex? : any
  
}