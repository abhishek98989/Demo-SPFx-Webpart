import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface IEventRecurrenceInfoProps {
  removeRecurrence:boolean
  display: boolean;
  recurrenceData: any,
  startDate: Date;
  endDate: Date;
  context: any;
  siteUrl: string;
  selectedKey: any;
  returnRecurrenceInfo:any;
  selectedRecurrenceRule: any;
  DueDate:any;
  returnRecurrenceData?: (startDate: Date, endDate: Date, recurrenceData: string,userTitel: string, userId: number, userEmail:string,PatternType:string,selectdateRangeOption:any,dataIndex:any) => void;
  deleteUserRecurence?:any
  useFor?:any
  recurrenceDataInfo?:any
  setAddForOther?:any
  dataIndex? : any
  setDeletedUserID?:any 
  removeChngeUserData?: any

}