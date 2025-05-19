export interface IEventRecurrenceInfoState {
  selectedKey:string;
  selectPatern:string;
  startDate: Date;
  endDate:Date;
  numberOcurrences:string;
  numberOfDays:string;
  disableNumberOfDays: boolean;
  disableNumberOcurrences: boolean;
  selectdateRangeOption:string;
  disableEndDate:boolean;
  selectedRecurrenceRule:string;
  removeRecurrence:boolean,
  PeopleId:any,
  PeopleName:any,
  PeopleEmail:any
  defaultSelectedUsers:any
  UserDataIndex : any
  parsedRecurrenceData?:any
}