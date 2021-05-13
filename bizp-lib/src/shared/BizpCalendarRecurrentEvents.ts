import * as moment from 'moment';

import {
  IBizpEventDataSpec,
  IBizpRecurrence,
  IBizpRecurrenceDateRange,
  IBizpDailyRecurrence,
  IBizpWeeklyRecurrence,
  IBizpMonthlyRecurrence,
  IBizpYearlyRecurrence
} from "./IBizpSharedInterface";

import {
  formatString,
  cloneObj,
  lastWeekdayOfMonth,
  lastWeekendDayOfMonth,
  lastSpecificDayOfMonth,
  getDayOfMonth,
  getWeekdayOfMonth,
  getWeekendOfMonth,
  getUtcTime
} from "./BizpBasesvc";


var wEvents: IBizpEventDataSpec[] = [];
const wom = ['first', 'second', 'third', 'fourth'];
const wd = ['su', 'mo', 'tu', 'we', 'th', 'fr', 'sa'];
const weekday = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
  // start is start time and end is end time. Both in Date format
  // It removes events of type 3 because they are deleted.
  // It returns events with start and end dates as Date type.
  // handling of event of type 0 and 4: it strips of z for all day event to make it non-UTC time. For regular events, it return unmodified.
  // precondition: All times are in ISO string format using site time. No zone conversion should be done.
export async function parseSeriesEventsSP(events: IBizpEventDataSpec[], seriesEvents: IBizpEventDataSpec[], start: string, end: string): Promise<IBizpEventDataSpec[]> {
  wEvents = wEvents.concat(events);
  wEvents = wEvents.concat(seriesEvents);

  let full: IBizpEventDataSpec[]=[];
  let returnEvents: IBizpEventDataSpec[];
  const argEnd = (end && (end.length > 0))? moment(end): null;
  const argStart = (start && (start.length > 0))? moment(start): null;
  for (let i = 0; i < seriesEvents.length; i++) {
    // series data contains the XML specification of the series
    let seriesEndDate: any = argEnd;

    // <windowEnd>2007-05-31T22:00:00Z</windowEnd> appears when the recurrence End by is apecified in the recurrence spec
    if (seriesEvents[i].recurrenceData.indexOf('<windowEnd>') != -1) {
      let wDtEnd = seriesEvents[i].recurrenceData.substring(seriesEvents[i].recurrenceData.indexOf("<windowEnd>") + 11);
      wDtEnd = wDtEnd.substring(0, wDtEnd.indexOf('<'));
      const calEndDate:any = moment(wDtEnd);
      seriesEndDate =  (argEnd && argEnd.isAfter(calEndDate)) ? moment(calEndDate): argEnd;
    }
    const calStartDate = moment(seriesEvents[i].startDate);
    const seriesStartDate = (argStart && argStart.isAfter(calStartDate)) ? argStart : calStartDate;
    // Note that there is no corresponding WindowStart clause to indicate the start date for the recurrence.
    // The start date for the series is the event start date eventStart
    returnEvents = await parseSeriesSP (seriesEvents[i],seriesStartDate, seriesEndDate);
    full = full.concat(returnEvents);
    console.log("recur Title: " + seriesEvents[i].title + "i:" + i +  " - Full length: ",full.length);
  }
  return full;
}

// precondition: wEvents contains all events from SP for the calendar
// improvement needed: it does not need to filter - just needs to find if any match exists. It sshould stop on the first match
function recurrenceExceptionExists(masterSeriesItemId:string, date:string) {
  const found = wEvents.filter((el,i) => {
    const str1:string = moment(el.recurrenceID).format('YYYY MM DD');
    const str2:string = moment(date).format('YYYY MM DD');
    if ((str1 == str2) && (el.masterSeriesItemID == masterSeriesItemId) ) {
      return el;
    }
  });
  return found.length > 0 ? true : false;
}

export function generateEvent(mDate:any,e:IBizpEventDataSpec): IBizpEventDataSpec{
  let mEventEnd = moment(mDate);
  mEventEnd.seconds(mEventEnd.seconds() + e.duration);
  const ni:IBizpEventDataSpec = cloneObj(e);
  ni.startDate = mDate.toISOString();
  ni.endDate = mEventEnd.toISOString();
  ni.ID = e.ID;
  return ni;
}

/****************************** Parse event ********************************************/
export async function parseSeriesSP(e: IBizpEventDataSpec, start: any, end: any): Promise<IBizpEventDataSpec[]> {
  const series: IBizpRecurrence = await parseEventToDataSP(e);
  let er: IBizpEventDataSpec[]=[]; // return unfolded events

  // calculate the start and end times based on the time window and series time frame
  const mEnd:any = end ? end: moment(e.endDate);
  const mStart:any = start ? start : moment(e.startDate);
  console.log("parseSeriestSP - event e: " + JSON.stringify(e));


  switch (series.rule) {

    case 'daily':
      createDailySeriesEvents(e,series,start,end,er);
      break;
    case 'weekly':
      createWeeklySeriesEvents(e,series,start,end,er);
      break;
    case 'monthly':
      createMonthlySeriesEvents(e,series,start,end,er);
      break;
    case 'yearly':
      createYearlySeriesEvents(e,series,start,end,er);
      break;
    default:

 }

  console.log("Recur start Date: " + mStart.toISOString() +  " Recur end Date: " + mEnd.toISOString());
  return er;
}

// assumptions: startTimeW and endTimeW are defined according to the time window
// assumption startTimeW and endTimeW are moment objects
function createDailySeriesEvents(e: IBizpEventDataSpec,series:IBizpRecurrence,startTimeW:any,endTimeW:any,er:IBizpEventDataSpec[]) {
  const dailySeries:IBizpDailyRecurrence = series.dailyRecurrence;
  if (dailySeries.pattern == 'everyweekday') {
    // for every weekday pattern, transform into weekly series request
    let w:IBizpWeeklyRecurrence = {
      dateRangeInfo:dailySeries.dateRangeInfo,
      frequency: '1',
      daysOption: {sunday:false,monday:true,tuesday:true,wednesday:true,thursday:true,friday:true,saturday:false}
    };
    const updatedSeries:IBizpRecurrence = {...series};
    updatedSeries.weeklyRecurrence = w;
    createWeeklySeriesEvents(e,updatedSeries,startTimeW,endTimeW,er);
    return;
  }
  let total:number;
  // determine target start and end time
  let dateRange:IBizpRecurrenceDateRange = dailySeries.dateRangeInfo;
  let seriesStartDate:any = moment(dateRange.startDate);
  let {targetStartTime,targetEndTime,initInstances,instanceLimit} = initSeriesProcessing(series,dailySeries.dateRangeInfo,startTimeW,endTimeW,e.duration);
  total = initInstances;

  let mInit:any = moment(targetStartTime);
  // adjust the target end time to the end of the day
  targetEndTime.hours(23).minutes(59);
  const frequency:number = parseInt(dailySeries.frequency);
  let loop:boolean = true;
  while (loop) {
    total++;
    if (!recurrenceExceptionExists(e.ID.toString(), mInit.toISOString()) &&
      mInit.isBetween(startTimeW,targetEndTime,'minutes','[]')) {
      er.push(generateEvent(mInit,e));
    }
    mInit.add(frequency,'days');
    // two cases of stopping the loop:
    // target end time is reached
    // If instaces are specified, number of required instances are reached
    if ((mInit.isAfter(targetEndTime)) || (instanceLimit > 0 && instanceLimit <= total)) loop = false;
  }
}

// assumptions: startTimeW and endTimeW are defined according to the time window
// assumption startTimeW and endTimeW are moment objects
function createWeeklySeriesEvents(e: IBizpEventDataSpec,series:IBizpRecurrence,startTimeW:any,endTimeW:any,er:IBizpEventDataSpec[]) {
  let total:number;
  const weeklySeries:IBizpWeeklyRecurrence = series.weeklyRecurrence;
  // determine target start and end time
  let dateRange:IBizpRecurrenceDateRange = weeklySeries.dateRangeInfo;
  let seriesStartDate:any = moment(dateRange.startDate);
  let {targetStartTime,targetEndTime,initInstances,instanceLimit} = initSeriesProcessing(series,weeklySeries.dateRangeInfo,startTimeW,endTimeW,e.duration);
  total = initInstances;
  let mInit:any = moment(targetStartTime);

  // adjust the target end time to the end of the day
  targetEndTime.hours(23).minutes(59);
  const frequency:number = parseInt(weeklySeries.frequency);
  let loop:boolean = true;
  let selectedDate:any;
  let initDay:number = mInit.day();
  while (loop) {
    selectedDate = moment(mInit);
    for (let i = initDay; (i < 7); i++) {
//    for (let daysCheck of Object.values(weeklySeries.daysOption)) {
      if ((weeklySeries.daysOption[weekday[i]]) && (instanceLimit == 0 || instanceLimit > total)) {
        total++;
        // iterDate.add((i - initDay),'days');
        console.log ("selectedDate: " + selectedDate.toString() + " start time: " + startTimeW.toString() + " end time: " +  targetEndTime.toString() + " time between: " + selectedDate.isBetween(startTimeW,targetEndTime,'minute','[]'));
        if ((!recurrenceExceptionExists(e.ID.toString(), selectedDate.toISOString())) &&
          selectedDate.isBetween(startTimeW,targetEndTime,'minute','[]')) {
          er.push(generateEvent(selectedDate,e));
        }
      }
      selectedDate.add(1,'days');
    }
    mInit.add(((7 * frequency) - initDay),'days');
    initDay = 0;
    if ((mInit.isAfter(targetEndTime))  || (instanceLimit > 0 && instanceLimit <= total)) loop = false;
  }
}

function createMonthlySeriesEvents(e: IBizpEventDataSpec,series:IBizpRecurrence,startTimeW:any,endTimeW:any,er:IBizpEventDataSpec[]) {
  let total:number;
  const monthlySeries:IBizpMonthlyRecurrence = series.monthlyRecurrence;
  // determine target start and end time
  let dateRange:IBizpRecurrenceDateRange = monthlySeries.dateRangeInfo;
  let seriesStartDate:any = moment(dateRange.startDate);
  let {targetStartTime,targetEndTime,initInstances,instanceLimit} = initSeriesProcessing(series,monthlySeries.dateRangeInfo,startTimeW,endTimeW,e.duration);
  total = initInstances;
  let mInit:any = moment(targetStartTime);

  const frequency:number = parseInt(monthlySeries.frequency);
  // adjust the target end time to the end of the day
  targetEndTime.hours(23).minutes(59);
  let selectedDate:any;
  let dayOfMonth:number = parseInt(monthlySeries.dayOfMonth);
  let loop:boolean = true;

  while (loop) {
    // set it to the first of the month
    mInit.date(1);
    total++;
    selectedDate = selectMonthlyDate(monthlySeries,mInit);
    // make sure the selected date is not before the series. This can happen on the
    // first iteration
    if (total==1 && selectedDate.isBefore(seriesStartDate)) {
      mInit.add(frequency,'months');
      selectedDate = selectMonthlyDate(monthlySeries,mInit);
    }
    if ((!recurrenceExceptionExists(e.ID.toString(), selectedDate.toISOString())) &&
      selectedDate.isBetween(startTimeW,targetEndTime,'minute','[]')) {
      er.push(generateEvent(selectedDate,e));
    }
    mInit.add(frequency,'months');


    if ((mInit.isAfter(targetEndTime))  || (instanceLimit > 0 && instanceLimit <= total)) loop = false;
  }
}

function createYearlySeriesEvents(e: IBizpEventDataSpec,series:IBizpRecurrence,startTimeW:any,endTimeW:any,er:IBizpEventDataSpec[]) {
  let total:number;
  const yearlySeries:IBizpYearlyRecurrence = series.yearlyRecurrence;
  // determine target start and end time
  let dateRange:IBizpRecurrenceDateRange = yearlySeries.dateRangeInfo;
  let seriesStartDate:any = moment(dateRange.startDate);
  let {targetStartTime,targetEndTime,initInstances,instanceLimit} = initSeriesProcessing(series,dateRange,startTimeW,endTimeW,e.duration);

  total = initInstances;
  const frequency:number = 1;
  // adjust the target end time to the end of the day
  targetEndTime.hours(23).minutes(59);
  let selectedDate:any;
  let loop:boolean = true;
  let mInit:any = moment(targetStartTime);
  while (loop) {
    // set it to the first of the month
    mInit.date(1);
    total++;
    selectedDate = selectYearlyDate(yearlySeries,mInit);
    // make sure the selected date is not before the series. This can happen on the
    // first iteration
    if (total==1 && selectedDate.isBefore(seriesStartDate)) {
      mInit.add(frequency,'years');
      selectedDate = selectYearlyDate(yearlySeries,mInit);
    }
    if ((!recurrenceExceptionExists(e.ID.toString(), selectedDate.toISOString())) &&
      selectedDate.isBetween(startTimeW,targetEndTime,'minute','[]')) {
      er.push(generateEvent(selectedDate,e));
    }
    mInit.add(frequency,'years');
    if ((mInit.isAfter(targetEndTime))  || (instanceLimit > 0 && instanceLimit <= total)) loop = false;
  }
}

function selectMonthlyDate(monthlySeries:IBizpMonthlyRecurrence,mInit:any):any {
  let day:number;
  let selectedDate:any = moment(mInit);
  switch (monthlySeries.pattern) {
    case "dayOfMonth": {
        const dayChoice:number = parseInt(monthlySeries.dayOfMonth);
        // if the selected day > last day of the the month, then set the selected day to the last day of the month
        day = (dayChoice > mInit.daysInMonth()) ? mInit.daysInMonth() : dayChoice;
        selectedDate.date(day);
      }
      break;
    case "orderInMonth": {
        const frequency:number = parseInt(monthlySeries.frequency);
        switch (monthlySeries.dayOption) {
          case "weekday":
            selectedDate.date(getWeekdayOfMonth(monthlySeries.orderInMonth,selectedDate));
            break;
          case "weekendday":
            selectedDate.date(getWeekendOfMonth(monthlySeries.orderInMonth,selectedDate),'date');
            break;
          case "day":
              //looking for the Nth day in the month...
              if (monthlySeries.orderInMonth == 'last') {

                selectedDate.date(mInit.daysInMonth());
              }
              //for first...fourth, add days to get to the Nth instance of this day
              else selectedDate.date(wom.indexOf(monthlySeries.orderInMonth)+1);
             break;
          default:
              // this is for the case of a specific day: sunday, monday ... or saturday
              // looking for a specific day of the week
              // find the day of the week
              let d:number;
              for (let i: number = 0; i < weekday.length; i++) { //get first instance of the specified day
                if (weekday[i] == monthlySeries.dayOption) {
                  d = i;
                }
              }
              selectedDate.date(getDayOfMonth(monthlySeries.orderInMonth,selectedDate,d));
            break;
        }
      }
      break;
    default:
      break;
  }
  return selectedDate;
}

function selectYearlyDate(yearlySeries:IBizpYearlyRecurrence,mInit:any):any {
  let selectedDate:any = moment(mInit);
  if (yearlySeries.pattern == 'byDay') {
    selectedDate.month(parseInt(yearlySeries.month)-1).date(yearlySeries.dayOfMonth);
  }
  else {
    selectedDate.month(parseInt(yearlySeries.dayPatternMonth)-1);
    switch (yearlySeries.dayOption) {
      case "weekday":
        selectedDate.date(getWeekdayOfMonth(yearlySeries.orderInMonth,selectedDate));
        break;
      case "weekendday":
        selectedDate.date(getWeekendOfMonth(yearlySeries.orderInMonth,selectedDate),'date');
        break;
      case "day":
          //looking for the Nth day in the month...
          if (yearlySeries.orderInMonth == 'last') {
            selectedDate.date(selectedDate.daysInMonth());
          }
          //for first...fourth, add days to get to the Nth instance of this day
          else selectedDate.date(wom.indexOf(yearlySeries.orderInMonth)+1);
         break;
      default:
          // this is for the case of a specific day: sunday, monday ... or saturday
          // looking for a specific day of the week
          // find the day of the week
          let d:number;
          for (let i: number = 0; i < weekday.length; i++) { //get first instance of the specified day
            if (weekday[i] == yearlySeries.dayOption) {
              d = i;
            }
          }
          selectedDate.date(getDayOfMonth(yearlySeries.orderInMonth,selectedDate,d));
        break;
    }
  }
  return selectedDate;
}

// initialize processing by determining
// start date: processing will start from this time
// end date: processing will end at this time
// instanceLimit: 0, if there is no instance limit. Otherwise, specified limit
// init instances: processing will ignore these instances before the start date
function initSeriesProcessing(series:IBizpRecurrence,dateRange: IBizpRecurrenceDateRange,startTW:any,endTW:any,duration:number):
          {targetStartTime:any,targetEndTime:any,initInstances:number,instanceLimit:number} {
  let targetStartTime:any;
  let targetEndTime:any;
  let instanceLimit:number = 0;
  let initInstances:number = 0;
  // determine target start and end time
  let seriesStartDate:any = moment(dateRange.startDate);

  // determine if the window starts before the series
  const isStartBeforeSeries:boolean = startTW.isSameOrBefore(seriesStartDate);
  targetStartTime = isStartBeforeSeries ? seriesStartDate : startTW;
  switch (dateRange.option) {
    case 'endDate':{
        const seriesEndDate:any = moment(dateRange.endDate);
        targetEndTime = seriesEndDate.isBefore(endTW) ? moment(seriesEndDate) : moment(endTW);
      }
      break;
    case 'noDate':
      targetEndTime = moment(endTW);
      break;
    case 'endAfter': {
        // series time frame involves instances
        instanceLimit = parseInt(dateRange.frequency);
        // series starts after the start of time window,
        // so ignore instances before the target start
        // init instances to be ignored and adjusted target start time
        if (isStartBeforeSeries) {
          initInstances = 0;
        }
        else {
          switch (series.rule) {
            case 'daily':
              ({targetStartTime,initInstances} = processDailyInstances(series.dailyRecurrence,startTW));
              break;
            case 'weekly':
              ({targetStartTime,initInstances} = processWeeklyInstances(series.weeklyRecurrence,startTW));
              break;
            case 'monthly':
              ({targetStartTime,initInstances} = processMonthlyInstances(series.monthlyRecurrence,startTW));
              break;
            case 'yearly':
              ({targetStartTime,initInstances} = processYearlyInstances(series.yearlyRecurrence,startTW));
              break;
            default:
              break;
          }
         }
      }
      break;
    default:
      break;
  }
  if (dateRange.option == 'endDate') {
    const seriesEndDate:any = moment(dateRange.endDate);
    targetEndTime = seriesEndDate.isBefore(endTW) ? moment(seriesEndDate) : moment(endTW);
  }
  else {
    targetEndTime = moment(endTW);
  }
  return {targetStartTime,targetEndTime,initInstances,instanceLimit};
}

function processDailyInstances(dailySeries:IBizpDailyRecurrence,startTW:any):
          {targetStartTime:any,initInstances:number} {
  // init total instances to be ignored and adjust the start date for processing
  let seriesStartDate:any = moment(dailySeries.dateRangeInfo.startDate);
  const frequency:number = parseInt(dailySeries.frequency);
  const days:number = startTW.diff(seriesStartDate,'days');
  let initInstances:number;
  let targetStartTime:any;

  if (dailySeries.pattern == 'every') {
    initInstances = days/frequency;
    targetStartTime = moment(seriesStartDate).add(initInstances * frequency,'days');
  }
  else {
    // pattern 'everyweekday'
    initInstances = days/7 * 5;
    targetStartTime = moment(seriesStartDate).add(days/7 * 7,'days');
  }
  return {targetStartTime,initInstances};
}

function processWeeklyInstances(weeklySeries:IBizpWeeklyRecurrence,startTW:any):
          {targetStartTime:any,initInstances:number} {
  // init total instances to be ignored and adjust the start date for processing
  let seriesStartDate:any = moment(weeklySeries.dateRangeInfo.startDate);
  const frequency:number = parseInt(weeklySeries.frequency);
  const days:number = startTW.diff(seriesStartDate,'days');
  let totalWeekdays:number = 0;
  const weeklyPeriod = 7*frequency;
  const periods:number = days/weeklyPeriod;
  for (let daysCheck of Object.values(weeklySeries.daysOption)) {
    if (daysCheck) totalWeekdays++;
  }
  const initInstances:number = periods * totalWeekdays;
  const targetStartTime = moment(seriesStartDate).add(periods * weeklyPeriod,'days');
  return {targetStartTime,initInstances};
}

function processMonthlyInstances(monthlySeries:IBizpMonthlyRecurrence,startTW:any):
          {targetStartTime:any,initInstances:number} {
  // init total instances to be ignored and adjust the start date for processing
  let seriesStartDate:any = moment(monthlySeries.dateRangeInfo.startDate);
  const frequency:number = parseInt(monthlySeries.frequency);
  const months:number = startTW.diff(seriesStartDate,'months');
  const initInstances:number = months/frequency;
  const firstDate:any = moment(startTW).date(1);
  const targetStartTime = months==0 ? seriesStartDate : firstDate;
  return {targetStartTime,initInstances};
}

function processYearlyInstances(yearlySeries:IBizpYearlyRecurrence,startTW:any):
          {targetStartTime:any,initInstances:number} {
  // init total instances to be ignored and adjust the start date for processing
  let seriesStartDate:any = moment(yearlySeries.dateRangeInfo.startDate);
  const years:number = startTW.diff(seriesStartDate,'years');
  const initInstances:number = years;
  const firstOfYear:any = moment(startTW).month(0).date(1);
  const targetStartTime = years==0 ? seriesStartDate : firstOfYear;
  return {targetStartTime,initInstances};
}
/******************************  parseEventToDataSP  ***************************************/
// Constructs new data structures from the event data returned by SP.
// New data structures are used by application
export async function parseEventToDataSP(event:IBizpEventDataSpec ): Promise<IBizpRecurrence> {

  let str:string; // temp string
  let arr:string[]; // temp string array

  const recurStr:string = event.recurrenceData;

  // construct DateRange object
  let dateRange = {
    startDate: "",
    endDate: "",
    frequency: "",
    option: ""
  };

  let recurrence:IBizpRecurrence = {
    rule: "",
    dailyRecurrence: undefined,
    weeklyRecurrence: undefined,
    monthlyRecurrence: undefined,
    yearlyRecurrence: undefined
  };

  dateRange.startDate = event.startDate;
  if (recurStr.indexOf('<windowEnd>') != -1) {
    let wDtEnd = recurStr.substring(recurStr.indexOf("<windowEnd>") + 11);
    wDtEnd = wDtEnd.substring(0, wDtEnd.indexOf('<'));
    dateRange.endDate = wDtEnd;
    dateRange.option = 'endDate';
  }
  if (recurStr.indexOf('<repeatInstances>') != -1) {
    dateRange.option = 'endAfter';
    str = recurStr.substring(recurStr.indexOf("<repeatInstances>") + 17);
    dateRange.frequency = str.substring(0, str.indexOf('<'));
  }
  if (recurStr.indexOf('<repeatForever>') != -1) {
    dateRange.option = 'noDate';
  }
  // construct daily rule
  arr = getPatternTokens(recurStr,'daily');
  if (arr.length != 0) {
    recurrence.rule = 'daily';
    recurrence.dailyRecurrence = constructDailyRecurrence (arr,dateRange);
  }
  // construct weekly rule
  arr = getPatternTokens(recurStr,'weekly');
  if (arr.length != 0) {
    recurrence.rule = 'weekly';
    recurrence.weeklyRecurrence = constructWeeklyRecurrence(arr,dateRange);
  }
  // construct monthly rule
  arr = getPatternTokens(recurStr,'monthly');
  if (arr.length != 0) {
    recurrence.rule = 'monthly';
    recurrence.monthlyRecurrence = constructMonthlyRecurrence(arr,dateRange);
  }
  // construct monthlyByDay rule
  arr = getPatternTokens(recurStr,'monthlyByDay');
  if (arr.length != 0) {
    recurrence.rule = 'monthly';
    recurrence.monthlyRecurrence = constructMonthlyByDayRecurrence(arr,dateRange);
  }
  // construct yearly rule
  arr = getPatternTokens(recurStr,'yearly');
  // Yearly option: specific month and date of the month for recurrence
  if (arr.length != 0) {
    recurrence.rule = 'yearly';
    recurrence.yearlyRecurrence = constructYearlyRecurrence(arr,dateRange);
  }
  // construct yearlyByDay rule
  arr = getPatternTokens(recurStr,'yearlyByDay');
  // Yearly option: specific month and date of the month for recurrence
  if (arr.length != 0) {
    recurrence.rule = 'yearly';
    recurrence.yearlyRecurrence = constructYearlyByDayRecurrence(arr,dateRange);
  }
  return recurrence;
}

function constructDailyRecurrence(tokens:string[], dateRange:IBizpRecurrenceDateRange): IBizpDailyRecurrence{
  let returnObj: IBizpDailyRecurrence = {
    dateRangeInfo: dateRange,
    frequency: "",
    pattern:""
  };
  if (tokens.indexOf("dayFrequency") != -1) {
    returnObj.pattern = 'every';
    // option: every <noOfTimes=frquency>
    returnObj.frequency = tokens[tokens.indexOf("dayFrequency") + 1];
  }
  else if (tokens.indexOf("weekday") != -1) {
    // option: every <weekday>
    // for a weekday, make all weekdays true
    returnObj.pattern = 'everyweekday';
  }
  return returnObj;
}

function constructWeeklyRecurrence(tokens:string[], dateRange:IBizpRecurrenceDateRange):IBizpWeeklyRecurrence{
  let returnObj: IBizpWeeklyRecurrence = {
    dateRangeInfo: dateRange,
    frequency: "",
    daysOption:{sunday:false,monday:false,tuesday:false,wednesday:false,
                thursday:false,friday:false,saturday:false}
  };
  returnObj.frequency = tokens[tokens.indexOf("weekFrequency") + 1];
  //init days to false
  returnObj.daysOption.sunday = false;
  returnObj.daysOption.monday = false;
  returnObj.daysOption.tuesday = false;
  returnObj.daysOption.wednesday = false;
  returnObj.daysOption.thursday = false;
  returnObj.daysOption.friday = false;
  returnObj.daysOption.saturday = false;

  // check for days enabled for the week
  if (tokens.indexOf(wd[0]) != -1 ) {
    returnObj.daysOption.sunday = true;
  }
  if (tokens.indexOf(wd[1]) != -1 ) {
    returnObj.daysOption.monday = true;
  }
  if (tokens.indexOf(wd[2]) != -1 ) {
    returnObj.daysOption.tuesday = true;
  }
  if (tokens.indexOf(wd[3]) != -1 ) {
    returnObj.daysOption.wednesday = true;
  }
  if (tokens.indexOf(wd[4]) != -1 ) {
    returnObj.daysOption.thursday = true;
  }
  if (tokens.indexOf(wd[5]) != -1 ) {
    returnObj.daysOption.friday = true;
  }
  if (tokens.indexOf(wd[6]) != -1 ) {
    returnObj.daysOption.saturday = true;
  }
  returnObj.dateRangeInfo = dateRange;
  return returnObj;
}

function constructMonthlyByDayRecurrence(tokens:string[], dateRange:IBizpRecurrenceDateRange):IBizpMonthlyRecurrence{
  let returnObj: IBizpMonthlyRecurrence = {
    pattern:"",frequency:"",orderInMonth:"",dayOfMonth:"",dayOption:"",
    dateRangeInfo:dateRange};

  returnObj.pattern = 'orderInMonth';
  // when a specific day(1-31) of the month is selected for recurrence
  returnObj.frequency = tokens[tokens.indexOf("monthFrequency") + 1];
  // weekdayOfMonth can be first,second,third,fourth,last
  returnObj.orderInMonth = tokens[tokens.indexOf("weekdayOfMonth") + 1];
  // do for weekday: any day (Mon-Fri)
  if (tokens.indexOf("weekday") != -1) {
    returnObj.dayOption = 'weekday';
  }
  else if (tokens.indexOf("weekend_day") != -1) {
    returnObj.dayOption = 'weekendday';
  }
  else if (tokens.indexOf("day") != -1) {
    //just looking for the Nth day in the month...
    returnObj.dayOption = 'day';
  }
  else {
    // looking for a specific day of the week
    returnObj.dayOption = getWeekdayString(tokens);
  }
  return returnObj;
}

function constructMonthlyRecurrence(tokens:string[], dateRange:IBizpRecurrenceDateRange):IBizpMonthlyRecurrence{
  let returnObj: IBizpMonthlyRecurrence = {
    pattern:"",frequency:"",orderInMonth:"",dayOfMonth:"",dayOption:"",
    dateRangeInfo:dateRange};

  returnObj.pattern = 'dayOfMonth';
  // when a specific day(1-31) of the month is selected for recurrence
  returnObj.frequency = tokens[tokens.indexOf("monthFrequency") + 1];
  returnObj.dayOfMonth = tokens[tokens.indexOf("day") + 1];
  return returnObj;
}

function constructYearlyRecurrence(tokens:string[], dateRange:IBizpRecurrenceDateRange):IBizpYearlyRecurrence{
  let returnObj: IBizpYearlyRecurrence = {
    dateRangeInfo: dateRange,
    pattern:"",
    month: "",
    dayOfMonth: "",
    orderInMonth:"",
    dayOption:"",
    dayPatternMonth:""
  };
 // option: specific month and date of the month for recurrence
  returnObj.pattern = 'byDay';
  // get the specific month
  const month:number = (parseInt(tokens[tokens.indexOf("month") + 1]));
  returnObj.month = month.toString();
  // get the specific date of the month
  returnObj.dayOfMonth = tokens[tokens.indexOf("day") + 1];
  return returnObj;
}

function constructYearlyByDayRecurrence(tokens:string[], dateRange:IBizpRecurrenceDateRange):IBizpYearlyRecurrence{
  let returnObj: IBizpYearlyRecurrence = {
    dateRangeInfo: dateRange,
    pattern:"",
    month: "",
    dayOfMonth: "",
    orderInMonth:"",
    dayOption:"",
    dayPatternMonth:""
  };
 // option: specific month and date of the month for recurrence
  returnObj.pattern = 'byDayPattern';
  // get the specific month
  const month:number = (parseInt(tokens[tokens.indexOf("month") + 1]));
  returnObj.dayPatternMonth = month.toString();
  // weekdayOfMonth can be first,second,third,fourth,last
  returnObj.orderInMonth = tokens[tokens.indexOf("weekdayOfMonth") + 1];
  if (tokens.indexOf("weekday") != -1) {
    returnObj.dayOption = 'weekday';
  } else if (tokens.indexOf("weekend_day") != -1) {
    returnObj.dayOption = 'weekendday';
  } else if (tokens.indexOf("day") != -1) {
    returnObj.dayOption = 'day';
  } else {
    // looking for a specific day of the week
    returnObj.dayOption = getWeekdayString(tokens);
  }
  return returnObj;
}

function getPatternTokens(patternString:string,tokenStr:string):string[]{
  const i:number = patternString.indexOf("<" + tokenStr + " ");
  let str: string; // temp string
  let arr:string[] = [];
  if (i != -1) {
    str = patternString.substring(i);

    str = str.substring(tokenStr.length+2, str.indexOf('/>') - 1);
    arr = formatString(str);
  }
  return arr;
}

function getWeekdayString(arr:string[]):string {
  let returnStr: string;
  if (arr.indexOf(wd[0]) != -1 ) {
    returnStr = 'sunday';
  }
  if (arr.indexOf(wd[1]) != -1 ) {
    returnStr = 'monday';
  }
  if (arr.indexOf(wd[2]) != -1 ) {
    returnStr = 'tuesday';
  }
  if (arr.indexOf(wd[3]) != -1 ) {
    returnStr = 'wednesday';
  }
  if (arr.indexOf(wd[4]) != -1 ) {
    returnStr = 'thursday';
  }
  if (arr.indexOf(wd[5]) != -1 ) {
    returnStr = 'friday';
  }
  if (arr.indexOf(wd[6]) != -1 ) {
    returnStr = 'saturday';
  }
  return returnStr;
}

// purpose of this is to strip of Z at the end of the date.
// When it is all day event, Sharepoint stores it in UTC format with the same day starting from 00:00:00 to 23:59:00
// If converted from UTC, it may show the wrong display date, so this function strips of the UTC form to retain the date
export function UTC_to_NonUTC(date: any, allDay: any) {
  if (typeof date == 'string') {
    if (allDay) {
      if (date.lastIndexOf('Z') == date.length - 1) {
        const dt = date.substring(0, date.length - 1);
        return new Date(dt);
      }
      else return new Date(date);
    }
    else {
      return new Date(date);
    }
  }
  return date;
}

/************************** Prepare XML for recurrence ************************* */
// returns SP recurrence formatted xml for the series
export async function prepareRecurrenceXML(data:IBizpRecurrence): Promise<string> {
  let recurrenceXML:string = "";
  switch (data.rule) {
    case "daily":
      recurrenceXML =  await prepareDailyRecurrence(data.dailyRecurrence);
      break;
    case "weekly":
      recurrenceXML =  await prepareWeeklyRecurrence(data.weeklyRecurrence);
      break;
    case "monthly":
      recurrenceXML =  await prepareMonthlyRecurrence(data.monthlyRecurrence);
      break;
    case "yearly":
      recurrenceXML =  await prepareYearlyRecurrence(data.yearlyRecurrence);
      break;
    default:
  }
  return recurrenceXML;
}
// returns SP recurrence formatted xml for date range
async function prepareDateRangeRecurrence(data:IBizpRecurrenceDateRange): Promise<string>  {
  try {
    const endDate = await getUtcTime(data.endDate);
    let dateRangeXML:string;
    switch (data.option) {
      case 'noDate':
        dateRangeXML = `<repeatForever>FALSE</repeatForever>`;
        break;
      case 'endAfter':
        dateRangeXML = `<repeatInstances>${data.frequency}</repeatInstances>`;
        break;
      case 'endDate':
        dateRangeXML = `<windowEnd>${endDate}</windowEnd>`;
        break;
      default:
        break;
    }
    return dateRangeXML;
  }
  catch(error) {
    // error stuff
    console.dir(error);
    return Promise.reject(error);
  }
}
// returns SP recurrence formatted xml for daily series
async function prepareDailyRecurrence(data:IBizpDailyRecurrence): Promise<string> {
  try {
    const dateRangeXML = await prepareDateRangeRecurrence(data.dateRangeInfo);
    const str = (data.pattern === 'every') ? `dayFrequency="${data.frequency.trim()}"/>` : `weekday="TRUE"/>`;
    const xml = `<recurrence><rule><firstDayOfWeek>su</firstDayOfWeek><repeat>` + `<daily ` + str + `</repeat>${dateRangeXML}</rule></recurrence>`;
 //     `<daily ${ data.pattern === 'every' ? `dayFrequency="${data.frequency.trim()}"/>` : `weekday="TRUE"/>`}</repeat>${dateRangeXML}</rule></recurrence>`;
    return xml;
  }
  catch(error) {
    // error stuff
    console.dir(error);
    return Promise.reject(error);
  }
}
// returns SP recurrence formatted xml for weekly series
async function prepareWeeklyRecurrence(data:IBizpWeeklyRecurrence): Promise<string> {
  let weekdays: string = '';
  if (data.daysOption.sunday) {
    weekdays = 'su="TRUE" ';
  }
  if (data.daysOption.monday) {
    weekdays = `${weekdays} mo="TRUE"`;
  }
  if (data.daysOption.tuesday) {
    weekdays = `${weekdays} tu="TRUE"`;
  }
  if (data.daysOption.wednesday) {
    weekdays = `${weekdays} we="TRUE"`;
  }
  if (data.daysOption.thursday) {
    weekdays = `${weekdays} th="TRUE"`;
  }
  if (data.daysOption.friday) {
    weekdays = `${weekdays} fr="TRUE"`;
  }
  if (data.daysOption.saturday) {
    weekdays = `${weekdays} sa="TRUE"`;
  }
  try {
    const dateRangeXML = await prepareDateRangeRecurrence(data.dateRangeInfo);
    const xml = `<recurrence><rule><firstDayOfWeek>su</firstDayOfWeek><repeat>` +
    `<weekly ${weekdays} weekFrequency="${data.frequency.trim()}" /></repeat>${dateRangeXML}</rule></recurrence>`;
    console.log(" dateRangeXML: " + dateRangeXML);
    console.log(" Full xml: " + xml);
    return xml;
  }
  catch(error) {
    // error stuff
    console.dir(error);
    return Promise.reject(error);
  }
}
// returns SP recurrence formatted xml for monthly series
async function prepareMonthlyRecurrence(data:IBizpMonthlyRecurrence): Promise<string> {
  try {
    const dateRangeXML = await prepareDateRangeRecurrence(data.dateRangeInfo);
    let recurrencePattern: string = '';
    if (data.pattern == 'dayOfMonth') {
      recurrencePattern = `<monthly  monthFrequency="${data.frequency}" day="${data.dayOfMonth}" /></repeat>${dateRangeXML}</rule></recurrence>`;
    }
    if (data.pattern == 'orderInMonth') {
      recurrencePattern = `<monthlyByDay weekdayOfMonth="${data.orderInMonth}" `;
      switch (data.dayOption) {
        case 'day':
          recurrencePattern = recurrencePattern + `day="TRUE"`;
          break;
        case 'weekday':
          recurrencePattern = recurrencePattern + `weekday="TRUE"`;
          break;
        case 'weekendday':
          recurrencePattern = recurrencePattern + `weekend_day="TRUE"`;
          break;
        case 'sunday':
          recurrencePattern = recurrencePattern + `su="TRUE"`;
          break;
        case 'monday':
          recurrencePattern = recurrencePattern + `mo="TRUE"`;
          break;
        case 'tuesday':
          recurrencePattern = recurrencePattern + `tu="TRUE"`;
          break;
        case 'wednesday':
          recurrencePattern = recurrencePattern + `we="TRUE"`;
          break;
        case 'thursday':
          recurrencePattern = recurrencePattern + `th="TRUE"`;
          break;
        case 'friday':
          recurrencePattern = recurrencePattern + `fr="TRUE"`;
          break;
        case 'saturday':
          recurrencePattern = recurrencePattern + `sa="TRUE"`;
          break;
        default:
          break;
      }
      recurrencePattern = recurrencePattern + ` monthFrequency="${data.frequency}" /></repeat>"${dateRangeXML}"</rule></recurrence>`;
    }
    const xml = `<recurrence><rule><firstDayOfWeek>su</firstDayOfWeek><repeat>` + recurrencePattern;
    console.log(" dateRangeXML: " + dateRangeXML);
    console.log(" Full xml: " + xml);
    return xml;
  }
  catch(error) {
    // error stuff
    console.dir(error);
    return Promise.reject(error);
  }
}
// returns SP recurrence formatted xml for yearly series
async function prepareYearlyRecurrence(data:IBizpYearlyRecurrence): Promise<string> {
  try {
    const dateRangeXML = await prepareDateRangeRecurrence(data.dateRangeInfo);
    let recurrencePattern: string = '';
    if (data.pattern == 'byDay') {
      recurrencePattern = `<yearly  yearFrequency="1" day="${data.dayOfMonth}" month="${data.month}" /></repeat>${dateRangeXML}</rule></recurrence>`;
    }
    if (data.pattern == 'byDayPattern') {
      recurrencePattern = `<yearlyByDay weekdayOfMonth="${data.orderInMonth}" month="${data.dayPatternMonth}" `;
      let dayOptionValue:string;
      switch (data.dayOption) {
        case 'day':
          dayOptionValue =  `day="TRUE"`;
          break;
        case 'weekday':
          dayOptionValue =  `weekday="TRUE"`;
          break;
        case 'weekendday':
          dayOptionValue =  `weekend_day="TRUE"`;
          break;
        case 'sunday':
          dayOptionValue =  `su="TRUE"`;
          break;
        case 'monday':
          dayOptionValue =  `mo="TRUE"`;
          break;
        case 'tuesday':
          dayOptionValue =  `tu="TRUE"`;
          break;
        case 'wednesday':
          dayOptionValue =  `we="TRUE"`;
          break;
        case 'thursday':
          dayOptionValue =  `th="TRUE"`;
          break;
        case 'friday':
          dayOptionValue =  `fr="TRUE"`;
          break;
        case 'saturday':
          dayOptionValue =  `sa="TRUE"`;
          break;
        default:
          break;
      }
      recurrencePattern = recurrencePattern + dayOptionValue + ` yearFrequency="1" /></repeat>${dateRangeXML}</rule></recurrence>`;
    }
    const xml = `<recurrence><rule><firstDayOfWeek>su</firstDayOfWeek><repeat>` + recurrencePattern;
    console.log(" dateRangeXML: " + dateRangeXML);
    console.log(" Full xml: " + xml);
    return xml;
  }
  catch(error) {
    // error stuff
    console.dir(error);
    return Promise.reject(error);
  }
}
