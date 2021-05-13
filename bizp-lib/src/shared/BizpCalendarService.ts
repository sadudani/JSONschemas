import * as strings from 'BizpcompsLibraryStrings';
import { WebPartContext } from "@microsoft/sp-webpart-base";
require("@pnp/common");
import { sp, Web, PermissionKind, RegionalSettings,SiteUser,IFieldInfo, Site } from "@pnp/sp/presets/all";
import { getGUID } from '@pnp/common';
import * as moment from 'moment';
import {
  IBizpCalendarRequest,
  IBizpPeriodOptionTypes,
  IUserPermissions,
  IBizpEventDataSpec,
  IBizpOrderOptionTypes,
  IBizpRecurrence,
  IBizpNonRecurrence,
  IBizpDailyRecurrence,
  IBizpWeeklyRecurrence,
  IBizpMonthlyRecurrence,
  IBizpYearlyRecurrence,
  IBizpRecurrenceDateRange,
  ICalendarEventCategory
} from "./IBizpSharedInterface";
import {
  deCodeHtmlEntities,
  generateRandomColor,
  getChoiceFieldOptions,
  getSPItemById,
  getUserProfilePictureUrl,
  getSiteRegionalSettingsTimeZone,
  getSiteTimeZoneHours,
  getUtcTime,
  toLocaleLongDateString,
  toLocaleShortDateString
} from "./BizpBasesvc";

import {parseSeriesEventsSP,prepareRecurrenceXML} from './BizpCalendarRecurrentEvents';

/*
export async function getSiteRegionalSettingsTimeZone(siteUrl: string) {
  let regionalSettings: RegionalSettings;
  try {
    const web = Web(siteUrl);
   // const r = await web();
    regionalSettings = await web.regionalSettings.timeZone.usingCaching().get();

  } catch (error) {
    return Promise.reject(error);
  }
  return regionalSettings;
}
*/
/*
export async function getSPCalendarCategoryOptions(siteUrl: string, listId: string): Promise<{ key: string, text: string }[]> {
  const web = Web(siteUrl);
  const categoryDropdownOption = await sp.web.lists.getById(listId).fields.getByTitle("Category")();
  const categoryDropdownOption1 = await sp.web.lists.getChoiceFieldOptions(this.props.siteUrl, this.props.listId, 'Category');
  console.log("calendar categories" + JSON.stringify(categoryDropdownOption));
  return categoryDropdownOption;
}
*/
let categoryColors: ICalendarEventCategory[] = [];

export async function initCategoryColors(siteUrl: string, listId: string){
  const categoryOptions:string[] = await getChoiceFieldOptions(siteUrl, listId, 'Category');
  categoryColors = [];
  for (const cat of categoryOptions) {
    // hsv (Hue, Saturation, Value). pass Saturation as 0.3 and Value as 0.99 for lighter color scheme
    categoryColors.push({ category: cat, color: generateRandomColor(0.3,0.99) });
  }
  console.log("Category colors: " + JSON.stringify(categoryColors));
}

export function getCategoryColor(cat:string):string {
  let c: string;
  let e:ICalendarEventCategory[];
  if (cat && cat.length > 0) {
    e = categoryColors.filter((value) => {
      return value.category == cat;
    });
    c = e[0].color;
  //  c = categoryColors[cat];
  } else {
    c = '#1a75ff'; // blue default
  }
  return c;
}

export async function getSPCalendarEvent(siteUrl: string, listId: string, eventId: number): Promise<IBizpEventDataSpec> {
  let returnEvent: IBizpEventDataSpec = undefined;
  try {
//    const siteTimeZoneHours: number = await getSiteTimeZoneHours(siteUrl);
    const web = Web(siteUrl);
    const event = await web.lists.getById(listId).items.usingCaching().getById(eventId)
      .select("RecurrenceID", "MasterSeriesItemID", "Id", "ID", "ParticipantsPickerId", "EventType", "Title", "Description", "EventDate",
      "EventDate/FieldValuesAsText","EndDate", "EndDate/FieldValuesAsText", "EventDate/FieldValuesAsText","Location", "Author/SipAddress", "Author/Title",
      "Geolocation", "fAllDayEvent", "fRecurrence", "RecurrenceData", "RecurrenceData", "Duration", "Category", "UID")
      .expand("Author","FieldValuesAsText")
      .get();

    returnEvent = {
      id: event.ID,
      ID: event.ID,
      eventType: event.EventType,
      title: deCodeHtmlEntities(event.Title),
      description: event.Description ? event.Description : '',
      startDate:   moment(event["FieldValuesAsText"].EventDate).toISOString(),
      endDate:     moment(event["FieldValuesAsText"].EndDate).toISOString(),
      location: event.Location,
      ownerEmail: event.Author.SipAddress,
      ownerPhoto: "",
      ownerInitial: '',
      color: getCategoryColor(event.Category),
      ownerName: event.Author.Title,
      attendes: event.ParticipantsPickerId,
      fAllDayEvent: event.fAllDayEvent,
      geolocation: { Longitude: event.Geolocation ? event.Geolocation.Longitude : 0, Latitude: event.Geolocation ? event.Geolocation.Latitude : 0 },
      category: event.Category,
      duration: event.Duration,
      UID: event.UID,
      recurrenceData: event.RecurrenceData ? await deCodeHtmlEntities(event.RecurrenceData) : "",
      fRecurrence: event.fRecurrence,
      recurrenceID: event.RecurrenceID,
      masterSeriesItemID: event.MasterSeriesItemID,
    };
  }
  catch (error) {
    return Promise.reject(error);
  }
  return returnEvent;
}

/**
   *
   * @param {string} siteUrl
   * @param {string} listId
   * @param {Date} eventStartDate
   * @param {Date} eventEndDate
   * @returns {Promise< IEventData[]>}
   * @memberof spservices
   */
export async function getSPCalendarEvents(request:IBizpCalendarRequest): Promise<IBizpEventDataSpec[]>
{
    const siteUrl:string = request.siteURL;
    const listId:string = request.listName;
    const requestStartDate:string = moment(request.eventStartDate).format('YYYY-MM-DD');
    const requestEndDate:string = moment(request.eventEndDate).format('YYYY-MM-DD');

    const siteTimeZoneHours: number = await getSiteTimeZoneHours(siteUrl);

    console.log("getSPCalendarEvents - Site: "+ siteUrl + "list: " + listId);

    console.log("getSPCalendarEvents - request StartDate: " + requestStartDate + " request EndDate: " + requestEndDate);
    if (!siteUrl || !listId) {
      return [];
    }
    let events: IBizpEventDataSpec[] = [];
    try {
      /*
      // Get Category Field Choices
      const categoryOptions:string[] = await getChoiceFieldOptions(siteUrl, listId, 'Category');
      let categoryColor: { category: string, color: string }[] = [];
      for (const cat of categoryOptions) {
        categoryColor.push({ category: cat, color: generateRandomColor() });
      }
      */
      const web = new(Web as any) (siteUrl);
      // get all single events within the requested time frame (eventType 1 or 4)

      let results = await web.lists.getById(listId).usingCaching().renderListDataAsStream(
        {
          ViewXml: `<View><ViewFields><FieldRef Name='RecurrenceData'/><FieldRef Name='Duration'/><FieldRef Name='Author'/><FieldRef Name='Category'/><FieldRef Name='Color'/><FieldRef Name='Description'/><FieldRef Name='ParticipantsPicker'/><FieldRef Name='Geolocation'/><FieldRef Name='ID'/><FieldRef Name='FieldValuesAsText/EndDate'/><FieldRef Name='FieldValuesAsText/EventDate'/><FieldRef Name='ID'/><FieldRef Name='Location'/><FieldRef Name='Title'/><FieldRef Name='fAllDayEvent'/><FieldRef Name='EventType'/><FieldRef Name='UID' /><FieldRef Name='fRecurrence' /></ViewFields>
          <Query>
          <Where>
          <And>
            <And>
              <Geq>
                <FieldRef Name='EventDate' />
                <Value IncludeTimeValue='false' Type='DateTime'>${requestStartDate}</Value>
              </Geq>
              <Leq>
                <FieldRef Name='EventDate' />
                <Value IncludeTimeValue='false' Type='DateTime'>${requestEndDate}</Value>
              </Leq>
            </And>
            <Or>
            <Eq>
              <FieldRef Name='EventType'/>
              <Value Type='Text'>0</Value>
            </Eq>
            <Eq>
              <FieldRef Name='EventType'/>
              <Value Type='Text'>4</Value>
            </Eq>
            </Or>
          </And>
          </Where>
          </Query>
          <RowLimit Paged=\"FALSE\">2000</RowLimit>
          </View>`
        });
      if (results && results.Row && results.Row.length > 0) {
        let event: any = '';
        for (event of results.Row) {
          const e:IBizpEventDataSpec = await createEventDataSpec(event);
          if (e) {
            events.push(e);
          }
        }
      }
      let seriesEvents: IBizpEventDataSpec[] = [];
      // deal with recurrent events
      results = await web.lists.getById(listId).usingCaching().renderListDataAsStream(
        {
          ViewXml: `<View><ViewFields><FieldRef Name='RecurrenceData'/><FieldRef Name='Duration'/><FieldRef Name='Author'/><FieldRef Name='Category'/><FieldRef Name='Color'/><FieldRef Name='Description'/><FieldRef Name='ParticipantsPicker'/><FieldRef Name='Geolocation'/><FieldRef Name='ID'/><FieldRef Name='FieldValuesAsText/EndDate'/><FieldRef Name='FieldValuesAsText/EventDate'/><FieldRef Name='ID'/><FieldRef Name='Location'/><FieldRef Name='Title'/><FieldRef Name='fAllDayEvent'/><FieldRef Name='EventType'/><FieldRef Name='UID' /><FieldRef Name='fRecurrence' /></ViewFields>
          <RowLimit Paged=\"FALSE\">2000</RowLimit>
          <Query>
            <Where>
              <Eq>
                <FieldRef Name='EventType'/>
                <Value Type='Text'>1</Value>
              </Eq>
            </Where>
          </Query>
          </View>`
        });

      if (results && results.Row && results.Row.length > 0) {
        let event: any = '';
        for (event of results.Row) {
          const e:IBizpEventDataSpec = await createEventDataSpec(event);
          if (e) {
            seriesEvents.push(e);
          }
        }
      }
      const sEvents = await parseSeriesEventsSP(events, seriesEvents,requestStartDate, requestEndDate);
      events = events.concat(sEvents);

      // Return Data
      return events;
    } catch (error) {
      console.dir(error);
      return Promise.reject(error);
    }
}

async function createEventDataSpec(eventSP): Promise<IBizpEventDataSpec> {
  try {
    console.log ("getSPCalendarEvents getEvents Event " + JSON.stringify(eventSP));
    console.log ("getSPCalendarEvents getEvents Event startDate: " + eventSP.EventDate + " endDate: " + eventSP.EndDate);

    const authorInitialsArray: string[] = eventSP.Author[0].title.split(' ');
    // initials is the first letter of the first two names. e.g. Initials for "Jay Alice Smith" will be JA
    const authorInitials: string = authorInitialsArray[0].charAt(0) + authorInitialsArray[authorInitialsArray.length - 1].charAt(0);
    const userPictureUrl = await getUserProfilePictureUrl(`i:0#.f|membership|${eventSP.Author[0].email}`);
    const attendees: number[] = [];
    let geo:string;
    if (eventSP.Geolocation && !eventSP.Geolocation.empty) {
      const first: number = eventSP.Geolocation.indexOf('(') + 1;
      const last: number = eventSP.Geolocation.indexOf('');
      geo = eventSP.Geolocation.substring(first, last);
    }
    else {
      geo = "";
    }
    const geolocation = geo.split(' ');
    /*
    const CategoryColorValue: any[] = colors.filter((value) => {
      return value.category == eventSP.Category;
    });
    */
    const isAllDayEvent: boolean = eventSP["fAllDayEvent.value"] === "1";

    for (const attendee of eventSP.ParticipantsPicker) {
      attendees.push(parseInt(attendee.id));
    }

    const mStartTime:any = moment(eventSP.EventDate);
    const mEndTime:any = moment(eventSP.EndDate);

    const t1 = await getUtcTime(mStartTime.toISOString());
    return ({
      id: eventSP.ID,
      ID: eventSP.ID,
      eventType: eventSP.EventType,
      title: deCodeHtmlEntities(eventSP.Title),
    description: eventSP.Description,
      startDate: mStartTime.toISOString(),
      endDate: mEndTime.toISOString(),
      location: eventSP.Location,
      ownerEmail: eventSP.Author[0].email,
      ownerPhoto: userPictureUrl ?
        `https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=${eventSP.Author[0].email}&UA=0&size=HR96x96` : '',
      ownerInitial: authorInitials,
//      color: CategoryColorValue.length > 0 ? CategoryColorValue[0].color : '#1a75ff', // blue default
      color: getCategoryColor(eventSP.Category),
      ownerName: eventSP.Author[0].title,
      attendes: attendees,
      fAllDayEvent: isAllDayEvent,
      geolocation: { Longitude: parseFloat(geolocation[0]), Latitude: parseFloat(geolocation[1]) },
      category: eventSP.Category,
      duration: eventSP.Duration,
      recurrenceData: eventSP.RecurrenceData ? deCodeHtmlEntities(eventSP.RecurrenceData) : "",
      fRecurrence: eventSP.fRecurrence,
      recurrenceID: eventSP.RecurrenceID ?  eventSP.RecurrenceID: undefined,
      masterSeriesItemID: eventSP.MasterSeriesItemID,
      UID: eventSP.UID.replace("{", "").replace("}", ""),
    });
  } catch (error) {
    console.dir(error);
    return Promise.reject(error);
  }
}

/* ***************************** create new SP event ***************************** */
export async function addNewSPCalendarEvent(siteUrl: string, listId: string, eventData: IBizpNonRecurrence,
  recurrentData:IBizpRecurrence ) {
  let results = null;
  let recurrenceXML:string = "";
  const attendees: number[] = [];
  let st:Date = new Date(eventData.startDate.toString());
  st.setSeconds(0);
  let et:Date = new Date(eventData.endDate.toString());
  et.setSeconds(0);
  let s:string;
  let e:string;

  s = await getUtcTime(st);
  if (eventData.eventType == '1') {
    recurrenceXML = await prepareRecurrenceXML(recurrentData);
    e = await getUtcTime(moment(et).add(998,'days').toDate());
  }
  else {
    e = await getUtcTime(et);
  }
  console.log("start: " + s + " end: " + e);
  console.log(" In addNewSPEvent - recurrenceXML: " + recurrenceXML);
  try {
    const web = new(Web as any) (siteUrl);
    results = await web.lists.getById(listId).items.add({
      UID: getGUID(),
      Title: eventData.title,
      Description: eventData.description,
      EventDate:s,
      EndDate: e,
      fAllDayEvent: (eventData.fAllDayEvent==true ? "true" : "false"),
      fRecurrence: (eventData.fRecurrence=="1" ? "true" : "false"),
      Category: eventData.category,
      ParticipantsPickerId: { results: attendees },
      EventType: eventData.eventType,
      RecurrenceData: (recurrenceXML == "") ? "" : deCodeHtmlEntities(recurrenceXML),
    });
  }
  catch (error) {
    console.log ("Error in saving: " + error);
    console.log (" HTTP - SiteURL " + siteUrl + " listId = " + listId);
    console.log(" Event Saved: " + JSON.stringify(eventData));
    return Promise.reject(error);
  }
  return results;
}

/* *********************** Update Event **************************** */
export async function updateSPCalendarEvent(siteUrl: string, listId: string,
                                            seriesSelected:boolean,
                                            updateEvent:IBizpEventDataSpec,
                                            eventData:IBizpNonRecurrence,
                                            recurrentData:IBizpRecurrence) {
  let results = null;
  const web = new(Web as any) (siteUrl);
  let ev:any;
  const attendees: number[] = [];

  let st:Date = new Date(eventData.startDate.toString());
  st.setSeconds(0);
  let et:Date = new Date(eventData.endDate.toString());
  et.setSeconds(0);
  let s:string = await getUtcTime(st);
  let e:string;

  switch(updateEvent.eventType) {
    case "0": {
        // single event was selected for editing. The event can be modified to become a series
        if (eventData.eventType == '1') {
          // single event changed to a series
          updateSPCalendarSeries(siteUrl, listId, updateEvent, eventData,recurrentData);
        }
        else {
          // single event was modified
          updateSPCalendarSingleEvent(siteUrl, listId, updateEvent, eventData);
        }
      }
      break;
    case "1": {
        if (seriesSelected) {
          // series was selected for editing
          if (eventData.eventType == '1') {
            // series details have changed
            updateSPCalendarSeries(siteUrl, listId, updateEvent, eventData,recurrentData);
          }
          else {
            // series has changed to a single event
            modifySPCalendarSeriesToSingleEvent(siteUrl, listId, updateEvent, eventData);
          }
        }
        else {
          // single event of a series was selected for editing (new exception event is created)
          createSPCalendarSeriesExceptionEvent(siteUrl, listId, updateEvent, eventData,recurrentData);
        }
      }
      break;
    case "4": {
        // exception event was selected was editing. This event is restricted to remain a single exception
        updateSPCalendarSingleEvent(siteUrl, listId, updateEvent, eventData);
      }
      break;
    default:
      break;
  }
}

// Modify a series to series
export async function updateSPCalendarSeries(siteUrl: string, listId: string,
                      originalEvent:IBizpEventDataSpec,eventData:IBizpNonRecurrence,seriesData:IBizpRecurrence) {
  let results = null;
  try {
    const web = new(Web as any) (siteUrl);
    let ev:any;
    const attendees: number[] = [];

    let st:Date = new Date(eventData.startDate.toString());
    st.setSeconds(0);
    let et:Date = new Date(eventData.endDate.toString());
    et.setSeconds(0);
    let s:string = await getUtcTime(st);
    const e = await getUtcTime(moment(et).add(998,'days').toDate());
    // recursive event
    console.log("Updating a series ID: " + originalEvent.ID);
    // delete all related (exceptions) items
    await deleteRecurrenceExceptions(originalEvent, siteUrl, listId);
    const recurrenceXML:string = await prepareRecurrenceXML(seriesData);
    ev = {
      Title: eventData.title,
      Description: eventData.description,
      EventDate:s,
      EndDate: e,
      fAllDayEvent: (eventData.fAllDayEvent==true ? "true" : "false"),
      fRecurrence: (eventData.fRecurrence=="1" ? "true" : "false"),
      EventType: eventData.eventType,
      Category: eventData.category,
      ParticipantsPickerId: { results: attendees },
      RecurrenceData: (recurrenceXML == "") ? "" : deCodeHtmlEntities(recurrenceXML),
    };
    if (originalEvent.UID) {
      ev.UID = originalEvent.UID;
    }
    if (originalEvent.masterSeriesItemID) {
      ev.MasterSeriesItemID = originalEvent.masterSeriesItemID;
    }
    console.log("Updating with event: " + JSON.stringify(ev));
    results = await web.lists.getById(listId).items.getById(originalEvent.ID).update(ev);
  }
  catch (error) {
    console.log ("Error in updateSPCalendarEvent: " + error);
    console.log (" HTTP - SiteURL " + siteUrl + " listId = " + listId);
    console.log(" Event update: " + JSON.stringify(eventData));
    return Promise.reject(error);
  }
}

// Modify a series to a single event
export async function modifySPCalendarSeriesToSingleEvent(siteUrl: string, listId: string,
              originalEvent:IBizpEventDataSpec,eventData:IBizpNonRecurrence) {
  let results = null;
  try {
    const web = new(Web as any) (siteUrl);
    let ev:any;
    const attendees: number[] = [];

    let st:Date = new Date(eventData.startDate.toString());
    st.setSeconds(0);
    let et:Date = new Date(eventData.endDate.toString());
    et.setSeconds(0);
    const s:string = await getUtcTime(st);
    const e = await getUtcTime(et);
    // recursive event
    console.log("Updating a series ID: " + originalEvent.ID);
    // delete all related (exceptions) items
    await deleteRecurrenceExceptions(originalEvent, siteUrl, listId);
    ev = {
      Title: eventData.title,
      Description: eventData.description,
      EventDate:s,
      EndDate: e,
      fAllDayEvent: (eventData.fAllDayEvent==true ? "true" : "false"),
      fRecurrence: (eventData.fRecurrence=="1" ? "true" : "false"),
      EventType: eventData.eventType,
      Category: eventData.category,
      ParticipantsPickerId: { results: attendees },
      RecurrenceData: "",
    };
    if (originalEvent.UID) {
      ev.UID = originalEvent.UID;
    }
    if (originalEvent.masterSeriesItemID) {
      ev.MasterSeriesItemID = "";
    }
    console.log("Updating with event: " + JSON.stringify(ev));
    results = await web.lists.getById(listId).items.getById(originalEvent.ID).update(ev);
  }
  catch (error) {
    console.log ("Error in modifySPCalendarSeriesToSingleEvent: " + error);
    console.log (" HTTP - SiteURL " + siteUrl + " listId = " + listId);
    console.log(" Event update: " + JSON.stringify(eventData));
    return Promise.reject(error);
  }
}

// Create an exception event of a series
export async function createSPCalendarSeriesExceptionEvent(siteUrl: string, listId: string,
                  originalEvent:IBizpEventDataSpec,eventData:IBizpNonRecurrence,seriesData:IBizpRecurrence) {
  let results = null;
  try {
    const web = new(Web as any) (siteUrl);
    let ev:any;
    const attendees: number[] = [];

    let st:Date = new Date(eventData.startDate.toString());
    st.setSeconds(0);
    let et:Date = new Date(eventData.endDate.toString());
    et.setSeconds(0);
    let s:string = await getUtcTime(st);
    const e = await getUtcTime(et);
    // recursive event
    console.log("Updating a series ID: " + originalEvent.ID);

    ev = {
      Title: eventData.title,
      Description: eventData.description,
      EventDate:s,
      EndDate: e,
      fAllDayEvent: eventData.fAllDayEvent,
      EventType: "4",
      fRecurrence:  "true",
      Category: eventData.category,
      ParticipantsPickerId: { results: attendees },
      RecurrenceData: "Exception event",
      RecurrenceID: s,
      MasterSeriesItemID: originalEvent.ID.toString(),
      UID: getGUID()
    };

    console.log("Updating with event: " + JSON.stringify(ev));
    results = await web.lists.getById(listId).items.add(ev);
  }
  catch (error) {
    console.log ("Error in createSPCalendarSeriesExceptionEvent: " + error);
    console.log (" HTTP - SiteURL " + siteUrl + " listId = " + listId);
    console.log(" Event update: " + JSON.stringify(eventData));
    return Promise.reject(error);
  }
}

// update an exception of a series, series remains the same,
// or update a single event
export async function updateSPCalendarSingleEvent(siteUrl: string, listId: string,
                      originalEvent:IBizpEventDataSpec,eventData:IBizpNonRecurrence) {
  let results = null;
  try {
    const web = new(Web as any) (siteUrl);
    let ev:any;
    const attendees: number[] = [];

    let st:Date = new Date(eventData.startDate.toString());
    st.setSeconds(0);
    let et:Date = new Date(eventData.endDate.toString());
    et.setSeconds(0);
    let s:string = await getUtcTime(st);
    const e = await getUtcTime(et);

    console.log("Updating a series ID: " + originalEvent.ID);

    ev = {
      Title: eventData.title,
      Description: eventData.description,
      EventDate:s,
      EndDate: e,
      fAllDayEvent: eventData.fAllDayEvent,
      Category: eventData.category,
      ParticipantsPickerId: { results: attendees }
    };
    console.log("Updating an event ID: " + JSON.stringify(ev));
    results = await web.lists.getById(listId).items.getById(originalEvent.ID).update(ev);
  }
  catch (error) {
    console.log (" Error in updateSPCalendarSingleEvent: " + error);
    console.log (" HTTP - SiteURL " + siteUrl + " listId = " + listId);
    console.log(" Event update: " + JSON.stringify(eventData));
    return Promise.reject(error);
  }
}

/* *********************** Delete Event **************************** */
export async function deleteSPCalendarEvent(siteUrl: string, listId: string,
                                  eventData: IBizpEventDataSpec,
                                  recurrenceSeriesEdited: boolean) {
  let results = null;
  try {
    const web = new(Web as any) (siteUrl);
    // Exception Recurrence eventtype = 4 ?  update to deleted Recurrence eventtype=3
    switch (eventData.eventType) {
      case '4': // Exception Recurrence Event
        console.log("single event of recurrence deleteing Title: " + eventData.title);
        results = await web.lists.getById(listId).items.getById(eventData.ID).update({
          Title: `Delete: ${eventData.title}`,
          EventType: '3',
        });
        // deleting simply will restore the event in the recurrence at the same date
        // await web.lists.getById(listId).items.getById(eventData.ID).delete();
        break;
      case '1': // recurrence Event
        // if  delete is a main recrrence delete all recurrences and main recurrence
        if (recurrenceSeriesEdited) {
          // delete execptions if exists before delete recurrence event
          await web.lists.getById(listId).items.getById(eventData.ID).delete();
        } else {
          results = await getSPItemById(siteUrl,listId,eventData.ID.toString());
          // This step converts to the proper format for input to SP UTC conversion
          // Alternatively, you can do the same using moment format
          const st:Date = new Date(eventData.startDate.toString());
          const st1:Date = new Date(eventData.endDate.toString());
          const s = await getUtcTime(st);
          const e = await getUtcTime(st1);
          console.log("deleteSPCalendarEvent: startDate: " + eventData.startDate + " s: " + s + " e: " + e);

          const results1 = await web.lists.getById(listId).items.add({
            UID: results.UID,
            Title: `Delete: ${results.Title}`,
            Description: results.Description,
            EventDate:s,
            EndDate: e,
            fAllDayEvent: results.fAllDayEvent,
            fRecurrence: results.fRecurrence,
            Category: results.Category,
            Location: results.Location,
            // set RecurrenceID to the date for which the event needs to be deleted
            // and make sure the hour minutes match the original time
            RecurrenceID: s,
            MasterSeriesItemID: results.ID.toString(),
            EventType: '3',
            RecurrenceData: results.RecurrenceData,
          });
        }

        break;
      case '0': // normal Event
        console.log("single event deleteing Title: " + eventData.title);
        await web.lists.getById(listId).items.getById(eventData.ID).delete();
        console.log("single event deleted Title: " + eventData.title);
        break;
    }
  } catch (error) {
    return Promise.reject(error);
  }
  return;
}

/* *********************** Utility functions *********************** */

export async function deleteRecurrenceExceptions(event: IBizpEventDataSpec, siteUrl: string, listId: string) {
  let results = null;
  try {
    const web = new(Web as any) (siteUrl);
    results = await web.lists.getById(listId).items
      .select('Id')
      .filter(`EventType eq '3' or EventType eq '4' and MasterSeriesItemID eq '${event.ID}' `)
      .get();
    if (results && results.length > 0) {
      for (const exceptionEvent of results) {
        await web.lists.getById(listId).items.getById(exceptionEvent.ID).delete();
      }
    }
  } catch (error) {
    return Promise.reject(error);
  }
  return;
}

/**************************** sorting by date ************************** */
// Sorts by start date
// order can be true for ascending, false for descending
export function sortByDate(items:IBizpEventDataSpec[], order:boolean) {
  return quickSortByDate(items,0,items.length-1,order);
}

export function quickSortByDate(items:IBizpEventDataSpec[], left:number, right:number,order:boolean) {
  var index:number;
  if (items.length > 1) {
      index = partition(items, left, right,order); //index returned from partition
      if (left < index - 1) { //more elements on the left side of the pivot
        quickSortByDate(items, left, index - 1,order);
      }
      if (index < right) { //more elements on the right side of the pivot
        quickSortByDate(items, index, right,order);
      }
  }
  return items;
}
function swap(items:IBizpEventDataSpec[], leftIndex:number, rightIndex:number){
  const temp:IBizpEventDataSpec = items[leftIndex];
  items[leftIndex] = items[rightIndex];
  items[rightIndex] = temp;
}
function partition(items:IBizpEventDataSpec[], left:number, right:number,order:boolean) {
  let pivot:IBizpEventDataSpec = items[Math.floor((right + left) / 2)], //middle element
      i = left, //left pointer
      j = right; //right pointer
  while (i <= j) {
      while (orderFn(items[i],pivot,order)) {
          i++;
      }
      while (orderFn(items[j],pivot,!order)) {
          j--;
      }
      if (i <= j) {
          swap(items, i, j); //sawpping two elements
          i++;
          j--;
      }
  }
  return i;
}
function orderFn(item1:IBizpEventDataSpec,item2:IBizpEventDataSpec,order:boolean):boolean {
  if (order) return (moment(item1.startDate).isBefore(moment(item2.startDate)));
  // otherwise
  return (moment(item1.startDate).isAfter(moment(item2.startDate)));
}

