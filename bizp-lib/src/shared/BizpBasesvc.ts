
import { WebPartContext } from "@microsoft/sp-webpart-base";
// import { sp } from "@pnp/sp";
import { sp,Web,Fields, IFields,IFieldInfo,PermissionKind,IWebInfosData } from "@pnp/sp/presets/all";
import { ISearchQuery, SearchResults, Search} from "@pnp/sp/search";

import { ICamlQuery } from "@pnp/sp/lists";
import {
  IDropdownOption
} from '@fluentui/react';
import {IBizpUserPermissions} from "./IBizpSharedInterface";

import * as moment from 'moment';
import { arrayToTree } from "performant-array-to-tree";

export const SPService = (spContext: WebPartContext) => {
  sp.setup({ spfxContext: spContext});

  // build the caml query object (in this example, we include Title field and limit rows to 5)
  const caml: ICamlQuery = {
    ViewXml: "<View><ViewFields><FieldRef Name='Title' /></ViewFields><RowLimit>5</RowLimit></View>",
  };

  // method that retrieves the View Query
  async function getViewQueryForList(listName:string,viewName:string):Promise<any> {
    // get view CAML query
    if (listName && viewName){
      const list = sp.web.lists.getByTitle(listName);
      return list.views.getByTitle(viewName).select("ViewQuery").get().then(v => {
          return v.ViewQuery;
      });
    } else {
      console.log('Data insufficient!');
      return null;
    }
  }

  async function getItemsByListView(listName:string, viewName:string):Promise<any> {
    let query: ICamlQuery = {
      ViewXml: "<View><ViewFields><FieldRef Name=viewName /></ViewFields><RowLimit>5</RowLimit></View>",
    };
    const list = sp.web.lists.getByTitle(listName);
    // get list items
    const result = await list.getItemsByCAMLQuery(query);

    return  result;
  }
};
/*
interface IWebInfosData {
    Configuration: number;
    Created: string;
    Description: string;
    Id: string;
    Language: number;
    LastItemModifiedDate: string;
    LastItemUserModifiedDate: string;
    ServerRelativeUrl: string;
    Title: string;
    WebTemplate: string;
    WebTemplateId: number;
}
*/
// Returns all sites and subsites under siteUrl (a site or site collection)
export async function getSPSites(siteUrl: string, includeLibs:boolean):Promise<any[]> {
  try {
    const text:string = "contentclass:STS_Web AND NOT WebTemplate:APP OR Path:" + siteUrl + " AND contentclass:STS_Site";
    const searcher = Search(siteUrl);
    const results: SearchResults = await searcher({
      Querytext: text,
      SelectProperties: ["Title","Path","Description","ParentLink","SiteLogo","ParentSiteTitle","WebTemplate"],
      RowLimit: 3000,
      TrimDuplicates: false
    });
    if (results) {
      console.log("sp search results: " + JSON.stringify(results.PrimarySearchResults));
    }
    let results1:any[];
    if (includeLibs) {
      const libs:any[] = await getLibsforSites(siteUrl,results.PrimarySearchResults);
      results1 = results.PrimarySearchResults.concat(...libs);
      console.log("Flat data with libraries: " + JSON.stringify(results1));
    }
    else {
      results1 = results.PrimarySearchResults;
    }
    // convert flat data to tree
    const tree = arrayToTree(
      results1,
      { id: "Path", parentId: "ParentLink", childrenField: "children" }
    );
    console.log("sp search results (tree): " + JSON.stringify(tree));
    return Promise.resolve(tree);
  }
  catch (error) {
     console.log ("Error: " + error);
    return Promise.reject(error);
  }
}

async function getLibsforSites(siteUrl:string,siteData:any[]):Promise<any[]> {
  // get all libraries under the site collection siteUrl
  let libs = await getSPLibs(siteUrl);
  let filteredLibs:any[]=[];
  let j:any;
  // get libraries for each site/subsite
  for(j in siteData) {
    let e = libs.filter((value) => {
      return (value.SiteId == siteData[j].SiteId) && (value.WebId == siteData[j].WebId);
    });
    filteredLibs.push(...e);
  }
  console.log("FilteredLibs: " +JSON.stringify(filteredLibs));
  return filteredLibs;
}

export async function getSPLibs(siteUrl:string,siteId?:string,webId?:string):Promise<any[]> {
  try {
    let text:string ;
    if (siteId == undefined || webId == undefined) {
      text = "contentclass:STS_List_DocumentLibrary";
    }
    else {
      text = "contentclass:STS_List_DocumentLibrary AND WebId:" + webId + " AND SiteId:" + siteId ;
    }
    const searcher = Search(siteUrl);
    const results: SearchResults = await searcher({
      Querytext: text,
      SelectProperties: ["Title","Path","ParentLink"],
      RowLimit: 3000,
      TrimDuplicates: false
    });
    if (results) {
      console.log("sp search results: " + JSON.stringify(results.PrimarySearchResults));
    }
    return Promise.resolve(results.PrimarySearchResults);
  }
  catch (error) {
    console.log ("Error: " + error);
   return Promise.reject(error);
  }
}

export async function getChoiceFieldOptions(siteUrl: string, listId: string, fieldInternalName: string): Promise<string[]> {
  let options:string[];
  try {
    const web = Web(siteUrl);
    const results:any = await web.lists.getById(listId)
      .fields
      .getByInternalNameOrTitle(fieldInternalName)
      .select("Title", "InternalName", "Choices")
      .get();
      console.log("results: ",JSON.stringify(results) );
    options = results.Choices;
  } catch (error) {
    return Promise.reject(error);
  }
  return options;
}

export async function getFieldDropdownOptions(siteUrl:string,listId:string,fieldInternalName: string): Promise<IDropdownOption[]> {
  let fieldOptions: IDropdownOption[] = [];
  let i:number;
  const options:string[] = await getChoiceFieldOptions(siteUrl, listId, 'Category');
  if (options && options.length > 0) {
    for(i = 0;i<options.length;i++) {
      fieldOptions.push({
        key: i.toString(),
        text: options[i]
      });
   }
  }
  return fieldOptions;
}

export async function getSPItemById(siteUrl: string, listId: string,id:string) {
  let results = null;
  try {

    const web = new(Web as any) (siteUrl);
    results = await web.lists.getById(listId).items.getById(id).select(
      "Id","ID","Title","Description","Location","EventDate","EndDate",
      "fAllDayEvent","Category","RecurrenceData",
      "fRecurrence","UID","RecurrenceID","MasterSeriesItemID"
    ).get();

    if (results) {
      console.log("getSPItemById First Event details for id: " + id + "event details :" + JSON.stringify(results));

    }
    return results;
  }
   catch (error) {
     console.log ("Error: " + error);
    return Promise.reject(error);
   }
}

export async function getUtcTime(date: string | Date): Promise<string> {
  let utcTime:string = "";
  try {
    if (date != undefined) {
      utcTime = await sp.web.regionalSettings.timeZone.localTimeToUTC(date);
    }
    return utcTime;
  }
  catch (error) {
    return Promise.reject(error);
  }
}

export async function getSiteTimeZoneHours(siteUrl: string): Promise<number> {
    let numberHours: number = 0;
    let siteTimeZoneBias: number;
    let siteTimeZoneDaylightBias: number;
    let currentDateTimeOffSet: number = new Date().getTimezoneOffset() / 60;

    try {
      const siteRegionalSettings: any = await getSiteRegionalSettingsTimeZone(siteUrl);
      /* This seems to create problem. Needs further investigation
      siteTimeZoneBias = siteRegionalSettings.Information.Bias;
      siteTimeZoneDaylightBias = siteRegionalSettings.Information.DaylightBias;
      */
      // Calculate  hour to current site
      // Formula to calculate the number of  hours need to get UTC Date.
      // numberHours = (siteTimeZoneBias / 60) + (siteTimeZoneDaylightBias / 60) - currentDateTimeOffSet;
      if (siteTimeZoneBias >= 0) {
        numberHours = ((siteTimeZoneBias / 60) - currentDateTimeOffSet) + siteTimeZoneDaylightBias / 60;
      } else {
        numberHours = ((siteTimeZoneBias / 60) - currentDateTimeOffSet);
      }

      numberHours = 0;
    }
    catch (error) {
      return Promise.reject(error);
    }
    return numberHours;
}

export async function getSiteRegionalSettingsTimeZone(siteUrl: string) {
  let s:any;
  try {
    const web = new(Web as any) (siteUrl);
    s = await web.regionalSettings.timeZone.usingCaching().get();
    console.log ("getSiteRegionalSettingsTimeZone" + JSON.stringify(s));
    // testing local offset
    const testDate = new Date();
    const localOffset = testDate.getTimezoneOffset();
    console.log ("Local Timezone Offset = " + localOffset);

    // printout: getSiteRegionalSettingsTimeZone{"odata.metadata":"https://m365x053591.sharepoint.com/sites/SPFX-Webpart/_api/$metadata#SP.ApiData.TimeZones/@Element","odata.type":"SP.TimeZone","odata.id":"https://m365x053591.sharepoint.com/sites/SPFX-Webpart/_api/web/regionalsettings/timezone","odata.editLink":"web/regionalsettings/timezone","Description":"(UTC-08:00) Pacific Time (US and Canada)","Id":13,"Information":{"Bias":480,"DaylightBias":-60,"StandardBias":0}}
    return s;
  } catch (error) {
    return Promise.reject(error);
  }
}

export async function getUserProfilePictureUrl(loginName: string) {
  let results: any = null;
  try {
    results = await sp.profiles.usingCaching().getPropertiesFor(loginName);
    return results.PictureUrl;
  } catch (error) {
    results = null;
    return results;
  }
}

export async function getUserPermissions(siteUrl: string, listId: string): Promise<IBizpUserPermissions> {
    let hasPermissionAdd: boolean = false;
    let hasPermissionEdit: boolean = false;
    let hasPermissionDelete: boolean = false;
    let hasPermissionView: boolean = false;
    let userPermissions: IBizpUserPermissions = undefined;
    try {
      const web = Web(siteUrl);
      const userEffectivePermissions = await web.lists.getById(listId).effectiveBasePermissions.get();
      // ...
      hasPermissionAdd = sp.web.lists.getById(listId).hasPermissions(userEffectivePermissions, PermissionKind.AddListItems);
      hasPermissionDelete = sp.web.lists.getById(listId).hasPermissions(userEffectivePermissions, PermissionKind.DeleteListItems);
      hasPermissionEdit = sp.web.lists.getById(listId).hasPermissions(userEffectivePermissions, PermissionKind.EditListItems);
      hasPermissionView = sp.web.lists.getById(listId).hasPermissions(userEffectivePermissions, PermissionKind.ViewListItems);
      userPermissions = { hasPermissionAdd: hasPermissionAdd, hasPermissionEdit: hasPermissionEdit, hasPermissionDelete: hasPermissionDelete, hasPermissionView: hasPermissionView };

    } catch (error) {
      return Promise.reject(error);
    }
    return userPermissions;
}

/*********************************** Calendar utilities ***************************** */
const wom = ['first', 'second', 'third', 'fourth'];
const wd = ['su', 'mo', 'tu', 'we', 'th', 'fr', 'sa'];
const weekday = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];

//
// removes quotes and = from the string
// creates an array of tokens
export function formatString(str: string): string[] {
  let arr = str.split("'");
  str = arr.join('');
  arr = str.split('"');
  str = arr.join('');
  arr = str.split('=');
  str = arr.join(' ');
  str.trim();
  return str.split(' ');
}

// returns last weekday of the month of mDate
// mDate must be a moment object
export function lastWeekdayOfMonth (mDate:any): number{
let temp:any=moment(mDate);
// get th elast date of the month
let selectedDate:number = mDate.daysInMonth();
// init to the last date of the month
temp.date(selectedDate);
const lastDay:number = temp.day();

if (lastDay == 0) {
  // last day is sunday, move back to Friday
  selectedDate = selectedDate - 2;
}
else if (lastDay == 1) {
  // last day is sunday, move back to Friday
  selectedDate = selectedDate - 1;
}
return selectedDate;
}

// returns last weekend day(sat or sun) of the month of mDate
// mDate must be a moment object
export function lastWeekendDayOfMonth (mDate:any): number{
let temp:any=moment(mDate);
// get th elast date of the month
let selectedDate:number = mDate.daysInMonth();
// init to the last date of the month
temp.date(selectedDate);
const lastDay:number = temp.day();

if ((lastDay != 0)&&(lastDay != 6)) {
  // last day is not a weekend day
  // move back to the sunday
  selectedDate = selectedDate - lastDay;
}
// Otherwise, it is a weekend day, nothing to do
return selectedDate;
}

// returns specific weekday(0-6) of the last week for the month of mDate
// mDate must be a moment object
export function lastSpecificDayOfMonth (mDate:any,weekDay:number): number {
let temp:any=moment(mDate);
// get the last date of the month
let selectedDate:number = temp.daysInMonth();
// init to the last date of the month
temp.date(selectedDate);
// get the last day (Sunday ... Saturday)
const lastDay:number = temp.day();

if (lastDay > weekDay) selectedDate = selectedDate - (lastDay-weekDay);
else if (lastDay < weekDay) selectedDate = selectedDate - (7 - (weekDay - lastDay));
// otherwise lastDay is the required day
return selectedDate;
}

// get day of month
// weedayOfMonth is one of wom, argDate is the argument date in moment, day is the weekday
// return day of month based on selected weekdayOfMonth
export function getDayOfMonth(weekdayOfMonth:string,argDate:any,day:number):number {
  let argDay:number = argDate.day();
  let aDate = argDate.date();
  let selectedDate:number;

  if (weekdayOfMonth == 'last') {
    selectedDate = lastSpecificDayOfMonth(argDate,day);
  }
  else {
    // go to the first required day
    if (argDay > day) selectedDate = (7 - (argDay - day)) + aDate;
    else selectedDate = (day - argDay) + aDate;
    // loop starts with the second instance
    for (let i: any = 0; i < wom.indexOf(weekdayOfMonth); i++) {
      selectedDate = selectedDate + 7;  //add a week to each instance to get the Nth instance
    }
  }
  return selectedDate;
}

// get week day of month specified by weekdayOfMonth
// weedayOfMonth is one of wom (first, second, third, fourth, last), argDate is the argument date in moment
// month is based on mDate
// return day of month based on selected weekdayOfMonth
export function getWeekdayOfMonth(weekdayOfMonth:string,argDate:any):number {
let selectedDate:number;
let mDate:any = moment(argDate);
// set it to the first day of the month
mDate.date(1);
// do for weekday: any day (Mon-Fri)
//find first weekday - if not saturday or sunday, then current date is a weekday
if (mDate.day() == 0) mDate.add(1,'days');// add one day to sunday
else if (mDate.day() == 6) mDate.add(2,'days'); //add two days to saturday
if (weekdayOfMonth == 'last') {
    // do for weekdayOfMonth = last
    mDate.date(lastWeekdayOfMonth(mDate),'date');
}
else {
    // do for weekdayOfMonth = first or second or third or fourth
    for (let i: any = 0; i < wom.indexOf(weekdayOfMonth); i++) {
        if (mDate.day() == 5) mDate.add(3,'days'); // for friday, add three days to get to monday
        else mDate.add(1,'days'); //otherwise, just add one day
    }
}
return mDate.date();
}

// get week day of month specified by weekdayOfMonth
// weedayOfMonth is one of wom (first, second, third, fourth, last), argDate is the argument date in moment
// month is based on mDate
// return day of month based on selected weekdayOfMonth
export function getWeekendOfMonth(weekdayOfMonth:string,argDate:any):number {
  let mDate:any = moment(argDate);
  const day:number = mDate.day();

  if (weekdayOfMonth == 'last') {
      mDate.date(lastWeekendDayOfMonth(mDate),'date');
  }
  else {
    //if not saturday or sunday, then add days to get to the first saturday
    if (day != 0 && day != 6) mDate.add((6 - day),'days');
        // loop starts with second weekend day
    for (let i: any = 0; i < wom.indexOf(weekdayOfMonth); i++) {
        // do for weekend dayOfMonth = first or second or third or fourth
        if (mDate.day() == 0) mDate.add(6,'days'); // for sunday, add six days to get to saturday
        else mDate.add(1,'days'); //otherwise, just add one day
    }
  }
  return mDate.date();
}

export function toLocaleLongDateString(date: Date) {
  return moment(date).format('LL');
}

export function toLocaleShortDateString(date: Date) {
  return moment(date).format('ll');
}

/********************************** Utility ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^*/
// copy an object hierarchically through the object nesting
export function cloneObj(obj: any): any {
  let copy: any;
  if (null == obj || "object" != typeof obj) return obj;
  if (obj instanceof Date) {
    copy = new Date();
    copy.setTime(obj.getTime());
    return copy;
  }
  if (obj instanceof Array) {
    copy = [];
    for (let i = 0, len = obj.length; i < len; i++) {
      copy[i] = cloneObj(obj[i]);
    }
    return copy;
  }
  if (obj instanceof Object) {
    copy = {};
    for (const attr in obj) {
      if (obj.hasOwnProperty(attr)) copy[attr] = cloneObj(obj[attr]);
    }
    return copy;
  }
  throw new Error("Unable to copy obj! Its type isn't supported.");
}

export function deCodeHtmlEntities(string: string) {
  const HtmlEntitiesMap = {
    "'": "&#39;",
    "<": "&lt;",
    ">": "&gt;",
    " ": "&nbsp;",
    "¡": "&iexcl;",
    "¢": "&cent;",
    "£": "&pound;",
    "¤": "&curren;",
    "¥": "&yen;",
    "¦": "&brvbar;",
    "§": "&sect;",
    "¨": "&uml;",
    "©": "&copy;",
    "ª": "&ordf;",
    "«": "&laquo;",
    "¬": "&not;",
    "®": "&reg;",
    "¯": "&macr;",
    "°": "&deg;",
    "±": "&plusmn;",
    "²": "&sup2;",
    "³": "&sup3;",
    "´": "&acute;",
    "µ": "&micro;",
    "¶": "&para;",
    "·": "&middot;",
    "¸": "&cedil;",
    "¹": "&sup1;",
    "º": "&ordm;",
    "»": "&raquo;",
    "¼": "&frac14;",
    "½": "&frac12;",
    "¾": "&frac34;",
    "¿": "&iquest;",
    "À": "&Agrave;",
    "Á": "&Aacute;",
    "Â": "&Acirc;",
    "Ã": "&Atilde;",
    "Ä": "&Auml;",
    "Å": "&Aring;",
    "Æ": "&AElig;",
    "Ç": "&Ccedil;",
    "È": "&Egrave;",
    "É": "&Eacute;",
    "Ê": "&Ecirc;",
    "Ë": "&Euml;",
    "Ì": "&Igrave;",
    "Í": "&Iacute;",
    "Î": "&Icirc;",
    "Ï": "&Iuml;",
    "Ð": "&ETH;",
    "Ñ": "&Ntilde;",
    "Ò": "&Ograve;",
    "Ó": "&Oacute;",
    "Ô": "&Ocirc;",
    "Õ": "&Otilde;",
    "Ö": "&Ouml;",
    "×": "&times;",
    "Ø": "&Oslash;",
    "Ù": "&Ugrave;",
    "Ú": "&Uacute;",
    "Û": "&Ucirc;",
    "Ü": "&Uuml;",
    "Ý": "&Yacute;",
    "Þ": "&THORN;",
    "ß": "&szlig;",
    "à": "&agrave;",
    "á": "&aacute;",
    "â": "&acirc;",
    "ã": "&atilde;",
    "ä": "&auml;",
    "å": "&aring;",
    "æ": "&aelig;",
    "ç": "&ccedil;",
    "è": "&egrave;",
    "é": "&eacute;",
    "ê": "&ecirc;",
    "ë": "&euml;",
    "ì": "&igrave;",
    "í": "&iacute;",
    "î": "&icirc;",
    "ï": "&iuml;",
    "ð": "&eth;",
    "ñ": "&ntilde;",
    "ò": "&ograve;",
    "ó": "&oacute;",
    "ô": "&ocirc;",
    "õ": "&otilde;",
    "ö": "&ouml;",
    "÷": "&divide;",
    "ø": "&oslash;",
    "ù": "&ugrave;",
    "ú": "&uacute;",
    "û": "&ucirc;",
    "ü": "&uuml;",
    "ý": "&yacute;",
    "þ": "&thorn;",
    "ÿ": "&yuml;",
    "Œ": "&OElig;",
    "œ": "&oelig;",
    "Š": "&Scaron;",
    "š": "&scaron;",
    "Ÿ": "&Yuml;",
    "ƒ": "&fnof;",
    "ˆ": "&circ;",
    "˜": "&tilde;",
    "Α": "&Alpha;",
    "Β": "&Beta;",
    "Γ": "&Gamma;",
    "Δ": "&Delta;",
    "Ε": "&Epsilon;",
    "Ζ": "&Zeta;",
    "Η": "&Eta;",
    "Θ": "&Theta;",
    "Ι": "&Iota;",
    "Κ": "&Kappa;",
    "Λ": "&Lambda;",
    "Μ": "&Mu;",
    "Ν": "&Nu;",
    "Ξ": "&Xi;",
    "Ο": "&Omicron;",
    "Π": "&Pi;",
    "Ρ": "&Rho;",
    "Σ": "&Sigma;",
    "Τ": "&Tau;",
    "Υ": "&Upsilon;",
    "Φ": "&Phi;",
    "Χ": "&Chi;",
    "Ψ": "&Psi;",
    "Ω": "&Omega;",
    "α": "&alpha;",
    "β": "&beta;",
    "γ": "&gamma;",
    "δ": "&delta;",
    "ε": "&epsilon;",
    "ζ": "&zeta;",
    "η": "&eta;",
    "θ": "&theta;",
    "ι": "&iota;",
    "κ": "&kappa;",
    "λ": "&lambda;",
    "μ": "&mu;",
    "ν": "&nu;",
    "ξ": "&xi;",
    "ο": "&omicron;",
    "π": "&pi;",
    "ρ": "&rho;",
    "ς": "&sigmaf;",
    "σ": "&sigma;",
    "τ": "&tau;",
    "υ": "&upsilon;",
    "φ": "&phi;",
    "χ": "&chi;",
    "ψ": "&psi;",
    "ω": "&omega;",
    "ϑ": "&thetasym;",
    "ϒ": "&Upsih;",
    "ϖ": "&piv;",
    "–": "&ndash;",
    "—": "&mdash;",
    "‘": "&lsquo;",
    "’": "&rsquo;",
    "‚": "&sbquo;",
    "“": "&ldquo;",
    "”": "&rdquo;",
    "„": "&bdquo;",
    "†": "&dagger;",
    "‡": "&Dagger;",
    "•": "&bull;",
    "…": "&hellip;",
    "‰": "&permil;",
    "′": "&prime;",
    "″": "&Prime;",
    "‹": "&lsaquo;",
    "›": "&rsaquo;",
    "‾": "&oline;",
    "⁄": "&frasl;",
    "€": "&euro;",
    "ℑ": "&image;",
    "℘": "&weierp;",
    "ℜ": "&real;",
    "™": "&trade;",
    "ℵ": "&alefsym;",
    "←": "&larr;",
    "↑": "&uarr;",
    "→": "&rarr;",
    "↓": "&darr;",
    "↔": "&harr;",
    "↵": "&crarr;",
    "⇐": "&lArr;",
    "⇑": "&UArr;",
    "⇒": "&rArr;",
    "⇓": "&dArr;",
    "⇔": "&hArr;",
    "∀": "&forall;",
    "∂": "&part;",
    "∃": "&exist;",
    "∅": "&empty;",
    "∇": "&nabla;",
    "∈": "&isin;",
    "∉": "&notin;",
    "∋": "&ni;",
    "∏": "&prod;",
    "∑": "&sum;",
    "−": "&minus;",
    "∗": "&lowast;",
    "√": "&radic;",
    "∝": "&prop;",
    "∞": "&infin;",
    "∠": "&ang;",
    "∧": "&and;",
    "∨": "&or;",
    "∩": "&cap;",
    "∪": "&cup;",
    "∫": "&int;",
    "∴": "&there4;",
    "∼": "&sim;",
    "≅": "&cong;",
    "≈": "&asymp;",
    "≠": "&ne;",
    "≡": "&equiv;",
    "≤": "&le;",
    "≥": "&ge;",
    "⊂": "&sub;",
    "⊃": "&sup;",
    "⊄": "&nsub;",
    "⊆": "&sube;",
    "⊇": "&supe;",
    "⊕": "&oplus;",
    "⊗": "&otimes;",
    "⊥": "&perp;",
    "⋅": "&sdot;",
    "⌈": "&lceil;",
    "⌉": "&rceil;",
    "⌊": "&lfloor;",
    "⌋": "&rfloor;",
    "⟨": "&lang;",
    "⟩": "&rang;",
    "◊": "&loz;",
    "♠": "&spades;",
    "♣": "&clubs;",
    "♥": "&hearts;",
    "♦": "&diams;"
  };
  let entityMap = HtmlEntitiesMap;
  let entity:string;
  let regex:RegExp;
  for (var key in entityMap) {
    entity = entityMap[key];
    regex = new RegExp(entity, 'g');
    string = string.replace(regex, key);
  }
  string = string.replace(/&quot;/g, '"');
  string = string.replace(/&amp;/g, '&');
  return string;
}
// hsv (Hue, Saturation, Value)
// s, and l are contained in the set [0, 1] and represent the two quantities
export function generateRandomColor(s:number,v:number):string {
  const hexCode:string = hsvToHex(Math.random(),s,v);
  return hexCode;
}

/**
 * Converts an HSV color value to RGB. Conversion formula
 * adapted from http://en.wikipedia.org/wiki/HSL_color_space.
 * https://martin.ankerl.com/2009/12/09/how-to-create-random-colors-programmatically/
 * https://en.wikipedia.org/wiki/HSL_and_HSV#Converting_to_RGB
 * Assumes h, s, and v are contained in the set [0, 1] and
 * returns r, g, and b in the set [0, 255].
 *
 * @param   {number}  h       The hue
 * @param   {number}  s       The saturation
 * @param   {number}  v       The lightness
 * @return  {Array}           The RGB representation
 */
  function hsvToHex(h:number, s:number, v:number):string {
    let r:number, g:number, b:number;

    const h_i = Math.floor(h*6);
    const f = (h*6) - h_i;
    const p:number = v * (1 - s);
    const q:number = v * (1 - (f*s));
    const t = v * (1 - ((1 - f) * s));
    switch (h_i) {
      case 0:
        r = v; g = t; b = p;
        break;
      case 1:
        r = q; g = v; b = p;
        break;
      case 2:
        r = p; g = v; b = t;
        break;
      case 3:
        r = p; g = q; b = v;
        break;
      case 4:
        r = t; g = p; b = v;
        break;
      case 5:
        r = v; g = p; b = q;
        break;
    }
  return (rgbToHex(Math.floor(r * 256), Math.floor(g * 256), Math.floor(b * 256)));
}

function componentToHex(c:number):string {
  var hex = c.toString(16);
  return hex.length == 1 ? "0" + hex : hex;
}

function rgbToHex(r:number, g:number, b:number):string {
  return "#" + componentToHex(r) + componentToHex(g) + componentToHex(b);
}
