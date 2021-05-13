import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneButton,PropertyPaneButtonType,
  PropertyPaneHorizontalRule,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import { sp } from "@pnp/sp";
import { setup as pnpSetup } from "@pnp/common";
import { graph } from "@pnp/graph";
import "@pnp/graph/groups";
import "@pnp/graph/calendars";
import '@pnp/graph/users';
import {ICalendar} from "@pnp/graph/calendars";
import { CalloutTriggers } from '@pnp/spfx-property-controls/lib/Callout';
import { PropertyFieldButtonWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldButtonWithCallout';
import { PropertyFieldSitePicker } from '@pnp/spfx-property-controls/lib/PropertyFieldSitePicker';
import { IPropertyFieldSite } from "@pnp/spfx-property-controls/lib/PropertyFieldSitePicker";
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { Checkbox,Stack} from '@fluentui/react';



import * as strings from 'RemindersWebPartStrings';
import Reminders from './components/Reminders';
import { IRemindersProps } from './components/IRemindersProps';
import { HoverCard } from 'office-ui-fabric-react';

export enum  IBizpCalendarSource {
  SPCal=0,
  O365Cal=1,
  googleCal=2
}

export interface IBizpCalendarSpec {
  source: IBizpCalendarSource;
  sites: IPropertyFieldSite[];
  lists: string | string[]; // Stores the list ID(s)
}
export interface IRemindersWebPartProps {
  description: string;
  title: string;
//  calendars: IBizpCalendarSpec[];
  SPCalendars:boolean;
  O365Calendars:boolean;
  sites: IPropertyFieldSite[];
  lists: string | string[]; // Stores the list ID(s)
  daysInFuture: number;
  daysInPast: number;
  myCal:any;
}

export default class RemindersWebPart extends BaseClientSideWebPart<IRemindersWebPartProps> {
  public constructor() {
    super();
  }
  private siteUrl: string;
  private listId: string; // id
  private selectedSites: IPropertyFieldSite[];
  private selectedSiteTitle:string;
  private siteId:string;

  private errorMessage:string;

  private async loadO365Calendars(): Promise<IPropertyPaneDropdownOption[]> {
    const cals: IPropertyPaneDropdownOption[] = [];
    try {
      const c = await graph.me.calendars();
      this.properties.myCal = c[0].name;
      console.log("my Calendars: " + this.properties.myCal);
      return cals;
    } catch (error) {
      this.errorMessage =  `${error.message} -  please check if valid.` ;
      this.context.propertyPane.refresh();
    }
    return cals;
  }
/*
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    console.log("onPropertyPaneFieldChanged: "+ propertyPath + " " +oldValue + " " +newValue);
    console.log("my Calendar: " + JSON.stringify(this.properties.myCal));
    switch(propertyPath) {
      case 'sites':
        if (this.properties.sites.length > 0 ) {
          console.log("Site url: "+ this.properties.sites[0].url);
          this.siteUrl = this.properties.sites[0].url;
        }
        break;
      case 'lists':
        if (this.properties.lists.length > 0 ) {
          console.log("oldValue: "+ oldValue);
          console.log("NewValue: "+ newValue);
          if (typeof this.properties.lists == "string") {
            this.listId = this.properties.lists;
            console.log("List Id: "+ this.properties.lists);
          }
          else {
            this.listId = this.properties.lists[0];
            console.log("List Id: "+ this.properties.lists[0]);
          }
        }
        break;
      case 'daysInPast':
        this.properties.daysInPast = newValue;
        break;
      case 'daysInFuture':
        this.properties.daysInFuture = newValue;
        break;
    }
    console.log("Webpart properties:  "+ JSON.stringify(this.properties));
  }
*/
  protected onListConfigurationChanged(propertyPath: string, oldValue: any, newValue: any): void {
//    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    console.log("onListConfigurationChanged: "+ propertyPath + " " +oldValue + " " +newValue);
    // console.log("my Calendar: " + JSON.stringify(this.properties.myCal));
    if (propertyPath === 'lists' && newValue) {
//      this.properties.lists = [];
      this.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      this.context.propertyPane.refresh();
    }
    else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    }

    if (this.properties.lists.length > 0 ) {
      console.log("oldValue: "+ oldValue);
      console.log("NewValue: "+ newValue);
      if (typeof this.properties.lists == "string") {
        this.listId = this.properties.lists;
        console.log("List Id: "+ this.properties.lists);
      }
      else {
        this.listId = this.properties.lists[0];
        console.log("List Id: "+ this.properties.lists[0]);
      }
      console.log("Webpart properties:  "+ JSON.stringify(this.properties));
    }
  }
  protected async onInit():Promise<void> {
    //    const c = await graph.me.calendars();
    //    this.properties.myCal = c[0].name;
    //    await this.loadO365Calendars();
    pnpSetup({ spfxContext: this.context });
    // init graph
    graph.setup({
      spfxContext: this.context
    });
//    this.properties.myCal = graph.me.calendar();
 //   const c = graph.users.getById('admin@M365x976266.onmicrosoft.com').calendars();
//    console.log("my Calendar: " + JSON.stringify(c));
    this.selectedSiteTitle = this.context.pageContext.web.title;
    this.siteId = this.context.pageContext.web.id.toString();
    this.siteUrl = this.context.pageContext.web.absoluteUrl;
    this.selectedSites =
      [{url:this.context.pageContext.web.absoluteUrl,title:this.selectedSiteTitle, id:this.siteId}];
    console.log(" Initial SelectedSites = " + JSON.stringify(this.selectedSites));
    if ((!this.properties.sites)||(this.properties.sites.length == 0)) {
      this.properties.sites = this.selectedSites;
    }
    console.log(" Initial propertySites = " + JSON.stringify(this.properties.sites));
    if (this.properties.lists && this.properties.lists.length > 0) {
      if (typeof this.properties.lists == "string") {
        this.listId = this.properties.lists;
      }
      else {
        this.listId = this.properties.lists[0];
      }
    }
    if (this.listId) {
      console.log("ListId:: " + this.listId);
    }
    else {
      console.log("List is not configured!");
    }
    if (this.properties.daysInFuture == undefined) {
      this.properties.daysInFuture = 21;
    }
    if (this.properties.daysInPast == undefined) {
      this.properties.daysInFuture = 14;
    }
    if (this.properties.SPCalendars == undefined && this.properties.O365Calendars==undefined) {
      this.properties.SPCalendars = true;
    }
    return Promise.resolve();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
  private validateSite(value: string): string {
    if (value === null ||
      value.trim().length === 0) {
      return 'Provide a valid site. No match found';
    }
    return '';
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  public render(): void {
    console.log("Rendering webpart:");
    const element: React.ReactElement<IRemindersProps> = React.createElement(
      Reminders,
      {
        description: this.properties.description,
        title: this.properties.title,
        siteUrl: this.siteUrl,
        list: this.listId,
        context:this.context,
        daysInFuture: this.properties.daysInFuture,
        daysInPast: this.properties.daysInPast
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    let sourceFields:any;
    let sourceSPFields:any = [
      PropertyFieldSitePicker('sites', {
        label: 'Select a site',
        initialSites: this.selectedSites,
        context: this.context,
        deferredValidationTime: 1000,
        multiSelect: false,
        onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
        properties: this.properties,
        onGetErrorMessage: this.validateSite.bind(this),
        key: 'sitesFieldId'
      }),
      PropertyFieldListPicker('lists', {
        label: 'Select a calendar list',
        includeHidden: false,
        selectedList: this.properties.lists,
        orderBy: PropertyFieldListPickerOrderBy.Title,
        disabled: false,
        multiSelect:true,
        onPropertyChange: this.onListConfigurationChanged.bind(this),
        properties: this.properties,
        context: this.context,
        onGetErrorMessage: null,
        baseTemplate: 106,
        deferredValidationTime: 0,
        key: 'listPickerFieldId'
      }),
      PropertyPaneHorizontalRule()
    ];
    let sourceO365Fields:any = [
      PropertyFieldSitePicker('sites', {
        label: 'Select a site',
        initialSites: this.selectedSites,
        context: this.context,
        deferredValidationTime: 1000,
        multiSelect: false,
        onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
        properties: this.properties,
        onGetErrorMessage: this.validateSite.bind(this),
        key: 'sitesFieldId'
      }),
      PropertyPaneHorizontalRule()
    ];
    if (this.properties.SPCalendars) sourceFields = sourceSPFields;
    if (this.properties.O365Calendars) sourceFields = sourceO365Fields;
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: "Applications",
              groupFields: [
                PropertyPaneButton('SP', {
                  text: "SharePoint",
                  buttonType: PropertyPaneButtonType.Primary,
                  icon : 'SharepointLogoInverse',
                  onClick: () => { this.properties.SPCalendars = true; this.properties.O365Calendars = false;}
                 }),

                 PropertyPaneButton('O365', {
                  text: "Office 365",
                  buttonType: PropertyPaneButtonType.Primary,
                  icon: 'OutlookLogo',
                  onClick: () => { this.properties.SPCalendars = false; this.properties.O365Calendars = true; }
                 }),

                 PropertyPaneHorizontalRule()
              ]
            },
            {
              groupName: strings.CalendarSelectionGroupName,
              groupFields: sourceFields
            },
            {
              groupName: " Event Date Range",
              groupFields: [
                PropertyPaneSlider('daysInFuture',{
                  label:"Max days in the future",
                  min:7,
                  max:60,
                  value:this.properties.daysInFuture,
                  showValue:true,
                }),
                PropertyPaneSlider('daysInPast',{
                  label:"Max days in the past",
                  min:0,
                  max:60,
                  value:this.properties.daysInPast,
                  showValue:true,
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
