import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneButton,PropertyPaneButtonType,
  PropertyPaneCheckbox,
  PropertyPaneHorizontalRule,
  PropertyPaneDropdown,
  PropertyPaneSlider,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'BizpOrgChartWebPartStrings';
import BizpOrgChart from './components/BizpOrgChart';
import { IBizpOrgChartProps } from './components/IBizpOrgChartProps';

import { ThemeProvider, IReadonlyTheme, ThemeChangedEventArgs } from '@microsoft/sp-component-base';

import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { Checkbox,Stack} from '@fluentui/react';

import { sp } from "@pnp/sp";
import { setup as pnpSetup } from "@pnp/common";
import { graph } from "@pnp/graph";
import { Logger, ConsoleListener, LogLevel } from "@pnp/logging";
import { PropertyFieldSitePicker } from '@pnp/spfx-property-controls/lib/PropertyFieldSitePicker';
import { IPropertyFieldSite } from "@pnp/spfx-property-controls/lib/PropertyFieldSitePicker";
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';


export interface IBizpOrgChartWebPartProps {
  description: string;
  title: string;
  SPSites:boolean;
  teams:boolean;
  sites: IPropertyFieldSite[];
  list: string | string[]; // Stores the list ID(s)
  displayLibs: boolean;
  displayLists: boolean;
  siteUrl:string;
  includeLibs:boolean;
  includeLists:boolean;
  layout:number;
}

export default class BizpOrgChartWebPart extends BaseClientSideWebPart<IBizpOrgChartWebPartProps> {
  public constructor() {
    super();
  }
  private themeProvider: ThemeProvider;
  private themeVariant: IReadonlyTheme | undefined;

  private site: string;
  private list: string; // id
  private selectedSites: IPropertyFieldSite[];
  private selectedSiteTitle:string;
  private siteId:string;

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    console.log("onPropertyPaneFieldChanged: "+ propertyPath + " " +oldValue + " " +newValue);
    console.log("Webpart properties:  "+ JSON.stringify(this.properties));
  }
  protected onListConfigurationChanged(propertyPath: string, oldValue: any, newValue: any): void {
    this.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

  private handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this.themeVariant = args.theme;
    this.render();
  }

  protected async onInit():Promise<void> {
    pnpSetup({ spfxContext: this.context });
    graph.setup({
      spfxContext: this.context
    });
    // Consume the new ThemeProvider service
    this.themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
    // If it exists, get the theme variant
    this.themeVariant = this.themeProvider.tryGetTheme();
    // Register a handler to be notified if the theme variant changes
    this.themeProvider.themeChangedEvent.add(this, this.handleThemeChangedEvent);
    console.debug("Theme: ",this.themeVariant);

    // subscribe a listener
    Logger.subscribe(new ConsoleListener());
    // set the active log level -- eventually make this a web part property
    Logger.activeLogLevel = LogLevel.Error;

    this.selectedSiteTitle = this.context.pageContext.web.title;
    this.siteId = this.context.pageContext.web.id.toString();
    this.site = this.context.pageContext.web.absoluteUrl;
    this.selectedSites =
          [{url:this.context.pageContext.web.absoluteUrl,title:this.selectedSiteTitle, id:this.siteId}];
    console.log(" Initial SelectedSites = " + JSON.stringify(this.selectedSites));
    if (!this.properties.sites) {
      this.properties.sites = this.selectedSites;
    }
    if (!this.properties.SPSites) {
      this.properties.SPSites = true;
      this.properties.teams = false;
    }
    if (!this.properties.description) {
      this.properties.description = strings.PropertyPaneDescription;

    }
    console.log(" Initial propertySites = " + JSON.stringify(this.properties.sites));

    return Promise.resolve();
  }

  private validateSite(value: string): string {
    if (value === null ||
      value.trim().length === 0) {
      return 'Provide a valid site. No match found';
    }
    return '';
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
        return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    let sourceFields:any;
    let sourceSPFields:any = [
      PropertyPaneTextField('siteUrl', {
        label: 'Site URL'
      }),
      PropertyPaneCheckbox('includeLibs', {
        text: 'Include Libraries'
      }),
      PropertyPaneCheckbox('includeLists', {
        text: 'Include listsL'
      }),
      PropertyPaneDropdown('layout', {
         label: 'Style',
         selectedKey: '1',
         options: [
           { key: '1', text: 'Simple' },
           { key: '2', text: 'Modern' },
           { key: '3', text: 'Explorer' },
           { key: '4', text: 'Fabric' }
         ]
       })
      // PropertyPaneTextField('description', {
      //   label: strings.DescriptionFieldLabel
      // })
    ];
    let sourceO365Fields:any = [
      PropertyFieldSitePicker('sites', {
        label: 'Select a team',
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
    if (this.properties.SPSites) sourceFields = sourceSPFields;
    if (this.properties.teams) sourceFields = sourceO365Fields;
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
                  onClick: () => { this.properties.SPSites = true; this.properties.teams = false;}
                 }),

                 PropertyPaneButton('Teams', {
                  text: "Teams",
                  buttonType: PropertyPaneButtonType.Primary,
                  icon: 'TeamsLogo',
                  onClick: () => { this.properties.SPSites = false; this.properties.teams = true; }
                 }),

                 PropertyPaneHorizontalRule()
              ]
            },
            {
              groupName: strings.siteURLSelectionGroupName,
              groupFields: sourceFields
            }
          ]
        }
      ]
    };
  }

  public render(): void {
    console.log("Org Chart webpart properties: " + JSON.stringify(this.properties));
    const element: React.ReactElement<IBizpOrgChartProps> = React.createElement(
      BizpOrgChart,
      {
        description: this.properties.description,
        title: this.properties.title,
        siteUrl: this.properties.siteUrl ? this.properties.siteUrl: null,
        list: this.properties.list? ((typeof this.properties.list == "string") ? this.properties.list : this.properties.list[0]):null,
        context:this.context,
        themeVariant:this.themeVariant,
        layout: this.properties.layout,
      }
    );

    ReactDom.render(element, this.domElement);
  }
}
