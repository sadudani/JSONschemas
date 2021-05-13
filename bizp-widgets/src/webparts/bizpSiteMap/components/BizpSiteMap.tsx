import * as React from 'react';
import {useState,useEffect} from 'react';
import { DefaultPalette, Stack, IStackStyles} from 'office-ui-fabric-react';

import styles from './BizpSiteMap.module.scss';
import { IBizpSiteMapProps } from './IBizpSiteMapProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as strings from 'BizpSiteMapWebPartStrings';
import {​​BizpWebpartHeader,BizpHierarchyDisplay }​​ from "bizp-lib";
import {​​IBizpMenuOptions }​​ from "bizp-lib/lib/shared/IBizpSharedInterface";

export default function BizpSiteMap(props: IBizpSiteMapProps) {
  const IdForHelp = "ReminderesHelp";
  const menuOpts:IBizpMenuOptions[] = [
    { key: 'help', text: strings.HelpMenuLabel,iconName:'StatusCircleQuestionMark'},
    { key: 'feedback', text: strings.FeedbackMenuLabel , iconName:'EmojiNeutral' },
    {key:'refresh',text:strings.RefreshMenuLabel, iconName: 'Refresh' }];
  const stackStyles: IStackStyles = {
    root: {
      background: DefaultPalette.themeTertiary,
      width: 1600,
    },
  };
  //refresh is used as a toggle. Anytime its value is toggled, it will force execute component
  const [refresh,setRefresh] = useState<boolean>(false); // refresh toggle

  function onRefresh():void {
    const r = refresh;
    setRefresh(!r);
  }

  console.log("Rendering Site Map Component, properties: " + JSON.stringify(props.siteUrl));
  return (
    <div className={ styles.bizpSiteMap } >
      <Stack styles={stackStyles} >
        <Stack styles={stackStyles}>
          <BizpWebpartHeader title="Site Tree Map"
            showTitle={true}
            menuOptions = {menuOpts}
            helpId = {IdForHelp}
            onRefresh = {onRefresh}
            themeVariant = {props.themeVariant}
            context = {props.context}
          >
          </BizpWebpartHeader>
        </Stack>
        <Stack styles={stackStyles}>
        <BizpHierarchyDisplay siteUrl={props.siteUrl} list={props.list} context = {props.context}
                              refresh = {refresh} displayLibs= {false} displayLists= {false} />
        </Stack>
      </Stack>
    </div>
  );
/*
  public render(): React.ReactElement<IBizpSiteMapProps> {
    return (
      <div className={ styles.bizpSiteMap }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
  */
}
