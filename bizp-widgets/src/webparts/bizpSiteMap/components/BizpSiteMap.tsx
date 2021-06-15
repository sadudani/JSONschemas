import * as React from 'react';
import {useState,useEffect} from 'react';
import { DefaultPalette, Stack, IStackStyles} from 'office-ui-fabric-react';

import styles from './BizpSiteMap.module.scss';
import { IBizpSiteMapProps } from './IBizpSiteMapProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as strings from 'BizpSiteMapWebPartStrings';
import {​​BizpWebpartHeader,BizpHierarchyDisplay,BizpSiteMapDisplay }​​ from "bizp-lib";
import {​​IBizpMenuOptions }​​ from "bizp-lib/lib/shared/IBizpSharedInterface";
import { getPropsWithDefaults } from '@fluentui/react';

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

  console.log("Rendering Site Map Component, properties: " + props.siteUrl);
  return (
    <div className={ styles.bizpSiteMap1 } >
    <div className={styles.container}>
      <div className={styles.row}>

          <BizpWebpartHeader title="Site Tree Map"
            showTitle={true}
            menuOptions = {menuOpts}
            helpId = {IdForHelp}
            onRefresh = {onRefresh}
            themeVariant = {props.themeVariant}
            context = {props.context}
          >
          </BizpWebpartHeader>

      </div>
      <div className={styles.row}>
          <div style={{ height: 600 }}>
            {
              (props.layout == 1) &&
              <BizpHierarchyDisplay siteUrl={props.siteUrl} list={props.list} context = {props.context}
              refresh = {refresh} displayLibs= {false} displayLists= {false} />
            }
            {
              (props.layout > 1) &&
              <BizpSiteMapDisplay siteUrl={props.siteUrl} list={props.list} context = {props.context}
              refresh = {refresh} displayLibs= {props.displayLibs} displayLists= {props.displayLists}
              layout = {props.layout} theme = {props.themeVariant}/>

            }
          </div>
      </div>
    </div>
    </div>
  );

  /* return (
    <div className={ styles.bizpSiteMap } >
    <div className={styles.container}>
      <div className={styles.row}>

          <BizpWebpartHeader title="Site Tree Map"
            showTitle={true}
            menuOptions = {menuOpts}
            helpId = {IdForHelp}
            onRefresh = {onRefresh}
            themeVariant = {props.themeVariant}
            context = {props.context}
          >
          </BizpWebpartHeader>

      </div>
      <div className={styles.row}>
        <div className={styles.column}>
          <div style={{ height: 600 }}>
            {
              (props.layout == 1) &&
              <BizpHierarchyDisplay siteUrl={props.siteUrl} list={props.list} context = {props.context}
              refresh = {refresh} displayLibs= {false} displayLists= {false} />
            }
            {
              (props.layout == 2) &&
              <BizpSiteMapDisplay siteUrl={props.siteUrl} list={props.list} context = {props.context}
              refresh = {refresh} displayLibs= {props.displayLibs} displayLists= {props.displayLists} />

            }
          </div>
        </div>
      </div>
    </div>
    </div>
  );
   */
 /*        <div className={ styles.bizpSiteMap } >
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
            {
              (props.layout == 1) &&
              <BizpHierarchyDisplay siteUrl={props.siteUrl} list={props.list} context = {props.context}
              refresh = {refresh} displayLibs= {false} displayLists= {false} />
            }
            {
              (props.layout == 2) &&
              <BizpSiteMapDisplay siteUrl={props.siteUrl} list={props.list} context = {props.context}
              refresh = {refresh} displayLibs= {props.displayLibs} displayLists= {props.displayLists} />

            }

          </Stack>
        </Stack>
      </div>
  );
  */
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
