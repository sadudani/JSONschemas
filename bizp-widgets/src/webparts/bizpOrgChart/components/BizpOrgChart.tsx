import * as React from 'react';
import styles from './BizpOrgChart.module.scss';
import { IBizpOrgChartProps } from './IBizpOrgChartProps';
import { escape } from '@microsoft/sp-lodash-subset';

import {useState,useEffect} from 'react';

import * as strings from 'BizpOrgChartWebPartStrings';
import {​​BizpWebpartHeader,BizpOrgChartDisplay }​​ from "bizp-lib";
import {​​IBizpMenuOptions }​​ from "bizp-lib/lib/shared/IBizpSharedInterface";
import { getPropsWithDefaults } from '@fluentui/react';

export default function BizpOrgChart(props: IBizpOrgChartProps) {
  const IdForHelp = "ReminderesHelp";
  const menuOpts:IBizpMenuOptions[] = [
    { key: 'help', text: strings.HelpMenuLabel,iconName:'StatusCircleQuestionMark'},
    { key: 'feedback', text: strings.FeedbackMenuLabel , iconName:'EmojiNeutral' },
    {key:'refresh',text:strings.RefreshMenuLabel, iconName: 'Refresh' }];

  //refresh is used as a toggle. Anytime its value is toggled, it will force execute component
  const [refresh,setRefresh] = useState<boolean>(false); // refresh toggle

  function onRefresh():void {
    const r = refresh;
    setRefresh(!r);
  }

  console.log("Rendering Org Chart Component, properties: " + props.siteUrl);
  return (
    <div className={ styles.bizpOrgChart } >
    <div className={styles.container}>
      <div className={styles.row}>

          <BizpWebpartHeader title="Organization Chart"
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
              <BizpOrgChartDisplay context = {props.context}
              refresh = {refresh} theme={props.themeVariant} />
            }
          </div>
      </div>
    </div>
    </div>
  );
}
