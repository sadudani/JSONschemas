import * as React from 'react';
import {useState,useEffect} from 'react';
import * as strings from 'BizpcompsLibraryStrings';
import styles from './BizpRecurrentEvent.module.scss';
import { IBizpRecurrentEventProps } from './IBizpRecurrentEventProps';
import { BizpDailyRecurrentEvent } from '../BizpDailyRecurrentEvent/BizpDailyRecurrentEvent';
import { BizpWeeklyRecurrentEvent } from '../BizpWeeklyRecurrentEvent/BizpWeeklyRecurrentEvent';
import { BizpMonthlyRecurrentEvent } from '../BizpMonthlyRecurrentEvent/BizpMonthlyRecurrentEvent';
import { BizpYearlyRecurrentEvent } from '../BizpYearlyRecurrentEvent/BizpYearlyRecurrentEvent';
import {
  IBizpDailyRecurrence,
  IBizpWeeklyRecurrence,
  IBizpMonthlyRecurrence,
  IBizpYearlyRecurrence,
} from '../../../../shared/IBizpSharedInterface';
import {
  ChoiceGroup,
  IChoiceGroupOption,
} from 'office-ui-fabric-react';

export function BizpRecurrentEvent(props: IBizpRecurrentEventProps) {
  // if it is not a new entry, set the rule to the series to be displayed
  const [rule,setRule] =  useState((props.eventSeries!=undefined)? props.eventSeries.rule: "daily");

  function returnDailyInfo(data:IBizpDailyRecurrence) {
    props.returnInfo({dailyRecurrence:data,rule:"daily"},false);
  }

  function returnWeeklyInfo(data:IBizpWeeklyRecurrence) {
    props.returnInfo({weeklyRecurrence:data,rule:"weekly"},false);
  }

  function returnMonthlyInfo(data:IBizpMonthlyRecurrence) {
    props.returnInfo({monthlyRecurrence:data,rule:"monthly"},false);
  }

  function returnYearlyInfo(data:IBizpYearlyRecurrence) {
    props.returnInfo({yearlyRecurrence:data,rule:"yearly"},false);
  }

  function onRuleChange(ev: React.SyntheticEvent<HTMLElement>, option: IChoiceGroupOption): void {
    setRule(option.key);
  }
  console.log("Rendering RecurrentEvent...");
  return (
    <div className={styles.divWrraper} >

      <div style={{ display: 'inline-block', verticalAlign: 'top' }}>
        <ChoiceGroup
          label={ "Repeat information" }
          selectedKey={rule}
          options={[
            {
              key: 'daily',
              iconProps: { iconName: 'CalendarDay' },
              text: strings.DailyLabel
            },
            {
              key: 'weekly',
              iconProps: { iconName: 'CalendarWeek' },
              text: strings.WeeklyLabel
            },
            {
              key: 'monthly',
              iconProps: { iconName: 'Calendar' },
              text:strings.MonthlyLabel

            },
            {
              key: 'yearly',
              iconProps: { iconName: 'Calendar' },
              text: strings.YearlyLabel,
            }
          ]}
          onChange={onRuleChange}
        />
      </div>
      {
        (rule=="daily") &&
        <BizpDailyRecurrentEvent event = {props.event} eventSeries = {props.eventSeries}  entryType = {props.entryType} startDateChange = {props.startDateChange} infoRequest = {props.infoRequest} returnInfo={returnDailyInfo}/>
      }
      {
        (rule=="weekly") &&
        <BizpWeeklyRecurrentEvent event = {props.event} eventSeries = {props.eventSeries} entryType = {props.entryType} startDateChange = {props.startDateChange} infoRequest = {props.infoRequest} returnInfo={returnWeeklyInfo}/>

      }
      {
        (rule=="monthly") &&
        <BizpMonthlyRecurrentEvent event = {props.event} eventSeries = {props.eventSeries} entryType = {props.entryType} startDateChange = {props.startDateChange} infoRequest = {props.infoRequest} returnInfo={returnMonthlyInfo}/>

      }
      {
        (rule=="yearly") &&
        <BizpYearlyRecurrentEvent event = {props.event} eventSeries = {props.eventSeries} entryType = {props.entryType} startDateChange = {props.startDateChange} infoRequest = {props.infoRequest} returnInfo={returnYearlyInfo}/>

      }

    </div>
  );
}
