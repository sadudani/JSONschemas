import * as React from 'react';
import {useState,useEffect} from 'react';
import * as strings from 'BizpcompsLibraryStrings';
import styles from './BizpDailyRecurrentEvent.module.scss';
import { IBizpDailyRecurrentEventProps } from './IBizpDailyRecurrentEventProps';
import { BizpRecurrentEventDateRange } from '../BizpRecurrentEventDateRange/BizpRecurrentEventDateRange';
import { IBizpRecurrenceDateRange} from '../../../../shared/IBizpSharedInterface';
import {
  Label,
  ChoiceGroup,
  IChoiceGroupOption,
  MaskedTextField,
}
from '@fluentui/react';

interface IBizpDailyItems {
  dayFrequency: string;
  disableDayFrequency: boolean;
  pattern: string;
  errMsgDayFrequency: string;
}

export function BizpDailyRecurrentEvent(props: IBizpDailyRecurrentEventProps) {

  const [dailyItems,setDailyItems] = useState<IBizpDailyItems>(
    {
      dayFrequency: "1",
      disableDayFrequency: false,
      pattern: "every",
      errMsgDayFrequency: ""
    }
  );

  useEffect(() =>
  {
      // initialize
    if ((props.eventSeries != undefined) &&
        (props.eventSeries.dailyRecurrence != undefined)){
      setDailyItems({...dailyItems,pattern:props.eventSeries.dailyRecurrence.pattern,dayFrequency:props.eventSeries.dailyRecurrence.frequency});
    }
  },[]
  );

  function onDayFrequencyChange(ev: React.SyntheticEvent<HTMLElement>, value: string) {
    ev.preventDefault();
    setTimeout(() => {
      if (Number(value.trim()) == 0 || Number(value.trim()) > 255) {
        setDailyItems({...dailyItems,dayFrequency:'1  ',errMsgDayFrequency:'Allowed values 1 to 255'});
      }
      else {
        setDailyItems({...dailyItems,dayFrequency:value,errMsgDayFrequency:""});
      }
    }, 2500);
  }

  function onPatternChange(ev: React.SyntheticEvent<HTMLElement>, option: IChoiceGroupOption): void {
    ev.preventDefault();
    setDailyItems({...dailyItems,pattern:option.key,disableDayFrequency:(option.key == 'every' ? false : true)});
  }

  function returnDateRangeInfo(dateRangeData:IBizpRecurrenceDateRange) {
    props.returnInfo({
      dateRangeInfo:dateRangeData,
      frequency:dailyItems.dayFrequency,
      pattern:dailyItems.pattern});
  }
  console.log("Rendering DailyRecurrentEvent...");
  return (
    <div >
      {
        <div>
          <div style={{ display: 'inline-block', float: 'right', paddingTop: '10px', height: '40px' }}>
          </div>
          <div style={{ width: '100%', paddingTop: '10px' }}>
            <Label>{ strings.PatternLabel }</Label>
            <ChoiceGroup

              defaultSelectedKey = 'every'
              options={[
                {
                  key: 'every',
                  text: strings.EveryLabel,
                  ariaLabel: strings.EveryLabel,

                  onRenderField: (props1, render) => {
                    return (
                      <div  >
                        {render!(props1)}
                        <MaskedTextField
                          styles={{ root: { display: 'inline-block', verticalAlign: 'top', width: '100px', paddingLeft: '10px' } }}
                          mask="999"
                          maskChar=' '
                          disabled={dailyItems.disableDayFrequency}
                          value={dailyItems.dayFrequency}
                          errorMessage={dailyItems.errMsgDayFrequency}
                          onChange={onDayFrequencyChange} />
                        <Label styles={{ root: { display: 'inline-block', verticalAlign: 'top', width: '60px', paddingLeft: '10px' } }}>{strings.daysLabel}</Label>
                      </div>
                    );
                  }
                },
                {
                  key: 'everyweekday',
                  text: strings.everyweekdayLabel,
                }
              ]}
              onChange={onPatternChange}
              required={true}
            />
          </div>
          <BizpRecurrentEventDateRange event = {props.event} dateRange = {(props.eventSeries == undefined)||(props.eventSeries.dailyRecurrence == undefined) ? undefined : props.eventSeries.dailyRecurrence.dateRangeInfo} entryType = {props.entryType} startDateChange = {props.startDateChange} returnInfo={returnDateRangeInfo} infoRequest = {props.infoRequest}/>
        </div>
      }
    </div>
  );
}

