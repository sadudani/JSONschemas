import * as React from 'react';
import {useState,useEffect} from 'react';
import * as strings from 'BizpcompsLibraryStrings';
import styles from './BizpWeeklyRecurrentEvent.module.scss';
import { IBizpWeeklyRecurrentEventProps} from './IBizpWeeklyRecurrentEventProps';
import { BizpRecurrentEventDateRange } from '../BizpRecurrentEventDateRange/BizpRecurrentEventDateRange';
import { IBizpRecurrenceDateRange,IBizpDaysCheck  } from '../../../../shared/IBizpSharedInterface';
import {
  Label,
  MaskedTextField,
  Checkbox
}
from '@fluentui/react';

interface IBizpWeeklyItems {
  weekFrequency:string;
  errMsgWeekFrequency:string;
  daysOption: IBizpDaysCheck;
}
export function BizpWeeklyRecurrentEvent(props: IBizpWeeklyRecurrentEventProps) {
  const [weeklyItems,setWeeklyItems] = useState<IBizpWeeklyItems>(
    {
      weekFrequency:"1",
      errMsgWeekFrequency:"",
      daysOption: {sunday:false,monday:true,tuesday:false,wednesday:false,thursday:false,friday:false,saturday:false}
    }
  );
  useEffect(() =>
    {
      // initialize category options
      if ((props.eventSeries != undefined) &&
          (props.eventSeries.weeklyRecurrence != undefined)){
            setWeeklyItems({...weeklyItems,weekFrequency:props.eventSeries.weeklyRecurrence.frequency,
                            daysOption:props.eventSeries.weeklyRecurrence.daysOption});
      }
    },[]
  );

  function onWeekFrequencyChange(ev: React.SyntheticEvent<HTMLElement>, value: string) {
    ev.preventDefault();
    setTimeout(() => {
      if (Number(value.trim()) == 0 || Number(value.trim()) > 255) {
        setWeeklyItems({...weeklyItems,weekFrequency:'1  ',errMsgWeekFrequency:strings.WeekFrequencyValidationMsg});
      }
      else {
        setWeeklyItems({...weeklyItems,weekFrequency:'1  ',errMsgWeekFrequency:""});
      }
    }, 2500);
  }

  const onCheckboxDayChange = (day:keyof IBizpDaysCheck) => (ev: React.FormEvent<HTMLElement>, isChecked: boolean) => {
    let checks:IBizpDaysCheck;
    checks = weeklyItems.daysOption;
    checks[day]=isChecked;
    setWeeklyItems({...weeklyItems,daysOption:{...checks}});
  };

  function returnDateRangeInfo(dateRangeData:IBizpRecurrenceDateRange) {
    props.returnInfo({dateRangeInfo:dateRangeData,frequency:weeklyItems.weekFrequency,daysOption:weeklyItems.daysOption});
  }

  console.log("Rendering WeeklyRecurrentEvent...");
  return (
    <div >
      {
        <div>
          <div style={{ display: 'inline-block', float: 'right', paddingTop: '10px', height: '40px' }}>

          </div>
          <div style={{ width: '100%', paddingTop: '10px' }}>
            <Label>{"Pattern"}</Label>
            <div>
              <Label styles={{ root: { display: 'inline-block', verticalAlign: 'top', width: '40px' } }}>{strings.EveryLabel}</Label>
              <MaskedTextField
                styles={{ root: { display: 'inline-block', verticalAlign: 'top', width: '100px', paddingLeft: '5px' } }}
                mask="999"
                maskChar=' '
                errorMessage={weeklyItems.errMsgWeekFrequency}
                value={weeklyItems.weekFrequency}
                onChange={onWeekFrequencyChange} />
              <Label styles={{ root: { display: 'inline-block', verticalAlign: 'top', width: '80px', paddingLeft: '10px' } }}>{strings.WeeksOnLabel}</Label>

            </div>
            <div style={{ marginTop: '10px' }}>
              <Checkbox label={strings.SundayLabel} className={styles.ckeckBoxInline} checked={weeklyItems.daysOption.sunday} onChange={onCheckboxDayChange("sunday")} />
              <Checkbox label={strings.MondayLabel} className={styles.ckeckBoxInline} checked={weeklyItems.daysOption.monday} onChange={onCheckboxDayChange("monday")} />
              <Checkbox label={strings.TuesdayLabel} className={styles.ckeckBoxInline} checked={weeklyItems.daysOption.tuesday} onChange={onCheckboxDayChange("tuesday")} />
              <Checkbox label={strings.WednesdayLabel} className={styles.ckeckBoxInline} checked={weeklyItems.daysOption.wednesday} onChange={onCheckboxDayChange("wednesday")} />
            </div>
            <div style={{ marginTop: '10px' }}>
              <Checkbox label={strings.ThursdayLabel} className={styles.ckeckBoxInline} checked={weeklyItems.daysOption.thursday} onChange={onCheckboxDayChange("thursday")} />
              <Checkbox label={strings.FridayLabel} className={styles.ckeckBoxInline} checked={weeklyItems.daysOption.friday} onChange={onCheckboxDayChange("friday")} />
              <Checkbox label={strings.SaturdayLabel} className={styles.ckeckBoxInline} checked={weeklyItems.daysOption.saturday} onChange={onCheckboxDayChange("saturday")} />
            </div>
          </div>
          <BizpRecurrentEventDateRange event = {props.event} dateRange = {(props.eventSeries == undefined)||(props.eventSeries.weeklyRecurrence == undefined) ? undefined : props.eventSeries.weeklyRecurrence.dateRangeInfo} entryType = {props.entryType} startDateChange = {props.startDateChange} returnInfo={returnDateRangeInfo} infoRequest={props.infoRequest}/>
        </div>
      }
    </div>
  );
}

