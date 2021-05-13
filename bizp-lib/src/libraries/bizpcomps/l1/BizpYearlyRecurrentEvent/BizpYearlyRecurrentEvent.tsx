import * as React from 'react';
import {useState,useEffect} from 'react';
import * as strings from 'BizpcompsLibraryStrings';
import styles from './BizpYearlyRecurrentEvent.module.scss';
import { IBizpYearlyRecurrentEventProps} from './IBizpYearlyRecurrentEventProps';
import { BizpRecurrentEventDateRange } from '../BizpRecurrentEventDateRange/BizpRecurrentEventDateRange';
import { IBizpRecurrenceDateRange} from '../../../../shared/IBizpSharedInterface';

import * as moment from 'moment';
import {
  Label,
  ChoiceGroup,
  IChoiceGroupOption,
  MaskedTextField,
  Dropdown,IDropdownOption
}
from '@fluentui/react';

interface IBizpYearlyItems {
  month: string;
  dayOfMonth: string;
  disableDayOfMonth: boolean;
  errMsgDayOfMonth: string;
  dayOption: string;
  dayPatternMonth: string;
  pattern: string;
  orderInMonth: string;
}


export function BizpYearlyRecurrentEvent(props: IBizpYearlyRecurrentEventProps) {
  const [month,setMonth] = useState (moment().format("M"));
  const [dayOfMonth,setDayOfMonth] = useState ((moment().date()).toString());
  const [disableDayOfMonth,setDisableDayOfMonth] = useState (false);
  const [errMsgDayOfMonth,setErrMsgDayOfMonth] = useState("");
  const [dayOption,setDayOption] = useState ("day");
  const [dayPatternMonth,setDayPatternMonth] = useState (moment().format("M"));

  const [pattern,setPattern] = useState("byDay");

  const [orderInMonth,setOrderInMonth] = useState ("first");

  const [yearlyItems,setYearlyItems] = useState<IBizpYearlyItems> (
    {
      month: moment().format("M"),
      dayOfMonth: (moment().date()).toString(),
      disableDayOfMonth: false,
      errMsgDayOfMonth: "",
      dayOption: "day",
      dayPatternMonth: moment().format("M"),
      pattern: "byDay",
      orderInMonth: "first"
    }
  );

  useEffect(() =>
    {
      // initialize category options
      if ((props.eventSeries != undefined) &&
          (props.eventSeries.yearlyRecurrence != undefined)){
            setYearlyItems({
              ...yearlyItems,
              month:props.eventSeries.yearlyRecurrence.month,
              dayOfMonth: props.eventSeries.yearlyRecurrence.dayOfMonth,
              disableDayOfMonth: (props.eventSeries.yearlyRecurrence.pattern == 'byDayPattern')?true:yearlyItems.disableDayOfMonth,
              orderInMonth: props.eventSeries.yearlyRecurrence.orderInMonth,
              dayOption: props.eventSeries.yearlyRecurrence.dayOption,
              dayPatternMonth:props.eventSeries.yearlyRecurrence.dayPatternMonth,
              pattern: props.eventSeries.yearlyRecurrence.pattern
            });
      }
    },[]
  );

 function onDayOfMonthChange(ev: React.SyntheticEvent<HTMLElement>, value: string) {
    ev.preventDefault();
    setTimeout(() => {
      let errorMessage = '';
      if (Number(value.trim()) == 0 || Number(value.trim()) > 31) {
        value = '1 ';
        errorMessage = strings.DayValidationMessage;
      }
      setYearlyItems({...yearlyItems,dayOfMonth:value,errMsgDayOfMonth:errorMessage});
    }, 2500);
  }

  function onMonthChange(ev: React.SyntheticEvent<HTMLElement>, item: IDropdownOption) {
    setYearlyItems({...yearlyItems,month:item.key.toString()});
  }

  function  onOrderInMonthChange(ev: React.FormEvent<HTMLDivElement>, item: IDropdownOption):void {
    setYearlyItems({...yearlyItems,orderInMonth:item.key.toString()});
  }

  function onDayOptionChange(ev: React.FormEvent<HTMLDivElement>, item: IDropdownOption):void {
    setYearlyItems({...yearlyItems,dayOption:item.key.toString()});
  }

  function onDayPatternMonthChange(ev: React.SyntheticEvent<HTMLElement>, item: IDropdownOption) {
    setYearlyItems({...yearlyItems,dayPatternMonth:item.key.toString()});
  }

  function onPatternChange(ev: React.SyntheticEvent<HTMLElement>, option: IChoiceGroupOption): void {
    ev.preventDefault();
    setYearlyItems({...yearlyItems,pattern:option.key,disableDayOfMonth:( option.key == 'byDay' ? false : true)});
  }

  function returnDateRangeInfo(dateRangeData:IBizpRecurrenceDateRange) {
    props.returnInfo({
      dateRangeInfo:dateRangeData,
      pattern:yearlyItems.pattern,
      month:yearlyItems.month,
      dayOfMonth:yearlyItems.dayOfMonth,
      orderInMonth:yearlyItems.orderInMonth,
      dayOption:yearlyItems.dayOption,
      dayPatternMonth:yearlyItems.dayPatternMonth
    });
  }

  console.log("Rendering YearlyRecurrentEvent...");
  return (
    <div>
      {
      <div>
      <div style={{ display: 'inline-block', float: 'right', paddingTop: '10px', height: '40px' }}></div>
      <div style={{ width: '100%', paddingTop: '10px' }}>
          <Label>{strings.PaternLabel}</Label>
          <ChoiceGroup
                selectedKey={yearlyItems.pattern}
                options={[
                  {
                    key: 'byDay',
                    text: strings.OnEveryLabel,
                    ariaLabel: strings.OnEveryLabel,
                    onRenderField: (props1, render) => {
                      return (
                        <div >
                          {render!(props1)}
                          <div style={{ display: 'inline-block', verticalAlign: 'top', width: '100px', paddingLeft: '10px' }}>
                            <Dropdown
                              selectedKey={yearlyItems.month}
                              onChange={onMonthChange}
                              disabled={yearlyItems.disableDayOfMonth}
                              options={[
                                { key: '1', text: strings.JanuaryLabel },
                                { key: '2', text: strings.FebruaryLabel },
                                { key: '3', text: strings.MarchLabel },
                                { key: '4', text: strings.AprilLabel },
                                { key: '5', text: strings.MayLabel },
                                { key: '6', text: strings.JuneLabel },
                                { key: '7', text: strings.JulyLabel },
                                { key: '8', text: strings.AugustLabel },
                                { key: '9', text: strings.SeptemberLabel },
                                { key: '10', text: strings.OctoberLabel },
                                { key: '11', text: strings.NovemberLabel },
                                { key: '12', text: strings.DecemberLabel },
                              ]}
                            />
                          </div>

                          <MaskedTextField
                            styles={{ root: { display: 'inline-block', verticalAlign: 'top', width: '100px', paddingLeft: '10px' } }}
                            mask="99"
                            maskChar=' '
                            disabled={yearlyItems.disableDayOfMonth}
                            value={yearlyItems.dayOfMonth}
                            errorMessage={yearlyItems.errMsgDayOfMonth}
                            onChange={onDayOfMonthChange} />
                        </div>
                      );
                    }
                  },

                  {
                    key: 'byDayPattern',
                    text: strings.OnTheLabel,
                    onRenderField: (props1, render) => {
                      return (
                        <div  >
                          {render!(props1)}
                          <div style={{ display: 'inline-block', verticalAlign: 'top', width: '80px', paddingLeft: '10px' }}>
                            <Dropdown
                              selectedKey={yearlyItems.orderInMonth}
                              onChange={onOrderInMonthChange}
                              disabled={!yearlyItems.disableDayOfMonth}
                              options={[
                                { key: 'first', text: strings.FirstLabel },
                                { key: 'second', text: strings.SecondLabel},
                                { key: 'third', text: strings.ThirdLabel },
                                { key: 'fourth', text: strings.FourthLabel },
                                { key: 'last', text: strings.LastLabel },

                              ]}
                            />
                          </div>
                          <div style={{ display: 'inline-block', verticalAlign: 'top', width: '100px', paddingLeft: '5px' }}>
                            <Dropdown
                              selectedKey={yearlyItems.dayOption}
                              disabled={!yearlyItems.disableDayOfMonth}
                              onChange={onDayOptionChange}
                              options={[
                                { key: 'day', text: strings.DayLabel },
                                { key: 'weekday', text: strings.WeekDayLabel },
                                { key: 'weekendday', text:strings.WeekendDayLabel },
                                { key: 'sunday', text: strings.SundayLabel},
                                { key: 'monday', text: strings.MondayLabel },
                                { key: 'tuesday', text: strings.TuesdayLabel },
                                { key: 'wednesday', text: strings.WednesdayLabel },
                                { key: 'thursday', text: strings.ThursdayLabel},
                                { key: 'friday', text: strings.FridayLabel },
                                { key: 'saturday', text: strings.SaturdayLabel },
                                ]}
                            />
                          </div>
                          <Label styles={{ root: { display: 'inline-block', verticalAlign: 'top', width: '30px', paddingLeft: '10px' } }}>{ "of"} </Label>
                          <div style={{ display: 'inline-block', verticalAlign: 'top', width: '100px', paddingLeft: '5px' }}>
                            <Dropdown
                              selectedKey={yearlyItems.dayPatternMonth}
                              onChange={onDayPatternMonthChange}
                              disabled={!yearlyItems.disableDayOfMonth}
                              options={[
                                { key: '1', text: strings.JanuaryLabel },
                                { key: '2', text: strings.FebruaryLabel },
                                { key: '3', text: strings.MarchLabel },
                                { key: '4', text: strings.AprilLabel },
                                { key: '5', text: strings.MayLabel },
                                { key: '6', text: strings.JuneLabel },
                                { key: '7', text: strings.JulyLabel },
                                { key: '8', text: strings.AugustLabel },
                                { key: '9', text: strings.SeptemberLabel },
                                { key: '10', text: strings.OctoberLabel },
                                { key: '11', text: strings.NovemberLabel },
                                { key: '12', text: strings.DecemberLabel },
                              ]}
                            />
                          </div>

                        </div>
                      );
                    }

                  }
                ]}
                onChange={onPatternChange}
                required={true}
          />
      </div>
      <BizpRecurrentEventDateRange event = {props.event} dateRange = {(props.eventSeries == undefined) || (props.eventSeries.yearlyRecurrence == undefined) ? undefined : props.eventSeries.yearlyRecurrence.dateRangeInfo} entryType = {props.entryType} startDateChange = {props.startDateChange} returnInfo={returnDateRangeInfo} infoRequest={props.infoRequest}/>
    </div>
    }
    </div>
  );
}
