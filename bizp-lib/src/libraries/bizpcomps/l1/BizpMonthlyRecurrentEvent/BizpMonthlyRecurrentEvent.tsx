import * as React from 'react';
import {useState,useEffect} from 'react';
import * as strings from 'BizpcompsLibraryStrings';
import styles from './BizpMonthlyRecurrentEvent.module.scss';
import { IBizpMonthlyRecurrentEventProps} from './IBizpMonthlyRecurrentEventProps';
import { IBizpRecurrenceDateRange} from '../../../../shared/IBizpSharedInterface';
import {
  Label,
  ChoiceGroup,
  IChoiceGroupOption,
  MaskedTextField,
  Dropdown,IDropdownOption
}
from '@fluentui/react';
import { BizpRecurrentEventDateRange } from '../BizpRecurrentEventDateRange/BizpRecurrentEventDateRange';

interface IBizpMonthlyItems {
  disableDayOfMonth:boolean;
  dayOfMonth: string;
  errMsgDayOfMonth: string;
  monthlyFrequency: string;
  errMonthlyFrequency: string;
  orderInMonth: string;
  dayOption: string;
  pattern: string;
}

export function BizpMonthlyRecurrentEvent(props: IBizpMonthlyRecurrentEventProps) {
  const [pattern,setPattern] = useState("dayOfMonth");
  const [disableDayOfMonth,setDisableDayOfMonth] = useState (false);
  const [dayOfMonth,setDayOfMonth] = useState ("1");
  const [errMsgDayOfMonth,setErrMsgDayOfMonth] = useState("");
  const [monthlyFrequency,setMonthlyFrequency] = useState ("1");
  const [errMonthlyFrequency,setErrMsgMonthlyFrequency] = useState("");
  const [orderInMonth,setOrderInMonth] = useState (strings.FirstLabel);
  const [dayOption,setDayOption] = useState (strings.DayLabel);

const [monthlyItems,setMonthlyItems] = useState<IBizpMonthlyItems>(
  {
    disableDayOfMonth:false,
    dayOfMonth: "1",
    errMsgDayOfMonth: "",
    monthlyFrequency: "1",
    errMonthlyFrequency: "",
    orderInMonth: strings.FirstLabel,
    dayOption: strings.DayLabel,
    pattern: "dayOfMonth"
  }
);

  useEffect(() =>
    {
      // initialize category options
      if ((props.eventSeries != undefined) &&
          (props.eventSeries.monthlyRecurrence != undefined)){
            setMonthlyItems({...monthlyItems,
              monthlyFrequency:props.eventSeries.monthlyRecurrence.frequency,
              pattern: props.eventSeries.monthlyRecurrence.pattern,
              dayOfMonth: props.eventSeries.monthlyRecurrence.dayOfMonth,
              orderInMonth: props.eventSeries.monthlyRecurrence.orderInMonth,
              dayOption: props.eventSeries.monthlyRecurrence.dayOption,
              disableDayOfMonth:(props.eventSeries.monthlyRecurrence.pattern == "dayOfMonth")?false:true,
            });
//        init();
      }
    },[]
  );

  function init() {
    setMonthlyFrequency(props.eventSeries.monthlyRecurrence.frequency);
    setPattern(props.eventSeries.monthlyRecurrence.pattern);
    setDayOfMonth(props.eventSeries.monthlyRecurrence.dayOfMonth);
    setOrderInMonth(props.eventSeries.monthlyRecurrence.orderInMonth);
    setDayOption(props.eventSeries.monthlyRecurrence.dayOption);
    if (props.eventSeries.monthlyRecurrence.pattern == "dayOfMonth") {
      setDisableDayOfMonth(false);
    }
    else {
      setDisableDayOfMonth(true);
    }

  }

 function onDayOfMonthChange(ev: React.SyntheticEvent<HTMLElement>, value: string) {
    ev.preventDefault();
    setTimeout(() => {
      let errorMessage = '';
      if (Number(value.trim()) == 0 || Number(value.trim()) > 31) {
        value = '1 ';
        errorMessage = strings.DayValidationMessage;
      }
      setMonthlyItems({...monthlyItems,dayOfMonth:value,errMsgDayOfMonth:errorMessage});
//      setDayOfMonth(value);
//     setErrMsgDayOfMonth(errorMessage);
    }, 2500);
  }

  function onMonthlyFrequencyChange(ev: React.SyntheticEvent<HTMLElement>, value: string) {
    ev.preventDefault();
    setTimeout(() => {
      let errorMessage = '';
      if (Number(value.trim()) == 0 || Number(value.trim()) > 12) {
        value = '1 ';
        errorMessage = strings.MonthValidationMessage;
      }
      setMonthlyItems({...monthlyItems,monthlyFrequency:value,errMsgDayOfMonth:errorMessage});

//      setMonthlyFrequency(value);
 //     setErrMsgDayOfMonth(errorMessage);
    }, 2500);
  }

  function  onOrderInMonthChange(ev: React.FormEvent<HTMLDivElement>, item: IDropdownOption):void {
//    setOrderInMonth(item.key.toString());
    setMonthlyItems({...monthlyItems,orderInMonth:item.key.toString()});

  }

  function onWeekdayChange(ev: React.FormEvent<HTMLDivElement>, item: IDropdownOption):void {
//    setDayOption(item.key.toString());
    setMonthlyItems({...monthlyItems,dayOption:item.key.toString()});

  }

  function onPatternChange(ev: React.SyntheticEvent<HTMLElement>, option: IChoiceGroupOption): void {
    ev.preventDefault();
//    setPattern(option.key);
 //   setDisableDayOfMonth( option.key == 'orderInMonth' ? true : false);
    setMonthlyItems({...monthlyItems,disableDayOfMonth:( option.key == 'orderInMonth' ? true : false),pattern:option.key});

  }

  function returnDateRangeInfo(dateRangeData:IBizpRecurrenceDateRange) {
    props.returnInfo({
      dateRangeInfo:dateRangeData,
      frequency:monthlyItems.monthlyFrequency,
      pattern:monthlyItems.pattern,
      dayOfMonth:monthlyItems.dayOfMonth,
      orderInMonth:monthlyItems.orderInMonth,
      dayOption:monthlyItems.dayOption
    });
  }
  console.log("Rendering MonthlyRecurrentEvent...");
  return (
    <div >
      {
        <div>
          <div style={{ display: 'inline-block', float: 'right', paddingTop: '10px', height: '40px' }}>
          </div>
          <div style={{ width: '100%', paddingTop: '10px' }}>
            <Label>{ "Pattern" }</Label>
            <ChoiceGroup
              selectedKey={monthlyItems.pattern}
              options={[
                {
                  key: 'dayOfMonth',
                  text: strings.DayLabel,
                  ariaLabel:  strings.DayLabel,

                  onRenderField: (props1, render) => {
                    return (
                      <div  >
                        {render!(props1)}
                        <MaskedTextField
                          styles={{ root: { display: 'inline-block', verticalAlign: 'top', width: '100px', paddingLeft: '10px' } }}
                          mask="99"
                          maskChar=' '
                          disabled={monthlyItems.disableDayOfMonth}
                          value={monthlyItems.dayOfMonth}
                          errorMessage={monthlyItems.errMsgDayOfMonth}
                          onChange={onDayOfMonthChange} />
                        <Label styles={{ root: { display: 'inline-block', verticalAlign: 'top', width: '65px', paddingLeft: '10px' } }}>{strings.OfEveryLabel}</Label>
                        <MaskedTextField
                          styles={{ root: { display: 'inline-block', verticalAlign: 'top', width: '100px', paddingLeft: '10px' } }}
                          mask="99"
                          maskChar=' '
                          disabled={monthlyItems.disableDayOfMonth}
                          value={monthlyItems.monthlyFrequency}
                          errorMessage={monthlyItems.errMonthlyFrequency}
                          onChange={onMonthlyFrequencyChange} />
                        <Label styles={{ root: { display: 'inline-block', verticalAlign: 'top', width: '100px', paddingLeft: '10px' } }}>{"month(s)"}</Label>
                      </div>
                    );
                  }
                },
                {
                  key: 'orderInMonth',
                  text: strings.TheLabel,
                  onRenderField: (props1, render) => {
                    return (
                      <div  >
                        {render!(props1)}
                        <div style={{ display: 'inline-block', verticalAlign: 'top', width: '90px', paddingLeft: '10px' }}>
                          <Dropdown
                            selectedKey={monthlyItems.orderInMonth}
                            onChange={onOrderInMonthChange}
                            disabled={!monthlyItems.disableDayOfMonth}
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
                            selectedKey={monthlyItems.dayOption}
                            disabled={!monthlyItems.disableDayOfMonth}
                            onChange={onWeekdayChange}
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
                        <Label styles={{ root: { display: 'inline-block', verticalAlign: 'top', width: '65px', paddingLeft: '10px' } }}>{strings.OfEveryLabel}</Label>
                        <MaskedTextField
                          styles={{ root: { display: 'inline-block', verticalAlign: 'top', width: '100px', paddingLeft: '10px' } }}
                          mask="99"
                          maskChar=' '
                          disabled={!monthlyItems.disableDayOfMonth}
                          value={monthlyItems.monthlyFrequency}
                          errorMessage={monthlyItems.errMonthlyFrequency}
                          onChange={onMonthlyFrequencyChange} />
                        <Label styles={{ root: { display: 'inline-block', verticalAlign: 'top', width: '80px', paddingLeft: '10px' } }}>{strings.MonthsLabel}</Label>
                      </div>
                    );
                  }

                }
              ]}
              onChange={onPatternChange}
              required={true}
            />
          </div>
          <BizpRecurrentEventDateRange event = {props.event} entryType = {props.entryType} dateRange = {(props.eventSeries == undefined)||(props.eventSeries.monthlyRecurrence == undefined) ? undefined : props.eventSeries.monthlyRecurrence.dateRangeInfo} startDateChange = {props.startDateChange} returnInfo={returnDateRangeInfo} infoRequest={props.infoRequest}/>
        </div>
      }
    </div>
  );
}

