import * as React from 'react';
import {useRef,useState,useEffect} from 'react';
import * as strings from 'BizpcompsLibraryStrings';
import styles from './BizpRecurrentEventDateRange.module.scss';
import { IBizpRecurrentEventDateRangeProps} from './IBizpRecurrentEventDateRangeProps';
import { toLocaleShortDateString} from '../../../../shared/BizpBasesvc';
import * as moment from 'moment';
import {
  Label,
  ChoiceGroup,
  IChoiceGroupOption,
  MaskedTextField,
  DatePicker, DayOfWeek, IDatePickerStrings
}
from '@fluentui/react';
import { IBizpEntryTypeEnum } from '../../../../shared/IBizpSharedInterface';

interface IBizpDateRangeItems {
  endDate: Date;
  disableEndDate:boolean;
  noOfOccurrences: string;
  errMsgNoOfOccurrences: string;
  disableNoOfOcurrences: boolean;
  dateRangeOption: string;
}
export function BizpRecurrentEventDateRange(props: IBizpRecurrentEventDateRangeProps) {
  const [dateRangeItems,setDateRangeItems] = useState <IBizpDateRangeItems> (
    {
      endDate: new Date(props.event.startDate),
      disableEndDate: false,
      noOfOccurrences: "1",
      errMsgNoOfOccurrences: "",
      disableNoOfOcurrences: false,
      dateRangeOption: "endAfter",
    });

  const DayPickerStrings: IDatePickerStrings = {
    months: [strings.JanuaryLabel, strings.FebruaryLabel, strings.MarchLabel, strings.AprilLabel, strings.MayLabel,
            strings.JuneLabel, strings.JulyLabel, strings.AugustLabel, strings.SeptemberLabel, strings.OctoberLabel, strings.NovemberLabel, strings.DecemberLabel],
    shortMonths: [strings.JanLabel, strings.FebLabel, strings.MarLabel, strings.AprLabel, strings.MayLabel, strings.JunLabel, strings.JulLabel,
      strings.AugLabel, strings.SepLabel, strings.OctLabel, strings.NovLabel, strings.DecLabel],
    days: [strings.SundayLabel, strings.MondayLabel, strings.TuesdayLabel, strings.WednesdayLabel, strings.ThursdayLabel, strings.FridayLabel, strings.SaturdayLabel],
    shortDays: [strings.ShortDay_SuLabel, strings.ShortDay_MoLabel, strings.ShortDay_TuLabel, strings.ShortDay_WeLabel, strings.ShortDay_ThLabel, strings.ShortDay_FrLabel, strings.ShortDay_SaLabel],
    goToToday: strings.GoToDayLabel,
    prevMonthAriaLabel: strings.PrevMonthLabel,
    nextMonthAriaLabel: strings.NextMonthLabel,
    prevYearAriaLabel: strings.PrevYearLabel,
    nextYearAriaLabel: strings.NextYearLabel,
    closeButtonAriaLabel: strings.CloseDateLabel,
    isRequiredErrorMessage: strings.IsRequiredLabel,
    invalidInputErrorMessage: strings.InvalidDateFormatMsg
  };

  useEffect(()=> {
    if ((props.entryType == IBizpEntryTypeEnum.editSeries) && (props.dateRange != undefined)) {
      let d: string | Date = props.event.startDate;
      if ((props.dateRange.option == 'endDate') && (typeof props.dateRange.endDate == 'string')) {
        d = new Date(props.dateRange.endDate);
      }
      setDateRangeItems({...dateRangeItems,endDate:d,
        noOfOccurrences:props.dateRange.frequency,dateRangeOption:props.dateRange.option});
    }
  },[]);

  const firstRun = useRef(true);

  // this is used to return the internal states
  useEffect(() => {
   // skip the first run
   if (firstRun.current) {
    firstRun.current = false;
    return;
    }
    props.returnInfo({
      startDate:props.event.startDate,
      endDate:dateRangeItems.endDate,
      frequency:dateRangeItems.noOfOccurrences,
      option:dateRangeItems.dateRangeOption
      });
  },[props.infoRequest]);

  function onDateRangeOptionChange(ev: React.SyntheticEvent<HTMLElement>, option: IChoiceGroupOption): void {
    ev.preventDefault();
    setDateRangeItems({...dateRangeItems,
      disableNoOfOcurrences:(option.key == 'endAfter' ? false : true),
      disableEndDate:(option.key == 'endDate' ? false : true),
      dateRangeOption:option.key});
  }

  function onStartDateChange(date: Date) {
    props.startDateChange(date);
    // prevent end date to be before the start date
    if (moment(date).isAfter(dateRangeItems.endDate)) {
      setDateRangeItems({...dateRangeItems,endDate:date});
    }
  }

  function onEndDateChange(date: Date) {
    setDateRangeItems({...dateRangeItems,endDate:date});
  }

  function onNoOfOcurrencesChange(ev: React.SyntheticEvent<HTMLElement>, value: string) {
    ev.preventDefault();
    setTimeout(() => {
      if (Number(value.trim()) == 0 || Number(value.trim()) > 999) {
        setDateRangeItems({...dateRangeItems,noOfOccurrences:'1  ',errMsgNoOfOccurrences:strings.InstancesValidationMsg});
      }
      else {
        setDateRangeItems({...dateRangeItems,noOfOccurrences:value.trim(),errMsgNoOfOccurrences:""});
      }
    }, 2500);

  }
  console.log("Rendering RecurrentEventDateRange...");
  return (
    <div style={{ paddingTop: '22px' }}>
    <Label>{ strings.DateRangeLabel }</Label>
    <div style={{ display: 'inline-block', verticalAlign: 'top', paddingRight: '35px', paddingTop: '10px' }}>
      <DatePicker
        firstDayOfWeek={DayOfWeek.Sunday}
        strings={DayPickerStrings}
        placeholder={strings.SelectDateLabel}
        ariaLabel={strings.SelectDateLabel}
        label={strings.StartDateLabel}
        value={props.event.startDate}
        onSelectDate={onStartDateChange}
        formatDate={toLocaleShortDateString}
      />

    </div>
    <div style={{ display: 'inline-block', verticalAlign: 'top', paddingTop: '10px' }}>
      <ChoiceGroup
        disabled = {props.entryType == IBizpEntryTypeEnum.viewEvent}
        selectedKey={dateRangeItems.dateRangeOption}
        onChange={onDateRangeOptionChange}
        options={[
          {
            key: 'noDate',
            text: strings.NoEndDate,
          },
          {
            key: 'endDate',
            text: strings.EndByLabel,
            onRenderField: (props1, render) => {
              return (
                <div  >
                  {render!(props1)}
                  <DatePicker
                    firstDayOfWeek={DayOfWeek.Sunday}
                    strings={DayPickerStrings}
                    placeholder={strings.SelectDateLabel}
                    ariaLabel= {strings.SelectDateLabel}
                    style={{ display: 'inline-block', verticalAlign: 'top', paddingLeft: '22px', }}
                    onSelectDate={onEndDateChange}
                    formatDate={toLocaleShortDateString}
                    value={dateRangeItems.endDate}
                    disabled={dateRangeItems.disableEndDate}
                    minDate={props.event.startDate}
                  />
                </div>
              );
            }
          },
          {
            key: 'endAfter',
            text: strings.EndAfterLabel,
            onRenderField: (props1, render) => {
              return (
                <div>
                  {render!(props1)}
                  <MaskedTextField
                    styles={{ root: { display: 'inline-block', verticalAlign: 'top', width: '100px', paddingLeft: '10px' } }}
                    mask="999"
                    maskChar=' '
                    value={dateRangeItems.noOfOccurrences}
                    disabled={dateRangeItems.disableNoOfOcurrences}
                    errorMessage={dateRangeItems.errMsgNoOfOccurrences}
                    onChange={onNoOfOcurrencesChange} />
                  <Label styles={{ root: { display: 'inline-block', verticalAlign: 'top', paddingLeft: '10px' } }}>{ "occurrences" }</Label>
                </div>
              );
            }
          },
        ]}
        required={true}
      />
    </div>
    </div>
  );
}
