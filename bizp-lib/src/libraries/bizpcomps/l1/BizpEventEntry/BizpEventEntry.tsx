import * as React from 'react';
import {useState,useEffect} from 'react';

import * as strings from 'BizpcompsLibraryStrings';
import styles from './BizpEventEntry.module.scss';
import {IBizpEventEntryProps} from './IBizpEventEntryProps';
import { getFieldDropdownOptions } from '../../../../shared/BizpBasesvc';
import * as moment from 'moment';

import {
    Panel, PanelType,
    ActionButton, IIconProps,
    TextField,
    Label,
    DefaultButton,
    PrimaryButton,
    Dropdown, IDropdownOption,
    MessageBar,
    MessageBarType,
    Spinner,
    SpinnerSize,
    Toggle
} from '@fluentui/react';

import { DateTimePicker, DateConvention, TimeConvention, TimeDisplayControlType } from '@pnp/spfx-controls-react/lib/DateTimePicker';
import { BizpRecurrentEvent } from '../BizpRecurrentEvent/BizpRecurrentEvent';
import { IBizpRecurrence,IBizpNonRecurrence, IBizpCRUDEnum, IBizpEntryTypeEnum } from '../../../../shared/IBizpSharedInterface';
import { generateRandomColor } from '../../../../shared/BizpBasesvc';

interface IBizpEventEntryItems {
  seriesChecked: boolean;
  entryType: IBizpEntryTypeEnum; /* entry types are 0 (new),1 (event),2 (event from series),3 (series) */
  headerText: string;
  categoryOptions: IDropdownOption[];
}
export function BizpEventEntry(props: IBizpEventEntryProps) {
  const [saving,setSaving] = useState(false); // constant state - not used currently
  const [saveDisabled,setSaveDisable] = useState(false); // constant state - not used currently
  const [hasError,setHasError] = useState(false); //constant state -  not used currently
  const [loading,setLoading] = useState(false); // constant state - not used currently
  const [showMoreLess,setShowMoreLess] = useState(false);
  const [categoryOptions,setCategoryOptions] = useState<IDropdownOption[]> ([]);
  // category options must be a state to render correctly after initialization
  const [savedEvent,setSavedEvent] = useState<IBizpNonRecurrence>(
    {
      title: "",
      description: "",
      startDate: new Date(),
      endDate: defaultEndDate(),
      eventType: "0",
      fAllDayEvent: false,
      duration: 60,
      fRecurrence: "0",
      category: "",
      attendees: [],
      location: "",
      geolocation: { Longitude: 0, Latitude: 0},
      color: ''
    }
  );

  const [requestRecurrence,setRequestRecurrence] = useState(false);
    // states that must set in a group to avoid multiple renderings
  const [entryItems,setEntryItems] = useState <IBizpEventEntryItems>(
    {
      seriesChecked: false,
      entryType: IBizpEntryTypeEnum.newEvent,
      headerText: "",
      categoryOptions: []
    }
  );
  const [series,setSeries] = useState<IBizpRecurrence>(props.eventSeries);
  // change the callback to save series information for appropriate action
  const [saveInfoFn,setSaveInfoFn] = useState <(recurrentData:IBizpRecurrence) => void> (MoveUpSeriesData);

  useEffect(() =>
    {
      init();
      initCategoryOptions();
      // initialize saved event
      initSavedEvent();
    },[]
  );

  const addFriendIcon: IIconProps = { iconName: 'AddFriend' };
  let errorMessage:string;

  function defaultEndDate():Date {
    let t = new Date();
    t.setHours(t.getHours()+1);
    return t;
  }

//  async function init(){
async function init(){
    console.log("entering Init-1: EventEntry");
    switch (props.panelMode) {
      case IBizpCRUDEnum.add:
        {
          setEntryItems({...entryItems,headerText:strings.AddNewEventLabel,entryType:IBizpEntryTypeEnum.newEvent});
          console.log("entering Init-3: EventEntry");
        }
        break;
      case IBizpCRUDEnum.edit:
        {
          if (props.event.eventType=='1') {
            console.log("entering Init-5: EventEntry");
            if (props.series) {
              setEntryItems({...entryItems,seriesChecked:true,headerText:strings.EditEventSeriesLabel,entryType:IBizpEntryTypeEnum.editSeries});
              console.log("entering Init-7: EventEntry");
            } else {
              setEntryItems({...entryItems,seriesChecked:true,headerText:strings.EditSingleEventLabel,entryType:IBizpEntryTypeEnum.editEventFromSeries});
              console.log("entering Init-9: EventEntry");
            }
          }
          else if (props.event.eventType=='0') {
            // edit a single event
            setEntryItems({...entryItems,headerText:strings.EditEventlabel,entryType:IBizpEntryTypeEnum.editEvent});
            console.log("entering Init-11: EventEntry");
          } else {
            // edit an exception
            setEntryItems({...entryItems,headerText:strings.EditSingleEventLabel,entryType:IBizpEntryTypeEnum.editEventFromSeries});
            console.log("entering Init-13: EventEntry");
          }
          console.log("entering Init-14: EventEntry");
        }
        break;
      default:
        {
          setEntryItems({...entryItems,headerText:strings.ViewEventLabel});
          console.log("entering Init-15: EventEntry");
        }
        break;
    }
    console.log("exiting Init: EventEntry");
  }

  async function initCategoryOptions(){
    console.log("initCategoryOptions");
    // initialize calendar category options
    const options:IDropdownOption[] = await getFieldDropdownOptions(props.siteUrl, props.listId,"Category");
    setCategoryOptions(options);
    console.log("exiting initCategoryOptions-1: EventEntry");
  }

  function initSavedEvent():void {
    if (props.event) {
      setSavedEvent(
        {
          title: props.event.title,
          description: props.event.description,
          startDate: new Date(props.event.startDate),
          endDate: new Date(props.event.endDate),
          eventType: props.event.eventType,
          fAllDayEvent: props.event.fAllDayEvent,
          duration: 60,
          fRecurrence: props.event.fRecurrence,
          category: props.event.category,
          attendees: props.event.attendes,
          location: props.event.location,
          geolocation: { Longitude: 0, Latitude: 0},
          color: props.event.color
        }
      );
    }
  }

  function onCancelEntry():void {
    props.onDismissEntry(false);
  }
  // this is a call back function to return the series information from subcomonents
  // It is used to pass information to the parent component
  // When invoked from subcomponents always pass initializing = false
  function MoveUpSeriesData(recurrentData:IBizpRecurrence, initializing:boolean = true) {
    // This is also called during saveInfoFn state initialization with initializing = true
    // so ignore the initialization call
    if (initializing) return;
    // When initializing = false, take action
    props.onSaveNewEntry(savedEvent,recurrentData);
  }
  // this is a call back function to return the series information from subcomonents
  // It is used to save the series information in this component during the entry operation
  // When invoked from subcomponents always pass initializing = false
  function saveSeries(recurrentData:IBizpRecurrence, initializing:boolean = true) {
    // this is a call back function to return the series information from subcomonents
    // This is also called during saveInfoFn state initialization with initializing = true
    // so ignore the call
    if (initializing) return;
    // When initializing = false, take action
    if (entryItems.seriesChecked  && (entryItems.entryType != IBizpEntryTypeEnum.editEventFromSeries))
      setSeries(recurrentData);
  }

  function onSave() {
    let recurData:IBizpRecurrence;
    if (entryItems.entryType == IBizpEntryTypeEnum.editEventFromSeries) {
      // no series data so just push the data up
      props.onSaveNewEntry(savedEvent,recurData);
      return;
    }

    if (savedEvent.eventType == '1') {
      // send a request to get the recurrence data
      if (showMoreLess) {
        setSaveInfoFn(MoveUpSeriesData);
        setRequestRecurrence(!requestRecurrence);
      }
      else {
        props.onSaveNewEntry(savedEvent,series);
      }
    }
    else {
      // no series data so just push the data up
      props.onSaveNewEntry(savedEvent,recurData);
    }
  }

  function categoryTextToKey(t:string) {
    const option:IDropdownOption = categoryOptions.find(item => item.text === t);
    return option.key;
  }
  // validate the title text
  function onGetErrorMessageTitle(value: string): string {
    let returnMessage: string = '';
    if (value.length === 0) {
      // validate
      returnMessage = strings.EventTitleErrorMessage;
    } else {
      setSavedEvent({...savedEvent,title:value});
    }
    return returnMessage;
  }

  function onChangeNewTaskTitle(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
                                newValue?: string) {
    if (newValue) {
      setSavedEvent({...savedEvent,title:newValue});
    }
  }

  function onAllDayEvent (ev:any, checked: boolean) {
    const e:IBizpNonRecurrence = {...savedEvent,fAllDayEvent:checked};
    ev.preventDefault();
    setSavedEvent(e);
  }

  function onStartDateChange(newDate: Date) {
    // calculate new end date
    let eDate:Date = moment(newDate).add(savedEvent.duration,'minutes').toDate();
    const e: IBizpNonRecurrence = {...savedEvent,startDate:newDate,endDate:eDate};
    setSavedEvent(e);
  }

  function onEndDateChange(newDate: Date) {
    const d = moment(newDate).diff(moment(savedEvent.startDate),'minutes');
    const e: IBizpNonRecurrence = {...savedEvent,endDate:newDate,duration:d};
    setSavedEvent(e);
  }

  function onShowMoreLess(ev:any) {
    if (showMoreLess) {
      // panel will hide the series info panel
      // set the function to get the series information
      setSaveInfoFn(saveSeries);
      // send a request to get the recurrence data
      setRequestRecurrence(!requestRecurrence);
    }
    else {
      // set the function to pass on series info to higher component
      setSaveInfoFn(MoveUpSeriesData);
    }
    setShowMoreLess(!showMoreLess);
  }

  function onDescriptionChange (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void {
    const e: IBizpNonRecurrence = {...savedEvent,description:newText};
    setSavedEvent(e);
  }

  function onCategoryChange(ev: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void {
    const e: IBizpNonRecurrence = {...savedEvent,category:item.text};
    setSavedEvent(e);
  }

  function onRepeatChange(ev: any, checked: boolean): void {
    const eventType:string = checked?"1":"0";
    const e:IBizpNonRecurrence = {...savedEvent,fRecurrence:eventType,eventType:eventType};
    setSavedEvent(e);
    setEntryItems({...entryItems,seriesChecked:checked});
  }

  function onRenderFooterContent() {
    return (
      <div >
        <DefaultButton onClick={onCancelEntry} style={{ marginBottom: '15px', float: 'right' }}>
          {strings.DialogCloseButtonLabel}
        </DefaultButton>
        <PrimaryButton
            disabled={saveDisabled}
            onClick={onSave}
            style={{ marginBottom: '15px', marginRight: '8px', float: 'right' }}>
            {strings.SaveButtonLabel}
        </PrimaryButton>
        {
          saving &&
          <Spinner size={SpinnerSize.medium} style={{ marginBottom: '15px', marginRight: '8px', float: 'right' }} />
        }
     </div>
    );
  }
  console.log("Rendering EventEntry...");
  return (
    <div>
    <Panel
      isOpen={props.showPanel}
      onDismiss={onCancelEntry}
      type={PanelType.medium}
      headerText={entryItems.headerText}
      isFooterAtBottom={true}
      onRenderFooterContent={onRenderFooterContent}
    >
      <div style={{ width: '100%' }}>
        {
          hasError &&
          <MessageBar messageBarType={MessageBarType.error}>
                {errorMessage}
          </MessageBar>
        }
        {
          !loading &&
          <div>
            <TextField
              label={strings.EventTitleLabel}
              value={savedEvent.title}
              onGetErrorMessage={onGetErrorMessageTitle}
              deferredValidationTime={500}
              onChange={onChangeNewTaskTitle}
              disabled={false}
            />

            <div style={{ display: 'inline-block', verticalAlign: 'top', width: '200px' }}>
                  <Toggle
                    checked={savedEvent.fAllDayEvent}
                    inlineLabel={true}
                    label={strings.AllDayEventLabel}
                    onChange={onAllDayEvent}
                  />
            </div>

            <DateTimePicker label= {strings.StartDateLabel}
                dateConvention={savedEvent.fAllDayEvent?DateConvention.Date:DateConvention.DateTime}
                timeConvention={TimeConvention.Hours12}
                timeDisplayControlType = {TimeDisplayControlType.Text}
                value={savedEvent.startDate}
                onChange={onStartDateChange}
            />
            <DateTimePicker label={strings.EndDateLabel}
                dateConvention={savedEvent.fAllDayEvent?DateConvention.Date:DateConvention.DateTime}
                timeConvention={TimeConvention.Hours12}
                timeDisplayControlType = {TimeDisplayControlType.Text}
                value={savedEvent.endDate}
                onChange={onEndDateChange}
                minDate={new Date(savedEvent.startDate.toString())}
            />
            <Label htmlFor="endTimeId">End Time</Label>
            <div style={{ display: 'inline-block', verticalAlign: 'top', width: '200px' }}>
              <ActionButton iconProps={addFriendIcon} allowDisabledFocus disabled={false} checked={showMoreLess} onClick={onShowMoreLess} >
                {showMoreLess?strings.ShowLessLabel:strings.ShowMoreLabel}
              </ActionButton>
            </div>
            {showMoreLess &&
              <div>
                <TextField
                  label= {strings.EventDescriptionLabel} multiline autoAdjustHeight
                  onChange={onDescriptionChange}
                  value={savedEvent.description}>
                </TextField>

                <Dropdown
                  label={strings.CategoryLabel}
                  selectedKey={savedEvent.category.length > 0 ? categoryTextToKey(savedEvent.category):""}
                  onChange={onCategoryChange}
                  options={categoryOptions}
                  placeholder={strings.CategoryPlaceHolder}
                  disabled={false}
                />
                {entryItems.entryType == IBizpEntryTypeEnum.editEvent &&
                  <div style={{ display: 'inline-block', verticalAlign: 'top', width: '200px' }}>
                    <Toggle
                      defaultChecked={false}
                      inlineLabel={true}
                      label={strings.Repeatlabel}
                      onText={strings.OnLabel}
                      offText={strings.OffLabel}
                      onChange={onRepeatChange}
                    />
                  </div>
                }
                { entryItems.entryType == IBizpEntryTypeEnum.editEventFromSeries &&
                  <div style={{ display: 'inline-block', verticalAlign: 'top', width: '200px' }}>
                    <Toggle
                      defaultChecked={true}
                      disabled
                      inlineLabel={true}
                      label={strings.Repeatlabel}
                      onText={strings.OnLabel}
                      offText={strings.OffLabel}
                      onChange={onRepeatChange}
                    />
                  </div>
                }
                { entryItems.entryType == IBizpEntryTypeEnum.editSeries &&
                  <div style={{ display: 'inline-block', verticalAlign: 'top', width: '200px' }}>
                    <Toggle
                      defaultChecked={true}
                      inlineLabel={true}
                      label={strings.Repeatlabel}
                      onText={strings.OnLabel}
                      offText={strings.OffLabel}
                      onChange={onRepeatChange}
                    />
                  </div>
                }
                { entryItems.entryType == IBizpEntryTypeEnum.newEvent &&
                  <div style={{ display: 'inline-block', verticalAlign: 'top', width: '200px' }}>
                    <Toggle
                      defaultChecked={false}
                      inlineLabel={true}
                      label={strings.Repeatlabel}
                      onText={strings.OnLabel}
                      offText={strings.OffLabel}
                      onChange={onRepeatChange}
                    />
                  </div>
                }

                { entryItems.seriesChecked  && (entryItems.entryType != IBizpEntryTypeEnum.editEventFromSeries) &&
                  <BizpRecurrentEvent event = {savedEvent} eventSeries = {series} entryType= {entryItems.entryType} startDateChange={onStartDateChange} infoRequest={requestRecurrence} returnInfo={saveInfoFn}/>
                }
              </div>
            }
          </div>
        }
      </div>
    </Panel>
    </div>
   );
}
