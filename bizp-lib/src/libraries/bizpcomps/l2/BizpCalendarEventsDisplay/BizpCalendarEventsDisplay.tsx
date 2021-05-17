import * as React from 'react';
import {useState,useEffect} from 'react';

import * as strings from 'BizpcompsLibraryStrings';
import styles from './BizpCalendarEventsDisplay.module.scss';
import {IBizpCalendarEventsDisplayProps} from './IBizpCalendarEventsDisplayProps';
import {​​BizpEventListDisplay }​​ from "../BizpEventListDisplay/BizpEventListDisplay";
import {​​BizpEventEntry }​​ from "../../l1/BizpEventEntry/BizpEventEntry";
import {
  DefaultButton,PrimaryButton,Spinner,SpinnerSize,Stack,IconButton,IIconProps,Selection,SelectionMode,
  IContextualMenuProps,IContextualMenuItem,
  Dialog,DialogType,DialogFooter,
  Label,
  initializeComponentRef
} from 'office-ui-fabric-react';

import {
  IBizpEventDataSpec,
  IBizpMenuOptions,
  IBizpRecurrence, IBizpNonRecurrence,
  IBizpCalendarRequest,
  IBizpUserPermissions,
  IBizpCRUDEnum } from '../../../../shared/IBizpSharedInterface';
import { getSPCalendarEvents,addNewSPCalendarEvent, updateSPCalendarEvent, deleteSPCalendarEvent,initCategoryColors, sortByDate} from '../../../../shared/BizpCalendarService';
import {parseEventToDataSP} from '../../../../shared/BizpCalendarRecurrentEvents';
import { SPService,getUserPermissions,cloneObj } from '../../../../shared/BizpBasesvc';
import * as moment from 'moment';

interface IBizpCalendarEventItems {
  updating: boolean; // used to show updating spinner
  addEventEnabled: boolean; // state of allowing add event
  editEventEnabled: boolean; // state of allowing edit event
  deleteEventEnabled:boolean; // state of allowing delete event
  updateDialog:boolean; // show dialog to confirm update
  seriesOptionEnabled: boolean; // stores if recurrent option is enabled
  seriesSelected: boolean; // stores if recurrent option is selected
//  selection: Selection; // current selection of items
}

export function BizpCalendarEventsDisplay(props: IBizpCalendarEventsDisplayProps) {
  const [data,setData] = useState(undefined);
  const [loadingError,setLoadingError] = useState({hasError:false,message:"Error message is empty"});
  const [loading,setLoading] = useState(false); // used to show loading spinner
  const [updateEvent,setUpdateEvent] = useState <IBizpEventDataSpec> (undefined);
  const [showEntryPanel,setShowEntryPanel] = useState<boolean>(false);
  const [mode,setMode] = useState<IBizpCRUDEnum>(IBizpCRUDEnum.view); // stores CRUD mode
  const [eventSeries,setEventSeries] = useState<IBizpRecurrence> (undefined);
  // labels and texts based on edit ot delete
  const [modeContent,setModeContent] = useState({
    dialogTitle:"",
    labelSeriesText:"",
    labelSeriesEventText:"",
    labelSingleEventText:"",
    spinnerLabel:"",
    confirmButtonText:""});
  const [calendarEventItems, setCalendarEventItems] = useState<IBizpCalendarEventItems> ({
    updating:false,
    addEventEnabled:false,
    editEventEnabled: false,
    deleteEventEnabled: false,
    updateDialog: false,
    seriesOptionEnabled: false,
    seriesSelected: false,
  });

  const [permissions,setPermissions] = useState <IBizpUserPermissions> ({
    hasPermissionDelete:false,
    hasPermissionEdit:false,
    hasPermissionAdd:false,
    hasPermissionView:false
  });

  const [selection,setSelection] = useState <Selection> (
    new Selection({
      onSelectionChanged: () =>{
          onNewSelection(selection);
      },
      selectionMode: SelectionMode.multiple
    })
  );

  const [selectionCount,setSelectionCount] = useState(0);

  const [refresh,setRefresh] = useState<boolean> (false);

  useEffect(() => {
    console.log ("In useEffect: CalendarsEventsDisplay");
    if (props.siteUrl && props.list) {
      initCategoryColors(props.siteUrl,props.list);
      init();
    }
  },[props.siteUrl,props.list,props.refresh,props.daysInFuture,props.daysInPast]
  );
  useEffect(() => {
    console.log ("In useEffect: CalendarsEventsDisplay - Initital run");
    if (props.siteUrl && props.list) {
      initCategoryColors(props.siteUrl,props.list);
      init();
    }
  },[]
  );
/*
  useEffect(() => {
    console.log ("In useEffect: CalendarsEventsDisplay");
    init();
  },[props.refresh]
  );
*/
  useEffect(() => {
    processNewSelection (selection);
  },[selectionCount]
  );

  const sample:IBizpEventDataSpec[] = [];

  const addEventIcon: IIconProps = { iconName: 'CircleAdditionSolid' };
  const editEventIcon: IIconProps = { iconName: 'EditSolid12' };
  const deleteEventIcon: IIconProps = { iconName: 'Delete' };

  async function init(){
    // init pnp context
    SPService(props.context);
    const p:IBizpUserPermissions = await getUserPermissions(props.siteUrl,props.list);
    setPermissions({...permissions,hasPermissionDelete:p.hasPermissionDelete,hasPermissionEdit:p.hasPermissionEdit,
      hasPermissionAdd:p.hasPermissionAdd,hasPermissionView:p.hasPermissionView});
    console.log("User permissions: " +JSON.stringify(p));
    // make sure no other call to update the calenderEventItems state is invoked in init
    setCalendarEventItems({...calendarEventItems,addEventEnabled:p.hasPermissionAdd});
//    setCalendarEventItems({...calendarEventItems,addEventEnabled:p.hasPermissionAdd,editEventEnabled:p.hasPermissionEdit,
//      deleteEventEnabled:p.hasPermissionDelete});
    if (props.list && p.hasPermissionView) {
      if (!loading) loadData();
    }
  }

  async function loadData(){
    let actual: IBizpEventDataSpec[];
    try {
      if (props.siteUrl && props.list) {
        setLoading(true);
        const request:IBizpCalendarRequest = {
          listName : props.list,
          siteURL : props.siteUrl,
          includeRecurringEvents : true,
          eventStartDate : moment().subtract(props.daysInPast,'d').toDate().toISOString(),
          eventEndDate : moment().add(props.daysInFuture,'d').toDate().toISOString()
        };
        actual = await getSPCalendarEvents (request);
        actual = sortByDate(actual,false);
        setData(actual);
      }
      else {
        setData(sample);
      }
      setLoading(false);
      setRefresh(!refresh);
      // no selection after loading
//      setSelection(undefined);
    }
    catch {
      setLoadingError({hasError:true,message:"got Error"});
      setLoading(false);
    }
  }

  function onAddEvent() {
    setMode(IBizpCRUDEnum.add);
    setShowEntryPanel(true);
    setCalendarEventItems({...calendarEventItems});
  }

  async function onSaveEntry(eventData: IBizpNonRecurrence,recurrentData:IBizpRecurrence){
    if (mode==IBizpCRUDEnum.add) {
      await addNewSPCalendarEvent(props.siteUrl,props.list,eventData,recurrentData);
    }
    else if (mode==IBizpCRUDEnum.edit) {
      await updateSPCalendarEvent(props.siteUrl,props.list,calendarEventItems.seriesSelected,updateEvent,eventData,recurrentData);
    }
    await loadData();
    setShowEntryPanel(false);
  }

  async function onDismissEntry(doRefresh: boolean){
    if (doRefresh === true) {
      await loadData();
    }
    setShowEntryPanel(false);
  }

  async function onConfirmUpdate(ev: React.MouseEvent<HTMLDivElement>){
    ev.preventDefault();
    try {
      setCalendarEventItems({...calendarEventItems,updating:true});
      // get selected items
      // delete one item at a time
      let itemList: any = selection.getSelection();
      // there will be only one selection
      for (const item of itemList) {
        setUpdateEvent({...item});
        if (mode==IBizpCRUDEnum.delete) {
          await deleteSPCalendarEvent(props.siteUrl, props.list, item, calendarEventItems.seriesSelected);
          // reload
          await loadData();
        } else if (mode==IBizpCRUDEnum.edit) {
          if (calendarEventItems.seriesSelected) {
            const e:IBizpRecurrence = await parseEventToDataSP(item);
            setEventSeries(cloneObj(e));
            setShowEntryPanel(true);
            setCalendarEventItems({...calendarEventItems,updating:false,updateDialog:false});
          }
          else {
            setShowEntryPanel(true);
            setCalendarEventItems({...calendarEventItems,updating:false,updateDialog:false});
          }
        }
      }
      setCalendarEventItems({...calendarEventItems,updating:false,updateDialog:false});
    } catch (error) {
      setCalendarEventItems({...calendarEventItems,updating:false});
    }
  }

  function onCloseUpdateDialog(ev: React.MouseEvent<HTMLDivElement>) {
    ev.preventDefault();
    setMode(IBizpCRUDEnum.view);
    setCalendarEventItems({...calendarEventItems,updateDialog:false});
  }

  function setupMode (m: IBizpCRUDEnum) {
    setMode(m);
    setCalendarEventItems({...calendarEventItems});
    switch (m) {
      case IBizpCRUDEnum.delete:
        setModeContent (
          {
            dialogTitle:strings.DeleteButtonLabel,
            labelSeriesText:strings.ConfirmDeleteSeriesMsg,
            labelSingleEventText:strings.ConfirmDeleteEventMsg,
            labelSeriesEventText:strings.ConfirmDeleteSeriesEventMsg,
            spinnerLabel:strings.SpinnerDeletingLabel,
            confirmButtonText:strings.DeleteButtonLabel
          }
        );
        break;
      case IBizpCRUDEnum.edit:
        setModeContent (
          {
            dialogTitle:strings.EditButtonLabel,
            labelSeriesText:strings.ConfirmUpdateSeriesMsg,
            labelSingleEventText:strings.ConfirmUpdateEventMsg,
            labelSeriesEventText:strings.ConfirmUpdateSeriesEventMsg,
            spinnerLabel:strings.SpinnerUpdatingLabel,
            confirmButtonText:strings.EditButtonLabel
          }
        );
        break;
      case IBizpCRUDEnum.add:
        break;
      case IBizpCRUDEnum.view:
        break;
      default:
        break;
    }
  }

  function onNewSelection (s:Selection) {
    setSelection(s);
    setSelectionCount(s.count);
  }

  function processNewSelection (s:Selection) {
    switch (s.count) {
      case 0: {
        setCalendarEventItems({...calendarEventItems,seriesOptionEnabled:false,deleteEventEnabled:false,editEventEnabled:false});
      }
      break;
      case 1: {
        let items:any = s.getSelection();
        if (items[0].eventType == '1') {
//          setCalendarEventItems({...calendarEventItems,selection:s,seriesOptionEnabled:true,deleteEventEnabled:true,editEventEnabled:true
 //                               });
          setCalendarEventItems({...calendarEventItems,seriesOptionEnabled:true,deleteEventEnabled:permissions.hasPermissionDelete,editEventEnabled:permissions.hasPermissionEdit
                                });
        }
        else {
//          setCalendarEventItems({...calendarEventItems,selection:s,seriesOptionEnabled:false,deleteEventEnabled:true,editEventEnabled:true});
          setCalendarEventItems({...calendarEventItems,seriesOptionEnabled:false,deleteEventEnabled:permissions.hasPermissionDelete,editEventEnabled:permissions.hasPermissionEdit});
         }
      }
      break;
      default: {
        setCalendarEventItems({...calendarEventItems,deleteEventEnabled:false,editEventEnabled:false});
      }
      break;
    }
  }

  const editMenuOptions:IBizpMenuOptions[] = [
    { key: 'recurrent', text: strings.EditSeriesLabel,iconName:'StatusCircleQuestionMark'},
    { key: 'non-recurrent', text: strings.EditOccurrenceLabel , iconName:'EmojiNeutral' }
  ];

  const deleteMenuOptions:IBizpMenuOptions[] = [
    { key: 'recurrent', text: strings.DeleteSeriesLabel,iconName:'StatusCircleQuestionMark'},
    { key: 'non-recurrent', text: strings.DeleteOccurrenceLabel , iconName:'EmojiNeutral' }
  ];

  const menuIcon: IIconProps = { iconName: 'ContextMenu' };

  function onSingleEvent(m:IBizpCRUDEnum) {
    setupMode(m);
    setCalendarEventItems({...calendarEventItems,seriesSelected:false,updateDialog:true});
  }

  const onEditSeriesEvent = (ev?: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>,
    item?: IContextualMenuItem):void => {
    setupMode(IBizpCRUDEnum.edit);
    switch (item.key) {
      case "recurrent":{
        setCalendarEventItems({...calendarEventItems,seriesSelected:true,updateDialog:true});
        break;
      }
      case "non-recurrent": {
        setCalendarEventItems({...calendarEventItems,seriesSelected:false,updateDialog:true});
        break;
      }
      default: {
        setCalendarEventItems({...calendarEventItems,updateDialog:true});
        break;
      }
    }
  };

  const onDeleteSeriesEvent = (ev?: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>,
                          item?: IContextualMenuItem):void => {
    setupMode(IBizpCRUDEnum.delete);
    switch (item.key) {
      case "recurrent":{
        setCalendarEventItems({...calendarEventItems,seriesSelected:true,updateDialog:true});
        break;
      }
      case "non-recurrent": {
        setCalendarEventItems({...calendarEventItems,seriesSelected:false,updateDialog:true});
        break;
      }
      default: {
        setCalendarEventItems({...calendarEventItems,updateDialog:true});
        break;
      }
    }
  };

  const editMenuProps: IContextualMenuProps = {
    items: editMenuOptions.map((val:IBizpMenuOptions, index):IContextualMenuItem =>{
      return {key: val.key, text:val.text, iconProps:{iconName: val.iconName},
      onClick: onEditSeriesEvent};
    }),
    directionalHintFixed: true,
  };

  const deleteMenuProps: IContextualMenuProps = {
    items: deleteMenuOptions.map((val:IBizpMenuOptions, index):IContextualMenuItem =>{
      return {key: val.key, text:val.text, iconProps:{iconName: val.iconName},
      onClick: onDeleteSeriesEvent};
    }),
    directionalHintFixed: true,
  };

  console.log("Rendering CalendarEventDisplay... ");
  return (
    <div>
      <p>
        Hover over the <i>title</i> of the reminder to see the details.
      </p>

      <Stack tokens={{ childrenGap: 8 }} horizontal>
            <IconButton iconProps={addEventIcon} title= {strings.AddNewEventLabel} ariaLabel={strings.AddNewEventLabel} disabled={!calendarEventItems.addEventEnabled} onClick={onAddEvent} allowDisabledFocus/>

            { calendarEventItems.editEventEnabled && !calendarEventItems.seriesOptionEnabled &&
              <IconButton iconProps={editEventIcon} title= {strings.EditEventlabel} ariaLabel= {strings.EditEventlabel}
                disabled={!calendarEventItems.deleteEventEnabled} onClick={() => {
                onSingleEvent(IBizpCRUDEnum.edit);}}
                allowDisabledFocus/>
            }
            { calendarEventItems.editEventEnabled && calendarEventItems.seriesOptionEnabled &&
              <IconButton
                title={strings.EditMenuLabel}
                ariaLabel= {strings.EditMenuLabel}
                menuProps={editMenuProps}
                iconProps={editEventIcon}
                disabled={!calendarEventItems.seriesOptionEnabled}
              />
            }
           { calendarEventItems.deleteEventEnabled && !calendarEventItems.seriesOptionEnabled &&
              <IconButton iconProps={deleteEventIcon} title= {strings.DeleteEventLabel} ariaLabel={strings.DeleteEventLabel}
                          disabled={!calendarEventItems.deleteEventEnabled} onClick={() => {
                          onSingleEvent(IBizpCRUDEnum.delete);}}
                          allowDisabledFocus/>
            }
            { calendarEventItems.deleteEventEnabled && calendarEventItems.seriesOptionEnabled &&
              <IconButton
                title= {strings.DeleteMenuLabel}
                ariaLabel= {strings.DeleteMenuLabel}
                menuProps={deleteMenuProps}
                iconProps={deleteEventIcon}
                disabled={!calendarEventItems.seriesOptionEnabled}
              />
            }

      </Stack>
      <BizpEventListDisplay displayData={data} selection = {selection} refresh = {refresh}></BizpEventListDisplay>
      { showEntryPanel &&
        <BizpEventEntry panelMode = {mode} onDismissEntry = {onDismissEntry} event = {updateEvent} series={calendarEventItems.seriesSelected}
        eventSeries = {eventSeries} onSaveNewEntry={onSaveEntry}  showPanel = {showEntryPanel} siteUrl = {props.siteUrl} listId = {props.list}/>
      }
      {
        calendarEventItems.updateDialog &&
        <div>
          <Dialog
            hidden={!calendarEventItems.updateDialog}
            dialogContentProps={{
            type: DialogType.normal,
            title: modeContent.dialogTitle,
            showCloseButton: false
            }}
            modalProps={{
            isBlocking: true,
            styles: { main: { maxWidth: 450 } }
            }}
          >
          <Label >{calendarEventItems.seriesOptionEnabled ? (calendarEventItems.seriesSelected?modeContent.labelSeriesText:modeContent.labelSeriesEventText) : modeContent.labelSingleEventText}</Label>
            {
             calendarEventItems.updating &&
              <Spinner size={SpinnerSize.medium} ariaLabel={modeContent.spinnerLabel} />
            }
             <DialogFooter>
                <PrimaryButton onClick={onConfirmUpdate} text={modeContent.confirmButtonText} disabled={calendarEventItems.updating} />
                <DefaultButton onClick={onCloseUpdateDialog} text="Cancel" />
              </DialogFooter>
          </Dialog>
        </div>
      }

    </div>

  );
}

