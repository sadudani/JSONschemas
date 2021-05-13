import * as React from 'react';
import {useState,useEffect} from 'react';


import { DefaultPalette, Stack, IStackStyles} from 'office-ui-fabric-react';
import { graph } from "@pnp/graph";
import "@pnp/graph/groups";
import "@pnp/graph/calendars";
import '@pnp/graph/users';

import { Group,User } from '@microsoft/microsoft-graph-types';

import styles from './Reminders.module.scss';
import * as strings from 'RemindersWebPartStrings';
import { IRemindersProps } from './IRemindersProps';
import {​​BizpWebpartHeader,BizpCalendarEventsDisplay }​​ from "bizp-lib";
import {​​IBizpMenuOptions }​​ from "bizp-lib/lib/shared/IBizpSharedInterface";
import { JSONParser } from '@pnp/odata';

export default function BizpReminders(props: IRemindersProps) {
  const IdForHelp = "ReminderesHelp";
  const menuOpts:IBizpMenuOptions[] = [
    { key: 'help', text: strings.HelpMenuLabel,iconName:'StatusCircleQuestionMark'},
    { key: 'feedback', text: strings.FeedbackMenuLabel , iconName:'EmojiNeutral' },
    {key:'refresh',text:strings.RefreshMenuLabel, iconName: 'Refresh' }];
  const stackStyles: IStackStyles = {
    root: {
      background: DefaultPalette.themeTertiary,
      width: 400,
    },
  };
  //refresh is used as a toggle. Anytime its value is toggled, it will force execute component
  const [refresh,setRefresh] = useState<boolean>(false); // refresh toggle
  const [groups,setGroups] = useState<Group[]>(null);
  const [user,setUser] = useState<User[]>(null);
  useEffect(() => {
    graph.groups.get<Group[]>().then(g => {
      setGroups(g);
    });
    graph.me().then(u =>{
      setUser(u[0]);
    });
    },[]
  );

  function onRefresh():void {
    const r = refresh;
    setRefresh(!r);
  }

  console.log("Groups: " + JSON.stringify(groups));
  console.log("User: " + JSON.stringify(user));
  console.log("Rendering Reminders....");
  return (
    <div className={ styles.reminders } >
      <Stack styles={stackStyles} >
        <Stack styles={stackStyles}>
          <BizpWebpartHeader title="My Reminders"
            context = {props.context}
            showTitle={true}
            menuOptions = {menuOpts}
            helpId = {IdForHelp}
            onRefresh = {onRefresh}
          >
          </BizpWebpartHeader>
        </Stack>
        <Stack styles={stackStyles}>
          <BizpCalendarEventsDisplay siteUrl={props.siteUrl} list={props.list} context={props.context}
                                     daysInFuture = {props.daysInFuture} daysInPast={props.daysInPast}
                                     refresh = {refresh}
          ></BizpCalendarEventsDisplay>
        </Stack>
      </Stack>
    </div>
  );
}








/*
export default class Reminders extends React.Component<IRemindersProps, IRemindersState> {
  constructor(props: IRemindersProps) {
    super(props);
    this.setState({
      refresh:false,
    });
  }
  private IdForHelp = "ReminderesHelp";
  private menuOpts:IBizpMenuOptions[] = [
      { key: 'help', text: strings.HelpMenuLabel,iconName:'StatusCircleQuestionMark'},
      { key: 'feedback', text: strings.FeedbackMenuLabel , iconName:'EmojiNeutral' },
      {key:'refresh',text:strings.RefreshMenuLabel, iconName: 'Refresh' }];

  private stackStyles: IStackStyles = {
    root: {
      background: DefaultPalette.themeTertiary,
      width: 400,
    },
  };
  protected onRefresh():void {
    const r = this.state.refresh;
    this.setState({refresh:!r});
  }

  public render(): React.ReactElement<IRemindersProps> {
    console.log("Rendering Reminders....");
    return (
      <div className={ styles.reminders } >
        <Stack styles={this.stackStyles} >
          <Stack styles={this.stackStyles}>
            <BizpWebpartHeader title="My Reminders"
              showTitle={true}
              menuOptions = {this.menuOpts}
              helpId = {this.IdForHelp}
              onRefresh = {this.onRefresh}
            >
            </BizpWebpartHeader>
          </Stack>
          <Stack styles={this.stackStyles}>
            <BizpCalendarEventsDisplay siteUrl={this.props.siteUrl} list={this.props.list} context={this.props.context}
                                        refresh = {this.state.refresh}
            ></BizpCalendarEventsDisplay>
          </Stack>
        </Stack>
      </div>
    );
  }
}
*/
