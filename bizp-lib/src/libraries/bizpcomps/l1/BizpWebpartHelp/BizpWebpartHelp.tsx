import * as React from 'react';
import styles from './BizpWebpartHelp.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import * as Fabric from 'office-ui-fabric-react';
import * as strings from 'BizpcompsLibraryStrings';
import { IBizpWebpartHelpProps } from './IBizpWebpartHelpProps';

export default function BizpWebpartHelp(props: IBizpWebpartHelpProps) {
  console.log("Rendering WebpartHelp...");
  return (
    <div >
       This help is for webpart with helpId={props.helpId}
   </div>
  );
}
