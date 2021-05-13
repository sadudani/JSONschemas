import * as React from 'react';
import styles from './BizpWebpartTitle.module.scss';
import { IBizpWebpartTitleProps } from './IBizpWebpartTitleProps';


export default function BizpWebpartTitle(props: IBizpWebpartTitleProps) {
  let headerText:any = "";
  if (props.showTitle) headerText = <div> {props.title}</div>;
  console.log("Rendering WebpartTitle...");
  return headerText;
}
