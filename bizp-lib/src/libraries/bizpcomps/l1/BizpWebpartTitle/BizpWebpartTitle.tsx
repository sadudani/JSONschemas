import * as React from 'react';
import styles from './BizpWebpartTitle.module.scss';
import { IBizpWebpartTitleProps } from './IBizpWebpartTitleProps';


export default function BizpWebpartTitle(props: IBizpWebpartTitleProps) {
  let headerText:any = "";
  if (props.showTitle) headerText =
  <div>
    <p className={styles.headerTitle}>{props.title}</p>
  </div>;
  console.log("Rendering WebpartTitle...");
  return headerText;
}
