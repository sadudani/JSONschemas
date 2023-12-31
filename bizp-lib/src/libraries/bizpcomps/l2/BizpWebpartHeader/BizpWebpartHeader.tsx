import * as React from 'react';
import styles from './BizpWebpartHeader.module.scss';
import * as strings from 'BizpcompsLibraryStrings';
import { DefaultPalette, Stack, IStackStyles} from 'office-ui-fabric-react';
import { IBizpWebpartHeaderProps } from './IBizpWebpartHeaderProps';
import BizpWebpartTitle from '../../l1/BizpWebpartTitle/BizpWebpartTitle';
import BizpWebpartMenu from '../../l1/BizpWebpartMenu/BizpWebpartMenu';

export function BizpWebpartHeader(props: IBizpWebpartHeaderProps) {
/*   const stackStyles: IStackStyles = {
    root: {
      background: DefaultPalette.themeTertiary,
      width: 400,
    },
  }; */
  const handleFocus = event => {
    event.preventDefault();
    const { target } = event;
    const extensionStarts = target.value.lastIndexOf('.');
    target.focus();
    target.setSelectionRange(0, extensionStarts);
  };
  console.log("Rendering WebpartHeader...");
  return (
    <div className={styles.bizpWebpartHeader}>
      <div className={styles.container}>
        <div className={styles.row}>
          <div className={styles.leftbox}>
          <BizpWebpartTitle title={props.title} showTitle={props.showTitle}/>
          </div>
          <div className={styles.rightbox}>
          <BizpWebpartMenu  menuOptions={props.menuOptions} helpId={props.helpId} onRefresh={props.onRefresh}
                              themeVariant={props.themeVariant} context={props.context}
          />
          </div>
        </div>
      </div>
    </div>

/*         <div >
        <Stack horizontal styles={stackStyles} >
          <Stack.Item align="start">
              <BizpWebpartTitle title={props.title} showTitle={props.showTitle}/>
          </Stack.Item>
          <Stack.Item align="end">
              <BizpWebpartMenu  menuOptions={props.menuOptions} helpId={props.helpId} onRefresh={props.onRefresh}
                                themeVariant={props.themeVariant} context={props.context}
              />
          </Stack.Item>
        </Stack>
      </div> */
  );
}
