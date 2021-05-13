import * as React from 'react';
import styles from './BizpWebpartMenu.module.scss';
import * as strings from 'BizpcompsLibraryStrings';
import * as Fabric from 'office-ui-fabric-react';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { Stack } from 'office-ui-fabric-react';

import { IBizpWebpartMenuProps } from './IBizpWebpartMenuProps';
import { BizpWebpartFeedback } from '../BizpWebpartFeedback/BizpWebpartFeedback';
import { IBizpMenuOptions } from '../../../../shared/IBizpSharedInterface';
import { IconButton, IIconProps, IContextualMenuItem, IContextualMenuProps} from 'office-ui-fabric-react';
import { Panel, PanelType} from 'office-ui-fabric-react';

export default function BizpWebpartMenu(props: IBizpWebpartMenuProps) {
  const { helpId, menuOptions } = props;

  const [showHelpPanel, setHelpPanel] = React.useState(false);
  const [feedbackSignal, setFeedbackSignal] = React.useState(false);
  const iconStyles = { marginRight: '8px' };

  const menuIcon: IIconProps = { iconName: 'ContextMenu' };

  const _onClickMenuItem = (ev?: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>,
                            item?: IContextualMenuItem):void => {
    switch (item.key) {
      case "help":{
        setHelpPanel(true);
        break;
      }
      case "feedback": {
        setFeedbackSignal(!feedbackSignal);
        break;
      }
      case "refresh": {
        props.onRefresh();
        break;
      }
      default: {
        break;
      }
    }
  };

  const menuProps: IContextualMenuProps = {
    items: menuOptions.map((val:IBizpMenuOptions, index):IContextualMenuItem =>{
      return {key: val.key, text:val.text, iconProps:{iconName: val.iconName},
      onClick: _onClickMenuItem};
    }),
    directionalHintFixed: true,
  };

  function closeHelpPanel() {
    setHelpPanel(false);
  }

  console.log("Rendering WebpartMenu...");
  return (
    <div>
      <div>
      <Stack tokens={{ childrenGap: 8 }} horizontal>
        <IconButton
          menuProps={menuProps}
          iconProps={menuIcon}
          title={strings.MenuLabel}
          ariaLabel={strings.MenuLabel}
        />
      </Stack>
      </div>
        <BizpWebpartFeedback feedbackId={"SiteHierarchyWebpart"} showCategory={true}
            themeVariant={props.themeVariant} context={props.context} openSignal={feedbackSignal}
        />
      {showHelpPanel &&
            <div>
              <Panel
                isOpen={ true }
                onDismiss= { closeHelpPanel }
                type={ PanelType.medium }
                headerText= {strings.HelpLabel}
              >
              <span className='ms-font-m'>Help Id: {props.helpId} Content goes here.</span>
              </Panel>
            </div>
      }

    </div>
  );
  /*
      {showFeedbackPanel &&
            <div>
              <Panel
                isOpen={ true }
                onDismiss= { closeFeedbackPanel }
                type={ PanelType.medium }
                headerText={strings.FeedbackLabel}
              >
              <span className='ms-font-m'>Content goes here.</span>
              </Panel>
            </div>
      }
      */
}

