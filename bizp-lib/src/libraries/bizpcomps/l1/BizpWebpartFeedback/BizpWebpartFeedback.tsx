import * as React from 'react';
import {useState,useEffect,useCallback} from 'react';
import styles from './BizpWebpartFeedback.module.scss';
import * as strings from 'BizpcompsLibraryStrings';
import { escape } from '@microsoft/sp-lodash-subset';

import * as Fabric from 'office-ui-fabric-react';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { Dialog, DialogFooter, DialogType } from 'office-ui-fabric-react/lib/Dialog';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';

import { sp } from '@pnp/sp/presets/all';
import { IPrincipalInfo} from "@pnp/sp";
import "@pnp/sp/sputilities";
import { IEmailProperties } from "@pnp/sp/sputilities";

import {LogHelper } from '../../../../shared/BizpLogger';
import { IBizpWebpartFeedbackProps } from './IBizpWebpartFeedbackProps';

const buttonStyles = { root: { marginRight: 8 } };

const dialogModalProps = {
  isBlocking: false,
  styles: { main: { maxWidth: 450 } },
};
const options: IDropdownOption[] = [
  {
    key: "general",
    text: "General"
  },
  {
    key: "typo",
    text: "Typo/Edit/Broken Link"
  },
  {
    key: "suggestion",
    text: "Suggestion"
  }
];

export function BizpWebpartFeedback(props: IBizpWebpartFeedbackProps) {
  const feedbackWebpartSite:string =  props.context.pageContext.web.title;
  const webpartSiteUrl:string = props.context.pageContext.web.absoluteUrl;

  const defaultCategory = {key:"general",text:"General"};
  const [initialized,setInitialized] = useState(false);
  const [isOpen, setIsOpen] = useState(false);
  const [isDialogVisible, setIsDialogVisible] = useState(false);
  const [txtValue, setTxtValue] = useState("");
  const [selectedCategory, setSelectedCategory] = useState<IDropdownOption>(defaultCategory);
  const[dialogContentProps,setDialogContentProps] = useState({
    type: DialogType.normal,
    title: strings.FeedbackSuccessMsg
  });

  useEffect(() => {
    // init selection object
    setInitialized(true);
    },[]
  );

  useEffect(() => {
    // init selection object
    if (initialized) setIsOpen(true);
    },[props.openSignal]
  );

  async function sendEmailToFeedbackSender(webpartTitle: string,  feedback:string, listitemid:string,
                  currentUserEmail:string, isSuccess:boolean): Promise<any> {

    if (feedback.indexOf("\n") > -1) {
      feedback = feedback.replace(/\n/g, '<br/>');
    }

    let body:string;
    if (isSuccess){
      body = "The feedback you provided on \"" + webpartTitle + "\" has been successfully submitted.</br></br> <i>\"" + feedback + "\"</i><br/><br>Webpart can be found here: <a href=\"" + webpartSiteUrl + "\">" + feedbackWebpartSite + "</a></br>";
    }
    else {
      body = "The feedback you provided on \"" + webpartTitle + "\" could not be submitted. Contact your portal administrator.</br></br> <i>\"" + feedback + "\"</i><br/><br>Webpart can be found here: <a href=\"" + webpartSiteUrl + "\">" + feedbackWebpartSite + "</a></br>";

    }

    if (currentUserEmail) {
      console.log("Email sending to User: " + currentUserEmail);
      const emailProps: IEmailProperties = {
        To: [currentUserEmail],
        Subject: "Feedback for " + webpartTitle,
        Body: body
      };
      await sp.utility.sendEmail(emailProps)
        .catch(e => {
          LogHelper.error("BizpWebpartFeedback",'sendEmail', e);
          throw e;
        });
    }
    return;
  }

  async function sendEmailToOwnerGroup(feedback:string, listitemid:string, category:string, currentUserName:string, currentUserEmail:string) {
    let webpartInfo = props.feedbackId;
    sp.setup({ spfxContext: props.context});
    let principals: IPrincipalInfo[] = await sp.utility.expandGroupsToPrincipals(["Feedback Owners"]);

    if (feedback.indexOf("\n") > -1) {
      feedback = feedback.replace(/\n/g, '<br/>');
    }

    var emails: string[] = [];
    for (var i = 0; i < principals.length; i++) {
      if (principals[i].Email) {
        emails.push(principals[i].Email);
      }
      else {
        LogHelper.warning("BizpWebpartFeedback", "sendEmailToOwnerGroup", `No email for ${principals[i].LoginName}`);
      }
    }

    console.log("Owner Emails: " + emails.join(";"));
    let isSuccess:boolean;

    if (emails && emails.length > 0) {
      const emailProps: IEmailProperties = {
        To: emails,
        Subject: "Feedback for " + webpartInfo + " (" + category + ")",
        Body: "<i>\"" + feedback + "\" </i><br/><br/>Category: " + category + " <br/><br/>Submitted by: " + currentUserName + " <a href=\"mailto:" +
        currentUserEmail + "\">" + currentUserEmail + "</a><br/><br/>Webpart can be found here: <a href=\"" + webpartSiteUrl + "\">" + feedbackWebpartSite + "</a></br>"
      };

      await sp.utility.sendEmail(emailProps)
        .catch(e => {
          LogHelper.error("BizpWebpartFeedback",'sendEmailToOwnerGroup', e);
          throw e;
        });
      LogHelper.info("BizpWebpartFeedback", "sendEmailToOwnerGroup", `Email Sent`);
      isSuccess = true;
      setDialogContentProps({
        type: DialogType.normal,
        title: strings.FeedbackSuccessMsg
      });
    }
    else {
      isSuccess = false;
      setDialogContentProps({
        type: DialogType.normal,
        title: strings.FeedbackAdminNotFoundMsg
      });
    }

    sendEmailToFeedbackSender(webpartInfo, feedback, listitemid, currentUserEmail,isSuccess );
  }

  const dismissPanel = React.useCallback(() => setIsOpen(false), [isOpen]);
  const hideDialog = useCallback(() => setIsDialogVisible(false), [isDialogVisible]);
  const hideDialogAndPanel = () => {
    setIsOpen(false);
    setIsDialogVisible(false);
  };

  const handleSubmit = (event) => {
    event.preventDefault();
    if (props.feedbackId == null){
      console.log("Feedback Id is null.");
    }
    sendEmailToOwnerGroup(txtValue, props.feedbackId, (props.showCategory) ? selectedCategory.text : "",
                          props.context.pageContext.user.displayName, props.context.pageContext.user.email);
    dismissPanel();
    setIsDialogVisible(true);
    setSelectedCategory(defaultCategory);
    setTxtValue("");
  };

  console.log("Rendering WebpartFeedback...");
  return (
    <div>
      <Panel
        isLightDismiss
        isOpen={isOpen}
        type={PanelType.medium}
        onDismiss={dismissPanel}
        headerText={strings.FeedbackLabel}
        closeButtonAriaLabel="Close"
      >
        <form onSubmit={handleSubmit}>
          <Dropdown label={(props.showCategory) ? strings.FeedbackCategoryLabel : ""} options={options} defaultSelectedKey={selectedCategory.key} hidden={(!props.showCategory)}
          onChange={(event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number) => {setSelectedCategory(option);}}/>
          <TextField name="feedbackTxt" multiline rows={8} value={txtValue} label={strings.FeedbackBoxLabel}
          onChange={(event: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) =>
          {setTxtValue(newValue);}}/>
          <br></br>
          <div>
          <p>{strings.FeedbackInstructions}</p>
          <DefaultButton onClick={dismissPanel}>{strings.CancelButtonLabel}</DefaultButton>
          <PrimaryButton type="submit" styles={buttonStyles} disabled={(txtValue.length > 10) ? false : true}>
            {strings.SubmitButtonLabel}
          </PrimaryButton>

          </div>
          </form>
      </Panel>
      <Dialog
        hidden={!isDialogVisible}
        onDismiss={hideDialog}
        dialogContentProps={dialogContentProps}
        modalProps={dialogModalProps}
      >
        <DialogFooter>
          <PrimaryButton onClick={hideDialogAndPanel} text={strings.OkButtonLabel} />
        </DialogFooter>
      </Dialog>
    </div>
  );
}
