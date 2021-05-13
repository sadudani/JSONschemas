import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IBizpWebpartFeedbackProps {
  context: WebPartContext;
  openSignal:boolean;
  feedbackId: string;
  showCategory:boolean;
  themeVariant: IReadonlyTheme;
}
