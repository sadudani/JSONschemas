
import { IBizpMenuOptions } from '../../../../shared/IBizpSharedInterface';
import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IBizpWebpartHeaderProps {
  title?: string;
  showTitle?: boolean;
  helpId?: string;
  feedbackId?: string;
  context:WebPartContext;
  menuOptions?: IBizpMenuOptions[];
  onRefresh?:()=>void;
  updateProperty?: (value: string) => void;
  getListData?: () => string;
  children: React.ReactNode;
  themeVariant?:IReadonlyTheme;
}
