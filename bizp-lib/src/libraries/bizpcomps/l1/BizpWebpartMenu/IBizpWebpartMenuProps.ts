
import { IBizpMenuOptions } from '../../../../shared/IBizpSharedInterface';
import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IBizpWebpartMenuProps {
  context: WebPartContext;
  menuOptions?: IBizpMenuOptions[];
  helpId?: string;
  themeVariant?:IReadonlyTheme;
  onRefresh?:()=>void;
}
