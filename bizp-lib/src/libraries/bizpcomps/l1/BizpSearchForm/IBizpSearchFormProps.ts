import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from '@microsoft/sp-component-base';

export interface IBizpSearchFormProps {
  searchFoundCount: number;
  searchFocusIndex: number;
  searchString: string;
  theme:IReadonlyTheme;
  selectPrevMatch: () => void;
  selectNextMatch: () => void;
  onSearchStringChange: (newSearch: string) => void;
}
