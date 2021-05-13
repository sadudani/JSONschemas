
import { Selection } from 'office-ui-fabric-react';
import { IBizpEventDataSpec } from '../../../../shared/IBizpSharedInterface';
export interface IBizpEventListDisplayProps {
  displayData?: IBizpEventDataSpec[]; // events to display
  selection:Selection;
  refresh:boolean;
}
