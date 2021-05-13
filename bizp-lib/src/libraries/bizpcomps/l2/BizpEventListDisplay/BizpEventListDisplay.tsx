import * as React from 'react';
import {useState,useEffect} from 'react';
import styles from './BizpEventListDisplay.module.scss';
import * as strings from 'BizpcompsLibraryStrings';
import {IBizpEventListDisplayProps} from './IBizpEventListDisplayProps';
import { IBizpEventDataSpec } from '../../../../shared/IBizpSharedInterface';
import * as moment from 'moment';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';

import {
  HoverCard, IPlainCardProps, HoverCardType, DefaultButton,PrimaryButton,
  DocumentCard,
  DocumentCardDetails,
  DocumentCardPreview,
  DocumentCardTitle,
  DocumentCardActivity,
  IDocumentCardActivityPerson,
  IDocumentCardPreviewProps,
  Icon
} from 'office-ui-fabric-react';
import { mergeStyles } from 'office-ui-fabric-react/lib/Styling';
import { DetailsList, buildColumns, IColumn, IDetailsListProps, IDetailsRowStyles, DetailsRow, Selection, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';

export function BizpEventListDisplay(props: IBizpEventListDisplayProps) {
  const [selectedItems,setSelectedItems] = useState(undefined);
  /*
  const [selection,setSelection] = useState<Selection>(
    // this function propagates selection to its parent component
    new Selection({
      onSelectionChanged: () =>{
          props.onNewSelection(selection);
      },
      selectionMode: SelectionMode.multiple
    })
  );
*/
  useEffect(() => {
    // init selection object
    setSelectedItems(props.selection.getSelection());
    },[]
  );

  const itemClass = mergeStyles({
    selectors: {
      '&:hover': {
        textDecoration: 'underline',
        cursor: 'pointer',
      },
    },
  });

  // const items: IExampleItem[] = createListItems(10);
  const buildSelectedColumn = (): IColumn[] => {
    const _columns = buildColumns(props.displayData).filter(
      column => column.name === 'title' || column.name === 'startDate',
    );
    for (var i in _columns)
    return _columns;
  };
  const columns: IColumn[] = buildSelectedColumn();

  // code taken from https://github.com/pnp/sp-dev-fx-webparts/blob/master/samples/react-calendar/src/webparts/calendar/components/Calendar.tsx
  const previewEventIcon: IDocumentCardPreviewProps = {
    previewImages: [
      {
        // previewImageSrc: event.ownerPhoto,
        previewIconProps: { iconName:  'Calendar', styles: { root: {  } }, className: styles.previewEventIcon },
        height: 43,
      }
    ]
  };

  const constructCardTimeDetails = (item: IBizpEventDataSpec): JSX.Element => {
    const singleDayEvent: boolean = (moment(item.startDate).format('MM/DD/YYYY') === moment(item.endDate).format('MM/DD/YYYY'));
    if (item.fAllDayEvent) {
      if (singleDayEvent) {
        // single all day event
        return (<span className={styles.DocumentCardTitleTime}>{moment(item.startDate).format("MMM D, YY ")}</span>);
      }
      else {
          // multi all day event
        return (<span className={styles.DocumentCardTitleTime}>{moment.utc(item.startDate).format("MMM D, YY ")} -
         {moment(item.endDate).format("MMM D, YY ")}</span>);
      }
    }
    else {
      if (singleDayEvent) {
          return (<span className={styles.DocumentCardTitleTime}>{moment(item.startDate).format("MMM D, YY ")}
                  {moment(item.startDate).format('h:m a')} - {moment(item.endDate).format('h:m a')}</span>);
      }
      else {
        return (
          <div>
          <span className={styles.DocumentCardTitleTime}>{moment(item.startDate).format("MMM D, YY ")} {moment(item.startDate).format('h:m a')} - {moment(item.endDate).format("MMM D, YY ")} {moment(item.endDate).format('h:m a')}</span>
        </div>
        );
      }
    }
  };

  const constructCardOwnerDetails = (item: IBizpEventDataSpec): JSX.Element => {
    const people: IDocumentCardActivityPerson[]=[{ name: item.ownerName, profileImageSrc: '', initials: item.ownerInitial }];
    if (item.ownerName) {
      // owner tag
      return (<DocumentCardActivity activity={item.category} people={people} />);
    }
  };

  const onRenderPlainCard = (item: IBizpEventDataSpec): JSX.Element => {
    return (
      <div className={itemClass} >
        <DocumentCard key={item.id.toString()} className={styles.Documentcard}   >
          <div>
            <DocumentCardPreview {...previewEventIcon} />
          </div>
          <DocumentCardDetails>
            <div className={styles.DocumentCardDetails}>
              <DocumentCardTitle title={item.title} shouldTruncate={true} className={styles.DocumentCardTitle} styles={{ root: { color: "black"} }} />
            </div>
            {constructCardTimeDetails(item)}
            {constructCardOwnerDetails(item)}
          </DocumentCardDetails>
        </DocumentCard>
      </div>
    );
  };

  const onRenderItemColumn = (item: IBizpEventDataSpec, index: number, column: IColumn): JSX.Element | React.ReactText => {
    const plainCardProps: IPlainCardProps = {
      onRenderPlainCard: onRenderPlainCard,
      renderData: item,
    };

    if (column.key === 'title') {
      return (
        <HoverCard plainCardProps={plainCardProps} instantOpenOnClick type={HoverCardType.plain}>
          <div className={itemClass} >
            {item.title}
          </div>
        </HoverCard>
      );
    }
    if (column.key === 'startDate') {
        return (moment(item.startDate).format("MMM D, YY "));
    }
    // this assumes that the item type is not an object. Otherwise, convert it to a string
    return item[column.key as keyof IBizpEventDataSpec];
  };

  const onRenderRow: IDetailsListProps['onRenderRow'] = props1 => {
    const customStyles: Partial<IDetailsRowStyles> = {};
    if (props1) {
      customStyles.root = { backgroundColor:props1.item.color  };
      return <DetailsRow {...props1} styles={customStyles} />;
    }
    return null;
  };

  console.log("Rendering EventListDisplay...");
  if (props.displayData && props.displayData.length > 0) {
    return (

      <div>
        <Fabric>
           <DetailsList setKey="hoverSet" data-is-scrollable="true" items={props.displayData}
           columns={columns} onRenderItemColumn={onRenderItemColumn} selection={ props.selection }
           onRenderRow={onRenderRow} />
        </Fabric>
      </div>
    );
  }
  else {
    return (
      <div>
        <p>
           No reminders to display.
        </p>
      </div>
    );
  }
}
