import * as React from 'react';
import {useState,useEffect} from 'react';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import {IIconProps} from 'office-ui-fabric-react';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { graph } from "@pnp/graph";
import { IUser, IUsers, User, Users, IPeople, People} from "@pnp/graph/users";

import styles from './BizpOrgChartDisplay.module.scss';
import {TreeContainer} from './BizpOrgChartDisplayStyles';
import * as strings from 'BizpcompsLibraryStrings';
import {IBizpOrgChartDisplayProps,IBizpUserData,IBizpOrgHierarchyData} from './IBizpOrgChartDisplayProps';
import { getUsers} from '../../../../shared/BizpBasesvc';
import {BizpSearchForm} from '../../l1/BizpSearchForm/BizpSearchForm';

import { TreeView, ITreeItem, TreeViewSelectionMode,ITreeItemAction,TreeItemActionsDisplayMode} from "@pnp/spfx-controls-react/lib/TreeView";
import SortableTree from 'react-sortable-tree';
import 'react-sortable-tree/style.css'; // This only needs to be imported once in your app
import { jsS } from '@pnp/common';

export function BizpOrgChartDisplay(props: IBizpOrgChartDisplayProps) {
  let navigateIcon: IIconProps = { iconName: 'NavigateBackMirrored' };
  let nodeId:number;

  const [siteTree,setSiteTree] = useState<any[]>([]);
  const [loaded, setLoaded] = useState(false);
  const [searchString, setSearchString] = useState("");
  const [searchFoundCount, setSearchFoundCount] = useState<number>(0);
  const [searchFocusIndex, setSearchFocusIndex] = useState<number>(0);

  /*  const treeData = [
      { title: 'Chicken', id: 1, type: "rootSiteNode", url:"https://officeworksventure.sharepoint.com/sites/portals365", expanded: true,
      children: [{ title: 'Fish', id: 2, type: "rootSiteNode", url:"https://officeworksventure.sharepoint.com/sites/portals365", expanded: true,
      children: [] }] }
    ]; */
  useEffect(() => {
    // init selection object
    if (!loaded) loadUsers();
    setLoaded(true);
    },[]
  );

  useEffect(() => {
    // init selection object
    loadUsers();
    },[props.refresh,props.layout]
  );

  async function loadUsers(){
    // init data
    graph.setup({
      spfxContext: props.context
    });
    const r:IBizpOrgHierarchyData[] = await getUsers();

    let sItems:ITreeItem[];
    nodeId = 0;
    sItems = constructOrgTree(r);

    setSiteTree(sItems);
    console.debug("New Tree: ",r);
  }

  function hasChildren(node:any) {
    return (typeof node === 'object')
        && (typeof node.children !== 'undefined')
        && (node.children.length > 0);
  }

  function constructOrgTree(r:IBizpOrgHierarchyData[]):any[] {
    let children:any[];
    let newTree:any[] = r.map((item,index) =>{
      let newItem:any = constructOrgNode(item);
      nodeId++;
      if (hasChildren(item)) {
        children = constructOrgTree(item.children);
        if (item.data.manager.id == null) {
          newItem.type = "Executive";
        }
        else {
          newItem.type = "Manager";
        }
        newItem.children = children;
      }
      return newItem;
    });
    return newTree;
  }

  function constructOrgNode(item:IBizpOrgHierarchyData): any {
    // select icon for the node
    let type:string;
    if (item.data.manager.id == null) {
      if (item.data.surname == null) {
        type = "Conference Room";
      }
      else if (item.data.mobilePhone == null) {
        type = "Non-Employee";
      }
      else {
        type = "Employee";
      }
    }
    else {
      type = "Employee";
    }
    let i = {
      'id': item.data.id,
      'title': item.data.displayName,
      'parent_id': item.data.manager.id,
      'type': type,
      'expanded': true,
      'data': item.data,
      children:null
    };
    return i;
  }

   function renderCustomTreeItem(item: ITreeItem): JSX.Element {
    return (

      <a href={item.key} target="_blank">
      <span>
        {
          item.iconProps &&
          <Icon iconName={item.iconProps.iconName} style={{ paddingRight: '4px' }} />
        }
        {item.label}
      </span></a>
    );
  }

  const onTreeNodeExpandCollapse =({ expanded }) => {
  // out = expanded ? 'expanded' : 'collapsed';
    console.log ("OnVisibilityToggle expanded:"+expanded);
  };

  // Case insensitive search of `node.title`
  const customSearchMethod = ({ node, searchQuery }) => searchQuery && node.title.toLowerCase().indexOf(searchQuery.toLowerCase()) > -1;
  const selectPrevMatch = () => searchFocusIndex !== null ? setSearchFocusIndex((searchFoundCount + searchFocusIndex - 1) % searchFoundCount) : setSearchFocusIndex(searchFoundCount - 1);
  const selectNextMatch = () => searchFocusIndex !== null ? setSearchFocusIndex((searchFocusIndex + 1) % searchFoundCount) : setSearchFocusIndex(0);
  const onSearchStringChange = (s) => setSearchString(s);

  const generateNodeProps= ({ node, path, treeIndex, lowerSiblingCounts, isSearchMatch, isSearchFocus}) => {
    let cName = `styles.level${path.length > 3 ? 3 : path.length}`;
    let cColor:string;
    console.log("generateNodeProps node: " + JSON.stringify(node));
    switch (path.length.toString()) {
      case '1':
        cName = "level1";
        cColor = 'black';
        break;
      case '2':
        cName = "level2";
        cColor = '#F44336';
        break;
      case '3':
        cName = "level3";
        cColor = '#4CAF50';
        break;
      default:
        cName = "level4";
        cColor = '#03A9F4';
        break;
    }
    const linkStyle:string = "color:" + cColor + ",width: 50, textDecoration: 'none'";
    const titleObj =
      (() => {
        switch (node['type']) {
          case 'Manager':
            return (<>
              <a style={{  color:'black',textDecoration: 'none' }} >
                <Icon iconName='People' style={{ paddingRight: '4px' }} />
                {node.title}
              </a>
            </>);
          case 'Executive':
            return (<>
              <a style={{  color:'black', textDecoration: 'none', fontSize:'[theme:fonts.small, default:1rem]' }} >
                <Icon iconName='AccountManagement' style={{ paddingRight: '4px' }} />
                {node.title}
              </a>
            </>);

          case 'Employee':
            return (<>
              <a style={{  color:'black', textDecoration: 'none', fontSize:'[theme:fonts.small, default:1rem]' }} >
                <Icon iconName='ReminderPerson' style={{ paddingRight: '4px' }} />
                {node.title}
              </a>
            </>);
          case 'Non-Employee':
            return (<>
              <a style={{color:'black', textDecoration: 'none' }} >
                <Icon iconName='ObjectRecognition' style={{ paddingRight: '4px' }} />
                {node.title}
              </a>
            </>);
          case 'Conference Room':
              return (<>
                <a style={{color:'black', textDecoration: 'none' }} >
                <Icon iconName='Room' style={{ paddingRight: '4px' }} />
                {node.title}
                </a>
              </>);
          case 'default':
            return (
              <>
                <a style={{  color:'black',textDecoration: 'none' }} title={node.title}><Icon iconName='CustomList' style={{ paddingRight: '4px' }} />{node.title}</a>
              </>
            );
        }
      })
    ;
    console.debug("generateNodeProps titleObj: ",titleObj);
    return {
      title: titleObj,
      className: cName
    };
  };

  console.log("Rendering Org Chart Component - layout:" + props.layout);
  return (<div className={styles.bizportals365SiteMapTreeView}>
    <div className={styles.container}>
      <div className={styles.row}>
        <BizpSearchForm  searchFoundCount = {searchFoundCount} searchFocusIndex = {searchFocusIndex}
                            searchString = {searchString} selectPrevMatch = {selectPrevMatch}
                            selectNextMatch = {selectNextMatch} onSearchStringChange = {onSearchStringChange}
                            theme = {props.theme}
        />
        <div className={styles.column}>
          {siteTree &&
          <>
            {
            <TreeContainer  style={{ height: 600}}>
              <SortableTree treeData={siteTree}
                onChange={treeData => setSiteTree(treeData)}
                searchMethod={customSearchMethod}
                searchQuery={searchString}
                searchFocusOffset={searchFocusIndex}
                canDrag={false}
                searchFinishCallback={matches => {
   //               onChange={a=>console.log("Row Change")}
                  setSearchFocusIndex(matches.length > 0 ? searchFocusIndex % matches.length : 0);
                  setSearchFoundCount(matches.length);
/*                   if (matches.length) {
                    setSearchFocusIndex(matches.length > 0 ? searchFocusIndex % matches.length : 0);
                    setSearchFoundCount(matches.length);
                  } */

                }}
                onVisibilityToggle={({ expanded }) => {
                  //                  out = expanded ? 'expanded' : 'collapsed';
                                    console.log ("OnVisibilityToggle expanded:"+expanded);
                                  }}
                reactVirtualizedListProps={{ width: 600, rowHeight: 44, style:{ color: '#ffffff', textDecoration: 'none' }}}
                generateNodeProps={generateNodeProps}
              />
            </TreeContainer>
            }
          </>
          }
        </div>
      </div>
    </div>
  </div>);
}
