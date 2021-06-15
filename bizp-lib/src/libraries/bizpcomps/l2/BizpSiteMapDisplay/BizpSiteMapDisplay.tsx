import * as React from 'react';
import {useState,useEffect} from 'react';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import {IIconProps} from 'office-ui-fabric-react';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import styles from './BizpSiteMapDisplay.module.scss';
import {TreeContainer} from './BizpSiteMapStyles';
import * as strings from 'BizpcompsLibraryStrings';
import {IBizpSiteMapDisplayProps,IBizpSiteData,IBizpSiteHierarchyData} from './IBizpSiteMapDisplayProps';
import { getSPSites} from '../../../../shared/BizpBasesvc';
import {BizpSearchForm} from '../../l1/BizpSearchForm/BizpSearchForm';

import { TreeView, ITreeItem, TreeViewSelectionMode,ITreeItemAction,TreeItemActionsDisplayMode} from "@pnp/spfx-controls-react/lib/TreeView";
import SortableTree from 'react-sortable-tree';
import 'react-sortable-tree/style.css'; // This only needs to be imported once in your app

import * as FabricUIThemeNodeContentRenderer from './BizpSiteMapThemeRenderers/BizpSiteMapFabricUITheme/index';
import * as ExplorerThemeNodeContentRenderer from './BizpSiteMapThemeRenderers/BizpSiteMapExplorerTheme/index';

export function BizpSiteMapDisplay(props: IBizpSiteMapDisplayProps) {
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

    if (!loaded) loadSites();
    setLoaded(true);
    },[]
  );

  useEffect(() => {
    // init selection object
    loadSites();
    },[props.siteUrl,props.refresh,props.layout]
  );

  async function loadSites(){
    // init data
    const r:any = await getSPSites(props.siteUrl,props.displayLibs);
    let sItems:ITreeItem[];
    nodeId = 0;
    sItems = constructSiteTree(r);

    setSiteTree(sItems);
    console.log("New Tree: " + JSON.stringify(sItems));
  }

  function hasChildren(node:any) {
    return (typeof node === 'object')
        && (typeof node.children !== 'undefined')
        && (node.children.length > 0);
  }

  function constructSiteTree(r:IBizpSiteHierarchyData[]):any[] {
    let children:any[];
    let newTree:any[] = r.map((item,index) =>{
      let newItem:any = constructSiteNode(item);
      nodeId++;
      if (hasChildren(item)) {
        children = constructSiteTree(item.children);
        newItem.children = children;
      }
      return newItem;
    });
    return newTree;
  }

  function constructSiteNode(item:IBizpSiteHierarchyData): any {
    // select icon for the node
    let iName:string;
    let type:string;
    switch (item.data.contentclass) {
      case "STS_List_DocumentLibrary":
        iName = "DocLibrary";
        type = "library";
        break;
      case "STS_Web":
        iName = "SharepointAppIcon16";
        type = "siteNode";
        break;
      case "STS_Site":
        iName = "SharepointLogoInverse";
        type = "rootSiteNode";
        break;
      default:
        iName = "GroupObject";
        type = "groupNode";
        break;
    }

    let i = {
      'id': nodeId,
      'title': item.data.Title,
      'parent_id': item.data.ParentLink,
      'type': type,
      'url': item.data.Path,
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

/*   function onTreeItemExpandCollapse(item: ITreeItem, isExpanded: boolean) {
    console.log((isExpanded ? "Item expanded: " : "Item collapsed: ") + item);
  } */

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
          case 'siteNode':
            return (<>
              <a style={{  color:'black',textDecoration: 'none' }} href={node['url']} target="_blank">
                <Icon iconName='SharepointLogoInverse' style={{ paddingRight: '4px' }} />
                {node.title}
              </a>
            </>);
          case 'rootSiteNode':
            return (<>
              <a style={{  color:'black', textDecoration: 'none', fontSize:'[theme:fonts.small, default:1rem]' }} href={node['url']} target="_blank">
                <Icon iconName='SharepointLogoInverse' style={{ paddingRight: '4px' }} />
                {node.title}
              </a>
            </>);
          case 'libraryNode':
            return (<><Icon iconName='FabricDocLibrary' style={{ color:'black',paddingRight: '4px' }} />{node.title}</>);
          case 'library':
            return (<>
              <a style={{color:'black', textDecoration: 'none' }} href={node['url']} target="_blank">
                <Icon iconName='FabricDocLibrary' style={{ paddingRight: '4px' }} />
                {node.title}
              </a>
            </>);
          case 'list':
          case 'default':
            return (
              <>
                <a style={{  color:'black',textDecoration: 'none' }} href={node['url']} target="_blank" title={node.title}><Icon iconName='CustomList' style={{ paddingRight: '4px' }} />{node.title}</a>
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

  console.log("Rendering Site Map Component layout:" + props.layout);
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
            {props.layout == 2 &&
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
              />s
            </TreeContainer>
            }
             {props.layout == 3 &&
            <TreeContainer style={{ height: 600 }}>
              <SortableTree treeData={siteTree}
                onChange={treeData => setSiteTree(treeData)}
                theme={ExplorerThemeNodeContentRenderer}
                searchMethod={customSearchMethod}
                searchQuery={searchString}
                searchFocusOffset={searchFocusIndex}
                canDrag={false}
                searchFinishCallback={matches => {
                  if (matches.length) {
                    setSearchFocusIndex(matches.length > 0 ? searchFocusIndex % matches.length : 0);
                    setSearchFoundCount(matches.length);
                  }
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
            {props.layout == 4 &&
            <TreeContainer style={{ height: 600}}>
              <SortableTree treeData={siteTree}
                onChange={treeData => setSiteTree(treeData)}
                theme={FabricUIThemeNodeContentRenderer}
                searchMethod={customSearchMethod}
                searchQuery={searchString}
                searchFocusOffset={searchFocusIndex}
                onVisibilityToggle={onTreeNodeExpandCollapse}
                canDrag={false}
                searchFinishCallback={matches => {
 //onChange={treeData => setSiteTree(treeData)}
                  if (matches.length) {
                    setSearchFocusIndex(matches.length > 0 ? searchFocusIndex % matches.length : 0);
                    setSearchFoundCount(matches.length);
                  }
                }}
                reactVirtualizedListProps={{ width: 600, rowHeight: 44, style:{ color: '#ffffff', textDecoration: 'none' }}}
                generateNodeProps={generateNodeProps}
              />
            </TreeContainer>
            }
          </>}
        </div>
      </div>
    </div>
  </div>);
}
