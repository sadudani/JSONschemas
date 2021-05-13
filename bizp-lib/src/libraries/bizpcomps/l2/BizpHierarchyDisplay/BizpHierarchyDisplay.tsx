import * as React from 'react';
import {useState,useEffect} from 'react';
import { sp,Web,IWebInfosData } from "@pnp/sp/presets/all";
import { ISearchQuery, SearchResults} from "@pnp/sp/search";

import styles from './BizpHierarchyDisplay.module.scss';
import * as strings from 'BizpcompsLibraryStrings';
import {IBizpHierarchyDisplayProps,IBizpSiteHierarchyData} from './IBizpHierarchyDisplayProps';
import { getSPSites,getSPLibs } from '../../../../shared/BizpBasesvc';

import { TreeView, ITreeItem, TreeViewSelectionMode,ITreeItemAction,TreeItemActionsDisplayMode} from "@pnp/spfx-controls-react/lib/TreeView";
import { IconButton, IIconProps, IContextualMenuItem, IContextualMenuProps} from 'office-ui-fabric-react';

export function BizpHierarchyDisplay(props: IBizpHierarchyDisplayProps) {
  let navigateIcon: IIconProps = { iconName: 'NavigateBackMirrored' };
  let sampleItems: ITreeItem[] = [
    {
      key: "R1",
      label: "Root",
      subLabel: "This is a sub label for node",
      iconProps: { iconName: 'SkypeCheck' },
      actions: [{
        title: "Get item",
        iconProps: {
          iconName: 'Warning',
          style: {
            color: 'salmon',
          },
        },
        id: "GetItem",
        actionCallback: async (treeItem: ITreeItem) => {
          console.log(treeItem);
        }
      }],
      children: [
        {
          key: "1",
          label: "Parent 1",
          selectable: false,
          children: [
            {
              key: "3",
              label: "Child 1",
              subLabel: "This is a sub label for node",
              actions: [{
                title:"Share",
                iconProps: {
                  iconName: 'Share'
                },
                id: "GetItem",
                actionCallback: async (treeItem: ITreeItem) => {
                  console.log(treeItem);
                }
              }],
              children: [
                {
                  key: "gc1",
                  label: "Grand Child 1",
                  actions: [{
                    title: "Get Grand Child item",
                    iconProps: {
                      iconName: 'Mail'
                    },
                    id: "GetItem",
                    actionCallback: async (treeItem: ITreeItem) => {
                      console.log(treeItem);
                    }
                  }]
                }
              ]
            },
            {
              key: "4",
              label: "Child 2",
              iconProps: { iconName: 'SkypeCheck' }
            }
          ]
        },
        {
          key: "2",
          label: "Parent 2"
        },
        {
          key: "5",
          label: "Parent 3",
          disabled: true
        },
        {
          key: "6",
          label: "Parent 4",
          selectable: true
        }
      ]
    },
    {
      key: "R2",
      label: "Root 2",
      children: [
        {
          key: "8",
          label: "Parent 5"
        }
      ]
    }
  ];
  const [siteTree,setSiteTree] = useState<ITreeItem[]>(sampleItems);

  useEffect(() => {
    // init selection object
    loadSites();
    },[]
  );
  useEffect(() => {
    // init selection object
    loadSites();
    },[props.siteUrl,props.refresh]
  );

  async function loadSites(){
    // init data
    const r:any = await getSPSites(props.siteUrl);
    const sItems:ITreeItem[] = constructSiteTree(r);
    setSiteTree(sItems);
    console.log("New Tree: " + JSON.stringify(sItems));
  }

  async function loadSPLibraries(site:ITreeItem):Promise<ITreeItem[]>{
    // init data
    // subLabel is the WebId
    const r:any = await getSPLibs(site.key,site.data.SiteId,site.data.WebId);
    console.log("libraries: " + JSON.stringify(r));
    const sItems:ITreeItem[] = constructLibTree(r);
    console.log("Library Tree: " + JSON.stringify(sItems));
    return sItems;
  }

  function hasChildren(node:any) {
    return (typeof node === 'object')
        && (typeof node.children !== 'undefined')
        && (node.children.length > 0);
  }

  function constructSiteTree(r:IBizpSiteHierarchyData[]):ITreeItem[] {
    let newTree:ITreeItem[] = r.map(item=>{
      let newItem:ITreeItem = constructSiteNode(item);
      if (hasChildren(item)) newItem.children = constructSiteTree(item.children);
      return newItem;
    });
    return newTree;
  }

  function constructSiteNode(item:IBizpSiteHierarchyData): ITreeItem {
    let i:ITreeItem = {
      key: item.data.Path,
      label: item.data.Title,
      subLabel: item.data.WebId,
      data: item.data,
      iconProps: {iconName: 'NavigateBackMirrored'} ,
      actions: [{
        title: "Go to site",
        iconProps: {
          iconName: 'Warning',
          style: {
            color: 'salmon',
          },
        },
        id: "GoToSite",
        actionCallback: async (treeItem: ITreeItem) => {
          console.log(treeItem);
        },
      }],
      children:[]
    };
//    const sItems:ITreeItem[] = await loadSPLibraries(i);
//    i.children = sItems;
    return i;
  }

  function constructLibTree(r:IBizpSiteHierarchyData[]):ITreeItem[] {
    let newTree:ITreeItem[] = r.map(item=>{
      let newItem:ITreeItem = constructLibNode(item);
      return newItem;
    });
    return newTree;
  }

  function constructLibNode(item:IBizpSiteHierarchyData): ITreeItem {
    let i = {
      key: item.data.UniqueId,
      label: item.data.Title,
      subLabel: item.data.WebId,
      data: item.data,
      iconProps: {iconName: 'NavigateBackMirrored'} ,
      actions: [{
        title: "Go to site",
        iconProps: {
          iconName: 'Warning',
          style: {
            color: 'salmon',
          },
        },
        id: "GoToSite",
        actionCallback: async (treeItem: ITreeItem) => {
          console.log(treeItem);
        },
        children:[]
      }],
    };
    loadSPLibraries(i);
    return i;
  }

  function renderCustomTreeItem(item: ITreeItem): JSX.Element {
    return (

      <a href={item.key} target="_blank">
      <span>
        {
          item.iconProps &&
          <i className={"ms-Icon ms-Icon--" + item.iconProps.iconName} style={{ paddingRight: '4px' }} />
        }
        {item.label}
      </span></a>
    );
  }
  function onTreeItemExpandCollapse(item: ITreeItem, isExpanded: boolean) {
    if (isExpanded) {
      loadSPLibraries(item);
    }
    console.log((isExpanded ? "Item expanded: " : "Item collapsed: ") + item);
  }

  return (
    <div className={styles.spfxPnpTreeview}>
    <TreeView
      items={siteTree}
      defaultExpanded={false}
      selectionMode={TreeViewSelectionMode.None}
      selectChildrenIfParentSelected={true}
      showCheckboxes={false}
      treeItemActionsDisplayMode={TreeItemActionsDisplayMode.Buttons}
      onExpandCollapse={onTreeItemExpandCollapse}
      onRenderItem={renderCustomTreeItem} />
    </div>
  );
}
