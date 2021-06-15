// import * as strings from 'BizpcompsLibraryStrings';
// console.log("String test: " + strings.days);
// import nodeContentRenderer from './node-content-renderer';
import {ExplorerThemeNodeContentRenderer as nodeContentRenderer} from './BizpNodeContentRenderer';
import {ExplorerThemeTreeNodeRenderer as treeNodeRenderer} from './BizpTreeNodeRenderer';

 module.exports = {
  nodeContentRenderer,
  treeNodeRenderer,
  scaffoldBlockPxWidth: 25,
  rowHeight: 25,
  slideRegionSize: 50,
};
