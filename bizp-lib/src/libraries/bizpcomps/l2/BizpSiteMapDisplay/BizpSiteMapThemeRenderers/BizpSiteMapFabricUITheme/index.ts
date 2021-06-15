
// import * as strings from 'BizpcompsLibraryStrings';
// console.log("String test: " + strings.days);
// import nodeContentRenderer from './node-content-renderer';
import {FabricThemeNodeContentRenderer as nodeContentRenderer} from './BizpFabricThemeNodeContentRenderer2';
import {FabricThemeTreeNodeRenderer as treeNodeRenderer} from './BizpTreeNodeRenderer2';

 module.exports = {
  nodeContentRenderer,
  treeNodeRenderer,
  scaffoldBlockPxWidth: 25,
  rowHeight: 40,
  slideRegionSize: 50,
};

