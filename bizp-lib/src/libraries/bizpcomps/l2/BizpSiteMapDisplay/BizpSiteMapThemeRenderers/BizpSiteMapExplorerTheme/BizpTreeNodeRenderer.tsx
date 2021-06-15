import * as React from 'react';
import * as PropTypes from 'prop-types';
import styles from './BizpTreeNode.module.scss';

interface IProps {
  treeIndex: number;
  treeId: string;
  swapFrom?: number;
  swapDepth?: number;
  swapLength?: number;
  scaffoldBlockPxWidth: number;
  lowerSiblingCounts: number[];
  listIndex: number;
  children: React.ReactNode;

  // Drop target
  connectDropTarget: (arg:any) => void;
  isOver: boolean;
  canDrop?: boolean;
  draggedNode: {};

  // used in dndManager
  getPrevRow: () => void;
  node: {};
  path: (string | number)[];
  rowDirection: string;
}

export class ExplorerThemeTreeNodeRenderer extends React.Component<IProps> {
  public static defaultProps = {
    swapFrom: null,
    swapDepth: null,
    swapLength: null,
    canDrop: false,
    draggedNode: null,
  };

  public render():any {
    const {
      children,
      listIndex,
      swapFrom,
      swapLength,
      swapDepth,
      scaffoldBlockPxWidth,
      lowerSiblingCounts,
      connectDropTarget,
      isOver,
      draggedNode,
      canDrop,
      treeIndex,
      treeId, // Delete from otherProps
      getPrevRow, // Delete from otherProps
      node, // Delete from otherProps
      path, // Delete from otherProps
      rowDirection,
      ...otherProps
    } = this.props;

    return connectDropTarget(
      <div {...otherProps} className={styles.node}>
        {React.Children.map(this.props.children, child => {
          if (!React.isValidElement<IProps>(child)) {
            return child;
          }
          return React.cloneElement(child, {
            isOver,
            canDrop,
            draggedNode,
            lowerSiblingCounts,
            listIndex,
            swapFrom,
            swapLength,
            swapDepth,
          });
        }
        )}
      </div>
    );
  }
}


