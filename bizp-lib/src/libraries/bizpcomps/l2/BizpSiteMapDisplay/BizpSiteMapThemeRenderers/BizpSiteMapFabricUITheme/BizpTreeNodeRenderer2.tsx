import * as React from 'react';
import {Children,cloneElement,Component} from 'react';
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


export class FabricThemeTreeNodeRenderer extends React.Component<IProps> {
/*   public static propTypes = {
    treeIndex: PropTypes.number.isRequired,
    treeId: PropTypes.string.isRequired,
    swapFrom: PropTypes.number,
    swapDepth: PropTypes.number,
    swapLength: PropTypes.number,
    scaffoldBlockPxWidth: PropTypes.number.isRequired,
    lowerSiblingCounts: PropTypes.arrayOf(PropTypes.number).isRequired,
    listIndex: PropTypes.number.isRequired,
    children: PropTypes.element.isRequired,

    // Drop target
    connectDropTarget: PropTypes.func.isRequired,
    isOver: PropTypes.bool.isRequired,
    canDrop: PropTypes.bool,
    draggedNode: PropTypes.shape({}),

    // used in dndManager
    getPrevRow: PropTypes.func.isRequired,
    node: PropTypes.shape({}).isRequired,
    path: PropTypes.arrayOf(
      PropTypes.oneOfType([PropTypes.string, PropTypes.number])
    ).isRequired,
    rowDirection: PropTypes.string.isRequired,
  };
 */
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
    /* const childrenWithExtraProp = React.Children.map(this.props.children, child => {
      return React.cloneElement(child, {
        isPlaying: child.props.title === currentPlayingTitle
      });
    }); */

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
        })}
      </div>
    );
  }
}

