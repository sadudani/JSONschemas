//import React, { Component } from 'react';
import * as React from 'react';
import * as PropTypes from 'prop-types';
import { DetailsRow,GroupedList,initializeIcons,Icon,Selection } from 'office-ui-fabric-react';
import styles from './BizpNodeContentRenderer.module.scss';
import './BizpNode.css';

export function FabricThemeNodeContentRenderer (props) {
  const {
    scaffoldBlockPxWidth,
    toggleChildrenVisibility,
    connectDragPreview,
    connectDragSource,
    isDragging,
    canDrop,
    canDrag,
    node,
    title,
    draggedNode,
    path,
    treeIndex,
    isSearchMatch,
    isSearchFocus,
    icons,
    buttons,
    className,
    style,
    didDrop,
    lowerSiblingCounts,
    listIndex,
    swapFrom,
    swapLength,
    swapDepth,
    treeId, // Not needed, but preserved for other renderers
    isOver, // Not needed, but preserved for other renderers
    parentNode, // Needed for dndManager
    rowDirection,
    ...otherProps
  } = props;
  initializeIcons();
  function isDescendant(older, younger) {
    return (
      !!older.children &&
      typeof older.children !== 'function' &&
      older.children.some(
        child => child === younger || isDescendant(child, younger)
      )
    );
  }
  function onRenderCell(nestingDepth, item, itemIndex) {
    return React.isValidElement(item) ? (<div className={styles.detailRow}>{item}</div>) : (
      <DetailsRow
        groupNestingDepth={nestingDepth}
        item={item}
        itemIndex={itemIndex}
        columns={getColumns()}
        selection={getSelections()}
        selectionMode={0}
        className={styles.detailRow}
      />
    );
  }
  function getItems() {
//    const { propsNode, propsPath, PropsTreeIndex, propsTitle } = props;
    const nodeTitle = title || node.title;

    return [typeof nodeTitle === 'object' ? nodeTitle : {
      description: typeof nodeTitle === 'function'
        ? nodeTitle({
          node,
          path,
          treeIndex,
        })
        : nodeTitle
    }];
  }
  function getColumns() {
    return Object.keys(getItems()[0]).slice(0, 3)
      .map(
        (key) => ({
          key,
          name: key,
          fieldName: key,
          minWidth: 300
        })
      );
  }
  function getSelections() {
    const selection = new Selection();
    selection.setItems(getItems());
    return selection;
  }
  // hack change the aria-label to be tree
  function nodeRef () {
    const virtualGrid = document.getElementsByClassName("ReactVirtualized__Grid");
    if (virtualGrid && virtualGrid.length > 0) {
      const gridElement = virtualGrid[0];
      if (gridElement.hasAttribute("aria-label") && gridElement.getAttribute("aria-label") === "grid") {
        gridElement.setAttribute("aria-label","tree");
      }
    }
  }

  // Construct th}e scaffold representing the structure of the tree
  function initRender():boolean{
    const scaffold = [];
    lowerSiblingCounts.forEach((lowerSiblingCount, i) => {
      scaffold.push(
        <div
          key={`pre_${1 + i}`}
          style={{ width: scaffoldBlockPxWidth }}
          className={styles.lineBlock}
        />
      );

      if (treeIndex !== listIndex && i === swapDepth) {
        // This row has been shifted, and is at the depth of
        // the line pointing to the new destination
        let highlightLineClass = '';

        if (listIndex === swapFrom + swapLength - 1) {
          // This block is on the bottom (target) line
          // This block points at the target block (where the row will go when released)
          highlightLineClass = styles.highlightBottomLeftCorner;
        } else if (treeIndex === swapFrom) {
          // This block is on the top (source) line
          highlightLineClass = styles.highlightTopLeftCorner;
        } else {
          // This block is between the bottom and top
          highlightLineClass = styles.highlightLineVertical;
        }

        scaffold.push(
          <div
            key={`highlight_${1 + i}`}
            style={{
              width: scaffoldBlockPxWidth,
              left: scaffoldBlockPxWidth * i,
            }}
            className={`${styles.absoluteLineBlock} ${highlightLineClass}`}
          />
        );
      }
    });
    return true;
  }
  const isLandingPadActive = !didDrop && isDragging;
  const isDraggedDescendant = draggedNode && isDescendant(draggedNode, node);
  const isLeafNode = !(toggleChildrenVisibility && node.children && node.children.length > 0);


  return (
      initRender() &&
      <div className={styles.rowStyle + (isLeafNode ? ` ${styles.rowLeafStyle}` : ` ${styles.rowTreeStyle}`) +
      (isSearchFocus ? ` ${styles.rowSearchFocus}` : '') +
      (isSearchMatch ? ` ${styles.rowSearchMatch}` : '')} {...otherProps}
      aria-label={node.ariaLabel}
      ref={nodeRef}
      >
      {// eslint-disable-next-line jsx-a11y/no-static-element-interactions,  jsx-a11y/click-events-have-key-events, no-unused-expressions
        <div className={styles.titleContainer}
          onClick={() =>
            {
              // eslint-disable-next-line no-unused-expressions
              if (!isLeafNode && !node.isTag)
                toggleChildrenVisibility({
                  node,
                  path,
                  treeIndex,
                });
            }
          }
        >
          {/* {scaffold} */}
          <div
            className={
              styles.row +
              (isLandingPadActive ? ` ${styles.rowLandingPad}` : '') +
              (isLandingPadActive && !canDrop
                ? ` ${styles.rowCancelPad}`
                : '') +
              (className ? ` ${className}` : '')
            }
            style={{
              opacity: isDraggedDescendant ? 0.5 : 1,
              ...style,
            }}
          >
            <GroupedList
              items={getItems()}
              onRenderCell={onRenderCell}
              selection={getSelections()}
              selectionMode={0}
            />
          </div>
        </div>
      }
        {!isLeafNode && !node.isTag && (
            <button
              type="button"
              aria-expanded={!!node.expanded}
              aria-label="Toggle the tree expand or collapse"
              className={
                node.expanded ? styles.collapseButton : styles.expandButton
              }
              style={{
                left: (lowerSiblingCounts.length - 0.7) * scaffoldBlockPxWidth,
              }}
              onClick={() =>
                toggleChildrenVisibility({
                  node,
                  path,
                  treeIndex,
                })
              }>
                {node.expanded ? (
                <Icon iconName="ChevronDown" className={styles.icon} />) : (
                <Icon iconName="ChevronRight" className={styles.icon} />)
                }
            </button>
          )}
      </div>
  );
}

FabricThemeNodeContentRenderer.defaultProps = {
  buttons: [],
  canDrag: false,
  canDrop: false,
  className: '',
  draggedNode: null,
  icons: [],
  isSearchFocus: false,
  isSearchMatch: false,
  parentNode: null,
  style: {},
  swapDepth: null,
  swapFrom: null,
  swapLength: null,
  title: null,
  toggleChildrenVisibility: null,
};

FabricThemeNodeContentRenderer.propTypes = {
  buttons: PropTypes.arrayOf(PropTypes.node),
  canDrag: PropTypes.bool,
  className: PropTypes.string,
  icons: PropTypes.arrayOf(PropTypes.node),
  isSearchFocus: PropTypes.bool,
  isSearchMatch: PropTypes.bool,
  listIndex: PropTypes.number.isRequired,
  lowerSiblingCounts: PropTypes.arrayOf(PropTypes.number).isRequired,
  node: PropTypes.shape({}).isRequired,
  path: PropTypes.arrayOf(
    PropTypes.oneOfType([PropTypes.string, PropTypes.number])
  ).isRequired,
  scaffoldBlockPxWidth: PropTypes.number.isRequired,
  style: PropTypes.shape({}),
  swapDepth: PropTypes.number,
  swapFrom: PropTypes.number,
  swapLength: PropTypes.number,
  title: PropTypes.oneOfType([PropTypes.func, PropTypes.node]),
  toggleChildrenVisibility: PropTypes.func,
  treeIndex: PropTypes.number.isRequired,
  treeId: PropTypes.string.isRequired,
  rowDirection: PropTypes.string.isRequired,

  // Drag and drop API functions
  // Drag source
  connectDragPreview: PropTypes.func.isRequired,
  connectDragSource: PropTypes.func.isRequired,
  didDrop: PropTypes.bool.isRequired,
  draggedNode: PropTypes.shape({}),
  isDragging: PropTypes.bool.isRequired,
  parentNode: PropTypes.shape({}), // Needed for dndManager
  // Drop target
  canDrop: PropTypes.bool,
  isOver: PropTypes.bool.isRequired,
};

