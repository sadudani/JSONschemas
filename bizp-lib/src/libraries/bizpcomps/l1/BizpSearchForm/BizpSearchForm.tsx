import * as React from 'react';
import {useState,useEffect} from 'react';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import {IIconProps} from 'office-ui-fabric-react';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import styles from './BizpSearchForm.module.scss';
import {SearchPrevButton,SearchNextButton,SearchInput,SearchCount} from './BizpSearchFormStyles';
import * as strings from 'BizpcompsLibraryStrings';
import {IBizpSearchFormProps} from './IBizpSearchFormProps';


export function BizpSearchForm(props: IBizpSearchFormProps) {
  console.log("Rendering Search Form Component - Search string: " + props.searchString);
  return (
    <form className = {styles.bizpSearchForm} style={{ display: 'flex', marginBottom: 10}} onSubmit={event => { event.preventDefault(); }}>
      <SearchInput id="siteMapFindBox" placeholder="Search..." theme={props.theme} value={props.searchString} onChange={event => props.onSearchStringChange(event.target.value)} />
      <SearchPrevButton disabled={!props.searchFoundCount} onClick={props.selectPrevMatch}>&lt; </SearchPrevButton>
      <SearchNextButton  disabled={!props.searchFoundCount} onClick={props.selectNextMatch}>&gt;</SearchNextButton>
      <SearchCount> &nbsp;{props.searchFoundCount > 0 ? props.searchFocusIndex + 1 : 0}&nbsp;/&nbsp;{props.searchFoundCount || 0} </SearchCount>
    </form>
  );
}
/* import styled from "styled-components";

const CustomizeButton = styled.button`
  outline: none;
  border: ${props => (props.primary ? "none" : "1px solid #d82b03")};
  cursor: pointer;
  font-family: Open Sans, sans-serif;
  font-size: ${props => (props.fontSize ? props.fontSize : "14px")};
  font-weight: bold;
  min-width: 165px;
  height: ${props => (props.height ? props.height : "30px")};
  width: ${props => props.width};
  background-color: ${props => (props.primary ? "#02a676" : "#fff")};
  border-radius: ${props => (props.borderRadius ? props.borderRadius : "14px")};
  color: ${props => (props.primary ? "#fff" : "#d82b03")};
  margin-top: 12px;
  padding: 5px 20px;

  &:hover {
    box-shadow: 0 1px 6px 0 rgba(32, 33, 36, 0.28);
  }
`;

export default CustomizeButton; */
