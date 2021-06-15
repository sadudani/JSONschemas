
import styled from 'styled-components';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

export const SearchPrevButton = styled.button`
  type: button;
`;
export const SearchNextButton = styled.button`
  type: submit;
`;
export const SearchInput = styled.input<{ theme: IReadonlyTheme; }>`
  type: "text";
  placeholder:"Search...";
  width: 100%;
  padding: .375rem .75rem;
  font-size: "[theme:fonts.small, default:1rem]";
//  font-weight: 400;
  line-height: 1.5;
//  color: #212529;
//  color: ${({ theme }) => theme.inputText || '#ccc'};
//  background-color:  ${({ theme }) => theme.inputBackground || '#fff'};
  color: "[theme:inputText, default:#ccc]";
  background-color: "[theme:inputBackground, default:#fff]";
  background-clip: padding-box;
  border: 1px solid #ced4da;
  -webkit-appearance: none;
  -moz-appearance: none;
  appearance: none;
  border-radius: .25rem;
  transition: border-color .15s ease-in-out,box-shadow .15s ease-in-out;
`;

export const SearchCount = styled.p`
  font-size: "[theme:fonts.small, default:1rem]";
  color: "[theme:bodySubtext, default:#ccc]";
  text-align: center;
`;
