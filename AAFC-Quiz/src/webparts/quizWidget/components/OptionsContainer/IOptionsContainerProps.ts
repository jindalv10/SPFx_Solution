export interface IOptionsContainerProps {
  disabled: boolean;
  selectedKey?: () => string;
  options: string | undefined;
  label?: string;
  multiSelect: boolean  | undefined;
  onChange: (ev: any, option: any, isMultiSel: boolean, pollId: string) => void;
  PollId: string;
}