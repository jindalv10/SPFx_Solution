import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface INominationProps {
  description: string;
  context: WebPartContext;
  currentUser?: string;
  //formType:string;
}
