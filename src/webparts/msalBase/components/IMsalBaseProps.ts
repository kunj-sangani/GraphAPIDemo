import * as Msal from "msal";
import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface IMsalBaseProps {
  description: string;
  msalObjcet: Msal.UserAgentApplication;
  context: WebPartContext;
}
