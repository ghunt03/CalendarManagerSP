import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import { MessageBarType } from "office-ui-fabric-react/lib/MessageBar";
export interface ICalendarMgrState {
  selectedUsers: Array<MicrosoftGraph.User>;
  selectedEvents: Array<MicrosoftGraph.Event>;
  showPanel: boolean;
  message: string;
  messageType: MessageBarType;
}

export const initialState = {
  selectedUsers: [],
  selectedEvents: [],
  showPanel: false,
  message: "",
  messageType: MessageBarType.info
};
