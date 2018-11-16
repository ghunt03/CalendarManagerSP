import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import { IColumn } from "office-ui-fabric-react";
export interface IEventListState {
  filterText: string;
  filterColumn: string;
  columns: Array<IColumn>;
  items: Array<MicrosoftGraph.Event>;
}


export const initialState = {
    filterText: "",
    filterColumn: "subject",
    items: [],
    columns: [
        {
          key: "subject",
          name: "Event",
          fieldName: "subject",
          minWidth: 200,
          maxWidth: 250,
          isResizable: true
        },
        {
          key: "start",
          name: "Start Date",
          fieldName: "start",
          minWidth: 140,
          maxWidth: 140,
          isResizable: true
        },
        {
          key: "end",
          name: "End Date",
          fieldName: "end",
          minWidth: 140,
          maxWidth: 140,
          isResizable: true
        },
        {
          key: "attendees",
          name: "Attendees",
          fieldName: "attendees",
          minWidth: 80,
          maxWidth: 80,
          isResizable: true
        }
      ]
};