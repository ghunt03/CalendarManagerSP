import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import { IColumn } from "office-ui-fabric-react";
export interface IInviteScreenState {
  showPanel: boolean;
  filterText: string;
  filterColumn: string;
  columns: Array<IColumn>;
  attendees: Array<IColumn>;
  items: Array<MicrosoftGraph.User>;
  selectedItems: Array<any>;
}

export const initialState = {
  showPanel: false,
  filterText: "",
  filterColumn: "displayName",
  items: [],
  selectedItems: [],
  columns: [
    {
      key: "displayName",
      name: "Display name",
      fieldName: "displayName",
      minWidth: 50,
      maxWidth: 100,
      isResizable: true
    },

    {
      key: "division",
      name: "Division",
      fieldName: "onPremisesExtensionAttributes.extensionAttribute1",
      minWidth: 50,
      maxWidth: 100,
      isResizable: true
    },

    {
      key: "department",
      name: "Department",
      fieldName: "department",
      minWidth: 50,
      maxWidth: 100,
      isResizable: true
    },
    {
      key: "companyName",
      name: "Company",
      fieldName: "companyName",
      minWidth: 50,
      maxWidth: 100,
      isResizable: true
    }
  ],
  attendees: [
    {
      key: "displayName",
      name: "Display name",
      fieldName: "displayName",
      minWidth: 50,
      maxWidth: 100,
      isResizable: true
    }
  ]
};
