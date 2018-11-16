import * as React from "react";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import { formatDateTime } from "../../utils/helpers";
import { getEvents } from "../../api/getEvents";
import { IEventListProps } from "./IEventListProps";
import { IEventListState, initialState } from "./IEventListState";
import {
  IRenderFunction,
  createRef
} from "office-ui-fabric-react/lib/Utilities";
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  CheckboxVisibility,
  SelectionMode,
  Selection,
  SearchBox,
  CommandBar,
  IDetailsHeaderProps
} from "office-ui-fabric-react";
import {
  ScrollablePane,
  IScrollablePane
} from "office-ui-fabric-react/lib/ScrollablePane";
import { Sticky, StickyPositionType } from "office-ui-fabric-react/lib/Sticky";
import { filterData, sortData } from "../../utils/detailListHelpers";
export default class EventList extends React.Component<
  IEventListProps,
  IEventListState
> {
  private _selection: Selection;
  private _items: Array<MicrosoftGraph.Event>;
  private _scrollablePane = createRef<IScrollablePane>();
  constructor(props: IEventListProps) {
    super(props);
    this.state = initialState;
    this._items = [];
    this.getData();
    this.onFilterChange = this.onFilterChange.bind(this);
    this._selection = new Selection({
      onSelectionChanged: () => {
        const { onSelectItems } = this.props;
        const selectedItems = this._selection.getSelection();
        onSelectItems(selectedItems);
      }
    });
  }

  private getData() {
    const { configurationProps } = this.props;
    getEvents(configurationProps).then(items => {
      this._items = items;
      this.setState({ items });
    });
  }

  private addCommandBarItems = [
    {
      key: "search",
      onRender: () => (
        <SearchBox
          placeholder="Search"
          className="searchBox"
          onChanged={this.onFilterChange.bind(this)}
        />
      )
    }
  ];

  public render() {
    const { items, columns } = this.state;
    return (
      <div>
        <div
          style={{
            height: "70vh",
            position: "relative"
          }}
        >
          <ScrollablePane componentRef={this._scrollablePane}>
            <Sticky stickyPosition={StickyPositionType.Header}>
              <CommandBar
                items={this.addCommandBarItems}
                className="commandbar"
              />
            </Sticky>
            <DetailsList
              items={items}
              columns={columns}
              selection={this._selection}
              onColumnHeaderClick={this._onColumnClick}
              setKey="id"
              checkboxVisibility={CheckboxVisibility.always}
              selectionMode={SelectionMode.multiple}
              layoutMode={DetailsListLayoutMode.justified}
              onRenderDetailsHeader={this._onRenderDetailsHeader}
              onRenderItemColumn={this._renderItemColumn}
            />
          </ScrollablePane>
        </div>
      </div>
    );
  }
  private _onRenderDetailsHeader(
    props: IDetailsHeaderProps,
    defaultRender?: IRenderFunction<IDetailsHeaderProps>
  ): JSX.Element {
    return (
      <Sticky stickyPosition={StickyPositionType.Header}>
        {defaultRender!({
          ...props
        })}
      </Sticky>
    );
  }

  private _renderItemColumn(item: any, index: number, column: IColumn) {
    const fieldContent = item[column.fieldName];
    switch (column.key) {
      case "start":
      case "end":
        return <span>{formatDateTime(fieldContent)}</span>;
      case "attendees":
        return <span>{fieldContent.length}</span>;
      default:
        return <span>{fieldContent}</span>;
    }
  }

  private onFilterChange(value) {
    const { filterColumn } = this.state;
    this.setState({
      items: filterData(value, this._items, filterColumn),
      filterText: value
    });
  }

  private _onColumnClick(
    event: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void {
    const { columns, items } = this.state;
    let sorted = sortData(columns, column, items);
    this.setState({
      items: sorted.items,
      columns: sorted.columns
    });
  }
}
