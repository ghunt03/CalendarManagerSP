import * as React from 'react';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { IInviteScreenProps } from './IInviteScreenProps';
import { IInviteScreenState, initialState } from './IInviteScreenState';
import { getUsers } from '../../api/getUsers';
import { IRenderFunction, createRef } from 'office-ui-fabric-react/lib/Utilities';
import {
  DetailsList,
  DetailsListLayoutMode,
  IDetailsHeaderProps,
  IColumn,
  CheckboxVisibility,
  SelectionMode,
  Selection,
  PrimaryButton,
  DefaultButton,
  SearchBox,
  CommandBar
} from 'office-ui-fabric-react';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { ScrollablePane, IScrollablePane } from 'office-ui-fabric-react/lib/ScrollablePane';
import { Sticky, StickyPositionType } from 'office-ui-fabric-react/lib/Sticky';
import { filterData, sortData, getUniqueValues } from '../../utils/detailListHelpers';

export default class InviteScreen extends React.Component<IInviteScreenProps, IInviteScreenState> {
  private _selection: Selection;
  private _removeSelection: Selection;
  private _items: Array<MicrosoftGraph.User>;
  private _scrollablePane = createRef<IScrollablePane>();
  constructor(props: IInviteScreenProps) {
    super(props);
    this.state = initialState;
    this._items = [];
    this.getData();
    this.onFilterChange = this.onFilterChange.bind(this);
    this._selection = new Selection();
    this._removeSelection = new Selection();
  }

  private getData() {
    const { configurationProps } = this.props;
    getUsers(configurationProps).then(items => {
      this._items = items;
      let companies = getUniqueValues(items, 'companyName');
      this.setState({ items });
    });
  }

  private addCommandBarItems = [
    {
      key: 'search',
      onRender: () => (
        <SearchBox
          placeholder="Search"
          className="searchBox"
          onChanged={this.onFilterChange.bind(this)}
        />
      )
    },
    {
      key: 'newItem',
      name: 'Add Attendee',
      onClick: () => this._onAddItems(),
      iconProps: { iconName: 'AddFriend' }
    }
  ];

  private removeCommandBarItems = [
    {
      key: 'newItem',
      name: 'Remove Attendees',
      onClick: () => this._onRemoveItems(),
      iconProps: { iconName: 'Delete' }
    }
  ];

  private _onRemoveItems = () => {
    let itemsToRemove = this._removeSelection.getSelection();
    let { selectedItems } = this.state;
    console.log(itemsToRemove);
    let emails = getUniqueValues(itemsToRemove, 'mail');
    this.setState({
      selectedItems: selectedItems.filter(i => !emails.includes(i['mail']))
    });
    this._removeSelection.setAllSelected(false);
  }

  private _onAddItems = () => {
    const selectedItems = this._selection.getSelection();
    this.setState({
      selectedItems: this.state.selectedItems.concat(selectedItems)
    });
    this._selection.setAllSelected(false);
  }

  private _onSendInvite = () => {
    const { onSendInvite } = this.props;
    const { selectedItems } = this.state;
    onSendInvite(selectedItems);
    this.setState({ showPanel: false });
  }

  public renderPanelContent() {
    const { items, columns, selectedItems, attendees } = this.state;
    return (
      <div className="ms-Grid">
        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm8 ms-md8 ms-lg8">
            <div
              style={{
                height: '70vh',
                position: 'relative'
              }}
            >
              <ScrollablePane componentRef={this._scrollablePane}>
                <Sticky stickyPosition={StickyPositionType.Header}>
                  <CommandBar items={this.addCommandBarItems} className="commandbar" />
                </Sticky>
                <DetailsList
                  items={items}
                  columns={columns}
                  selection={this._selection}
                  onColumnHeaderClick={this._onColumnClick}
                  setKey="mail"
                  checkboxVisibility={CheckboxVisibility.always}
                  selectionMode={SelectionMode.multiple}
                  layoutMode={DetailsListLayoutMode.justified}
                  onRenderDetailsHeader={this._onRenderDetailsHeader}
                  onRenderItemColumn={this._renderItemColumn}
                />
              </ScrollablePane>
            </div>
          </div>
          <div className="ms-Grid-col ms-sm4 ms-md4 ms-lg4">
            <div
              style={{
                height: '70vh',
                position: 'relative'
              }}
            >
              <ScrollablePane componentRef={this._scrollablePane}>
                <Sticky stickyPosition={StickyPositionType.Header}>
                  <CommandBar items={this.removeCommandBarItems} className="commandbar" />
                </Sticky>
                <DetailsList
                  items={selectedItems}
                  selection={this._removeSelection}
                  columns={attendees}
                  setKey="mail"
                  checkboxVisibility={CheckboxVisibility.always}
                  selectionMode={SelectionMode.multiple}
                  layoutMode={DetailsListLayoutMode.justified}
                  onRenderDetailsHeader={this._onRenderDetailsHeader}
                  onRenderItemColumn={this._renderItemColumn}
                />
              </ScrollablePane>
            </div>
          </div>
        </div>
      </div>
    );
  }

  public render() {
    const { showPanel } = this.state;
    return (
      <div>
        <PrimaryButton text="Invite" title="Invite" onClick={this._onOpenPanel} />
        <Panel
          isOpen={showPanel}
          onDismiss={this._onClosePanel}
          type={PanelType.custom}
          customWidth="1000px"
          headerText="Invite"
          onRenderFooterContent={this._onRenderPanelFooterContent}
        >
          <div>{this.renderPanelContent()}</div>
        </Panel>
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
      case 'division':
        return <span>{item.onPremisesExtensionAttributes.extensionAttribute1}</span>;
      default:
        return <span>{fieldContent}</span>;
    }
  }

  private onFilterChange = value => {
    const { filterColumn } = this.state;
    this.setState({
      items: filterData(value, this._items, filterColumn),
      filterText: value
    });
  }

  private _onColumnClick = (event: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const { columns, items } = this.state;
    let sorted = sortData(columns, column, items);
    this.setState({
      items: sorted.items,
      columns: sorted.columns
    });
  }

  private _onClosePanel = (): void => {
    this.setState({ showPanel: false });
  }

  private _onOpenPanel = (): void => {
    this.setState({ showPanel: true });
  }

  private _onRenderPanelFooterContent = (): JSX.Element => {
    return (
      <div>
        <PrimaryButton onClick={this._onSendInvite.bind(this)} style={{ marginRight: '8px' }}>
          Send Invite
        </PrimaryButton>
        <DefaultButton onClick={this._onClosePanel}>Cancel</DefaultButton>
      </div>
    );
  }
}
