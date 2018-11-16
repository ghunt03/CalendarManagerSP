import * as React from "react";
import styles from "./CalendarMgr.module.scss";
import { ICalendarMgrProps } from "./ICalendarMgrProps";
import { ICalendarMgrState, initialState } from "./ICalendarMgrState";
import { sendInvites } from "../api/sendInvites";
import InviteScreen from "./InviteScreen/InviteScreen";
import EventList from "./EventList/EventList";
import { MessageBar, MessageBarType } from "office-ui-fabric-react";

export default class CalendarMgr extends React.Component<
  ICalendarMgrProps,
  ICalendarMgrState
> {
  constructor(props: ICalendarMgrProps, state: ICalendarMgrState) {
    super(props);
    this.state = initialState;
  }
  public onSelectEvents = selectedEvents => {
    this.setState({ selectedEvents });
  }

  public onSelectUsers = selectedUsers => {
    this.setState({ selectedUsers });
  }

  public onCommandBarPress = action => {
    if (action === "OpenPanel") {
      this.setState({ showPanel: true });
    }
  }

  public sendInvite = invitees => {
    const { selectedEvents } = this.state;
    if (invitees.length < 1) {
      this.setState({
        message: "Need to select at least one attendee before sending invite",
        messageType: MessageBarType.error
      });
    } else {
      sendInvites(this.props, selectedEvents, invitees)
        .then(res => {
          this.setState({
            selectedEvents: [],
            showPanel: false,
            message: "Appointments have been sent",
            messageType: MessageBarType.success
          });
        })
        .catch(err => {
          this.setState({
            message: err,
            messageType: MessageBarType.error
          });
        });
    }
  }

  private displayMessage = () => {
    const { message, messageType } = this.state;
    if (messageType === MessageBarType.success) {
      return (
        <MessageBar messageBarType={messageType}>
          <div>{message}</div>
        </MessageBar>
      );
    } else {
      return <MessageBar messageBarType={messageType}>{message}</MessageBar>;
    }
  }

  public render(): React.ReactElement<ICalendarMgrProps> {
    const { message } = this.state;
    return (
      <div className={styles.calendarMgr}>
        <div className={styles.container}>
          <div className="ms-Grid">
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                {message !== "" ? this.displayMessage() : null}
              </div>
            </div>
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                <EventList
                  onSelectItems={this.onSelectEvents}
                  configurationProps={this.props}
                />
              </div>
            </div>
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                <InviteScreen
                  configurationProps={this.props}
                  onSendInvite={this.sendInvite}
                />
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
