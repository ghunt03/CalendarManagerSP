import { MSGraphClient } from "@microsoft/sp-http";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";

interface IGraphBatchRequest {
  id: string;
  method: string;
  url: string;
  body: any;
  headers: any;
}

interface IGraphBatchBody {
  requests: IGraphBatchRequest[];
}

const patchEvent = (context, groupId, selectedEvents, selectedUsers) => {
  return new Promise((resolve, reject) => {
    if (selectedUsers.length < 1) {
      reject("Error: Missing invitees");
    } else if (selectedEvents.length < 1) {
      reject("Error: Missing selected events");
    } else {
      // prep data
      let requestBody: IGraphBatchBody = { requests: [] };

      let invitees = selectedUsers.map(u => {
        return {
          type: "required",
          emailAddress: {
            name: u.displayName,
            address: u.mail
          }
        };
      });

      let body = {
        attendees: invitees
      };
      selectedEvents.forEach((event, i) => {
        let requestUrl: string = `/groups/${groupId}/events/${event.id}`;
        requestBody.requests.push({
          id: i.toString(),
          method: "PATCH",
          url: requestUrl,
          body: body,
          headers: {
            "Content-Type": "application/json"
          }
        });
      });
      console.log(requestBody);
      context.msGraphClientFactory.getClient().then(client => {
        client.api(`$batch`).post(requestBody, (err, res) => {
          if (err) {
            console.error(err);
            reject(err);
            return;
          }
          resolve(res);
        });
      });
    }
  });
};

export const sendInvites = (
  { context, groupId },
  selectedEvents,
  selectedUsers
): Promise<any[]> => {
  return patchEvent(context, groupId, selectedEvents, selectedUsers)
    .then(data => data)
    .catch(err => err);
};
