import { MSGraphClient } from "@microsoft/sp-http";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";

function getEventData(context, groupId) {
  return new Promise((resolve, reject)  =>{
    context.msGraphClientFactory.getClient().then(client => {
      client
        .api(`groups/${groupId}/events`)
        .get((err, res) => {
          if (err) {
            reject(err);
          }
          if (res === null) {
            reject("No results found");
          }
          let users: [MicrosoftGraph.Event] = res.value;
          resolve(users);
        });
    });
  });
}

export function getEvents({ context, groupId }): Promise<any[]> {
  return getEventData(context, groupId)
    .then(data => {
      console.log(data);
      return data;
    })
    .catch(err => err);
}
