import { MSGraphClient } from "@microsoft/sp-http";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";

function getUserData(context) {
  return new Promise((resolve, reject) => {
    context.msGraphClientFactory.getClient().then(client => {
      client
        .api("users")
        .version("beta")
        .filter("country eq 'Australia'")
        .get((err, res) => {
          if (err) {
            reject(err);
          }
          if (res === null) {
            reject("No results found");
          }
          let users: [MicrosoftGraph.User] = res.value;
          resolve(users);
        });
    });
  });
}

export function getUsers({ context }): Promise<any[]> {
  return getUserData(context)
    .then(data => {
      console.log(data);
      return data;
    })
    .catch(err => err);
}
