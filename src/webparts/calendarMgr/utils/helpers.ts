import * as moment from "moment";

export function formatDate(date): string {
  return moment(date).format("YYYY-MM-DD");
}

export function formatDateTime(date): string {
  return moment.utc(date['dateTime']).local().format("YYYY-MM-DD hh:mm A");
}
