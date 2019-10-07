export interface IGamesListItem {
  ['@odata.type']?: string;
  ['@odata.etag']?: string;
  Id: number;
  Title: string;
  platform: string;
  datePurchased: string;
  dateLastPlayed: string;
  comments: string;
}