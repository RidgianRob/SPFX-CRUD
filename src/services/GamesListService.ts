import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions
} from '@microsoft/sp-http';

import { IGamesListItem } from '../models';
import SphttpclientWebPart from '../webparts/sphttpclient/SphttpclientWebPart';

const LIST_API_ENDPOINT: string = `/_api/web/lists/getbytitle('Games List')`;
const SELECT_QUERY: string = `$select=Id,Title,platform,datePurchased,dateLastPlayed,comments`;
const QUERY_ORDER_MAX: string = `&$orderby=Id desc&$top=1`;

export class GamesListService {

  private _spHttpOptions: any = {
    getNoMetadata: <ISPHttpClientOptions> {
      headers: { 'ACCEPT': 'application/json; odata.metadata=none' }
    },
    getFullMetadata: <ISPHttpClientOptions> {
      headers: { 'ACCEPT': 'application/json; odata.metadata=full' }
    },
    postNoMetadata: <ISPHttpClientOptions> {
      headers: {
        'ACCEPT': 'application/json; odata.metadata=none',
        'CONTENT-TYPE' : 'application/json'
      }
    },
    updateNoMetadata: <ISPHttpClientOptions> {
      headers: {
        'ACCEPT': 'application/json; odata.metadata=none',
        'CONTENT-TYPE': 'application/json',
        'X-HTTP-Method': 'MERGE'
      }
    },
    deleteNoMetadata: <ISPHttpClientOptions> {
      headers: {
        'ACCEPT': 'application/json; odata.metadata=none',
        'X-HTTP-Method': 'DELETE'
      }
    }
  };

  constructor(private siteAbsUrl: string, private client: SPHttpClient) {}

  public getGames(): Promise<IGamesListItem[]> {
    let promise: Promise<IGamesListItem[]> = new Promise<IGamesListItem[]>((resolve, reject) => {
      let query: string = `${this.siteAbsUrl}${LIST_API_ENDPOINT}/items?${SELECT_QUERY}`;

      this.client.get(
        query,
        SPHttpClient.configurations.v1,
        this._spHttpOptions.getNoMetadata
      )
        .then((response: SPHttpClientResponse): Promise<{value: IGamesListItem[]}> => {
          return response.json();
        })
        .then((response: { value: IGamesListItem[] }) => {
          resolve(response.value);
        })
        .catch((error: any) => {
          reject(error);
        });
    });

    return promise;
  }

  public getGame(introSourceId: number): Promise<IGamesListItem> {
    let promise: Promise<IGamesListItem> = new Promise<IGamesListItem>((resolve, reject) => {
      let query: string = `${this.siteAbsUrl}${LIST_API_ENDPOINT}/items(${introSourceId})?${SELECT_QUERY}`;
      this.client.get(
        query,
        SPHttpClient.configurations.v1,
        this._spHttpOptions.getFullMetadata
      )
        .then((response: SPHttpClientResponse): Promise<IGamesListItem> => {
          return response.json();
        })
        .then((response: IGamesListItem) => {
          resolve(response);
        })
        .catch((error: any) => {
          reject(error);
        });
    });

    return promise;
  }

  public getLastGame(): Promise<IGamesListItem> {
    let promise: Promise<IGamesListItem> = new Promise<IGamesListItem>((resolve, reject) => {
      let query: string = `${this.siteAbsUrl}${LIST_API_ENDPOINT}/items?${SELECT_QUERY}${QUERY_ORDER_MAX}`;
      this.client.get(
        query,
        SPHttpClient.configurations.v1,
        this._spHttpOptions.getFullMetadata
      )
        .then((response: SPHttpClientResponse): Promise<any> => {
          return response.json();
        })
        .then((response: any) => {
          resolve(response.value[0]);
        })
        .catch((error: any) => {
          reject(error);
        });
    });

    return promise;
  }

  public _getItemEntityType(): Promise<string> {
    let promise: Promise<string> = new Promise<string>((resolve, reject) => {
      this.client.get(`${this.siteAbsUrl}${LIST_API_ENDPOINT}?$select=ListItemEntityTypeFullName`,
        SPHttpClient.configurations.v1,
        this._spHttpOptions.getNoMetadata
      )
        .then((response: SPHttpClientResponse): Promise<{ ListItemEntityTypeFullName: string}> => {
          return response.json();
        })
        .then((response: { ListItemEntityTypeFullName: string }): void => {
          resolve(response.ListItemEntityTypeFullName);
        })
        .catch((error: any) => {
          reject(error);
        });
    });

    return promise;
  }

  public createGame(newGame: IGamesListItem): Promise<void> {
    let promise: Promise<void> = new Promise<void>((resolve, reject) => {
      this._getItemEntityType()
        .then((spEntityType: string) => {
          // create list item
          let newListItem: IGamesListItem = newGame;

          // add SP required metadata
          newListItem['@odata.type'] = spEntityType;

          // build request
          let requestDetails: any = this._spHttpOptions.postNoMetadata;
          requestDetails.body = JSON.stringify(newListItem);

          // create item
          return this.client.post(`${this.siteAbsUrl}${LIST_API_ENDPOINT}/items`,
            SPHttpClient.configurations.v1,
            requestDetails
          );
        })
        .then((response: SPHttpClientResponse): Promise<IGamesListItem> => {
          return response.json();
        })
        .then((newSPListItem: IGamesListItem): void => {
          resolve();
        })
        .catch((error: any) => {
          reject(error);
        });
    });

    return promise;
  }

  public updateGame(gameToUpdate: IGamesListItem): Promise<void> {
    let promise: Promise<void> = new Promise<void>((resolve, reject) => {
      let requestDetails: any = this._spHttpOptions.updateNoMetadata;

      requestDetails.headers['IF-MATCH'] = gameToUpdate['@odata.etag'];
      requestDetails.body = JSON.stringify(gameToUpdate);

      this.client.post(`${this.siteAbsUrl}${LIST_API_ENDPOINT}/items(${gameToUpdate.Id})`,
        SPHttpClient.configurations.v1,
        requestDetails
      )
      .then(() => {
        resolve();
      });
    });

    return promise;
  }

  public deleteGame(gameToDelete: IGamesListItem): Promise<void> {
    let promise: Promise<void> = new Promise<void>((resolve, reject) => {
      let requestDetails: any = this._spHttpOptions.deleteNoMetadata;

      // Check to make sure we're updating the latest version
      requestDetails.headers['IF-MATCH'] = gameToDelete['@odata.etag'];
      requestDetails.body = JSON.stringify(gameToDelete);

      this.client.post(`${this.siteAbsUrl}${LIST_API_ENDPOINT}/items(${gameToDelete.Id})`,
        SPHttpClient.configurations.v1,
        requestDetails
      )
      .then(() => {
        resolve();
      });
    });

    return promise;
  }

}