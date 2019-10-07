import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SphttpclientWebPart.module.scss';
import * as strings from 'SphttpclientWebPartStrings';

// Interfaces (models)
import { IGamesListItem } from '../../models';

// Services
import { GamesListService } from '../../services';

export interface ISphttpclientWebPartProps {
  description: string;
}

export default class SphttpclientWebPart extends BaseClientSideWebPart<ISphttpclientWebPartProps> {

  private gamesListService: GamesListService;

  private gamesDetailElement: HTMLElement;

  protected onInit(): Promise<any> {
    this.gamesListService = new GamesListService (
      this.context.pageContext.web.absoluteUrl,
      this.context.spHttpClient
    );

    return Promise.resolve();
  }

  public render(): void {
    if (!this.renderedOnce) {
      this.domElement.innerHTML = `
        <div class="${ styles.sphttpclient }">
          <div class="${ styles.container }">
            <div class="${ styles.row }">
              <div class="${ styles.column }">
                <span class="${ styles.title }">Games List</span>
                <p class="${ styles.subTitle }">Demonstrating SharePoint HTTP Client.</p>
                <button id="getGames" class="${styles.button}">Get Games</button>
                <button id="getGame" class="${styles.button}">Get Game</button>
                <button id="getLastGame" class="${styles.button}">Get Last Game</button>
                <button id="createGame" class="${styles.button}">Create Game</button>
                <button id="updateGame" class="${styles.button}">Update Game</button>
                <button id="deleteGame" class="${styles.button}">Delete Game</button>
                <div id="claasIntroSources"></div>
              </div>
            </div>
          </div>
        </div>`;

        this.gamesDetailElement = document.getElementById('claasIntroSources');

        document.getElementById('getGames')
          .addEventListener('click', () => {
            this._getGames();
          });

          document.getElementById('getGame')
            .addEventListener('click', () => {
              this._getGame();
          });

          document.getElementById('getLastGame')
            .addEventListener('click', () => {
              this._getLastGame();
          });

          document.getElementById('createGame')
            .addEventListener('click', () => {
              this._createGame();
          });

          document.getElementById('updateGame')
            .addEventListener('click', () => {
              this._updateLastGame();
          }); 

          document.getElementById('deleteGame')
            .addEventListener('click', () => {
              this._deleteLastGame();
          }); 

    }
  }

  private _renderGames(element: HTMLElement, game: IGamesListItem[]): void {
    let gamesList: string = '';

    if (game && game.length && game.length > 0) {
      game.forEach((game: IGamesListItem) => {
        gamesList = gamesList + `<tr>
          <td>${game.Id}</td>
          <td>${game.Title}</td>
          <td>${game.platform}</td>
          <td>${game.datePurchased}</td>
          <td>${game.dateLastPlayed}</td>
          <td>${game.comments}</td>
        </tr>`;
      });
    }

    element.innerHTML = `<table border=1>
      <tr>
        <th>Id</th>
        <th>Title</th>
        <th>Platform</th>
        <th>Date Purchased</th>
        <th>Date Last Played</th>
        <th>Comments</th>
        <tbody>${gamesList}</tbody>
      </tr>
    </table>`;
  }

  private _getGames(): void {
    this.gamesListService.getGames()
      .then((games: IGamesListItem[]) => {
        this._renderGames(this.gamesDetailElement, games);
      });

  }

  private _getGame(): void {
    this.gamesListService.getGame(57)
      .then((games: IGamesListItem) => {
        this._renderGames(this.gamesDetailElement, [games]);
      });

  }

  private _getLastGame(): void {
    this.gamesListService.getLastGame()
      .then((games: IGamesListItem) => {
        this._renderGames(this.gamesDetailElement, [games]);
      });

  }

  private _createGame(): void {
    const newGame: IGamesListItem = <IGamesListItem> {
      Title: 'Half Life 3',
      platform: 'PS5',
      datePurchased: '2057-08-30',
      dateLastPlayed: '2057-08-30',
      comments: 'Never gonna happen...'
    };

    this._renderGames(this.gamesDetailElement, null);

    this.gamesListService.createGame(newGame)
      .then(() => {
        this._getGames();
      });

  }

  private _updateLastGame(): void {
    this.gamesListService.getLastGame()
      .then((games: IGamesListItem) => {
        games.dateLastPlayed = new Date().toISOString();
          return this.gamesListService.updateGame(games);
      });
  }

  private _deleteLastGame(): void {
    this.gamesListService.getLastGame()
      .then((games: IGamesListItem) => {
          return this.gamesListService.deleteGame(games);
      })
      .then(() => {
        this._getGames();
      });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
