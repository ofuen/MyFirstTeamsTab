import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './MyFirstTeamsTabWebPart.module.scss';
import * as strings from 'MyFirstTeamsTabWebPartStrings';

export interface IMyFirstTeamsTabWebPartProps {
  description: string;
}

export default class MyFirstTeamsTabWebPart extends BaseClientSideWebPart<IMyFirstTeamsTabWebPartProps> {

  // This variable has been added
  private _teamsContext: microsoftTeams.Context;

  protected onInit(): Promise<any> {
    let retVal: Promise<any> = Promise.resolve();
    if (this.context.microsoftTeams) {
      retVal = new Promise((resolve, _reject) => {
        this.context.microsoftTeams.getContext(context => {
          this._teamsContext = context;
          resolve();
        });
      });
    }
    return retVal;
  }

  public render(): void {

    let title: string = '';
    let subTitle: string = '';
    let siteTabTitle: string = '';

    if (this._teamsContext) {
      // We have Teams context for the web part
      title = "Welcome to Teams!";
      subTitle = "Building custom enterprise tabs for your business.";
      siteTabTitle = "We are in the context of following Teams team: " + this._teamsContext.teamName;
    }
    else {
      // We are rendered in normal SharePoint context
      title = "Welcome to SharePoint!";
      subTitle = "Customize SharePoint experiences using Web Parts.";
      siteTabTitle = "We are in the context of following SharePoint site: " + this.context.pageContext.web.title;
    }

    this.domElement.innerHTML = `
    <div class="${ styles.myFirstTeamsTab}">
      <div class="${ styles.container}">
        <div class="${ styles.row}">
          <div class="${ styles.column}">
            <span class="${ styles.title}">${title}</span>
            <p class="${ styles.subTitle}">${subTitle}</p>
            <p class="${ styles.description}">${siteTabTitle}</p>
            <p class="${ styles.description}">Description property value - ${escape(this.properties.description)}</p>
            <a href="https://aka.ms/spfx" class="${ styles.button}">
              <span class="${ styles.label}">Learn more</span>
            </a>
          </div>
        </div>
      </div>
    </div>`;
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
