import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneButton,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SpfxMgtWebPart.module.scss';
import * as strings from 'SpfxMgtWebPartStrings';
import { MgtPeoplePicker, MgtTeamsChannelPicker, Providers, SharePointProvider } from '@microsoft/mgt';

export interface ISpfxMgtWebPartProps {
  description: string;
}

export default class SpfxMgtWebPart extends BaseClientSideWebPart<ISpfxMgtWebPartProps> {

  protected async onInit() {
    Providers.globalProvider = new SharePointProvider(this.context);
  }
  public render(): void {
    this.domElement.innerHTML = `
    <mgt-person person-query="me" show-name show-email></mgt-person>
    <mgt-agenda></mgt-agenda>
    <mgt-people></mgt-people>
    <mgt-person-card person-query="me"></mgt-person-card>
    <mgt-todo></mgt-todo>
    <mgt-teams-channel-picker></mgt-teams-channel-picker>
    <button class="button">Get SelectedChannel</button>
    <div class="output"></div>  
   `;
  
  }

  
    }
  //protected get dataVersion(): Version {
    //return Version.parse('1.0');
 // }

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
                }),
                PropertyPaneButton('button',{label: strings.DescriptionFieldLabel})
              ]
            }
          ]
        }
      ]
    };
  }
}
