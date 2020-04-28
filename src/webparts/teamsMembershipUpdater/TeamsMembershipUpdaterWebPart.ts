import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import * as strings from 'TeamsMembershipUpdaterWebPartStrings';
import TeamsMembershipUpdater from './components/TeamsMembershipUpdater';
import { ITeamsMembershipUpdaterProps } from './components/ITeamsMembershipUpdaterProps';
import { IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';

export interface ITeamsMembershipUpdaterWebPartProps {
  description: string;
  items: IDropdownOption[];
  context: WebPartContext;
}


export default class TeamsMembershipUpdaterWebPart extends BaseClientSideWebPart <ITeamsMembershipUpdaterWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ITeamsMembershipUpdaterProps> = React.createElement(
      TeamsMembershipUpdater,
      {
        description: this.properties.description,
        items: [],
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
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
