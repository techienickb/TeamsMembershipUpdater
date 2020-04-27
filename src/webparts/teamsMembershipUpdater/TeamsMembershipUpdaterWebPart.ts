import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { Guid } from '@microsoft/sp-core-library';
import * as strings from 'TeamsMembershipUpdaterWebPartStrings';
import TeamsMembershipUpdater from './components/TeamsMembershipUpdater';
import { ITeamsMembershipUpdaterProps } from './components/ITeamsMembershipUpdaterProps';
import { sp } from "@pnp/sp/presets/all";
import { graph } from "@pnp/graph/presets/all";
import { IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';

export interface ITeamsMembershipUpdaterWebPartProps {
  description: string;
  items: IDropdownOption[];
}


export default class TeamsMembershipUpdaterWebPart extends BaseClientSideWebPart <ITeamsMembershipUpdaterWebPartProps> {

  public teams = [];

  public render(): void {
    const element: React.ReactElement<ITeamsMembershipUpdaterProps> = React.createElement(
      TeamsMembershipUpdater,
      {
        description: this.properties.description,
        items: []
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
  
    // other init code may be present
  
    sp.setup(this.context);
    graph.setup({spfxContext: this.context});

    //this.teams = await graph.me.joinedTeams();
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
