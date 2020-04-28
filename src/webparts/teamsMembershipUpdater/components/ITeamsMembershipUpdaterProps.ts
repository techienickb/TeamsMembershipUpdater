import { Guid } from '@microsoft/sp-core-library';
import { IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IDetailsListItem {
  key: number;
  name: string;
  id: Guid;
}

export interface ITeamsMembershipUpdaterProps {
  description: string;
  items: IDropdownOption[];
  context: WebPartContext;
}
