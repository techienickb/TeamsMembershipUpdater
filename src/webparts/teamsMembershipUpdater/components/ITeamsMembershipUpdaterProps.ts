import { Guid } from '@microsoft/sp-core-library';
import { IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';

export interface IDetailsListItem {
  key: number;
  name: string;
  id: Guid;
}

export interface ITeamsMembershipUpdaterProps {
  description: string;
  items: IDropdownOption[];
}
