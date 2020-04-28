import { Guid } from "@microsoft/sp-core-library";

export interface IUserItem {
    displayName: string;
    mail: string;
    userPrincipalName: string;
    id: Guid;
  }

  export interface ITeamItem {
    id: Guid;
    displayName: string;
  }