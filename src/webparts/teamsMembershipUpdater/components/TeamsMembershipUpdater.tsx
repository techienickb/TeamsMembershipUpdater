import * as React from 'react';
import styles from './TeamsMembershipUpdater.module.scss';
import { ITeamsMembershipUpdaterProps } from './ITeamsMembershipUpdaterProps';
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
import { ProgressIndicator } from 'office-ui-fabric-react/lib/ProgressIndicator';
import { Separator } from 'office-ui-fabric-react/lib/Separator';
import { PrimaryButton } from 'office-ui-fabric-react';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Text } from 'office-ui-fabric-react/lib/Text';
import { ITeamsMembershipUpdaterWebPartProps } from '../TeamsMembershipUpdaterWebPart';
import { CSVReader } from 'react-papaparse';
import { MSGraphClient } from '@microsoft/sp-http';


//Providers.globalProvider = new MsalProvider({ clientId: '3dd6b876-66b0-4ac2-86b3-0ed32bf07d3e' });

export enum Stage {
  LoadingTeams,
  CheckingOwnership,
  LoadingCurrentMembers,
  ComparingMembers,
  RemovingOrphendMembers,
  AddingNewMembers,
  LoggingDone,
  Done
}

export interface ITeamsMembershipUpdaterState {
  items: IDropdownOption[];
  selectionDetails: IDropdownOption;
  csvdata: any[];
  csvcolumns: IColumn[];
  csvSelected: IDropdownOption;
  csvItems: IDropdownOption[];
  me: string;
  groupOwners: string[];
  groupMembers: string[];
  stage: Stage;
}

export default class TeamsMembershipUpdater extends React.Component<ITeamsMembershipUpdaterProps, ITeamsMembershipUpdaterState> {
  private _datacolumns: IColumn[];
  private _data: null;

  constructor(props: ITeamsMembershipUpdaterWebPartProps) {
    super(props);

    this.state = {
      items: props.items,
      selectionDetails: null,
      csvdata: null,
      csvcolumns: [],
      csvSelected: null,
      csvItems: [],
      me: null,
      groupOwners: [],
      groupMembers: [],
      stage: Stage.LoadingTeams
    };
  }


  public handleOnDrop = (data) => {
    var h = data[0].meta.fields;
    this._data = data.map(r => { return r.data; });
    this._datacolumns = h.map(r => { return { key: r.replace(' ', ''), name: r, fieldName: r, isResizable: true }; });
    this.setState({...this.state, csvcolumns: this._datacolumns, csvdata: this._data, csvItems: h.map(r => { return  { key: r.replace(' ', ''), text: r };}) });
    //this._datacolumns = h.map(r => { return { key: r[0]  }})
    //this.setState({ ...this.state, selectedColumn: null, loaded: true, download: false, header: h, rows: d1, attributes: !(h.includes('attribute2') && h.includes('attribute3') && h.includes('office') && h.includes('displayName') && h.includes('department')) });
  }

  public handleOnError = (err, file, inputElem, reason) => {
    console.error(err);
  }

  public handleOnRemoveFile = (data) => {
    this._data = null;
    this.setState({...this.state, csvdata: null });
  }

  public onChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    if (!item) return;
    this.setState({ ...this.state, stage: Stage.CheckingOwnership});
    this.props.context.msGraphClientFactory.getClient().then((client: MSGraphClient): void => {
      client.api(`groups/${item.key}/owners`).version("v1.0").get((err, res) => {
        if (err) {
          console.error(err);
          return;
        }
        let _owners: Array<string> = new Array<string>();
        let b: boolean = false;
        b = res.value.forEach(element => {
          _owners.push(element.userPrincipalName);
          if (element.userPrincipalName == this.state.me) b = true;
        });
        if (b == true) this.setState({ ...this.state, selectionDetails: item, groupOwners: _owners });
        else alert("You are not an owner of this group, select another");
        this.setState({ ...this.state, stage: Stage.Done});
      });
    });
  }//2fad99b2-6b0b-4282-9df9-2d5f53db3e22

  public onEmailChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    this.setState({ ...this.state, csvSelected: item });
  }

  public onRun = (e) => {
    e.preventDefault();
    this.setState({ ...this.state, stage: Stage.LoadingCurrentMembers});
    this.props.context.msGraphClientFactory.getClient().then((client: MSGraphClient): void => {
      client.api(`groups/${this.state.selectionDetails.key}/members`).version("v1.0").get((err, res) => {
        if (err) {
          console.error(err);
          return;
        }
        let _members: Array<string> = res.value.map(element => {
          return element.userPrincipalName;
        });
        this.setState({ ...this.state, groupMembers: _members, stage: Stage.ComparingMembers });

        let _delete: Array<string> = new Array<string>();

        //Loop through _members comparing if user is in csv/owners or not, add to _delete, remove from _members arrays
        //update state and call the graph api to remove users


        //will be called inside the next api call completion

        let _add: Array<string> = new Array<string>();

        //Loop through csv, find missing emails in _members and add to _add array
        //update the state and call the graph api to add users


        //Finally log with a Sharepoint list this exection and result, maybe


      });
    });
  }
  
  public componentDidMount(): void {
    this.props.context.msGraphClientFactory.getClient().then((client: MSGraphClient): void => {
      client.api("me/joinedTeams").version("v1.0").get((err, res) => {
       if (err) {
         console.error(err);
         return;
       }
       var teams: Array<IDropdownOption> = res.value.map((item: any) => {
         return { key: item.id, text: item.displayName };
       });
       this.setState({...this.state, items: teams, stage: Stage.Done });
      });
      client.api("me").version("v1.0").get((err, res) => {
        if (err) {
          console.error(err);
          return;
        }
        this.setState({...this.state, me: res.value.userPrincipalName });
       });
   });

  }

  public render(): React.ReactElement<ITeamsMembershipUpdaterProps> {
    const { items, selectionDetails, csvItems, csvSelected, csvdata, csvcolumns, stage } = this.state;
    return (
      <div className={ styles.teamsMembershipUpdater }>
        <div className={ styles.container }>
          <Text variant="xLarge">{this.props.description}</Text>
          {stage == Stage.LoadingTeams && <ProgressIndicator label="Loading Teams" description="Loading the teams you are a member of" /> }
          {stage == Stage.CheckingOwnership && <ProgressIndicator label="Checking Team Ownership" description="Checking to make sure you are an owner of this team" /> }
          {stage == Stage.LoadingCurrentMembers && <ProgressIndicator label="Loading Current Members" description="Generating a list of current members" /> }
          {stage == Stage.ComparingMembers && <ProgressIndicator label="Comparing Current Members" description="Comparing the current members with the csv file" /> }
          {stage == Stage.RemovingOrphendMembers && <ProgressIndicator label="Removing Orphend Members" description="Removing members who are not owners or in the csv file (orphend)" /> }
          {stage == Stage.AddingNewMembers && <ProgressIndicator label="Adding New Members" description="Adding members who are new in the csv file" /> }
          {stage == Stage.LoggingDone && <ProgressIndicator label="Logging this request" description="Logging this request for stats purposes" /> }
          <Dropdown label="1. Select the Team (you need to be an owner, it will be checked)" selectedKey={selectionDetails ? selectionDetails.key : undefined}
            onChange={this.onChange}
            placeholder="Select an option"
            options={items} disabled={items.length == 0}
            />
          <Dropdown label="2. Select the Email Addresss Column" selectedKey={csvSelected ? csvSelected.key : undefined}
          onChange={this.onEmailChange}
          placeholder="Select an option"
          options={csvItems} disabled={!csvdata}
          />
          <PrimaryButton text="3. Update Membership" onClick={this.onRun} allowDisabledFocus disabled={!csvdata || items.length == 0 || stage != Stage.Done} />
          
          <Separator>CSV File</Separator>
          <CSVReader onDrop={this.handleOnDrop} onError={this.handleOnError} addRemoveButton config={{ header: true, skipEmptyLines: true }} onRemoveFile={this.handleOnRemoveFile}><span>Drop CSV file here or click to upload.</span></CSVReader>

          {csvdata && <DetailsList
              items={csvdata}
              columns={csvcolumns}
              setKey="set"
              layoutMode={DetailsListLayoutMode.justified}
              selectionMode={SelectionMode.none}
            />}
        </div>
      </div>
    );
  }  
}
