import * as React from 'react';
import styles from './TeamsMembershipUpdater.module.scss';
import { ITeamsMembershipUpdaterProps } from './ITeamsMembershipUpdaterProps';
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
import { ProgressIndicator } from 'office-ui-fabric-react/lib/ProgressIndicator';
import { Separator } from 'office-ui-fabric-react/lib/Separator';
import { PrimaryButton, MessageBar, MessageBarType, Link } from 'office-ui-fabric-react';
import { List } from 'office-ui-fabric-react/lib/List';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Text } from 'office-ui-fabric-react/lib/Text';
import { ITeamsMembershipUpdaterWebPartProps } from '../TeamsMembershipUpdaterWebPart';
import { CSVReader } from 'react-papaparse';
import { MSGraphClient, SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
//Providers.globalProvider = new MsalProvider({ clientId: '3dd6b876-66b0-4ac2-86b3-0ed32bf07d3e' });

export enum Stage {
  LoadingTeams,
  CheckingOwnership,
  LoadingCurrentMembers,
  ComparingMembers,
  RemovingOrphendMembers,
  AddingNewMembers,
  LoggingDone,
  Done,
  ErrorOwnership
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
  groupMembers: Array<MicrosoftGraph.User>;
  stage: Stage;
  logs: Array<string>;
  errors: Array<string>;
  logurl: string;
}

export default class TeamsMembershipUpdater extends React.Component<ITeamsMembershipUpdaterProps, ITeamsMembershipUpdaterState> {
  private _datacolumns: IColumn[];
  private _data: any[];

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
      stage: Stage.LoadingTeams,
      logs: [],
      errors: [],
      logurl: null
    };
  }

  public addError = (e: string, o: any):void => {
    console.error(e, o);
    let _log: Array<string> = this.state.errors;
    _log.push(e);
    this.setState({...this.state, errors: _log });
  }

  public addLog = (e: string): void => {
    let _log: Array<string> = this.state.logs;
    _log.push(e);
    this.setState({...this.state, logs: _log });
  }

  public handleOnDrop = (data) => {
    var h = data[0].meta.fields;
    this._data = data.map(r => { return r.data; });
    this._datacolumns = h.map(r => { return { key: r.replace(' ', ''), name: r, fieldName: r, isResizable: true }; });
    this.setState({...this.state, csvcolumns: this._datacolumns, csvdata: this._data, csvItems: h.map(r => { return  { key: r.replace(' ', ''), text: r };}), logs: [], errors: [], logurl: null });
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
      this.setState({ ...this.state, stage: Stage.CheckingOwnership, logs: [], errors: [], logurl: null });
      this.props.context.msGraphClientFactory.getClient().then((client: MSGraphClient): void => {
        client.api(`groups/${item.key}/owners`).version("v1.0").get((err, res) => {
          if (err) {
            this.addError(err.message, err);
            return;
          }
          let _owners: Array<string> = new Array<string>();
          let b: boolean = false;
          res.value.forEach(element => {
            _owners.push(element.userPrincipalName);
            if (element.userPrincipalName == this.state.me) b = true; 
          });
          if (b) this.setState({ ...this.state, selectionDetails: item, groupOwners: _owners, stage: Stage.Done });
          else this.setState({ ...this.state, stage: Stage.ErrorOwnership });
        });
      });
  }

  public onEmailChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    this.setState({ ...this.state, csvSelected: item });
  }

  public onRun = (e) => {
    this.setState({ ...this.state, stage: Stage.LoadingCurrentMembers});
    this.props.context.msGraphClientFactory.getClient().then((client: MSGraphClient): void => {
      client.api(`groups/${this.state.selectionDetails.key}/members`).version("v1.0").get(async (err, res) => {
        if (err) {
          this.addError(err.message, err);
          return;
        }

        let _members:Array<MicrosoftGraph.User> = res.value;
        this.setState({ ...this.state, groupMembers: _members, stage: Stage.ComparingMembers });

        this.addLog(`Found ${_members.length} members existing in the group`);

        let _delete: Array<MicrosoftGraph.User> = new Array<MicrosoftGraph.User>();

        _members = _members.filter(m => {
          if (this._data.some(value => value[this.state.csvSelected.text] === m.mail) || this.state.groupOwners.some(value => value === m.userPrincipalName)) return m;
          else { _delete.push(m); this.addLog(`Will delete ${m.mail}`); }
        });

        let req = { requests: Array<any>() };

        this.setState({ ...this.state, stage: Stage.RemovingOrphendMembers });
        _delete.forEach(e1 => {
          req.requests.push({
            id: `${req.requests.length + 1}`,
            method: "DELETE",
            url: `groups/${this.state.selectionDetails.key}/members/${e1.id}/$ref`
          });
        });

        let newMembers: Array<string> = new Array<string>();
        
        this._data.forEach(async e2 => {
          if (_members.some(m => m.mail === e2[this.state.csvSelected.text]) == false) {
            newMembers.push(e2[this.state.csvSelected.text]);
            this.addLog(`Will add ${e2[this.state.csvSelected.text]}`);
          }
        });

        if (req.requests.length > 0) { 
          await client.api("$batch").version("v1.0").post(req, (er, re) => { 
            if (err) { this.addError(err.message, err); return; }
            if (re) re.reponses.forEach(e3 => { if (e3.body.error) this.addError(e3.body.error.toString(), e3.body.error); });
            this.addLog(`Deleting Done`);
            if (newMembers.length == 0) {
              this.Done();
            }
            this.addMembers(newMembers, client);
          });
        } else if (newMembers.length == 0) {
          this.Done();
        } else this.addMembers(newMembers, client);
      });
    });
  }

  public Done = (): void => {
    this.setState({ ...this.state, stage: Stage.LoggingDone });

    this.props.context.spHttpClient.get(`${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Team Membership Update Log')?$select=ListItemEntityTypeFullName`, SPHttpClient.configurations.v1)
    .then((res: SPHttpClientResponse): Promise<{ ListItemEntityTypeFullName: string; }> => {
      return res.json();
    })
    .then((web: {ListItemEntityTypeFullName: string}): void => {
      const p = {
        //"__metadata": { "type": web.ListItemEntityTypeFullName },
        "Title": `${this.state.selectionDetails.text} update ${new Date().toString()}`,
        "Logs" : this.state.logs.join(", \n"),
        "Errors": this.state.errors.join(", \n")
      };

      this.props.context.spHttpClient.post(`${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Team Membership Update Log')/items`, SPHttpClient.configurations.v1, {
        body: JSON.stringify(p)
      }).then((res: SPHttpClientResponse): Promise<{ w: any; }> => {
        return res.json();
      })
      .then((w: any): void => {
        this.setState({ ...this.state, stage: Stage.Done, logurl: 'https://cf.sharepoint.com/sites/cloudservicesteam/Lists/Team Membership Update Log/my.aspx' });
      });
    });
  }

  public addMembers = (newMembers: string[], client: MSGraphClient): void => {
    this.setState({ ...this.state, stage: Stage.AddingNewMembers });
    let req: any = { requests: Array<any>() };
    newMembers.forEach(e => {
      req.requests.push({
        id: `${req.requests.length + 1}`,
        method: "GET",
        url: `users/${e}?$select=id`
      });
    });
    this.addLog(`Getting Object IDs for ${newMembers.length} Members to Add from Graph`);
    client.api("$batch").version("v1.0").post(req, (er, re) => { 
      if (er) { this.addError(er.message, er); return; }
      req.requests = new Array<any>();
      if (re) {
        re.responses.forEach(e => {
          if (e.body.error) this.addError(e.body.error.toString(), e.body.error);
          else { 
            req.requests.push({
              id: `${req.requests.length + 1}`,
              method: "POST",
              url: `groups/${this.state.selectionDetails.key}/members/$ref`,
              headers: { "Content-Type": "application/json" },
              body: { "@odata.id": `https://graph.microsoft.com/v1.0/directoryObjects/${e.body.id}` }
            });
          }
        });
        this.addLog(`Adding ${req.requests.length} Members`);
        client.api("$batch").version("v1.0").post(req, (err, res) => {
          if (err) { this.addError(err.message, err); return;}
          req.requests = new Array<any>();
          if (res) {
            res.responses.forEach(e => {
              if (e.body.error) this.addError(e.body.error.toString(), e.body.error);
            });
            this.addLog("Adding Done");
            this.Done();
          }
          this.addLog("Adding Done");
          this.Done();
        });
      }
    });
  }
  
  public componentDidMount(): void {
    this.props.context.msGraphClientFactory.getClient().then((client: MSGraphClient): void => {
      let req =  { requests: [
        { id: "1", method: "GET", url: "me" },
        { id: "2", method: "GET", url: "me/joinedTeams" }
      ] };
      client.api("$batch").version("v1.0").post(req , (err, res) => {
        if (err) {
          this.addError(err.message, err);
          return;
        }
        let teams: Array<IDropdownOption> = res.responses[1].body.value.map((item: any) => {
          return { key: item.id, text: item.displayName };
        });
        this.setState({...this.state, me: res.responses[0].body.userPrincipalName, items: teams, stage: Stage.Done });

      });
   });

  }

  public render(): React.ReactElement<ITeamsMembershipUpdaterProps> {
    const { items, csvItems, csvdata, csvcolumns, stage, csvSelected, logurl, logs, errors } = this.state;
    return (
      <div className={ styles.teamsMembershipUpdater }>
        <div className={ styles.container }>
          <Text variant="xLarge">{this.props.description}</Text>
          {stage == Stage.Done && logurl != null &&   
            <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
            Membership has been updated, it can take up to an hour for Teams to reflect this! 
            <Link href={logurl}>History</Link>
          </MessageBar>}
          {stage == Stage.LoadingTeams && <ProgressIndicator label="Loading Teams" description="Loading the teams you are a member of" /> }
          {stage == Stage.LoadingCurrentMembers && <ProgressIndicator label="Loading Current Members" description="Generating a list of current members" /> }
          {stage == Stage.ComparingMembers && <ProgressIndicator label="Comparing Current Members" description="Comparing the current members with the csv file" /> }
          {stage == Stage.RemovingOrphendMembers && <ProgressIndicator label="Removing Orphend Members" description="Removing members who are not owners or in the csv file (orphend)" /> }
          {stage == Stage.AddingNewMembers && <ProgressIndicator label="Adding New Members" description="Adding members who are new in the csv file" /> }
          {stage == Stage.LoggingDone && <ProgressIndicator label="Logging this request" description="Logging this request for stats purposes" /> }
          <Dropdown label="1. Select the Team (you need to be an owner, it will be checked)" onChange={this.onChange} placeholder="Select an option" options={items} disabled={items.length == 0} />
          {stage == Stage.CheckingOwnership && <ProgressIndicator label="Checking Team Ownership" description="Checking to make sure you are an owner of this team" /> }
          {stage == Stage.ErrorOwnership && <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>You are not an owner of this group. Please select another.</MessageBar>}
          <div style={{padding: "5px 0"}}>
            <span>2. Select a CSV File</span>
            <CSVReader onDrop={this.handleOnDrop} onError={this.handleOnError} addRemoveButton config={{ header: true, skipEmptyLines: true }} onRemoveFile={this.handleOnRemoveFile}><span>Drop CSV file here or click to upload.</span></CSVReader>
          </div>
          <Dropdown label="3. Select the Email Addresss Column" onChange={this.onEmailChange} placeholder="Select an option" options={csvItems} disabled={!csvdata} />
          <PrimaryButton text="4. Update Membership" onClick={this.onRun} allowDisabledFocus disabled={!csvdata || items.length == 0 || stage != Stage.Done || !csvSelected} />

          <Separator>CSV Preview</Separator>
          {csvdata && <DetailsList
              items={csvdata}
              columns={csvcolumns}
              setKey="set"
              layoutMode={DetailsListLayoutMode.justified}
              selectionMode={SelectionMode.none}
            />}
          {logs.length > 0 && (<><Separator>Logs</Separator><List items={logs} onRenderCell={this._onRenderCell} /></>)}
          {errors.length > 0 && (<><Separator>Errors</Separator><List items={errors} onRenderCell={this._onRenderCell} /></>)}
        </div>
      </div>
    );
  }  
  
  private _onRenderCell = (item: any, index: number | undefined): JSX.Element => {
    return (
      <div data-is-focusable={true}>
        <div style={{padding: 2}}>
          {item}
        </div>
      </div>
    );
  }
}
