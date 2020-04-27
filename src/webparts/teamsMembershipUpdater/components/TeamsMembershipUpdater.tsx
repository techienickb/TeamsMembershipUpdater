import * as React from 'react';
import styles from './TeamsMembershipUpdater.module.scss';
import { ITeamsMembershipUpdaterProps } from './ITeamsMembershipUpdaterProps';
import { DetailsList, DetailsListLayoutMode, Selection, IColumn, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
import { ProgressIndicator } from 'office-ui-fabric-react/lib/ProgressIndicator';
import { Separator } from 'office-ui-fabric-react/lib/Separator';
import { PrimaryButton } from 'office-ui-fabric-react';
import { Dropdown, DropdownMenuItemType, IDropdownOption, IDropdownProps } from 'office-ui-fabric-react/lib/Dropdown';
import { Text, ITextProps } from 'office-ui-fabric-react/lib/Text';
import { escape } from '@microsoft/sp-lodash-subset';
import { ITeamsMembershipUpdaterWebPartProps } from '../TeamsMembershipUpdaterWebPart';
import { CSVReader } from 'react-papaparse';


export interface ITeamsMembershipUpdaterState {
  items: IDropdownOption[];
  selectionDetails: IDropdownOption;
  csvdata: [];
  csvcolumns: IColumn[];
  csvSelected: IDropdownOption;
  csvItems: IDropdownOption[];
}

export default class TeamsMembershipUpdater extends React.Component<ITeamsMembershipUpdaterProps, ITeamsMembershipUpdaterState> {
  private _selection: Selection;
  private _columns: IColumn[];
  private _datacolumns: IColumn[];
  private _data: null;

  constructor(props: ITeamsMembershipUpdaterWebPartProps) {
    super(props);

    this._columns = [
      { key: 'column1', name: 'Name', fieldName: 'name', minWidth: 100, maxWidth: 200, isResizable: true },
    ];

    this.state = {
      items: props.items,
      selectionDetails: null,
      csvdata: null,
      csvcolumns: [],
      csvSelected: null,
      csvItems: []
    };
  }


  public handleOnDrop = (data) => {
    console.log(data);
    var h = data[0].meta.fields;
    this._data = data.map(r => { return r.data; });
    this._datacolumns = h.map(r => { return { key: r.replace(' ', ''), name: r, fieldName: r, isResizable: true }; });
    console.log(this._datacolumns);
    console.log(this._data);
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
    this.setState({ ...this.state, selectionDetails: item });
  };

  public onEmailChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    this.setState({ ...this.state, csvSelected: item });
  };

  public onRun = (e) => {
    e.preventDefault();
  }
  

  public render(): React.ReactElement<ITeamsMembershipUpdaterProps> {
    const { items, selectionDetails, csvItems, csvSelected, csvdata, csvcolumns } = this.state;
    return (
      <div className={ styles.teamsMembershipUpdater }>
        <div className={ styles.container }>
          <Text variant="xLarge">{this.props.description}</Text>
          {items.length == 0 && <ProgressIndicator label="Loading Teams" description="Loading the teams you are an owner of" /> }
          <Dropdown label="1. Select the Team (you own)" selectedKey={selectionDetails ? selectionDetails.key : undefined}
            onChange={this.onChange}
            placeholder="Select an option"
            options={items} disabled={items.length == 0}
            />
          <Dropdown label="2. Select the Email Addresss Column" selectedKey={csvSelected ? csvSelected.key : undefined}
          onChange={this.onEmailChange}
          placeholder="Select an option"
          options={csvItems} disabled={!csvdata}
          />
          <PrimaryButton text="3. Update Membership" onClick={this.onRun} allowDisabledFocus disabled={!csvdata || items.length == 0} />
          
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
