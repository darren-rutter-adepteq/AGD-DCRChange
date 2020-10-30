import * as React from 'react';

import {Label, DetailsList, DefaultButton, Selection, IColumn, Persona, PersonaSize, Icon, SelectionMode,
        mergeStyles, PrimaryButton, ILabelStyles, IStyle} from 'office-ui-fabric-react';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import {PeoplePicker} from "@pnp/spfx-controls-react/lib/PeoplePicker";

import {sp} from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/lists/web";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/files";
import "@pnp/sp/site-users";

export interface IApproversState {
  addUsers: number[];
  originalDraftApprovers: number[];
  endDraftApprovers: number[];
  approversList: IDetailsListItem[];
  addUserNumber: number;
  selectedApprover: Selection;
  disableAddApprover: boolean;
  showSaveButton: boolean;
  showCancelButton: boolean;
  afterSave: boolean;

  approverListColumns: IColumn[];

  dcrApprovers: number[];
}

export interface IDetailsListItem {
  key: number;
  name: string;
  value: number;
}

export interface IApproversProps {
  showLabel: boolean;
  context: WebPartContext;
  selectedDocument?: Selection;
  documentId?: number;
  mode: string;
  draftId?: number;
  selectedDCR?: Selection;
  updateApprovers?: any;
  dcrId?: number;
}

export class Approvers extends React.Component<IApproversProps, IApproversState> {
  private _approverListColumns: IColumn[];
  private _approverNameOnlyColumns: IColumn[];
  private headerLabelStyle: ILabelStyles;
  private headerStyle:IStyle;

  private iconStyles = mergeStyles({
    fontSize: 20,
    height: 20,
    width: 20,
    textAlign: 'center',
    verticalAlign: 'center'
  })

  public constructor (props: IApproversProps, state: IApproversState) {
    super(props);

    this.headerStyle = {
      fontSize: "20px"
    };

    this.headerLabelStyle = { root: this.headerStyle};

    this._approverNameOnlyColumns = [
      {
        key: 'column1', name: 'Name', fieldName: 'Approver.Title', minWidth: 200, onRender: (item) => {
          console.log("APPROVER RENDER: ", item);
          if (item['Title'] == "New") {
            this.setState({addUserNumber: this.state.addUserNumber + 1})
            return (
              <PeoplePicker context={this.props.context} ensureUser={true} personSelectionLimit={5} selectedItems={this._getPeoplePickerItems}></PeoplePicker>
            )
          } else {
            return (
              <Persona text={item['Title']} size={PersonaSize.size24}></Persona>
            );
          }
        }
      }
    ];

    this._approverListColumns = [
      {
        key: 'column1', name: 'Name', fieldName: 'Approver.Title', minWidth: 150, onRender: (item) => {
          console.log("APPROVER RENDER: ", item);
          if (item['Title'] == "New") {
            this.setState({addUserNumber: this.state.addUserNumber + 1})
            return (
              <PeoplePicker context={this.props.context} ensureUser={true} personSelectionLimit={5} selectedItems={this._getPeoplePickerItems}></PeoplePicker>
            )
          } else {
            return (
              <Persona text={item['Title']} size={PersonaSize.size24}></Persona>
            );
          }
        }
      },
      { key: 'column2', name: 'Source', fieldName: 'Source', minWidth: 200 },
      { key: 'column3', name: '', fieldName: '', minWidth: 100, maxWidth: 100, onRender: (item) => {
        if (item['Source'] == "Document Category Default" || item['Source'] == "Document Default" || item['Source'] == "Document Specific Default" || item['Source'] == "Document Owner") {
          return (
            <Icon iconName='LockSolid' className={this.iconStyles}/>
          )
        } else if ((item['Source'] == "One-time" && item['Title'] != "New") || (item['Source'] == "DCR" && item['Title'] != "New")) {
          return (<DefaultButton text="remove" onClick={this._removeApprover}></DefaultButton>)
        } else {
          return (<span>&nbsp;</span>)
        }
      }}
    ];

    this.state = {
      addUsers: [],
      originalDraftApprovers: [],
      endDraftApprovers: [],
      approversList: [],
      addUserNumber: 0,
      selectedApprover: new Selection({selectionMode: SelectionMode.single }),
      disableAddApprover: false,
      showSaveButton: false,
      afterSave: false,
      approverListColumns: this._approverListColumns,
      dcrApprovers: [],
      showCancelButton: false
    }

    this._getPeoplePickerItems = this._getPeoplePickerItems.bind(this);
    this._saveApprovers = this._saveApprovers.bind(this);
    this._cancelAddApprovers = this._cancelAddApprovers.bind(this);
  }

  public async getData() {
    let currentDocumentCategory;
    let currentDocumentApprovers;
    let currentDocumentOwner;
    console.log("SELECTED DOCUMENT PROPS: ", this.props.selectedDocument);
    if (this.props.selectedDocument) {
      currentDocumentCategory = this.props.selectedDocument.getSelection()[0]['Category']['Title'];
      currentDocumentOwner = this.props.selectedDocument.getSelection()[0]['DocumentOwner'];
    } else {
      let response = await sp.web.lists.getByTitle('Document List').items.getById(this.props.documentId).select("*,Category/Title,DocumentSpecificApprover/Title,DocumentSpecificApprover/Id,DocumentOwner/Id,DocumentOwner/Title").expand("Category,DocumentSpecificApprover,DocumentOwner").get()
      console.log("DOCUMENT RESPONSE: ", response);
      currentDocumentCategory = response['Category']['Title'];
      currentDocumentApprovers = response['DocumentSpecificApprover'];
      currentDocumentOwner = response['DocumentOwner'];
    }
    
    let approvers = [];
    console.log("CURRENT DOCUMENT OWNER: ", currentDocumentOwner);
    if (currentDocumentOwner != null) {
      currentDocumentOwner['Source'] = "Document Owner";
      approvers.push(currentDocumentOwner)
    }

    let documentCategory = await sp.web.lists.getByTitle('Document Category').items.select('*,Member_x0020_Responsibility/Id,Member_x0020_Responsibility/Title').expand('Member_x0020_Responsibility').filter(`Title eq '${currentDocumentCategory}'`).get()
    console.log("DOCUMENT CATEGORY: ", documentCategory[0])
    if (documentCategory[0]['Member_x0020_Responsibility'] != null) {
      documentCategory[0]['Member_x0020_Responsibility'].forEach((key, value) => {
        key['Source'] = "Document Category Default";
        approvers.push(key);
      })
    }

    if ((this.props.selectedDocument && this.props.selectedDocument.getSelection()[0]['DocumentSpecificApproverId']) || currentDocumentApprovers) {
      if (this.props.selectedDocument) {
        currentDocumentApprovers = this.props.selectedDocument.getSelection()[0]['DocumentSpecificApprover'];
      }
      
      currentDocumentApprovers.forEach((key, value) => {
        console.log("DRAFTKEY: ", key, "DRAFTVALUE: ", value);
        key['Source'] = "Document Default";
        approvers.push(key);
      })
    }
    let originalDraftApprovers = [];
    let currentDCRApprovers = [];

    if (this.props.draftId && this.props.mode == "One-time") {
      let draftApprovers = await sp.web.lists.getByTitle('Draft List').items.getById(this.props.draftId).select("*,Approvers/Id,Approvers/Title").expand("Approvers").get();
      console.log("DRAFT APPROVERS: ", draftApprovers['Approvers']);
      if (draftApprovers['Approvers'] != null) {
        draftApprovers['Approvers'].forEach((key, value) => {
          key['Source'] = "One-time";
          approvers.push(key);
          originalDraftApprovers.push(key['Id']);
        })
      }
    } else if (this.props.selectedDCR && this.props.mode == "DCR") {
      console.log("*** TO DO GET DCR APPROVERS ***");

      let DCRApprovers = await sp.web.lists.getByTitle("DCR Register").items.getById(this.props.dcrId).select("*,Approvers/Id,Approvers/Title").expand("Approvers").get();//this.props.selectedDCR.getSelection()[0]['Approvers'];
      console.log("DCR APPROVERS: ", DCRApprovers['Approvers'])
      if (DCRApprovers['Approvers'] != null) {
        DCRApprovers['Approvers'].forEach((key, value) => {
          console.log("DCRKEY: ", key, "DCRVALUE: ", value);
          key['Source'] = "DCR";
          currentDCRApprovers.push(key['Id']);
          approvers.push(key);
        })
      }
      this.props.updateApprovers(currentDCRApprovers);
    }

    this.setState({
      dcrApprovers: [].concat(currentDCRApprovers),
      approversList: [].concat(approvers),
      originalDraftApprovers: originalDraftApprovers
    })
  }

  public async componentDidUpdate(prevProps) {
    console.log("COMPARE PROPS: ", prevProps, this.props);
    console.log("COMPONENT DID UPDATE: ", prevProps.dcrId, this.props.dcrId)
    if ((prevProps.dcrId != this.props.dcrId) || (prevProps.documentId != this.props.documentId)) {
      await this.getData()
    }

    if (this.state.afterSave) {
      await this.getData()
      this.setState({
        afterSave: false
      })
    }
  }

  public async componentDidMount() {
    // GET DOCUMENT CATEGORY APPROVERS
    /*let documentId = this.props.documentId;
    let documentUsers = await sp.web.lists.getByTitle('Approver List').items.select('*,Approver/Id,Approver/Title').expand('Approver').filter(`DocumentId eq '${documentId}'`).get();
    console.log("DOCUMENT USERS: ", documentUsers);

    this.setState({
      approversList: [].concat(documentUsers)
    })*/
    this.getData()
  }

  private async _getPeoplePickerItems(items: any[]) {
    console.log("GET PEOPLE PICKER ITEMS: ", this, items);
      let currentUsers = [];

      items.forEach((key, value) => {
        console.log("PEOPLE PICKER: ", key, value);
        currentUsers.push(key['id']);
      })
    if (this.props.mode == "One-time") {
      console.log("ONE TIME USERS TO ADD: ", currentUsers);
      if (currentUsers.length) {
        this.setState({
          showSaveButton: true,
          addUsers: [].concat(currentUsers),
          endDraftApprovers: [].concat(this.state.originalDraftApprovers).concat(currentUsers)
        })
      } else {
        this.setState({
          showSaveButton: false,
          addUsers: [].concat(currentUsers)
        })
      }
    } else if (this.props.mode == 'DCR') {
      let totalDCRUsers = [].concat(this.state.dcrApprovers).concat(currentUsers);
      this.props.updateApprovers(totalDCRUsers);
    }
  }

  private _addApprover = (event) => {
    this.setState({approverListColumns: this._approverNameOnlyColumns})
    let approvers = this.state.approversList;
    let newItem = {value: 0, name: 'NewItemTest', key: 0, Title: 'New', Source: this.props.mode}
    approvers.push(newItem);

    this.setState({
      approversList: [].concat(approvers),
      disableAddApprover: true,
      showCancelButton: true
    })
  }

  private _removeApprover = async () => {
    console.log("REMOVE APPROVER THIS: ", this);
    let approverToRemove = this.state.selectedApprover.getSelection()[0]['Id'];
    console.log("SELECTION: ", this.state.selectedApprover.getSelection()[0]['Id']);
    let approvers;
    if (this.props.mode == "One-time") {
      approvers = this.state.originalDraftApprovers
    } else if (this.props.mode == "DCR") {
      approvers = this.state.dcrApprovers
    }
    
    var index = approvers.indexOf(approverToRemove)
    if (index !== -1) {
      approvers.splice(index, 1);
    }
    if (this.props.mode == "One-time") {
      this.setState({
        originalDraftApprovers: approvers
      });
      let approverReturn = await sp.web.lists.getByTitle('Draft List').items.getById(this.props.draftId).update({
        ApproversId: {"results": this.state.originalDraftApprovers }
      });
    } else if (this.props.mode == "DCR") {
      this.setState({
        dcrApprovers: approvers
      });
      let approverReturn = await sp.web.lists.getByTitle("DCR Register").items.getById(this.props.dcrId).update({
        ApproversId: {"results": this.state.dcrApprovers}
      })
    }
    
    this.setState({
      afterSave: true,
      showSaveButton: false,
      disableAddApprover: false
    })
  }

  private _saveApprovers = async (event) => {
    console.log("SAVE APPROVER THIS: ", this)
    console.log("SAVE APPROVERS: ", this.state.endDraftApprovers);
    let approverReturn = await sp.web.lists.getByTitle('Draft List').items.getById(this.props.draftId).update({
      ApproversId: {"results": this.state.endDraftApprovers }
    });
    console.log("Approver Return: ", approverReturn);
    this.setState({
      afterSave: true,
      showSaveButton: false,
      disableAddApprover: false,
      showCancelButton: false,
      approverListColumns: this._approverListColumns
    })
  }

  private _cancelAddApprovers = () => {
    let originalApproverList = this.state.approversList;
    originalApproverList.pop()
    this.setState({
      approverListColumns: this._approverListColumns,
      approversList: originalApproverList,
      showCancelButton: false,
      disableAddApprover: false,
      showSaveButton: false,
      addUsers: []
    });
  }

  public render(): React.ReactElement<IApproversProps> {
    const {approversList, selectedApprover, disableAddApprover, showSaveButton, approverListColumns, showCancelButton} = this.state;
    return (
      <div>
        {(this.props.showLabel) ? 
        <Label styles={this.headerLabelStyle}>Approvers</Label>
        : <span>&nbsp;</span> }
        <DetailsList items={approversList} columns={approverListColumns} selection={selectedApprover}></DetailsList>
        <br />
        {(this.props.mode != 'view') ? 
         (disableAddApprover) ? 
        <DefaultButton text="Add Approvers" disabled iconProps={{iconName: 'add'}} onClick={this._addApprover}></DefaultButton> 
        : <DefaultButton text="Add Approvers" iconProps={{iconName: 'add'}} onClick={this._addApprover}></DefaultButton>
        : <span>&nbsp;</span> }
        {(showSaveButton) ? <PrimaryButton text="Save" iconProps={{iconName: 'save'}} onClick={this._saveApprovers}></PrimaryButton> : <span>&nbsp;</span>}
        {(showCancelButton) ? <DefaultButton text="Cancel" iconProps={{iconName: 'cancel'}} onClick={this._cancelAddApprovers}></DefaultButton> : <span></span>}
      </div>
    )
  }

}