import * as React from 'react';

import {Label, TextField, Stack, Dropdown, DetailsList, DefaultButton, ILabelStyles, IDropdownOption,
        IColumn, Persona, PersonaSize, Icon, mergeStyles, IStackTokens, Selection, SelectionMode,
        Separator, PrimaryButton, Spinner, SpinnerSize, Panel, PanelType} from 'office-ui-fabric-react'

import {sp} from "@pnp/sp";
import {Web} from "@pnp/sp/webs";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/lists/web";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/site-users";
import "@pnp/sp/security/web";
import "@pnp/sp/security";
import "@pnp/sp/site-users/web";

import {PeoplePicker} from "@pnp/spfx-controls-react/lib/PeoplePicker";

import { WebPartContext } from "@microsoft/sp-webpart-base";

import {getGUID} from "@pnp/common";

//import {Approvers} from "./Approvers"

import {Approvers} from 'agd-dcr-library';

export interface IDetailsListItem {
  key: number;
  name: string;
  value: number;
}

export interface ICreateNewDCRState {
  text: string;

  dcrRequest: string;
  dcrRequestValidated: boolean;
  dcrRequestError: string;

  dcrReason: string;
  dcrReasonValidated: boolean;
  dcrReasonError: string;

  dcrPriority: string;
  dcrPriorityValidated: boolean;
  dcrPriorityError: string;

  dcrMagnitude: string;
  dcrMagnitudeValidated: boolean;
  dcrMagnitudeError: string;

  dcrStatus: string;
  dcrStatusValidated: boolean;
  dcrStatusError: string;

  addUsers: number[];

  isDataFetched: boolean;

  approversListItems: IDetailsListItem[];
  approversSelection: Selection;

  currentlySaving: boolean;

  isOpen: boolean;

  approverIds: number[];
}

export interface ICreateNewDCRProps {
  description: string;
  headerLabelStyle: ILabelStyles;
  headerLabel: string;
  context: WebPartContext;
  approversListItems: IDetailsListItem[];
  documentId: number;
  afterSuccess: any;
  onClose: any;
  selectedDocument: Selection;
  draftId?: number;
  currentUser: number;
}

export class CreateNewDCR extends React.Component<ICreateNewDCRProps, ICreateNewDCRState> {
  private _magnitudeOptions: IDropdownOption[] = [];
  private _priorityOptions: IDropdownOption[] = [];
  private _statusOptions: IDropdownOption[] = [];
  private _approverListColumns: IColumn[];

  private _stackTokens: IStackTokens = {
    childrenGap: 10,
  };

  private stackStyles = mergeStyles({
    display: 'flex',
    justifyContent: 'space-between',
    flexBasis: "100%"
  });

  private stackItemStyles = mergeStyles({
    minWidth: "150px"
  });

  private iconStyles = mergeStyles({
    fontSize: 20,
    height: 20,
    width: 20,
    textAlign: 'center',
    verticalAlign: 'center'
  })

  private dcrReasonErrorDefault = "Please enter a Reason";
  private dcrRequestErrorDefault = "Please enter a Request";
  private dcrPriorityErrorDefault = "Please select a Priority";
  private dcrMagnitudeErrorDefault = "Please select a Magnitude";
  private dcrStatusErrorDefault = "Please select a Status";

  public constructor (props: ICreateNewDCRProps, state: ICreateNewDCRState) {
    super(props);

    this._approverListColumns = [
      {
        key: 'column1', name: 'Name', fieldName: 'Approver.Title', minWidth: 200, onRender: (item) => {
          console.log("APPROVER RENDER: ", item);
          if (item['Title'] == "New") {
            return (
              <PeoplePicker context={this.props.context} ensureUser={true} selectedItems={this._getPeoplePickerItems}></PeoplePicker>
            )
          } else {
            return (
              <Persona text={item['Title']} size={PersonaSize.size24}></Persona>
            );
          }
        }
      },
      { key: 'column2', name: 'Source', fieldName: 'Source', minWidth: 200 },
      { key: 'column3', name: '', fieldName: '', minWidth: 50, maxWidth: 50, onRender: (item) => {
        if (item['Source'] == "Document Category Default" || item['Source'] == "Document Default" || item['Source'] == "Document Specific Default") {
          return (
            <Icon iconName='LockSolid' className={this.iconStyles}/>
          )
        } else {
          return (<span>&nbsp;</span>)
        }
      }}
    ];

    this.state = {
      text: "Hello",
      dcrRequest: "",
      dcrRequestValidated: false,
      dcrRequestError: this.dcrRequestErrorDefault,
      
      dcrReason: "",
      dcrReasonValidated: false,
      dcrReasonError: this.dcrReasonErrorDefault,

      dcrPriority: "",
      dcrPriorityValidated: false,
      dcrPriorityError: this.dcrPriorityErrorDefault,
      
      dcrMagnitude: "",
      dcrMagnitudeValidated: false,
      dcrMagnitudeError: this.dcrMagnitudeErrorDefault,

      dcrStatus: "Open",
      dcrStatusValidated: true,
      dcrStatusError: "",

      addUsers: [],

      approversListItems: this.props.approversListItems,
      approversSelection: new Selection({selectionMode: SelectionMode.none}),

      isDataFetched: false,

      currentlySaving: false,

      isOpen: true,

      approverIds: []
    }

    //this.getDropDownValues()

    this.magnitudeDropDownChanged = this.magnitudeDropDownChanged.bind(this);
    this.priorityDropDownChanged = this.priorityDropDownChanged.bind(this);
    this.statusDropDownChanged = this.statusDropDownChanged.bind(this);
    this._getPeoplePickerItems = this._getPeoplePickerItems.bind(this);
    this.saveDCRChanges = this.saveDCRChanges.bind(this);
    this.discardDCRChanges = this.discardDCRChanges.bind(this);
    this.getDropDownValues = this.getDropDownValues.bind(this);
    this._panelDismiss = this._panelDismiss.bind(this);
    this.setApprovers = this.setApprovers.bind(this);
  }


  private _getPeoplePickerItems(items: any[]) {
    console.log("GET PEOPLE PICKER ITEMS: ", this, items);
    let currentUsers = this.state.addUsers;
    currentUsers.push(items[0]['id']);
    console.log("ONE TIME USERS TO ADD: ", currentUsers);
    this.setState({
      addUsers: [].concat(currentUsers)
    });
  }

  public async componentDidMount() {
    console.log("COMPONENT DID MOUNT");
    await this.getDropDownValues()
  }

  public async getDropDownValues() {

    let list = sp.web.lists.getByTitle("DCR Register");
    let batch = sp.web.createBatch();

    //list.items.inBatch(batch).add(

    list.fields.getByTitle("Priority").inBatch(batch)()
      .then((response) => {
        console.log("GET DROP DOWN DETAILS - PRIORITY: ", response);
        for(const key in response['Choices']) {
          this._priorityOptions.push({key: response['Choices'][key], text: response['Choices'][key]});
        }
      });

    list.fields.getByTitle("Magnitude").inBatch(batch)()
      .then((response) => {
        for(const key in response['Choices']) {
          this._magnitudeOptions.push({key: response['Choices'][key], text: response['Choices'][key]});
        }
      })

    list.fields.getByTitle('Status').inBatch(batch)()
      .then((response) => {
        for (const key in response['Choices']) {
          this._statusOptions.push({key: response['Choices'][key], text: response['Choices'][key]})
        }
      })
    
      const batchResponse = await batch.execute();

      if (this.props.selectedDocument.getSelection()[0]['DocumentOwnerId'] == this.props.currentUser) {
        console.log("USER IS DOCUMENT OWNER!!");
        this.setState({
          dcrStatus: "Accepted"
        })
      }

      this.setState({
        isDataFetched: true
      })
  }

  private _dcrRequestFieldChanged = (event, newValue) => {
    console.log("DCR REQUEST FIELD CHANGED: ", newValue);

    this.setState({
      dcrRequest: newValue
    });

    if (newValue == "") {
      this.setState({
        dcrRequestValidated: false,
        dcrRequestError: "Please enter a Request"
      })
    } else {
      this.setState({
        dcrRequestValidated: true,
        dcrRequestError: ""
      })
    }
  }

  private _dcrReasonFieldChanged = (event, newValue) => {
    console.log("DCR REASON FIELD CHANGED: ", newValue);

    this.setState({
      dcrReason: newValue
    });

    if (newValue == "") {
      this.setState({
        dcrReasonValidated: false,
        dcrReasonError: "Please enter a Reason"
      })
    } else {
      this.setState({
        dcrReasonValidated: true,
        dcrReasonError: ""
      })
    }
  }

  public priorityDropDownChanged(event, value) {
    console.log("DCR PRIORITY FIELD CHANGED: ", value);
    this.setState({
      dcrPriority: value.text
    });
    if (value.text == "") {
      this.setState({
        dcrPriorityValidated: false,
        dcrPriorityError: "Please select a Priority"
      });
    } else {
      this.setState({
        dcrPriorityValidated: true,
        dcrPriorityError: ""
      })
    }
    
  }

  public statusDropDownChanged(event, value) {
    console.log("NEW STATUS SELECTED", value)
    this.setState({
      dcrStatus: value.text
    });
    if (value.text == "") {
      this.setState({
        dcrStatusValidated: false,
        dcrStatusError: "Please select a Priority"
      });
    } else {
      this.setState({
        dcrStatusValidated: true,
        dcrStatusError: ""
      })
    }
  }

  public magnitudeDropDownChanged(event, value) {
    console.log("NEW MAGNITUDE SELECTED: ", value)
    this.setState({
      dcrMagnitude: value.text
    });

    if (value.text == "") {
      this.setState({
        dcrMagnitudeValidated: false,
        dcrMagnitudeError: "Please select a Magnitude"
      })
    } else {
      this.setState({
        dcrMagnitudeValidated: true,
        dcrMagnitudeError: ""
      })
    }
  }

  private _addDCRApprover = (event) => {
    let approverItems = this.props.approversListItems;
    let newItem = {value: 0 , name: 'NewItemTest', key: 0, Source: 'One-time', Title: 'New'};
    approverItems.push(newItem);

    this.setState({
      approversListItems: [].concat(approverItems)
    });
  }

  private _addApproverDefaultItem(oApproverItem, newDCRID) {
    let returnObject = {}
    console.log("APPROVER ITEM: ", oApproverItem);
    if (!/New/.test(oApproverItem['Title'])) {
      returnObject['Title'] = getGUID(),
      returnObject['Source'] = oApproverItem['Source'],
      returnObject['ApproverId'] = oApproverItem['Id'],
      returnObject['Linked_x0020_DCRId'] = newDCRID
    }
    return returnObject;
  }

  private _addApproverOneTimeItem(nUserID, newDCRID) {
    let returnObject = {}
    returnObject['Title'] = getGUID(),
    returnObject['Source'] = 'One-time',
    returnObject['ApproverId'] = nUserID,
    returnObject['Linked_x0020_DCRId'] = newDCRID
    return returnObject
  }

  private setApprovers(aApprovers) {
    console.log("SET APPROVERS: ", aApprovers)
    this.setState({approverIds: aApprovers})
  }

  private async saveDCRChanges() {
    try {
      const {dcrRequestValidated, dcrReasonValidated, dcrStatusValidated, dcrMagnitudeValidated, dcrPriorityValidated,
             dcrRequest, dcrReason, dcrStatus, dcrMagnitude, dcrPriority, currentlySaving, approverIds} = this.state
      let newReferenceNumber = 0;
      let newDCRID = 0;
      let newDCRReference = "";
      if (dcrRequestValidated && dcrReasonValidated && dcrMagnitudeValidated && dcrPriorityValidated && dcrStatus) {
        let recentDCRNumberResponse = await sp.web.lists.getByTitle("DCR Register").items.select("ID, Reference").orderBy("Created", false).top(1).get()
        let reference = recentDCRNumberResponse[0]['Reference'];
        let num = parseInt(reference.match(/\d+/));
        newReferenceNumber = num + 1;

        let newDCRItem = await sp.web.lists.getByTitle("DCR Register").items.add(
          {
            Reference: `DCR ${newReferenceNumber}`,
            Title: dcrRequest,
            Reason: dcrReason,
            Status: dcrStatus,
            Magnitude: dcrMagnitude,
            Priority: dcrPriority,
            Linked_x0020_DocumentId: this.props.documentId,
            ApproversId: {"results": approverIds}
          });

        newDCRID = newDCRItem.data.ID;
        newDCRReference = newDCRItem.data.Reference;
        //let list = sp.web.lists.getByTitle("Approver List");
        //let batch =sp.web.createBatch();

        /*for (let i=0, len=this.state.approversListItems.length; i<len; i++) {
          //promises.push(this._addApproverDefaultItem(this.state.approversListItems[i], newDCRID));
          list.items.inBatch(batch).add(this._addApproverDefaultItem(this.state.approversListItems[i], newDCRID));
        }

        for (let i=0, len=this.state.addUsers.length; i<len; i++) {
          //promises.push(this._addApproverOneTimeItem(this.state.addUsers[i], newDCRID));
          list.items.inBatch(batch).add(this._addApproverOneTimeItem(this.state.addUsers[i], newDCRID))  
        }*/
        try {
          this.setState({currentlySaving: true})
          //const batchResponse = await batch.execute();
          
          this.props.afterSuccess(`DCR record ${newDCRReference} has been successfully created`)

          this.setState({
            addUsers: [],
            dcrReason: "",
            dcrReasonValidated: false,
            dcrReasonError: this.dcrReasonErrorDefault,
            dcrRequest: "",
            dcrRequestValidated: false,
            dcrRequestError: this.dcrRequestErrorDefault,
            dcrStatus: "",
            dcrStatusValidated: false,
            dcrStatusError: this.dcrStatusErrorDefault,
            dcrMagnitude: "",
            dcrMagnitudeValidated: false,
            dcrMagnitudeError: this.dcrMagnitudeErrorDefault,
            dcrPriority: "",
            dcrPriorityValidated: false,
            dcrPriorityError: this.dcrPriorityErrorDefault,
            currentlySaving: false,
            isOpen: false
          });
        } catch (e) {
          console.error(e);
          this.setState({
            addUsers: [],
            dcrReason: "",
            dcrReasonValidated: false,
            dcrReasonError: this.dcrReasonErrorDefault,
            dcrRequest: "",
            dcrRequestValidated: false,
            dcrRequestError: this.dcrRequestErrorDefault,
            dcrStatus: "",
            dcrStatusValidated: false,
            dcrStatusError: this.dcrStatusErrorDefault,
            dcrMagnitude: "",
            dcrMagnitudeValidated: false,
            dcrMagnitudeError: this.dcrMagnitudeErrorDefault,
            dcrPriority: "",
            dcrPriorityValidated: false,
            dcrPriorityError: this.dcrPriorityErrorDefault,
            isOpen: false
          });
        }
      }
/*.then((response) => {
                    console.log("ADD ONE TIME USERS RESPONSE: ", response);
                    
                  })*/

      /*for (let i=0, len=this.state.addUsers.length; i<len; i++) {
        sp.web.lists.getByTitle('Approver List').items.add({
          Title: getGUID(),
          Source: 'One-time',
          ApproverId: this.state.addUsers[i],
          Linked_x0020_DCRId: this.state.dcrId
        });
      }*/

      /*sp.web.lists.getByTitle('DCR Register').items.getById(this.state.dcrId).update({
        Title: this.state.dcrRequest,
        Reason: this.state.dcrReason,
        Magnitude: this.state.dcrMagnitude,
        Priority: this.state.dcrPriority
      }).then((result):void => {
        console.log("DCR UPDATE COMPLETE", result);
      });*/
    } catch(e) {
      
    }
  }

  private discardDCRChanges() {
    console.log("DISCARD DCR CHANGES THIS: ", this);
    this.setState({isOpen: false})
  }

  private _panelDismiss() {
    this.setState({isOpen: false})
  }

  public render(): React.ReactElement<ICreateNewDCRProps> {
    console.log("CREATE DCR STATE", this.state)
    const {isDataFetched, dcrRequest, dcrRequestError, dcrReason, dcrReasonError, dcrMagnitude, dcrMagnitudeError, 
      dcrPriorityError, dcrPriority, dcrStatus, dcrStatusError,approversListItems, currentlySaving,
      approversSelection, isOpen } = this.state;

      let _panelFooter;
      _panelFooter = () => (
        <div style={{paddingLeft: "20px"}}>
          <Separator></Separator>
          <PrimaryButton text="Save changes" onClick={this.saveDCRChanges}/>
          <DefaultButton style={{marginLeft: "10px"}} text="Discard changes" onClick={this.discardDCRChanges}/>
          {(currentlySaving) ? 
            <Spinner label="Saving..." size={SpinnerSize.small}></Spinner> : <span>&nbsp;</span>
            }
          <br />
          <br />
          <br />
        </div>
      );

    return ( 
      (isDataFetched && isOpen) ? <Panel isOpen={true} type={PanelType.medium} isFooterAtBottom={true} isBlocking={false} onDismiss={this._panelDismiss} onRenderFooter={_panelFooter}>
            <Label styles={this.props.headerLabelStyle}>{this.props.headerLabel}</Label>
            <TextField label="Request" multiline rows={3} value={dcrRequest} onChange={this._dcrRequestFieldChanged} required errorMessage={dcrRequestError}/>
            <br />
            <TextField label="Reason" multiline rows={3} value={dcrReason} onChange={this._dcrReasonFieldChanged} required errorMessage={dcrReasonError}/>
            <br />
            <Stack horizontal tokens={this._stackTokens} className={this.stackStyles}>
              <Stack.Item className={this.stackItemStyles}>
                <Dropdown options={this._statusOptions} disabled selectedKey={dcrStatus} onChange={this.statusDropDownChanged} label="Status"  required errorMessage={dcrStatusError}  ></Dropdown>
              </Stack.Item >
              <Stack.Item className={this.stackItemStyles}>
                <Dropdown options={this._magnitudeOptions} selectedKey={dcrMagnitude} onChange={this.magnitudeDropDownChanged} label="Magnitude" required errorMessage={dcrMagnitudeError}></Dropdown>
              </Stack.Item>
              <Stack.Item className={this.stackItemStyles}>
                <Dropdown options={this._priorityOptions} selectedKey={dcrPriority} onChange={this.priorityDropDownChanged} label="Priority" required errorMessage={dcrPriorityError}></Dropdown>
              </Stack.Item> 
            </Stack>
            <br />
            <Approvers showLabel={true} context={this.props.context} selectedDocument={this.props.selectedDocument} mode={'DCR'}
                updateApprovers={this.setApprovers}></Approvers>
            {/*<Label>Approver List</Label> 
            <DetailsList items={approversListItems} columns={this._approverListColumns} selection={approversSelection}></DetailsList>
            <br />
            <DefaultButton text="Add Approver" iconProps={{iconName: 'Add'}} onClick={this._addDCRApprover}/>*/}
            <br /> 
      </Panel> : <span>&nbsp;</span>
    )
  }
}