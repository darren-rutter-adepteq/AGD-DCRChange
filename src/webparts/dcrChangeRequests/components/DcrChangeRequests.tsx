import * as React from 'react';
import styles from './DcrChangeRequests.module.scss';
import { IDcrChangeRequestsProps } from './IDcrChangeRequestsProps';
import { escape } from '@microsoft/sp-lodash-subset';

import * as moment from 'moment';

import {DetailsList, SearchBox, Panel, Pivot, PivotItem, TextField, Stack, Dropdown, Label,
        DefaultButton, Separator, PrimaryButton, IColumn, Selection, SelectionMode, PanelType, IStackTokens,
        IStackStyles, IStackItemTokens, mergeStyles, IDropdownOption, Persona, PersonaSize, autobind,
        MessageBarType, MessageBar, ConstrainMode, Sticky, StickyPositionType, ScrollablePane, ScrollbarVisibility} from 'office-ui-fabric-react';

import {PeoplePicker, PrincipalType} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import {Pagination} from "@pnp/spfx-controls-react/lib/Pagination";
import {SPHttpClient, SPHttpClientResponse} from "@microsoft/sp-http";

import {getGUID} from "@pnp/common";

import {useConstCallback} from "@uifabric/react-hooks";

//import {Approvers} from "./Approvers";

import {Approvers, DCRPanel} from 'agd-dcr-library';

import {sp} from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users";
import { PagedItemCollection } from '@pnp/sp/items';

export interface IDcrChangeRequestsState {
  dcrListItems: IDetailsListItem[];
  allDCRListItems: IDetailsListItem[];

  filteredItems: IDetailsListItem[];

  approversListItems: IDetailsListItem[];

  historyItems: IDetailsListItem[];
  
  isOpen?: boolean;
  messageBarShow?: boolean;
  panelDCRDetails?: boolean;
  panelDocumentDetails?: boolean;

  messageBarType: MessageBarType;
  messageBarValue: string;

  documentId: number;
  dcrId: number;
  dcrRequest: string;
  dcrReason: string;
  dcrStatus: string;
  dcrMagnitude: string;
  dcrPriority: string;

  addUsers: number[];

  totalPagesNeeded: number;
  currentSelectedPage: number;
  showPagination: boolean;

  currentSearchString: string;

  currentUser: number;
}

export interface IDetailsListItem {
  key: number;
  name: string;
  value: number;
}



export default class DcrChangeRequests extends React.Component<IDcrChangeRequestsProps, IDcrChangeRequestsState> {
  private _allItems: IDetailsListItem[] = [];
  private _dcrListColumns: IColumn[];
  private _approverListColumns: IColumn[];
  private _historyListColumns: IColumn[];
  private _stackItemTokens: IStackItemTokens;
  private _stackTokens: IStackTokens;
  private _magnitudeOptions: IDropdownOption[] = [];
  private _priorityOptions: IDropdownOption[] = [];
  private _statusOptions: IDropdownOption[] = [];
  private stackStyle;
  private dcrSelection: Selection;
  private approversSelection: Selection;
  private pagedItemCollection: PagedItemCollection<any[]>

  private stackStyles = mergeStyles({
    display: 'flex',
    justifyContent: 'space-between',
    flexBasis: "100%"
  });

  private stackItemStyles = mergeStyles({
    minWidth: "150px"
  });

  public constructor(props: IDcrChangeRequestsProps, state: IDcrChangeRequestsState) {
    super(props);

    this.dcrSelection = new Selection({onSelectionChanged: this._onDCRSelectionChanged, selectionMode: SelectionMode.single})
    this.approversSelection = new Selection({canSelectItem: this._canSelectItemApprovers})

    this.state = {
      dcrListItems: this._allItems,
      allDCRListItems: this._allItems,
      approversListItems: this._allItems,
      filteredItems: this._allItems,
      historyItems: this._allItems,
      dcrId: 0,
      dcrRequest: "",
      dcrReason: "",
      dcrStatus: "",
      dcrMagnitude: "",
      dcrPriority: "",
      documentId: 0,
      panelDCRDetails: false,
      addUsers: [],
      messageBarType: MessageBarType.error,
      messageBarValue: "",
      totalPagesNeeded: 1,
      currentSelectedPage: 1,
      showPagination: false,
      currentSearchString: "",
      currentUser: 0
    };

    this._stackItemTokens = {
      margin: 10
    };

    this._stackTokens = {
      childrenGap: 10,
    };

    this.stackStyle = mergeStyles({
      display: 'none'
    });

    this._dcrListColumns = [
      { key: 'column0', name: 'ID', fieldName: 'Reference', minWidth: 100, maxWidth: 100 },
      { key: 'column5', name: 'Document', minWidth: 150, onRender: (item) => {
        return (item['Linked_x0020_Document']['Description']);
      }},
      { key: 'column7', name: 'Date Created', fieldName: 'Created', minWidth: 100, onRender: (item) => {
        let newDate = '';
        if (/(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})/.test(item['Created'])) {
          newDate = moment(item['Created']).format("Do MMM YYYY HH:mm:ss")
        }
        return ( newDate );
      }},
      { key: 'column1', name: 'Change', fieldName: 'Request', minWidth: 150, maxWidth: 400 },
      { key: 'column2', name: 'Reason', fieldName: 'Reason', minWidth: 150, maxWidth: 400 },
      { key: 'column3', name: 'Magnitude', fieldName: 'Magnitude', minWidth: 100, maxWidth: 100 }, 
      { key: 'column4', name: 'Priority', fieldName: 'Priority', minWidth: 100, maxWidth: 100},
      { key: 'column6', name: 'Rasied By', minWidth: 150, onRender: (item) => {
        return (
          <Persona text={item['Author']['Title']} size={PersonaSize.size24}></Persona>
        )
      }}
    ];

    this._approverListColumns = [
      { 
        key: 'column1', name: 'Name', fieldName: 'Approver.Title', minWidth: 100 , onRender: (item) => {
          console.log("APPROVER RENDER: ", item, item['Approver']['Title']);
          if (item['Approver']['Title'] != "New") {
            return (
              <Persona text={item['Approver']['Title']} size={PersonaSize.size24}></Persona>
            );
          } else {
            return (
              <PeoplePicker context={this.props.context} ensureUser={true} selectedItems={this._getPeoplePickerItems}></PeoplePicker>
            );
          }
        }
      },
      {key: 'column2', name: 'Source', fieldName: 'Source', minWidth: 250}
    ];

    this._historyListColumns = [
      { key: 'column3', name: 'Date', fieldName: 'Created', minWidth: 80, onRender: (item) => {
        let newDate = '';
        if (/(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})/.test(item['Created'])) {
          newDate = moment(item['Created']).format("Do MMM YYYY HH:mm:ss")
        }
        return(
          newDate
        );
      }},
      { key: 'column1', name: 'Event Type', fieldName: 'EventType', minWidth: 100},
      { key: 'column2', name: 'Event', fieldName: 'EventValue', minWidth: 300, onRender: (item) => {
        return (
          <div dangerouslySetInnerHTML={{__html: item['EventValue']}}></div>
        )
      }}
    ]

    this.magnitudeDropDownChanged = this.magnitudeDropDownChanged.bind(this);
    this.priorityDropDownChanged = this.priorityDropDownChanged.bind(this);
    this.discardDCRChanges = this.discardDCRChanges.bind(this);
  }

  public async componentDidMount() {
    console.log("COMPONENT DID MOUNT");
    let user = await sp.web.currentUser.get();
    this.setState({currentUser: user['Id']});
    this.updateDimensions();
    window.addEventListener("resize", this.updateDimensions.bind(this));
    this.getSPDCRData();
    this.getDropDownValues();
  }

  public async componentDidUpdate() {
    if (!this.state.showPagination) {
      this.setState({
        showPagination: true
      });
    }
  }

  public updateDimensions() {
  }

  public getDropDownValues() {
    let reactHandler = this;
    this.props.context.spHttpClient.get(`${this.props.siteurl}/_api/web/lists/getbytitle('DCR Register')/fields?$filter=Title eq 'Priority' or Title eq 'Magnitude' or Title eq 'Status'`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        response.json()
          .then((responseJSON: any) => {
            responseJSON.value.forEach(item => {
              if (item.Title.toLowerCase() == "magnitude" || item.Title.toLowerCase() == "priority") {
                this[`_${item.Title.toString().toLowerCase()}Options`].push({key: 'empty', text: ' '});
              }
              for(const key in item.Choices) {
                console.log(this, this._magnitudeOptions, this._priorityOptions);
                console.log(this, item.Title.toLowerCase(), this[`_${item.Title.toLowerCase()}Options`]);
                this[`_${item.Title.toString().toLowerCase()}Options`].push({key: item.Choices[key], text: item.Choices[key]});
                //this._magnitudeOptions.push(console.log("CHOICE VALUE: ", key, item.Choices[key]))
              }
            });
            console.log(this._magnitudeOptions, this._priorityOptions);
          });
      });
  }

  private _addDCRApprover = (event) => {
    let approverItems = this.state.approversListItems;
    let newItem = {value: 0 , name: 'NewItemTest', key: 0, Source: 'One-time', Approver: {Title: 'New'}};
    approverItems.push(newItem);

    this.setState({
      approversListItems: [].concat(approverItems)
    });
  }

  private _dcrRequestFieldChanged = (event, newValue) => {
    console.log("DCR REQUEST FIELD CHANGED: ", newValue);

    this.setState({
      dcrRequest: newValue
    });
  }

  private _dcrReasonFieldChanged = (event, newValue) => {
    console.log("DCR REASON FIELD CHANGED: ", newValue);

    this.setState({
      dcrReason: newValue
    });
  }

  @autobind
  private _getPeoplePickerItems(items: any[]) {
    console.log("GET PEOPLE PICKER ITEMS: ", this, items);
    let currentUsers = this.state.addUsers;
    currentUsers.push(items[0]['id']);
    this.setState({
      addUsers: [].concat(currentUsers)
    });
  }

  public _onDCRSelectionChanged = (): void => {
    this.onDCRSelectionChanged();
  }

  private _canSelectItemApprovers(item: any): boolean {
    console.log("CAN SELECT ITEM APPROVERS: ", item, item['Source']);
    if (item['Source'] === "Document Category Default" || item['Source'] === "Document Specific Default") {
      return false;
    } else {
      return true;
    }
  }

  public  getSPDCRData = async () => {
    let dcrItems = await sp.web.lists.getByTitle("DCR Register").items.top(this.props.itemsPerPage).select("*,Approvers/Title,Approvers/Id,Author/Title,Author/Id,Linked_x0020_Document/Title,Linked_x0020_Document/Description").expand("Approvers,Author,Linked_x0020_Document").orderBy('ActionedDate', true).orderBy('Status', false).orderBy('Created', false).getPaged();
    this.pagedItemCollection = dcrItems;
    if (this.pagedItemCollection.hasNext) { this.setState({totalPagesNeeded: this.state.totalPagesNeeded + 1})}
    this.setState({
      dcrListItems: dcrItems['results'], 
      allDCRListItems: [].concat(this.state.allDCRListItems).concat(dcrItems['results']),
      filteredItems: dcrItems['results'],
      showPagination: false
    });

    /*let reactHandler = this;
    this.props.context.spHttpClient.get(`${this.props.siteurl}/_api/web/lists/getbytitle('DCR Register')/items?$select=*,Approvers/Title,Approvers/Id&$expand=Approvers`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        console.log("DCR RESPONSE: ", response);
        response.json()
          .then((responseJSON: any) => {
            console.log("DCR LIST: ", responseJSON.value);
            reactHandler.setState({
              dcrListItems: responseJSON.value,
              filteredItems: responseJSON.value
            });
            this.dcrSelection.setItems(reactHandler.state.dcrListItems, false);
          });
      });*/
  }

  public priorityDropDownChanged(event, value) {
    /*let {dcrListItems, dcrSelection} = this.state;
    let selectedIndex = dcrSelection.getSelectedIndices()[0];
    let newDCRItems = [].concat(dcrListItems);
    newDCRItems[selectedIndex]['Priority'] = value.text;*/
    this.setState({
      dcrPriority: value.text
    });
  }

  public magnitudeDropDownChanged(event, value) {
    /*let {dcrListItems, dcrSelection} = this.state;
    let selectedIndex = dcrSelection.getSelectedIndices()[0];
    let newDCRItems = [].concat(dcrListItems);
    newDCRItems[selectedIndex]['Magnitude'] = value.text; */
    this.setState({
      dcrMagnitude: value.text
    });
  }

  public async onDCRSelectionChanged() {
    console.log("DCR SELECTION: ", this.dcrSelection.getSelection());
    const reactHandler = this;

    if (this.dcrSelection.getSelectedCount() == 1) {
      const dcrItem = this.dcrSelection.getSelection()[0];

      
      this.props.context.spHttpClient.get(`${this.props.siteurl}/_api/web/lists/getbytitle('Approver List')/items()?$expand=Approver&$select=*,Approver/Title&$filter=Linked_x0020_DCRId eq ${dcrItem['ID']}`, SPHttpClient.configurations.v1)
          .then((response: SPHttpClientResponse) => {
            response.json()
              .then((responseJSON: any) => {
                console.log("GET APPROVERS RESPONSE JSON", responseJSON);
                reactHandler.setState({
                  'approversListItems': responseJSON.value
                });
              });
          });

      let history = await sp.web.lists.getByTitle("DCR Event").items.filter(`DCRId eq ${dcrItem['Id']}`).orderBy(`Created`, false).get();

      this.setState({
        historyItems: [].concat(history)
      })

      console.log("DCR HISTORY: ", history);

    
      this.setState({
        isOpen: true,
        messageBarShow: false,
        panelDCRDetails: true,
        panelDocumentDetails: false,
        dcrId: dcrItem['Id'],
        dcrRequest: dcrItem['Request'],
        dcrReason: dcrItem['Reason'],
        dcrStatus: dcrItem['Status'],
        dcrMagnitude: dcrItem['Magnitude'],
        dcrPriority: dcrItem['Priority'],
        documentId: dcrItem['Linked_x0020_DocumentId']
      });
    } else {
      this.setState({
        isOpen: false
      })
    }
  }

  @autobind
  private async saveDCRChanges() {
    try {
      /*for (let i=0, len=this.state.addUsers.length; i<len; i++) {
        sp.web.lists.getByTitle('Approver List').items.add({
          Title: getGUID(),
          Source: 'One-time',
          ApproverId: this.state.addUsers[i],
          Linked_x0020_DCRId: this.state.dcrId
        });
      }*/

      let updateReturn = await sp.web.lists.getByTitle('DCR Register').items.getById(this.state.dcrId).update({
        Request: this.state.dcrRequest,
        Reason: this.state.dcrReason,
        Magnitude: this.state.dcrMagnitude,
        Priority: this.state.dcrPriority,
        ApproversId: {"results": this.state.addUsers}
      })

      let dcrList = this.state.dcrListItems;
      let allDCRList = this.state.allDCRListItems;

      let newRecordInfo = await sp.web.lists.getByTitle('DCR Register').items.getById(this.state.dcrId).select("*,Approvers/Title,Approvers/Id,Author/Title,Author/Id,Linked_x0020_Document/Title,Linked_x0020_Document/Description").expand("Approvers,Author,Linked_x0020_Document").get();
      let indexOfObjectInList = dcrList.map((item) =>{return item['ID'];}).indexOf(this.state.dcrId);
      let indexOfObjectInFullList = allDCRList.map((item) => {return item['ID'];}).indexOf(this.state.dcrId);

      dcrList[indexOfObjectInList] = newRecordInfo;
      allDCRList[indexOfObjectInFullList] = newRecordInfo;

      this.setState({
        allDCRListItems: [].concat(allDCRList),
        dcrListItems: [].concat(dcrList),
        messageBarShow: true,
        messageBarType: MessageBarType.success,
        messageBarValue: "DCR record has been successfully updated",
        addUsers: [],
        isOpen: false
      });
    } catch(e) {
      this.setState({
        messageBarShow: true,
        isOpen: false,
        messageBarType: MessageBarType.error,
        messageBarValue: "An error occurred while update the DCR Record"
      });
    }
  }

  private setAddUsers = (usersToAdd) => {
    this.setState({
      addUsers: usersToAdd
    })
  }

  private discardDCRChanges() {
    this.setState({
      isOpen: false
    });
  }

  public async filterItems(searchString) {
    console.log("FILTERING ITEMS", searchString);
    console.log(this.state.dcrListItems);
    let newFilteredList = [];
    let filterString = `(substringof('${searchString}', Request)) or (substringof('${searchString}', Reason)) or (substringof('${searchString}', Magnitude)) or (substringof('${searchString}', Priority)) or (substringof('${searchString}', Reference)) or (substringof('${searchString}', Status))`
    let getItems = await sp.web.lists.getByTitle("DCR Register").items.top(this.props.itemsPerPage).filter(filterString).select('*,Approvers/Title,Approvers/Id').expand("Approvers").getPaged();
    this.pagedItemCollection = getItems;

    if (this.pagedItemCollection.hasNext) {
      this.setState({totalPagesNeeded: 2});
    } else {
      this.setState({totalPagesNeeded: 1});
    }
    
    this.setState({
      showPagination: false,
      allDCRListItems: [].concat(getItems['results']),
      dcrListItems: [].concat(getItems['results']),
      currentSelectedPage: 1,
      currentSearchString: searchString
    })
  }

  public clearItems() {
    this.setState({
      allDCRListItems: [].concat(),
      showPagination: false,
      currentSelectedPage: 1,
      totalPagesNeeded: 1,
      currentSearchString: ''
    });
    this.getSPDCRData();
  }

  private DCRListHeaderRender = (props, defaultRender) => {
    return (
      <Sticky stickyPosition={StickyPositionType.Header} isScrollSynced>
        {defaultRender!({...props})}
      </Sticky>
    )
  }

  private _renderRow = (props, defaultRender) => {
    return defaultRender({styles: { cell: {whiteSpace: 'normal'}}, ...props});
  }

  private _paginationChanged = async (pageNumber) => {
    this.setState({currentSelectedPage: pageNumber});
    if (pageNumber == this.state.totalPagesNeeded) {
      if (await this.pagedItemCollection.hasNext) {
          let newItems = await this.pagedItemCollection.getNext();
          this.pagedItemCollection = newItems;
          if (await this.pagedItemCollection.hasNext) {
            this.setState({
              totalPagesNeeded: this.state.totalPagesNeeded + 1
            })
          }
          this.setState({
            dcrListItems: [].concat(newItems['results']),
            allDCRListItems: [].concat(this.state.allDCRListItems).concat(newItems['results']),
            showPagination: false,
          });

        } else {
          let documentStart;
          if (pageNumber == 1) {
            documentStart = 0;
          } else {
            documentStart = ((pageNumber-1) * this.props.itemsPerPage);
          }
          let dcrListItems = [];
          for (let i=documentStart, len=documentStart + this.props.itemsPerPage; i<len; i++) {
            if (this.state.allDCRListItems[i] != undefined) {
              dcrListItems.push(this.state.allDCRListItems[i]);
            }
          }
          this.setState({
            dcrListItems: [].concat(dcrListItems)
          })
        }
    } else {
      let documentStart;
      if (pageNumber == 1) {
        documentStart = 0;
      } else {
        documentStart = ((pageNumber - 1) * this.props.itemsPerPage);
      }
      let dcrListItems = [];
      for (var j=documentStart, len2=documentStart + this.props.itemsPerPage; j<len2; j++) {
        if (this.state.allDCRListItems[j] != undefined) {
          dcrListItems.push(this.state.allDCRListItems[j]);
        }
      }
      this.setState({
        dcrListItems: [].concat(dcrListItems)
      })
    }
  }

  public render(): React.ReactElement<IDcrChangeRequestsProps> {
    console.log("THIS: ", this);
    const {dcrListItems, isOpen, dcrRequest, dcrReason, dcrMagnitude, dcrPriority,
           approversListItems, messageBarShow, messageBarType, messageBarValue,
           filteredItems, dcrStatus, documentId, dcrId, historyItems, showPagination,
           currentSelectedPage, totalPagesNeeded, currentSearchString} = this.state;
    const reactHander = this;
    return (
      <div className={ styles.dcrChangeRequests }>
        {(messageBarShow) ? <MessageBar messageBarType={messageBarType}>{messageBarValue}</MessageBar> : <span>&nbsp;</span>}
        <div style={{height: '80vh', position: 'relative'}}>
        <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
          <Sticky stickyPosition={StickyPositionType.Header}>
            <SearchBox value={currentSearchString} onSearch={newValue => this.filterItems(newValue)} onClear={() => this.clearItems()}/>
          </Sticky>
          <DetailsList items={dcrListItems} columns={this._dcrListColumns} selection={this.dcrSelection}
            onRenderDetailsHeader={this.DCRListHeaderRender} onRenderRow={this._renderRow}></DetailsList>
          {(showPagination) ? <Pagination currentPage={currentSelectedPage} totalPages={totalPagesNeeded} hideFirstPageJump={true} hideLastPageJump={true} onChange={(page) => this._paginationChanged(page)}></Pagination> : ''} 
        </ScrollablePane>
        </div>
        <Panel isOpen={isOpen} type={PanelType.medium} isFooterAtBottom={true} isBlocking={false}>
          <Pivot>
            <PivotItem headerText="Details">
              <DCRPanel description="test" context={this.props.context} documentId={documentId} headerLabelStyle={{root: {fontSize: "20px"}}}
                headerLabel="Some Info" onClose={() => {}} afterSuccess={() => {}} mode="edit" currentUser={this.state.currentUser} hidePanel={() => {}} 
                approversListItems={[]} selectedDocument={this.approversSelection}></DCRPanel>
              {/*<br />
              <TextField label="Request" multiline rows={3} value={dcrRequest} maxLength={255} onChange={this._dcrRequestFieldChanged}/>
              <br />
              <TextField label="Reason" multiline rows={3} value={dcrReason} maxLength={255} onChange={this._dcrReasonFieldChanged}/>
              <br />
              <Stack horizontal tokens={this._stackTokens} className={this.stackStyles}>
                <Stack.Item className={this.stackItemStyles}>
                  <Dropdown options={this._statusOptions} disabled label="Status" selectedKey={dcrStatus} dropdownWidth={500}></Dropdown>
                </Stack.Item >
                <Stack.Item className={this.stackItemStyles}>
                  <Dropdown options={this._magnitudeOptions} selectedKey={dcrMagnitude} onChange={this.magnitudeDropDownChanged} label="Magnitude"></Dropdown>
                </Stack.Item>
                <Stack.Item className={this.stackItemStyles}>
                  <Dropdown options={this._priorityOptions} selectedKey={dcrPriority} onChange={this.priorityDropDownChanged} label="Priority"></Dropdown>
                </Stack.Item> 
              </Stack>
              <br />
              <Approvers showLabel={true} context={this.props.context} mode={'DCR'} documentId={documentId} updateApprovers={this.setAddUsers} selectedDCR={this.dcrSelection} dcrId={dcrId}></Approvers>
              <br />
              <Separator></Separator>
              <Stack horizontal>
                <PrimaryButton text="Save changes" onClick={this.saveDCRChanges}/>
                <DefaultButton text="Discard changes" onClick={this.discardDCRChanges}/>
              </Stack>
              {/*<Label>Approver List</Label> 
              <DetailsList items={approversListItems} columns={this._approverListColumns} selection={approversSelection}></DetailsList>
              <br />
              <DefaultButton text="Add Approver" iconProps={{iconName: 'Add'}} onClick={this._addDCRApprover}/>
              <br />
              {/*<DefaultButton title="Add Selected Users" onClick={this.addSelectedUsers}></DefaultButton>*/}
            </PivotItem>
            <PivotItem headerText="History">
                <DetailsList items={historyItems} columns={this._historyListColumns} selectionMode={SelectionMode.none} onRenderRow={this._renderRow}></DetailsList>
                <DefaultButton text="Close" onClick={this.discardDCRChanges}></DefaultButton>
            </PivotItem>
          </Pivot>
        </Panel>
      </div>
    );
  }
}
