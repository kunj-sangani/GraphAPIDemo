import * as React from 'react';
import styles from './MsalBase.module.scss';
import { IMsalBaseProps } from './IMsalBaseProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  IUsers, IMail, IGroup, ISites, IAllChats, IAllMessages, IEvents, IRecentFiles, IPeopleIWork, IItemModifiedByMe,
  IItemtrendingaroundme, IUsersSchedule, IEachSchedule
} from "../../../Interfaces/AllInterfaces";
import { HttpClient, HttpClientResponse, IHttpClientOptions } from "@microsoft/sp-http";
import { TagPicker, IBasePicker, ITag } from 'office-ui-fabric-react/lib/Pickers';
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  SelectionMode,
  IColumn,
} from 'office-ui-fabric-react/lib/DetailsList';

export interface IMsalBaseState {
  allUsersInOrg?: IUsers[];
  currentUser?: IUsers;
  allMails?: IMail[];
  allGroups?: IGroup[];
  allSites?: ISites[];
  allChats?: IAllChats[];
  allMessages?: IAllMessages[];
  allEvents?: IEvents[];
  allRecentFiles?: IRecentFiles[];
  allPeopleIWork?: IPeopleIWork[];
  allItemModifiedByMe?: IItemModifiedByMe[];
  selectedUser?: string;
  allItemtrendingaroundme?: IItemtrendingaroundme[];
  allUsersSchedule?: IUsersSchedule[];
}

const columns: IColumn[] = [
  {
    key: 'subject',
    name: 'subject',
    minWidth: 1,
    maxWidth: 150,
    isResizable: true,
    onRender: (item: IMail) => {
      return <a href={item.webLink} target="__blank">{item.subject}</a>;
    }
  }, {
    key: 'receivedDateTime',
    name: 'receivedDateTime',
    minWidth: 1,
    maxWidth: 150,
    isResizable: true,
    onRender: (item: IMail) => {
      return <span>{item.receivedDateTime}</span>;
    }
  }, {
    key: 'from',
    name: 'from',
    minWidth: 1,
    maxWidth: 150,
    isResizable: true,
    onRender: (item: IMail) => {
      return <span>{item.from.emailAddress.name}</span>;
    }
  }, {
    key: 'bodyPreview',
    name: 'bodyPreview',
    minWidth: 5,
    maxWidth: 150,
    isResizable: true,
    onRender: (item: IMail) => {
      return <span>{item.bodyPreview}</span>;
    }
  }
];

const groupcolumns: IColumn[] = [
  {
    key: 'displayName',
    name: 'displayName',
    minWidth: 5,
    maxWidth: 150,
    isResizable: true,
    onRender: (item: IGroup) => {
      return <span>{item.displayName}</span>;
    }
  }, {
    key: 'description',
    name: 'description',
    minWidth: 5,
    maxWidth: 150,
    isResizable: true,
    onRender: (item: IGroup) => {
      return <span>{item.description}</span>;
    }
  }
];

const sitescolumns: IColumn[] = [
  {
    key: 'displayName',
    name: 'Name',
    minWidth: 5,
    maxWidth: 150,
    isResizable: true,
    onRender: (item: ISites) => {
      return <a href={item.webUrl} target="__blank">{item.displayName}</a>;
    }
  }, {
    key: 'description',
    name: 'description',
    minWidth: 5,
    maxWidth: 150,
    isResizable: true,
    onRender: (item: ISites) => {
      return <span>{item.description}</span>;
    }
  }, {
    key: 'createdDateTime',
    name: 'created Date',
    minWidth: 5,
    maxWidth: 150,
    isResizable: true,
    onRender: (item: ISites) => {
      return <span>{item.createdDateTime}</span>;
    }
  }, {
    key: 'lastModifiedDateTime',
    name: 'Modified Date',
    minWidth: 5,
    maxWidth: 150,
    isResizable: true,
    onRender: (item: ISites) => {
      return <span>{item.lastModifiedDateTime}</span>;
    }
  }
];

const messageColumn: IColumn[] = [
  {
    key: 'From',
    name: 'From User',
    minWidth: 5,
    maxWidth: 150,
    isResizable: true,
    onRender: (item: IAllMessages) => {
      return <span>{item.from.user.displayName}</span>;
    }
  }, {
    key: 'Body',
    name: 'Body',
    minWidth: 5,
    maxWidth: 150,
    isResizable: true,
    onRender: (item: IAllMessages) => {
      return <span>{item.body.content}</span>;
    }
  }
];

const evntsColumn: IColumn[] = [
  {
    key: 'subject',
    name: 'subject',
    minWidth: 1,
    maxWidth: 150,
    isResizable: true,
    onRender: (item: IEvents) => {
      return <a href={item.webLink} target="__blank">{item.subject}</a>;
    }
  }, {
    key: 'bodyPreview',
    name: 'bodyPreview',
    minWidth: 1,
    maxWidth: 150,
    isResizable: true,
    onRender: (item: IEvents) => {
      return <span>{item.bodyPreview}</span>;
    }
  }, {
    key: 'start',
    name: 'start',
    minWidth: 1,
    maxWidth: 150,
    isResizable: true,
    onRender: (item: IEvents) => {
      return <span>{item.start.dateTime}</span>;
    }
  }, {
    key: 'end',
    name: 'end',
    minWidth: 1,
    maxWidth: 150,
    isResizable: true,
    onRender: (item: IEvents) => {
      return <span>{item.end.dateTime}</span>;
    }
  }
];

const recentFilesColumn: IColumn[] = [
  {
    key: 'name',
    name: 'name',
    minWidth: 1,
    maxWidth: 150,
    isResizable: true,
    onRender: (item: IRecentFiles) => {
      return <a href={item.webUrl} target="__blank">{item.name}</a>;
    }
  }, {
    key: 'createdDateTime',
    name: 'createdDateTime',
    minWidth: 1,
    maxWidth: 150,
    isResizable: true,
    onRender: (item: IRecentFiles) => {
      return <span>{item.fileSystemInfo.createdDateTime}</span>;
    }
  }, {
    key: 'lastModifiedDateTime',
    name: 'lastModifiedDateTime',
    minWidth: 1,
    maxWidth: 150,
    isResizable: true,
    onRender: (item: IRecentFiles) => {
      return <span>{item.fileSystemInfo.lastModifiedDateTime}</span>;
    }
  }
];

const peopleIWorkColumn: IColumn[] = [
  {
    key: 'DisplayName',
    name: 'Display Name',
    minWidth: 1,
    maxWidth: 150,
    isResizable: true,
    onRender: (item: IPeopleIWork) => {
      return <span>{item.displayName}</span>;
    }
  }, {
    key: 'Email',
    name: 'Email',
    minWidth: 1,
    maxWidth: 150,
    isResizable: true,
    onRender: (item: IPeopleIWork) => {
      return <span>{item.scoredEmailAddresses[0].address}</span>;
    }
  }
];

const itemModifiedByMeColumn: IColumn[] = [
  {
    key: 'Name',
    name: 'Name',
    minWidth: 1,
    maxWidth: 150,
    isResizable: true,
    onRender: (item: IItemModifiedByMe) => {
      return <a href={item.resourceReference.webUrl} target="__blank">{item.resourceVisualization.title}</a>;
    }
  },
  {
    key: 'lastAccessedDateTime',
    name: 'last Accessed Date',
    minWidth: 1,
    maxWidth: 150,
    isResizable: true,
    onRender: (item: IItemModifiedByMe) => {
      return <span>{item.lastUsed.lastAccessedDateTime}</span>;
    }
  },
  {
    key: 'lastModifiedDateTime',
    name: 'last Modified Date',
    minWidth: 1,
    maxWidth: 150,
    isResizable: true,
    onRender: (item: IItemModifiedByMe) => {
      return <span>{item.lastUsed.lastModifiedDateTime}</span>;
    }
  }
];

const itemtrendingaroundmeColumn: IColumn[] = [
  {
    key: 'Name',
    name: 'Name',
    minWidth: 1,
    maxWidth: 150,
    isResizable: true,
    onRender: (item: IItemtrendingaroundme) => {
      return <a href={item.resourceReference.webUrl} target="__blank">{item.resourceVisualization.title}</a>;
    }
  },
  {
    key: 'containerDisplayName',
    name: 'container',
    minWidth: 1,
    maxWidth: 150,
    isResizable: true,
    onRender: (item: IItemtrendingaroundme) => {
      return <a href={item.resourceVisualization.containerWebUrl} target="__blank">{item.resourceVisualization.containerDisplayName}</a>;
    }
  }
];

const usersScheduleColumn: IColumn[] = [
  {
    key: 'status',
    name: 'status',
    minWidth: 1,
    maxWidth: 150,
    isResizable: true,
    onRender: (item: IEachSchedule) => {
      return <span>{item.status}</span>;
    }
  }, {
    key: 'start',
    name: 'start Date',
    minWidth: 1,
    maxWidth: 150,
    isResizable: true,
    onRender: (item: IEachSchedule) => {
      return <span>{item.start.dateTime}</span>;
    }
  }, {
    key: 'end',
    name: 'end Date',
    minWidth: 1,
    maxWidth: 150,
    isResizable: true,
    onRender: (item: IEachSchedule) => {
      return <span>{item.end.dateTime}</span>;
    }
  }
];


export default class MsalBase extends React.Component<IMsalBaseProps, IMsalBaseState> {

  public AllChatsColumns: IColumn[] = [
    {
      key: 'displayName',
      name: 'Name',
      minWidth: 5,
      maxWidth: 150,
      isResizable: true,
      onRender: (item: IAllChats) => {
        let userName = this.state.allUsersInOrg.
          filter(val => { return val.mail !== this.props.context.pageContext.user.email; }).
          filter((val) => { return item.id.indexOf(val.id) > 0; });
        return <a onClick={() => this.getTeamsMessage(item.id)} >{userName[0].displayName}</a>;
      }
    }, {
      key: 'createdDateTime',
      name: 'created Date',
      minWidth: 5,
      maxWidth: 150,
      isResizable: true,
      onRender: (item: IAllChats) => {
        return <span>{item.createdDateTime}</span>;
      }
    }, {
      key: 'lastModifiedDateTime',
      name: 'Modified Date',
      minWidth: 5,
      maxWidth: 150,
      isResizable: true,
      onRender: (item: IAllChats) => {
        return <span>{item.lastModifiedDateTime}</span>;
      }
    }
  ];

  constructor(props: IMsalBaseProps) {
    super(props);
    this.state = {
      allUsersInOrg: [] as IUsers[]
    };
  }

  public componentDidMount() {
    this.props.msalObjcet.acquireTokenSilent({ scopes: ['Directory.Read.All', 'User.Read', 'User.Read.All'] }).then((val) => {
      this.callMSGraph('https://graph.microsoft.com/v1.0/users', val.accessToken).then((allusers) => {
        this.setState({
          allUsersInOrg: allusers.value
        });
      }).catch((error) => {

      });
    }).catch((error) => {
      console.error(error);
    });
  }

  public callMSGraph = (url: string, token: string): Promise<any> => {
    return new Promise<any>((resolve, reject) => {
      this.props.context.httpClient.fetch(url, HttpClient.configurations.v1, {
        headers: { Authorization: `Bearer ${token}` }
      }).then((response: HttpClientResponse) => {
        response.json().then((users) => {
          resolve(users);
        });
      }).catch((responseError) => {
        reject(responseError);
      });
    });
  }

  public callMSGraphPOST = (url: string, token: string, body: string): Promise<any> => {
    const requestHeaders: Headers = new Headers();
    requestHeaders.append('Content-type', 'application/json');
    requestHeaders.append('Cache-Control', 'no-cache');
    requestHeaders.append('Authorization', `Bearer ${token}`);
    const httpClientOptions: IHttpClientOptions = {
      body: body,
      headers: requestHeaders
    };
    return new Promise<any>((resolve, reject) => {
      this.props.context.httpClient.post(url, HttpClient.configurations.v1, httpClientOptions).then((response: HttpClientResponse) => {
        response.json().then((schedule) => {
          resolve(schedule);
        });
      }).catch((responseError) => {
        reject(responseError);
      });
    });
  }

  public userMail = () => {
    if (!this.state.allMails) {
      this.props.msalObjcet.acquireTokenSilent({ scopes: ['Mail.Read'] }).then((val) => {
        this.callMSGraph('https://graph.microsoft.com/v1.0/me/messages', val.accessToken).then((allmails) => {
          this.setState({
            allMails: allmails.value
          });
        }).catch((errormail) => {
          console.log(errormail);
        });
      }).catch((error) => {
        console.error(error);
      });
    } else {
      this.setState({
        allMails: null
      });
    }
  }

  public userGroups = () => {
    if (this.state.selectedUser) {
      if (!this.state.allGroups) {
        let userId = this.state.allUsersInOrg.filter((val) => { return val.displayName === this.state.selectedUser; })[0].id;
        this.props.msalObjcet.acquireTokenSilent({ scopes: ['Directory.Read.All', 'User.Read.All'] }).then((val) => {
          this.callMSGraph(`https://graph.microsoft.com/v1.0/users/${userId}/joinedTeams`, val.accessToken).then((allothergroup) => {
            this.setState({
              allGroups: allothergroup.value
            });
          }).catch((errormail) => {
            console.log(errormail);
          });
        }).catch((error) => {
          console.error(error);
        });
      } else {
        this.setState({
          allGroups: null
        });
      }
    } else {
      if (!this.state.allGroups) {
        this.props.msalObjcet.acquireTokenSilent({ scopes: ['Directory.Read.All', 'User.Read.All'] }).then((val) => {
          this.callMSGraph('https://graph.microsoft.com/v1.0/me/memberOf', val.accessToken).then((allothergroup) => {
            this.setState({
              allGroups: allothergroup.value
            });
          }).catch((errormail) => {
            console.log(errormail);
          });
        }).catch((error) => {
          console.error(error);
        });
      } else {
        this.setState({
          allGroups: null
        });
      }
    }
  }

  public sharePointSites = () => {
    if (!this.state.allSites) {
      this.props.msalObjcet.acquireTokenSilent({ scopes: ['Sites.Read.All'] }).then((val) => {
        this.callMSGraph('https://graph.microsoft.com/v1.0/sites?search=*', val.accessToken).then((allSharePointSites) => {
          this.setState({
            allSites: allSharePointSites.value
          });
        }).catch((errorsites) => {
          console.log(errorsites);
        });
      }).catch((error) => {
        console.error(error);
      });
    } else {
      this.setState({
        allSites: null
      });
    }
  }

  public getTeamsChat = () => {
    if (!this.state.allChats) {
      this.props.msalObjcet.acquireTokenSilent({ scopes: ['Chat.Read'] }).then((val) => {
        this.callMSGraph('https://graph.microsoft.com/beta/me/chats', val.accessToken).then((allTeamsChat) => {
          this.setState({
            allChats: allTeamsChat.value
          });
        }).catch((errorsites) => {
          console.log(errorsites);
        });
      }).catch((error) => {
        console.error(error);
      });
    } else {
      this.setState({
        allChats: null,
        allMessages: null
      });
    }
  }

  public getTeamsMessage = (chatID: string) => {
    this.props.msalObjcet.acquireTokenSilent({ scopes: ['Chat.Read'] }).then((val) => {
      this.callMSGraph(`https://graph.microsoft.com/beta/me/chats/${chatID}/messages`, val.accessToken).then((allTeamsMessage) => {
        this.setState({
          allMessages: allTeamsMessage.value
        });
      }).catch((errorsites) => {
        console.log(errorsites);
      });
    }).catch((error) => {
      console.error(error);
    });
  }

  public getMyEvents = () => {
    let startDate = new Date;
    let EndDate = new Date;
    EndDate.setDate(EndDate.getDate() + 7);
    if (!this.state.allEvents) {
      this.props.msalObjcet.acquireTokenSilent({ scopes: ['Calendars.Read'] }).then((val) => {
        this.callMSGraph(`https://graph.microsoft.com/v1.0/me/calendarview?startdatetime=${startDate.toISOString()}&enddatetime=${EndDate.toISOString()}`, val.accessToken).then((allEvents) => {
          this.setState({
            allEvents: allEvents.value
          });
        }).catch((errorevents) => {
          console.log(errorevents);
        });
      }).catch((error) => {
        console.error(error);
      });
    } else {
      this.setState({
        allEvents: null
      });
    }
  }

  public getMyRecentFiles = () => {
    if (!this.state.allRecentFiles) {
      this.props.msalObjcet.acquireTokenSilent({ scopes: ['Files.Read', 'Files.Read.All', 'Sites.Read.All'] }).then((val) => {
        this.callMSGraph(`https://graph.microsoft.com/v1.0/me/drive/recent`, val.accessToken).then((allRecentFiles) => {
          this.setState({
            allRecentFiles: allRecentFiles.value
          });
        }).catch((errorevents) => {
          console.log(errorevents);
        });
      }).catch((error) => {
        console.error(error);
      });
    } else {
      this.setState({
        allRecentFiles: null
      });
    }
  }

  public getPeopleIWork = () => {
    if (this.state.selectedUser) {
      if (!this.state.allPeopleIWork) {
        let userId = this.state.allUsersInOrg.filter((val) => { return val.displayName === this.state.selectedUser; })[0].id;
        this.props.msalObjcet.acquireTokenSilent({ scopes: ['People.Read.All'] }).then((val) => {
          this.callMSGraph(`https://graph.microsoft.com/v1.0/users/${userId}/people`, val.accessToken).then((allPeopleIWork) => {
            this.setState({
              allPeopleIWork: allPeopleIWork.value
            });
          }).catch((errorevents) => {
            console.log(errorevents);
          });
        }).catch((error) => {
          console.error(error);
        });
      } else {
        this.setState({
          allPeopleIWork: null
        });
      }
    } else {
      if (!this.state.allPeopleIWork) {
        this.props.msalObjcet.acquireTokenSilent({ scopes: ['People.Read', 'People.Read.All'] }).then((val) => {
          this.callMSGraph(`https://graph.microsoft.com/v1.0/me/people`, val.accessToken).then((allPeopleIWork) => {
            this.setState({
              allPeopleIWork: allPeopleIWork.value
            });
          }).catch((errorevents) => {
            console.log(errorevents);
          });
        }).catch((error) => {
          console.error(error);
        });
      } else {
        this.setState({
          allPeopleIWork: null
        });
      }
    }
  }

  public getItemModifiedByMe = () => {
    if (this.state.selectedUser) {
      if (!this.state.allItemModifiedByMe) {
        let userId = this.state.allUsersInOrg.filter((val) => { return val.displayName === this.state.selectedUser; })[0].id;
        this.props.msalObjcet.acquireTokenSilent({ scopes: ['Sites.Read.All'] }).then((val) => {
          this.callMSGraph(`https://graph.microsoft.com/beta/users/${userId}/insights/used`, val.accessToken).then((allItemModifiedByMe) => {
            this.setState({
              allItemModifiedByMe: allItemModifiedByMe.value
            });
          }).catch((errorevents) => {
            console.log(errorevents);
          });
        }).catch((error) => {
          console.error(error);
        });
      } else {
        this.setState({
          allItemModifiedByMe: null
        });
      }
    }
    else {
      if (!this.state.allItemModifiedByMe) {
        this.props.msalObjcet.acquireTokenSilent({ scopes: ['Sites.Read.All'] }).then((val) => {
          this.callMSGraph(`https://graph.microsoft.com/beta/me/insights/used`, val.accessToken).then((allItemModifiedByMe) => {
            this.setState({
              allItemModifiedByMe: allItemModifiedByMe.value
            });
          }).catch((errorevents) => {
            console.log(errorevents);
          });
        }).catch((error) => {
          console.error(error);
        });
      } else {
        this.setState({
          allItemModifiedByMe: null
        });
      }
    }
  }

  public getItemtrendingaroundme = () => {
    if (!this.state.allItemtrendingaroundme) {
      let userId = this.state.allUsersInOrg.filter((val) => { return val.displayName === this.state.selectedUser; })[0].id;
      this.props.msalObjcet.acquireTokenSilent({ scopes: ['Sites.Read.All'] }).then((val) => {
        this.callMSGraph(`https://graph.microsoft.com/beta/users/${userId}/insights/trending`, val.accessToken).then((allItemtrendingaroundme) => {
          this.setState({
            allItemtrendingaroundme: allItemtrendingaroundme.value
          });
        }).catch((errorevents) => {
          console.log(errorevents);
        });
      }).catch((error) => {
        console.error(error);
      });
    } else {
      this.setState({
        allItemtrendingaroundme: null
      });
    }
  }

  public getUsersSchedule = () => {
    if (!this.state.allUsersSchedule) {
      let userMail = this.state.allUsersInOrg.filter((val) => { return val.displayName === this.state.selectedUser; })[0].mail;
      let startDate = new Date;
      let endDate = new Date;
      endDate.setDate(endDate.getDate() + 7);
      let requestBody = { "Schedules": [`${userMail}`], "StartTime": { "dateTime": `${startDate.toISOString()}`, "timeZone": "Pacific Standard Time" }, "EndTime": { "dateTime": `${endDate.toISOString()}`, "timeZone": "Pacific Standard Time" }, "availabilityViewInterval": "30" };
      this.props.msalObjcet.acquireTokenSilent({ scopes: ['Calendars.Read'] }).then((val) => {
        this.callMSGraphPOST(`https://graph.microsoft.com/v1.0/me/calendar/getschedule`, val.accessToken, JSON.stringify(requestBody)).then((allUsersSchedule) => {
          // console.log(allUsersSchedule.value);
          this.setState({
            allUsersSchedule: allUsersSchedule.value
          });
        }).catch((errorevents) => {
          console.log(errorevents);
        });
      }).catch((error) => {
        console.error(error);
      });
    } else {
      this.setState({
        allUsersSchedule: null
      });
    }
  }

  private getTextFromItem(item: ITag): string {
    return item.name;
  }

  private onFilterChanged = (filterText: string, tagList: ITag[]): ITag[] => {
    return filterText
      ? this.state.allUsersInOrg
        .filter(tag => tag.displayName.toLowerCase().indexOf(filterText.toLowerCase()) === 0).map(item => ({ key: item.displayName, name: item.displayName }))
      : [];
  }

  private userItemSelected = (item: ITag): ITag | null => {
    this.setState({
      selectedUser: item.name
    });
    return item;
  }

  private userItemChanged = (item: ITag[]): ITag[] | null => {
    if (item.length > 0) {
      this.setState({
        selectedUser: item[0].name
      });
    } else {
      this.setState({
        selectedUser: null
      });
    }
    return item;
  }

  public userInfo = (selectedUser?: string) => {
    if (!this.state.currentUser) {
      let currentUser;
      if (selectedUser) {
        currentUser = this.state.allUsersInOrg.filter(
          (v) => { return v.displayName === selectedUser; }
        );
      } else {
        currentUser = this.state.allUsersInOrg.filter(
          (v) => { return v.displayName === this.props.context.pageContext.user.displayName; }
        );
      }
      this.setState({
        currentUser: currentUser[0]
      });
    } else {
      this.setState({
        currentUser: null
      });
    }
  }

  private _getKey(item: any, index?: number): string {
    return item.key;
  }

  private _onItemInvoked(item: any): void {
    //alert(`Item invoked: ${item.name}`);
  }

  public render(): React.ReactElement<IMsalBaseProps> {
    return (
      <div className={styles.msalBase} >
        <div>
          <div className={styles.row}>
            <TagPicker
              removeButtonAriaLabel="Remove"
              onResolveSuggestions={this.onFilterChanged}
              getTextFromItem={this.getTextFromItem}
              pickerSuggestionsProps={{
                suggestionsHeaderText: 'Suggested Persons',
                noResultsFoundText: 'No Person Found',
              }}
              itemLimit={1}
              disabled={false}
              onItemSelected={this.userItemSelected}
              onChange={this.userItemChanged}
            />
          </div>
        </div>
        <div style={this.state.selectedUser ? { display: "none" } : { display: "block" }}>
          <div className={styles.row}>
            <div className={styles.title}><h1>Welcome to Graph API demo {this.props.context.pageContext.user.displayName}</h1></div>
            <a className={styles.button} onClick={() => this.userInfo()}>
              <span className={styles.label}>My Details</span>
            </a>
            <a className={styles.button} onClick={this.userMail}>
              <span className={styles.label}>My Emails</span>
            </a>
            <a className={styles.button} onClick={this.userGroups}>
              <span className={styles.label}>My Groups</span>
            </a>
            <a className={styles.button} onClick={this.sharePointSites}>
              <span className={styles.label}>SharePoint Sites I have access</span>
            </a>
            <a className={styles.button} onClick={this.getTeamsChat}>
              <span className={styles.label}>Teams Chat</span>
            </a>
            <a className={styles.button} onClick={this.getMyEvents}>
              <span className={styles.label}>My Events for Next Week</span>
            </a>
            <a className={styles.button} onClick={this.getMyRecentFiles}>
              <span className={styles.label}>My recent Files</span>
            </a>
            <a className={styles.button} onClick={this.getPeopleIWork}>
              <span className={styles.label}>People I work with</span>
            </a>
            <a className={styles.button} onClick={this.getItemModifiedByMe}>
              <span className={styles.label}>Items viewed modified by me</span>
            </a>
          </div>
          <div className={styles.row}>
            {this.state.currentUser && Object.keys(this.state.currentUser).map((userval) => {
              return (<div>
                <span>{userval} : </span>
                <span>{this.state.currentUser[userval]}</span>
              </div>);
            })}
            {this.state.allMails &&
              <DetailsList
                items={this.state.allMails}
                compact={false}
                columns={columns}
                selectionMode={SelectionMode.none}
                getKey={this._getKey}
                setKey="none"
                layoutMode={DetailsListLayoutMode.justified}
                isHeaderVisible={true}
                onItemInvoked={this._onItemInvoked}
              />
            }
            {this.state.allGroups &&
              <DetailsList
                items={this.state.allGroups}
                compact={false}
                columns={groupcolumns}
                selectionMode={SelectionMode.none}
                getKey={this._getKey}
                setKey="none"
                layoutMode={DetailsListLayoutMode.justified}
                isHeaderVisible={true}
                onItemInvoked={this._onItemInvoked}
              />
            }
            {this.state.allSites &&
              <DetailsList
                items={this.state.allSites}
                compact={false}
                columns={sitescolumns}
                selectionMode={SelectionMode.none}
                getKey={this._getKey}
                setKey="none"
                layoutMode={DetailsListLayoutMode.justified}
                isHeaderVisible={true}
                onItemInvoked={this._onItemInvoked}
              />
            }
            {this.state.allChats &&
              <DetailsList
                items={this.state.allChats}
                compact={false}
                columns={this.AllChatsColumns}
                selectionMode={SelectionMode.none}
                getKey={this._getKey}
                setKey="none"
                layoutMode={DetailsListLayoutMode.justified}
                isHeaderVisible={true}
                onItemInvoked={this._onItemInvoked}
              />
            }
            {this.state.allMessages &&
              <DetailsList
                items={this.state.allMessages}
                compact={false}
                columns={messageColumn}
                selectionMode={SelectionMode.none}
                getKey={this._getKey}
                setKey="none"
                layoutMode={DetailsListLayoutMode.justified}
                isHeaderVisible={true}
                onItemInvoked={this._onItemInvoked}
              />
            }
            {this.state.allEvents &&
              <DetailsList
                items={this.state.allEvents}
                compact={false}
                columns={evntsColumn}
                selectionMode={SelectionMode.none}
                getKey={this._getKey}
                setKey="none"
                layoutMode={DetailsListLayoutMode.justified}
                isHeaderVisible={true}
                onItemInvoked={this._onItemInvoked}
              />
            }
            {this.state.allRecentFiles &&
              <DetailsList
                items={this.state.allRecentFiles}
                compact={false}
                columns={recentFilesColumn}
                selectionMode={SelectionMode.none}
                getKey={this._getKey}
                setKey="none"
                layoutMode={DetailsListLayoutMode.justified}
                isHeaderVisible={true}
                onItemInvoked={this._onItemInvoked}
              />
            }
            {this.state.allPeopleIWork &&
              <DetailsList
                items={this.state.allPeopleIWork}
                compact={false}
                columns={peopleIWorkColumn}
                selectionMode={SelectionMode.none}
                getKey={this._getKey}
                setKey="none"
                layoutMode={DetailsListLayoutMode.justified}
                isHeaderVisible={true}
                onItemInvoked={this._onItemInvoked}
              />
            }
            {this.state.allItemModifiedByMe &&
              <DetailsList
                items={this.state.allItemModifiedByMe}
                compact={false}
                columns={itemModifiedByMeColumn}
                selectionMode={SelectionMode.none}
                getKey={this._getKey}
                setKey="none"
                layoutMode={DetailsListLayoutMode.justified}
                isHeaderVisible={true}
                onItemInvoked={this._onItemInvoked}
              />
            }
          </div>
        </div>
        <div style={this.state.selectedUser ? { display: "block" } : { display: "none" }}>
          <div className={styles.row}>
            <div className={styles.title}><h1>Welcome to Graph API demo {this.state.selectedUser}</h1></div>
            <a className={styles.button} onClick={() => this.userInfo(this.state.selectedUser)}>
              <span className={styles.label}>My Details</span>
            </a>
            <a className={styles.button} onClick={this.userGroups}>
              <span className={styles.label}>Users Groups</span>
            </a>
            <a className={styles.button} onClick={this.getItemtrendingaroundme}>
              <span className={styles.label}>Items trending around me</span>
            </a>
            <a className={styles.button} onClick={this.getUsersSchedule}>
              <span className={styles.label}>Users Schedule</span>
            </a>
            <a className={styles.button} onClick={this.getPeopleIWork}>
              <span className={styles.label}>People User work with</span>
            </a>
            <a className={styles.button} onClick={this.getItemModifiedByMe}>
              <span className={styles.label}>Items viewed modified by User</span>
            </a>
          </div>
          <div className={styles.row}>
            {this.state.currentUser && Object.keys(this.state.currentUser).map((userval) => {
              return (<div>
                <span>{userval} : </span>
                <span>{this.state.currentUser[userval]}</span>
              </div>);
            })}
            {this.state.allGroups &&
              <DetailsList
                items={this.state.allGroups}
                compact={false}
                columns={groupcolumns}
                selectionMode={SelectionMode.none}
                getKey={this._getKey}
                setKey="none"
                layoutMode={DetailsListLayoutMode.justified}
                isHeaderVisible={true}
                onItemInvoked={this._onItemInvoked}
              />
            }
            {this.state.allItemtrendingaroundme &&
              <DetailsList
                items={this.state.allItemtrendingaroundme}
                compact={false}
                columns={itemtrendingaroundmeColumn}
                selectionMode={SelectionMode.none}
                getKey={this._getKey}
                setKey="none"
                layoutMode={DetailsListLayoutMode.justified}
                isHeaderVisible={true}
                onItemInvoked={this._onItemInvoked}
              />
            }
            {this.state.allPeopleIWork &&
              <DetailsList
                items={this.state.allPeopleIWork}
                compact={false}
                columns={peopleIWorkColumn}
                selectionMode={SelectionMode.none}
                getKey={this._getKey}
                setKey="none"
                layoutMode={DetailsListLayoutMode.justified}
                isHeaderVisible={true}
                onItemInvoked={this._onItemInvoked}
              />
            }
            {this.state.allItemModifiedByMe &&
              <DetailsList
                items={this.state.allItemModifiedByMe}
                compact={false}
                columns={itemModifiedByMeColumn}
                selectionMode={SelectionMode.none}
                getKey={this._getKey}
                setKey="none"
                layoutMode={DetailsListLayoutMode.justified}
                isHeaderVisible={true}
                onItemInvoked={this._onItemInvoked}
              />
            }
            {this.state.allUsersSchedule &&
              <DetailsList
                items={this.state.allUsersSchedule[0].scheduleItems}
                compact={false}
                columns={usersScheduleColumn}
                selectionMode={SelectionMode.none}
                getKey={this._getKey}
                setKey="none"
                layoutMode={DetailsListLayoutMode.justified}
                isHeaderVisible={true}
                onItemInvoked={this._onItemInvoked}
              />
            }
          </div>
        </div>
      </div>
    );
  }
}
