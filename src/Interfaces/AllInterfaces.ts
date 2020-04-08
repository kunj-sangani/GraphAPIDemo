
export interface IUsers {
    businessPhones: string[];
    displayName: string;
    givenName: string;
    jobTitle: string;
    mail: string;
    mobilePhone: string;
    officeLocation: string;
    preferredLanguage: string;
    surname: string;
    userPrincipalName: string;
    id: string;
}

export interface IMail {
    subject: string;
    bodyPreview: string;
    receivedDateTime: string;
    webLink: string;
    body: {
        contentType: string;
        content: string;
    };
    from: IUserMailDetails;
    toRecipients: IUserMailDetails[];
}

export interface IUserMailDetails {
    emailAddress: {
        name: string;
        address: string;
    };
}

export interface IGroup {
    displayName: string;
    description: string;
}

export interface ISites {
    displayName: string;
    webUrl: string;
    createdDateTime: String;
    lastModifiedDateTime: string;
    description?: string;
}

export interface IAllChats {
    id: string;
    createdDateTime: String;
    lastModifiedDateTime: string;
}

export interface IAllMessages {
    body: {
        contentType: String;
        content: string;
    };
    from: {
        user: {
            displayName: string;
        }
    };
}

export interface IEvents {
    subject: String;
    bodyPreview: string;
    webLink: string;
    body: {
        contentType: string;
        content: string;
    };
    start: {
        dateTime: string;
        timeZone: string;
    };
    end: {
        dateTime: string;
        timeZone: string;
    };
}

export interface IRecentFiles {
    name: string;
    webUrl: string;
    fileSystemInfo?: {
        createdDateTime: String;
        lastModifiedDateTime: string;
    };
}

export interface IPeopleIWork {
    displayName: string;
    scoredEmailAddresses: IPeopleEmail[];
}

export interface IPeopleEmail {
    address: string;
}

export interface IItemModifiedByMe {
    lastUsed: {
        lastAccessedDateTime: string;
        lastModifiedDateTime: string;
    };
    resourceVisualization: {
        title: string;
    };
    resourceReference: {
        webUrl: string;
    };
}

export interface IItemtrendingaroundme {
    resourceVisualization: {
        title: string;
        containerDisplayName: string;
        containerWebUrl: string;
    };
    resourceReference: {
        webUrl: string;
    };
}

export interface IUsersSchedule {
    scheduleItems: IEachSchedule[];
}

export interface IEachSchedule {
    status: string;
    start: {
        dateTime: string;
        timeZone: string;
    };
    end: {
        dateTime: string;
        timeZone: string;
    };
}