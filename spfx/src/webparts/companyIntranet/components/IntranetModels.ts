export interface INewsItem {
  Id: number;
  Title: string;
  NewsSummary?: string;
  NewsLink?: {
    Url: string;
    Description?: string;
  };
  NewsImage?: string;
  Modified: string;
}

/** Raw item from the SharePoint "Company Events" list. */
export interface IEventItem {
  Id: number;
  Title: string;
  EventDate: string;
  EventEndDate?: string;
  EventLocation?: string;
  EventDescription?: string;
  RegistrationLink?: {
    Url: string;
    Description?: string;
  };
}

/** Normalised event rendered by the web part (calendar or company list). */
export interface IIntranetEvent {
  id: string;
  title: string;
  start: string;
  end?: string;
  location?: string;
  description?: string;
  link?: string;
  source: 'calendar' | 'company';
  organizer?: string;
  isAllDay?: boolean;
}

export interface IGraphDateTime {
  dateTime: string;
  timeZone?: string;
}

export interface IGraphCalendarEvent {
  id: string;
  subject?: string;
  start: IGraphDateTime;
  end: IGraphDateTime;
  location?: {
    displayName?: string;
  };
  bodyPreview?: string;
  webLink?: string;
  isAllDay?: boolean;
  organizer?: {
    emailAddress?: {
      name?: string;
    };
  };
}

export interface IGraphCalendarResponse {
  value: IGraphCalendarEvent[];
}

export interface IQuickLink {
  label: string;
  url: string;
  icon?: string;
}

/** Row from the "Company Stats" list — a KPI number. */
export interface IStatItem {
  Id: number;
  Title: string;
  Value: string;
  Icon?: string;
  InHero?: boolean;
}

/** Row from the "Featured Apps" list. */
export interface IFeaturedApp {
  Id: number;
  Title: string;
  Description?: string;
  Icon?: string;
  Url?: {
    Url: string;
    Description?: string;
  };
}

/** Row from the "Departments" list. */
export interface IDepartment {
  Id: number;
  Title: string;
  Description?: string;
  Icon?: string;
  Head?: string;
  Url?: {
    Url: string;
    Description?: string;
  };
}

/** Item from the "Client Logos" picture library. */
export interface IClientLogo {
  Id: number;
  Title?: string;
  File?: {
    ServerRelativeUrl: string;
  };
}

/** Row from the "Leadership" list. Title = name; order the carousel with SortOrder. */
export interface ILeaderItem {
  Id: number;
  Title: string;
  Role?: string;
  SortOrder?: number;
  LeaderPhoto?: {
    Url: string;
    Description?: string;
  };
  LeaderLinkedIn?: {
    Url: string;
    Description?: string;
  };
}

/** Row from the "Celebrations" list: one row per employee. Title = name. */
export interface ICelebrationItem {
  Id: number;
  Title: string;
  PersonBirthday?: string;
  PersonJoinDate?: string;
  PersonEmail?: string;
  PersonDepartment?: string;
  PersonPhoto?: {
    Url: string;
    Description?: string;
  };
}
