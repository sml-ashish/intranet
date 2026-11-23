import type { WebPartContext } from '@microsoft/sp-webpart-base';
import type { IQuickLink } from './IntranetModels';

export interface ICompanyIntranetProps {
  title: string;
  subtitle: string;
  logoUrl: string;
  newsListTitle: string;
  eventsListTitle: string;
  eventSource: 'calendar' | 'company' | 'both';
  quickLinksJson: string;
  youtubeUrl: string;
  youtubeTitle: string;
  videoUrl: string;
  videoPosterUrl: string;
  videoSubtitle: string;
  heroImagesListTitle: string;
  heroLinks: IQuickLink[];
  statsListTitle: string;
  featuredAppsListTitle: string;
  departmentsListTitle: string;
  clientLogosListTitle: string;
  leadersListTitle: string;
  celebrationsListTitle: string;
  ctaHeading: string;
  ctaText: string;
  ctaButtonText: string;
  ctaUrl: string;
  hrContact: string;
  itContact: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  userDisplayName: string;
  context: WebPartContext;
}
