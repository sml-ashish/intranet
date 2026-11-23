import * as React from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'CompanyIntranetWebPartStrings';
import CompanyIntranet from './components/CompanyIntranet';
import { ICompanyIntranetProps } from './components/ICompanyIntranetProps';
import type { IQuickLink } from './components/IntranetModels';

/** Left-hand hero links: label + icon defaults (URLs are filled in the property pane). */
const HERO_LINK_DEFAULTS: Array<{ label: string; icon: string }> = [
  { label: 'Time Sheets', icon: 'Clock' },
  { label: 'Time Off', icon: 'Airplane' },
  { label: 'Benefits', icon: 'Heart' },
  { label: 'Projects', icon: 'ProjectCollection' },
  { label: 'Helpdesk', icon: 'Headset' },
  { label: 'Directory', icon: 'ContactList' },
  { label: 'Org Chart', icon: 'Org' },
  { label: 'Expenses', icon: 'Money' }
];

export interface ICompanyIntranetWebPartProps {
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
  heroLink1Label: string;
  heroLink1Url: string;
  heroLink2Label: string;
  heroLink2Url: string;
  heroLink3Label: string;
  heroLink3Url: string;
  heroLink4Label: string;
  heroLink4Url: string;
  heroLink5Label: string;
  heroLink5Url: string;
  heroLink6Label: string;
  heroLink6Url: string;
  heroLink7Label: string;
  heroLink7Url: string;
  heroLink8Label: string;
  heroLink8Url: string;
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
}

export default class CompanyIntranetWebPart extends BaseClientSideWebPart<ICompanyIntranetWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _root: Root | undefined;

  public render(): void {
    const element: React.ReactElement<ICompanyIntranetProps> = React.createElement(
      CompanyIntranet,
      {
        title: this.properties.title || strings.DefaultTitle,
        subtitle: this.properties.subtitle || strings.DefaultSubtitle,
        logoUrl: this.properties.logoUrl || '',
        newsListTitle: this.properties.newsListTitle || strings.DefaultNewsListTitle,
        eventsListTitle: this.properties.eventsListTitle || strings.DefaultEventsListTitle,
        eventSource: this.properties.eventSource || 'both',
        quickLinksJson: this.properties.quickLinksJson || '',
        youtubeUrl: this.properties.youtubeUrl || '',
        youtubeTitle: this.properties.youtubeTitle || '',
        videoUrl: this.properties.videoUrl || '',
        videoPosterUrl: this.properties.videoPosterUrl || '',
        videoSubtitle: this.properties.videoSubtitle || '',
        heroImagesListTitle: this.properties.heroImagesListTitle || 'Hero Images',
        heroLinks: this._getHeroLinks(),
        statsListTitle: this.properties.statsListTitle || strings.DefaultStatsListTitle,
        featuredAppsListTitle: this.properties.featuredAppsListTitle || strings.DefaultFeaturedAppsListTitle,
        departmentsListTitle: this.properties.departmentsListTitle || strings.DefaultDepartmentsListTitle,
        clientLogosListTitle: this.properties.clientLogosListTitle || strings.DefaultClientLogosListTitle,
        leadersListTitle: this.properties.leadersListTitle || 'Leadership',
        celebrationsListTitle: this.properties.celebrationsListTitle || 'Celebrations',
        ctaHeading: this.properties.ctaHeading || strings.DefaultCtaHeading,
        ctaText: this.properties.ctaText || strings.DefaultCtaText,
        ctaButtonText: this.properties.ctaButtonText || strings.DefaultCtaButtonText,
        ctaUrl: this.properties.ctaUrl || strings.DefaultCtaUrl,
        hrContact: this.properties.hrContact || '',
        itContact: this.properties.itContact || '',
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context
      }
    );

    if (!this._root) {
      this._root = createRoot(this.domElement);
    }

    this._root.render(element);
  }

  /** Reads the 8 label/URL pairs from the property pane; blank label falls back to the default. */
  private _getHeroLinks(): IQuickLink[] {
    const props: Record<string, string | undefined> =
      this.properties as unknown as Record<string, string | undefined>;

    return HERO_LINK_DEFAULTS.map((entry: { label: string; icon: string }, index: number): IQuickLink => {
      const n: number = index + 1;
      return {
        label: (props[`heroLink${n}Label`] || '').trim() || entry.label,
        url: (props[`heroLink${n}Url`] || '').trim(),
        icon: entry.icon
      };
    });
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    this._root?.unmount();
    this._root = undefined;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BrandingGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneTextField('subtitle', {
                  label: strings.SubtitleFieldLabel,
                  multiline: true
                }),
                PropertyPaneTextField('logoUrl', {
                  label: strings.LogoUrlFieldLabel,
                  description: strings.LogoUrlFieldDescription
                })
              ]
            },
            {
              groupName: strings.DataGroupName,
              groupFields: [
                PropertyPaneTextField('newsListTitle', {
                  label: strings.NewsListTitleFieldLabel,
                  description: strings.NewsListTitleFieldDescription
                }),
                PropertyPaneTextField('eventsListTitle', {
                  label: strings.EventsListTitleFieldLabel,
                  description: strings.EventsListTitleFieldDescription
                }),
                PropertyPaneChoiceGroup('eventSource', {
                  label: strings.EventSourceFieldLabel,
                  options: [
                    { key: 'calendar', text: strings.EventSourceCalendar },
                    { key: 'company', text: strings.EventSourceCompany },
                    { key: 'both', text: strings.EventSourceBoth }
                  ]
                }),
                PropertyPaneTextField('statsListTitle', {
                  label: strings.StatsListTitleFieldLabel,
                  description: strings.StatsListTitleFieldDescription
                }),
                PropertyPaneTextField('featuredAppsListTitle', {
                  label: strings.FeaturedAppsListTitleFieldLabel,
                  description: strings.FeaturedAppsListTitleFieldDescription
                }),
                PropertyPaneTextField('departmentsListTitle', {
                  label: strings.DepartmentsListTitleFieldLabel,
                  description: strings.DepartmentsListTitleFieldDescription
                }),
                PropertyPaneTextField('clientLogosListTitle', {
                  label: strings.ClientLogosListTitleFieldLabel,
                  description: strings.ClientLogosListTitleFieldDescription
                }),
                PropertyPaneTextField('leadersListTitle', {
                  label: 'Leadership list title',
                  description: 'List with Title (name), Role, Photo, LinkedIn and SortOrder columns. Default: Leadership'
                }),
                PropertyPaneTextField('celebrationsListTitle', {
                  label: 'Celebrations list title',
                  description: 'List with Title (name), Birthday, Join date, Email, Department and Photo columns. Default: Celebrations'
                })
              ]
            },
            {
              groupName: strings.MediaGroupName,
              groupFields: [
                PropertyPaneTextField('youtubeUrl', {
                  label: strings.YouTubeUrlFieldLabel,
                  description: strings.YouTubeUrlFieldDescription
                }),
                PropertyPaneTextField('youtubeTitle', {
                  label: strings.YouTubeTitleFieldLabel,
                  description: 'Headline shown on the video (works for YouTube and uploaded video).'
                }),
                PropertyPaneTextField('videoSubtitle', {
                  label: 'Video sub-line',
                  description: 'Small line under the headline, e.g. "See how your benefits have changed!"'
                }),
                PropertyPaneTextField('videoUrl', {
                  label: 'Uploaded video file (optional)',
                  description: 'Link to an MP4 stored in SharePoint. If filled, it is used instead of the YouTube link.'
                }),
                PropertyPaneTextField('videoPosterUrl', {
                  label: 'Video cover image (optional)',
                  description: 'Link to a JPG/PNG shown before play. YouTube videos use their own thumbnail if this is empty.'
                })
              ]
            },
            {
              groupName: 'Hero background slideshow',
              groupFields: [
                PropertyPaneTextField('heroImagesListTitle', {
                  label: 'Hero images library name',
                  description: 'Document library holding the crane photos. Default: Hero Images'
                })
              ]
            },
            {
              groupName: 'Hero links (left side)',
              groupFields: HERO_LINK_DEFAULTS.reduce(
                (fields: ReturnType<typeof PropertyPaneTextField>[], entry: { label: string; icon: string }, index: number) => {
                  const n: number = index + 1;
                  fields.push(
                    PropertyPaneTextField(`heroLink${n}Label`, { label: `Link ${n} text`, placeholder: entry.label }),
                    PropertyPaneTextField(`heroLink${n}Url`, { label: `Link ${n} address (URL)`, placeholder: 'https://… or /sites/…' })
                  );
                  return fields;
                },
                []
              )
            },
            {
              groupName: strings.QuickLinksGroupName,
              groupFields: [
                PropertyPaneTextField('quickLinksJson', {
                  label: strings.QuickLinksFieldLabel,
                  description: strings.QuickLinksFieldDescription,
                  multiline: true,
                  rows: 6
                })
              ]
            },
            {
              groupName: strings.CtaGroupName,
              groupFields: [
                PropertyPaneTextField('ctaHeading', {
                  label: strings.CtaHeadingFieldLabel
                }),
                PropertyPaneTextField('ctaText', {
                  label: strings.CtaTextFieldLabel,
                  multiline: true,
                  rows: 3
                }),
                PropertyPaneTextField('ctaButtonText', {
                  label: strings.CtaButtonTextFieldLabel
                }),
                PropertyPaneTextField('ctaUrl', {
                  label: strings.CtaUrlFieldLabel
                }),
                PropertyPaneTextField('hrContact', {
                  label: strings.HrContactFieldLabel
                }),
                PropertyPaneTextField('itContact', {
                  label: strings.ItContactFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
