import * as React from 'react';
import { createRoot } from 'react-dom/client';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import CompanyIntranet from '../src/webparts/companyIntranet/components/CompanyIntranet';
import type { ICompanyIntranetProps } from '../src/webparts/companyIntranet/components/ICompanyIntranetProps';

/**
 * Standalone preview harness.
 *
 * Renders the real SML Cosmos web part outside SharePoint with a mocked
 * SPFx context, so the design, animations and layout can be reviewed before
 * anything is deployed. The component detects localhost and switches to demo
 * content automatically.
 *
 * Optional query parameters:
 *   ?video=https://youtu.be/...   preview the featured video section
 *   ?logo=https://.../logo.png    preview a custom company logo
 */
const params: URLSearchParams = new URLSearchParams(window.location.search);

const context = {
  pageContext: {
    web: {
      absoluteUrl: 'https://demo.sharepoint.com/sites/cosmos',
      serverRelativeUrl: ''
    }
  }
} as unknown as WebPartContext;

const props: ICompanyIntranetProps = {
  title: 'SML Cosmos',
  subtitle: 'Your Digital Universe.',
  logoUrl: params.get('logo') || '',
  newsListTitle: 'Company News',
  eventsListTitle: 'Company Events',
  eventSource: 'both',
  quickLinksJson: '',
  youtubeUrl: params.get('video') || '',
  youtubeTitle: 'SML Cosmos — year in space',
  statsListTitle: 'Company Stats',
  featuredAppsListTitle: 'Featured Apps',
  departmentsListTitle: 'Departments',
  clientLogosListTitle: 'Client Logos',
  ctaHeading: 'Safety first. Report an incident',
  ctaText: 'See something unsafe? One click reaches the safety team — every report keeps the crew home safe.',
  ctaButtonText: 'Report an incident',
  ctaUrl: '/Lists/QuickLinks',
  hrContact: 'HR: hr@smlcosmos.com · ext. 201',
  itContact: 'IT support: it@smlcosmos.com · ext. 100',
  isDarkTheme: false,
  environmentMessage: 'Local preview',
  userDisplayName: 'Aarav',
  context
};

const container: HTMLElement | null = document.getElementById('preview-root');

if (container) {
  createRoot(container).render(<CompanyIntranet {...props} />);
}
