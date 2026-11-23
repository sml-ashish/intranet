import * as React from 'react';
import styles from './CompanyIntranet.module.scss';
import type { ICompanyIntranetProps } from './ICompanyIntranetProps';
import type {
  IClientLogo,
  IDepartment,
  IEventItem,
  IFeaturedApp,
  IGraphCalendarEvent,
  IGraphCalendarResponse,
  IIntranetEvent,
  INewsItem,
  IQuickLink,
  IStatItem
} from './IntranetModels';
import BrandLogo from './BrandLogo';
import LeadershipSection from './LeadershipSection';
import CelebrationsSection from './CelebrationsSection';
import {
  MSGraphClientV3,
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import { Icon } from '@fluentui/react/lib/Icon';
import { initializeIcons } from '@fluentui/react/lib/Icons';

initializeIcons();

interface INewsResponse {
  value: INewsItem[];
}

interface ICompanyEventsResponse {
  value: IEventItem[];
}

interface ICompanyIntranetState {
  news: INewsItem[];
  companyEvents: IEventItem[];
  calendarEvents: IIntranetEvent[];
  quickLinks: IQuickLink[];
  stats: IStatItem[];
  featuredApps: IFeaturedApp[];
  departments: IDepartment[];
  clientLogos: IClientLogo[];
  isNewsLoading: boolean;
  isCompanyLoading: boolean;
  isCalendarLoading: boolean;
  isStatsLoading: boolean;
  newsError: string | undefined;
  companyError: string | undefined;
  calendarError: string | undefined;
  heroImages: string[];
  heroIndex: number;
  isVideoPlaying: boolean;
}

/** Default shortcut tiles — used when no JSON override and no list rows exist. */
const DEFAULT_QUICK_LINKS: IQuickLink[] = [
  { label: 'HR & Payroll', url: '/sitepages/hr', icon: 'People' },
  { label: 'IT Support', url: '/Lists/QuickLinks', icon: 'Settings' },
  { label: 'Safety', url: '/Lists/QuickLinks', icon: 'Shield' },
  { label: 'Policies', url: '/Shared Documents/Policies', icon: 'DocumentSet' },
  { label: 'Documents', url: '/Shared Documents', icon: 'FolderOpen' },
  { label: 'Fleet', url: '/Lists/QuickLinks', icon: 'DeliveryTruck' }
];

/** List read with no property-pane override — staff manage these in SharePoint. */
const QUICK_LINKS_LIST_TITLE = 'Quick Links';
const CALENDAR_LOOKAHEAD_DAYS = 30;

/** Demo content used only when the web part runs outside a real SharePoint page. */

/** ISO string for a local time `dayOffset` days from today at `hour`:`minute`. */
function demoDate(dayOffset: number, hour: number, minute: number = 0): string {
  const d: Date = new Date();
  d.setDate(d.getDate() + dayOffset);
  d.setHours(hour, minute, 0, 0);
  return d.toISOString();
}

const DEMO_NEWS: INewsItem[] = [
  {
    Id: 1,
    Title: 'SML Cosmos opens its new digital workplace',
    NewsSummary:
      'The new intranet brings news, events, documents and your calendar together on one home page.',
    Modified: new Date(Date.now() - 86400000).toISOString()
  },
  {
    Id: 2,
    Title: 'Constellation Day — everyone is invited',
    NewsSummary:
      'Join us for an afternoon of awards, food and fun as we celebrate another year of moving the world forward.',
    Modified: new Date(Date.now() - 3 * 86400000).toISOString()
  },
  {
    Id: 3,
    Title: 'Safety first: quarterly fleet inspection begins',
    NewsSummary:
      'All vehicles will be inspected this month. Schedule your slot with the safety team to avoid delays.',
    Modified: new Date(Date.now() - 6 * 86400000).toISOString()
  }
];

const DEMO_COMPANY_EVENTS: IEventItem[] = [
  {
    Id: 1,
    Title: 'Company Town Hall',
    EventDate: demoDate(1, 15, 0),
    EventEndDate: demoDate(1, 16, 0),
    EventLocation: 'Main Auditorium',
    EventDescription: 'Quarterly results, new projects and open Q&A with the leadership team.'
  },
  {
    Id: 2,
    Title: 'Constellation Day Celebration',
    EventDate: demoDate(5, 17, 0),
    EventEndDate: demoDate(5, 20, 0),
    EventLocation: 'Rooftop Lawn',
    EventDescription: 'Awards, food, music and the annual team photograph.'
  },
  {
    Id: 3,
    Title: 'Quarterly Business Review',
    EventDate: demoDate(9, 11, 0),
    EventEndDate: demoDate(9, 12, 30),
    EventLocation: 'Meeting Room 2',
    EventDescription: 'Review of delivery, safety and customer satisfaction for the quarter.'
  }
];

const DEMO_CALENDAR_EVENTS: IIntranetEvent[] = [
  {
    id: 'cal-demo-1',
    title: 'Sprint planning — Project Cosmos',
    start: demoDate(1, 10, 0),
    end: demoDate(1, 11, 0),
    location: 'Microsoft Teams',
    description: 'Plan the next sprint with the design and platform teams.',
    source: 'calendar',
    isAllDay: false
  },
  {
    id: 'cal-demo-2',
    title: '1:1 with manager',
    start: demoDate(2, 16, 0),
    end: demoDate(2, 16, 30),
    location: 'Cabin 12',
    source: 'calendar',
    isAllDay: false
  },
  {
    id: 'cal-demo-3',
    title: 'Design review: intranet homepage',
    start: demoDate(4, 14, 0),
    end: demoDate(4, 15, 0),
    location: 'Microsoft Teams',
    description: 'Walk through the new landing page animations and content blocks.',
    source: 'calendar',
    isAllDay: false
  }
];

const DEMO_QUICK_LINKS: IQuickLink[] = DEFAULT_QUICK_LINKS;

const DEMO_STATS: IStatItem[] = [
  { Id: 1, Title: 'Fleet utilisation', Value: '92%', Icon: 'DeliveryTruck', InHero: true },
  { Id: 2, Title: 'Active sites', Value: '14', Icon: 'Location', InHero: true },
  { Id: 3, Title: 'Days without incident', Value: '213', Icon: 'Shield', InHero: true },
  { Id: 4, Title: 'Fleet size', Value: '250+', Icon: 'DeliveryTruck' },
  { Id: 5, Title: 'Years of operation', Value: '38', Icon: 'Clock' },
  { Id: 6, Title: 'People', Value: '1,200+', Icon: 'People' },
  { Id: 7, Title: 'Projects delivered', Value: '5,400', Icon: 'FavoriteStar' }
];

const DEMO_APPS: IFeaturedApp[] = [
  { Id: 1, Title: 'Safety Portal', Description: 'Report hazards, near-misses and toolbox talks.', Icon: 'Shield', Url: { Url: '/Lists/QuickLinks' } },
  { Id: 2, Title: 'Crane Booking', Description: 'Reserve cranes and operators for your site.', Icon: 'Calendar', Url: { Url: '/Lists/QuickLinks' } },
  { Id: 3, Title: 'HR Self-Service', Description: 'Leave, payslips and personal details in one place.', Icon: 'People', Url: { Url: '/Lists/QuickLinks' } },
  { Id: 4, Title: 'IT Support', Description: 'Raise a ticket and track progress.', Icon: 'Phone', Url: { Url: '/Lists/QuickLinks' } }
];

/** An application shown in the carousel below the hero. */
interface IAppCarouselItem {
  name: string;
  description: string;
  url: string;
  /** Optional image URL (e.g. /SiteAssets/logos/nxss.png). Falls back to the built-in mark. */
  logoUrl?: string;
  /** Built-in mark used when no logoUrl is set. */
  mark: 'pulse' | 'nxss' | 'weather';
  color: string;
}

/**
 * Applications in the carousel. Replace each `url` with the real site address
 * and, optionally, set `logoUrl` to use your own logo image.
 */
const APP_CAROUSEL_ITEMS: IAppCarouselItem[] = [
  {
    name: 'Customer Pulse',
    description: 'Customer feedback and satisfaction insights.',
    url: 'https://customerpulse.example.com',
    mark: 'pulse',
    color: '#dc2626'
  },
  {
    name: 'NXSS',
    description: 'Next-generation service and support suite.',
    url: 'https://nxss.example.com',
    mark: 'nxss',
    color: '#166534'
  },
  {
    name: 'Weather App',
    description: 'Live site weather and lift-readiness forecasts.',
    url: 'https://weather.example.com',
    mark: 'weather',
    color: '#0369a1'
  }
];

function AppMark({ mark }: { mark: IAppCarouselItem['mark'] }): React.ReactElement {
  const common = { viewBox: '0 0 32 32', fill: 'none', stroke: '#fff', strokeWidth: 2.4, strokeLinecap: 'round' as const, strokeLinejoin: 'round' as const, 'aria-hidden': true };
  if (mark === 'pulse') {
    return (<svg {...common}><polyline points="3,17 9,17 12,8 17,24 21,13 24,17 29,17" /></svg>);
  }
  if (mark === 'weather') {
    return (
      <svg {...common}>
        <circle cx="12" cy="12" r="4" />
        <path d="M12 3v2M3 12h2M5.6 5.6l1.4 1.4M18.4 5.6L17 7" />
        <path d="M11 25h12a5 5 0 0 0 .6-9.96A7 7 0 0 0 10.2 16.4 4.3 4.3 0 0 0 11 25z" />
      </svg>
    );
  }
  return (
    <svg {...common}>
      <path d="M7 24V8l18 16V8" />
    </svg>
  );
}

function AppCarousel(): React.ReactElement {
  // Each group is repeated so it is wider than the screen; two identical groups
  // scroll as one seamless, endless loop.
  const group: IAppCarouselItem[] = [...APP_CAROUSEL_ITEMS, ...APP_CAROUSEL_ITEMS, ...APP_CAROUSEL_ITEMS];

  const renderGroup = (suffix: string, hidden: boolean): React.ReactElement => (
    <ul className={styles.appGroup} aria-hidden={hidden ? 'true' : undefined}>
      {group.map((app: IAppCarouselItem, index: number): React.ReactElement => (
        <li key={`${app.name}-${suffix}-${index}`} className={styles.appItem}>
          <a
            className={styles.appLink}
            href={app.url}
            target="_blank"
            rel="noopener noreferrer"
            aria-label={`Open ${app.name}`}
            tabIndex={hidden ? -1 : undefined}
            title={app.name}
          >
            {app.logoUrl ? (
              <img className={styles.appLogoImg} src={app.logoUrl} alt="" />
            ) : (
              <>
                <span className={styles.appLogo} style={{ background: app.color }}>
                  <AppMark mark={app.mark} />
                </span>
                <span className={styles.appWordmark}>{app.name}</span>
              </>
            )}
          </a>
        </li>
      ))}
    </ul>
  );

  return (
    <section className={`${styles.section} ${styles.appMarqueeSection}`} id="applications" data-reveal="1">
      <div className={styles.appMarquee} aria-label="Our applications">
        <div className={styles.appMarqueeTrack}>
          {renderGroup('a', false)}
          {renderGroup('b', true)}
        </div>
      </div>
    </section>
  );
}

/** "About us" nav item with a dropdown menu. */
function NavAbout(): React.ReactElement {
  const [open, setOpen] = React.useState<boolean>(false);
  const ref = React.useRef<HTMLDivElement>(null);

  React.useEffect((): (() => void) => {
    const onDown = (e: MouseEvent): void => {
      if (ref.current && !ref.current.contains(e.target as Node)) {
        setOpen(false);
      }
    };
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === 'Escape') {
        setOpen(false);
      }
    };
    document.addEventListener('mousedown', onDown);
    document.addEventListener('keydown', onKey);
    return (): void => {
      document.removeEventListener('mousedown', onDown);
      document.removeEventListener('keydown', onKey);
    };
  }, []);

  return (
    <div className={styles.navDropdown} ref={ref}>
      <button
        type="button"
        className={`${styles.navLink} ${styles.navDropdownBtn}`}
        aria-haspopup="true"
        aria-expanded={open}
        onClick={(): void => setOpen(!open)}
      >
        About us
        <Icon iconName="ChevronDown" className={`${styles.navDropdownChevron} ${open ? styles.navDropdownChevronOpen : ''}`} />
      </button>
      {open && (
        <div className={styles.navMenu} role="menu">
          <a className={styles.navMenuItem} role="menuitem" href="#leadership" onClick={(): void => setOpen(false)}>
            Leadership
          </a>
        </div>
      )}
    </div>
  );
}


const DEMO_DEPARTMENTS: IDepartment[] = [
  { Id: 1, Title: 'Operations', Description: 'Fleet, logistics and day-to-day site delivery.', Icon: 'DeliveryTruck', Head: 'ops@smlcosmos.com', Url: { Url: '/Lists/QuickLinks' } },
  { Id: 2, Title: 'Projects & Engineering', Description: 'Project planning, lifts and engineering studies.', Icon: 'Globe', Head: 'projects@smlcosmos.com', Url: { Url: '/Lists/QuickLinks' } },
  { Id: 3, Title: 'Safety & Compliance', Description: 'Training, audits and incident follow-up.', Icon: 'Shield', Head: 'safety@smlcosmos.com', Url: { Url: '/Lists/QuickLinks' } }
];

const DEMO_LOGOS: IClientLogo[] = [
  { Id: 1, Title: 'Acme Logistics' },
  { Id: 2, Title: 'Northwind Steel' },
  { Id: 3, Title: 'Contoso Shipping' }
];

export default class CompanyIntranet extends React.Component<ICompanyIntranetProps, ICompanyIntranetState> {

  private _rootRef = React.createRef<HTMLElement>();
  private _observer: IntersectionObserver | undefined;
  private _countObserver: IntersectionObserver | undefined;
  private _observedReveals = new Set<Element>();
  private _observedCounts = new Set<HTMLElement>();
  private _fallbackTimer: number | undefined;
  private _slideTimer: number | undefined;

  constructor(props: ICompanyIntranetProps) {
    super(props);

    this.state = {
      news: [],
      companyEvents: [],
      calendarEvents: [],
      quickLinks: [],
      stats: [],
      featuredApps: [],
      departments: [],
      clientLogos: [],
      isNewsLoading: false,
      isCompanyLoading: false,
      isCalendarLoading: false,
      isStatsLoading: false,
      newsError: undefined,
      companyError: undefined,
      calendarError: undefined,
      heroImages: [],
      heroIndex: 0,
      isVideoPlaying: false
    };
  }

  public componentDidMount(): void {
    this._ensureFonts();
    this._loadNews().then(() => undefined).catch(() => undefined);
    this._loadCompanyEvents().then(() => undefined).catch(() => undefined);
    this._loadCalendarEvents().then(() => undefined).catch(() => undefined);
    this._loadQuickLinks().then(() => undefined).catch(() => undefined);
    this._loadStats().then(() => undefined).catch(() => undefined);
    this._loadFeaturedApps().then(() => undefined).catch(() => undefined);
    this._loadDepartments().then(() => undefined).catch(() => undefined);
    this._loadClientLogos().then(() => undefined).catch(() => undefined);
    this._loadHeroImages().then(() => undefined).catch(() => undefined);
    this._setupReveal();
  }

  public componentDidUpdate(prevProps: ICompanyIntranetProps): void {
    if (prevProps.newsListTitle !== this.props.newsListTitle) {
      this._loadNews().then(() => undefined).catch(() => undefined);
    }

    if (prevProps.eventSource !== this.props.eventSource) {
      this._loadCompanyEvents().then(() => undefined).catch(() => undefined);
      this._loadCalendarEvents().then(() => undefined).catch(() => undefined);
    }

    if (prevProps.eventSource === this.props.eventSource
      && prevProps.eventsListTitle !== this.props.eventsListTitle) {
      this._loadCompanyEvents().then(() => undefined).catch(() => undefined);
    }

    if (prevProps.quickLinksJson !== this.props.quickLinksJson) {
      this._loadQuickLinks().then(() => undefined).catch(() => undefined);
    }

    if (prevProps.statsListTitle !== this.props.statsListTitle) {
      this._loadStats().then(() => undefined).catch(() => undefined);
    }

    if (prevProps.featuredAppsListTitle !== this.props.featuredAppsListTitle) {
      this._loadFeaturedApps().then(() => undefined).catch(() => undefined);
    }

    if (prevProps.departmentsListTitle !== this.props.departmentsListTitle) {
      this._loadDepartments().then(() => undefined).catch(() => undefined);
    }

    if (prevProps.clientLogosListTitle !== this.props.clientLogosListTitle) {
      this._loadClientLogos().then(() => undefined).catch(() => undefined);
    }

    if (prevProps.heroImagesListTitle !== this.props.heroImagesListTitle) {
      this._loadHeroImages().then(() => undefined).catch(() => undefined);
    }

    if (prevProps.youtubeUrl !== this.props.youtubeUrl || prevProps.videoUrl !== this.props.videoUrl) {
      this.setState({ isVideoPlaying: false });
    }

    // Sections rendered late (after a list loads) need observing too.
    this._setupReveal();
  }

  public componentWillUnmount(): void {
    if (this._observer) {
      this._observer.disconnect();
      this._observer = undefined;
    }
    if (this._countObserver) {
      this._countObserver.disconnect();
      this._countObserver = undefined;
    }
    if (this._fallbackTimer !== undefined) {
      window.clearTimeout(this._fallbackTimer);
      this._fallbackTimer = undefined;
    }
    this._stopSlideshow();
  }

  /** Loads Poppins + Open Sans once; falls back to Segoe UI if the CDN is blocked. */
  private _ensureFonts(): void {
    if (typeof document === 'undefined' || document.getElementById('smlCosmosFonts')) {
      return;
    }
    const link: HTMLLinkElement = document.createElement('link');
    link.id = 'smlCosmosFonts';
    link.rel = 'stylesheet';
    link.href =
      'https://fonts.googleapis.com/css2?family=Open+Sans:wght@400;500;600;700&family=Poppins:wght@500;600;700&display=swap';
    document.head.appendChild(link);
  }

  /**
   * Staggered reveal-on-scroll for each section, plus count-up for stat
   * numbers. Safe to call repeatedly — sections that appear after data loads
   * are picked up on the next call. Falls back to showing everything
   * immediately when Intersection Observer is unavailable or the user
   * prefers reduced motion.
   */
  private _setupReveal(): void {
    const root: HTMLElement | null = this._rootRef.current;
    if (!root) {
      return;
    }

    const prefersReducedMotion: boolean =
      typeof window.matchMedia === 'function'
        ? window.matchMedia('(prefers-reduced-motion: reduce)').matches
        : false;

    if (prefersReducedMotion || typeof IntersectionObserver === 'undefined') {
      this._revealAllNow();
      return;
    }

    if (!this._observer) {
      this._observer = new IntersectionObserver(
        (entries: IntersectionObserverEntry[]) => {
          entries.forEach((entry: IntersectionObserverEntry) => {
            if (entry.isIntersecting) {
              entry.target.classList.add(styles.revealed);
              if (this._observer) {
                this._observer.unobserve(entry.target);
              }
            }
          });
        },
        { threshold: 0.12, rootMargin: '0px 0px -40px 0px' }
      );

      this._countObserver = new IntersectionObserver(
        (entries: IntersectionObserverEntry[]) => {
          entries.forEach((entry: IntersectionObserverEntry) => {
            if (entry.isIntersecting) {
              this._animateCount(entry.target as HTMLElement);
              if (this._countObserver) {
                this._countObserver.unobserve(entry.target);
              }
            }
          });
        },
        { threshold: 0.4 }
      );
    }

    const nodes: NodeListOf<Element> = root.querySelectorAll('[data-reveal]');
    for (let i: number = 0; i < nodes.length; i++) {
      const node: Element = nodes.item(i);
      if (!this._observedReveals.has(node)) {
        this._observedReveals.add(node);
        this._observer.observe(node);
      }
    }

    const counters: NodeListOf<HTMLElement> = root.querySelectorAll('[data-count]');
    for (let i: number = 0; i < counters.length; i++) {
      const counter: HTMLElement = counters.item(i);
      if (!this._observedCounts.has(counter)) {
        this._observedCounts.add(counter);
        if (this._countObserver) {
          this._countObserver.observe(counter);
        }
      }
    }

    // Safety net so content is never stuck invisible. Re-armed on every call
    // so sections rendered late (after a list loads) are also covered.
    if (this._fallbackTimer !== undefined) {
      window.clearTimeout(this._fallbackTimer);
    }
    this._fallbackTimer = window.setTimeout(() => this._revealAllNow(), 2500);
  }

  /** Reveal/animate everything currently in the DOM (fallback path). */
  private _revealAllNow(): void {
    const root: HTMLElement | null = this._rootRef.current;
    if (!root) {
      return;
    }

    const nodes: NodeListOf<Element> = root.querySelectorAll('[data-reveal]');
    for (let i: number = 0; i < nodes.length; i++) {
      nodes.item(i).classList.add(styles.revealed);
    }

    const counters: NodeListOf<HTMLElement> = root.querySelectorAll('[data-count]');
    for (let i: number = 0; i < counters.length; i++) {
      this._animateCount(counters.item(i));
    }
  }


  /**
   * Animates "92%" / "1,200+" style values from 0 when the stat scrolls in.
   * The element already renders its final value, so this only ever enhances.
   */
  private _animateCount(el: HTMLElement): void {
    // Skip if this element is already animating (observer + fallback can overlap).
    if (el.getAttribute('data-count-animating') === '1') {
      return;
    }

    const raw: string = el.getAttribute('data-count') || el.textContent || '';
    const cleaned: string = raw.replace(/,/g, '');
    const match: RegExpExecArray | null = /^(\d+(?:\.\d+)?)(.*)$/.exec(cleaned);

    if (!match || typeof window.requestAnimationFrame !== 'function') {
      return;
    }

    const prefersReducedMotion: boolean =
      typeof window.matchMedia === 'function'
        ? window.matchMedia('(prefers-reduced-motion: reduce)').matches
        : false;

    if (prefersReducedMotion) {
      return;
    }

    const target: number = parseFloat(match[1]);
    const suffix: string = match[2];
    const decimals: number = (match[1].split('.')[1] || '').length;
    const useComma: boolean = raw.indexOf(',') !== -1;
    const duration: number = 1100;
    const startTime: number = performance.now();

    const format = (value: number): string => {
      const fixed: string = value.toFixed(decimals);
      if (!useComma) {
        return fixed;
      }
      return Number(fixed).toLocaleString('en-US', {
        minimumFractionDigits: decimals,
        maximumFractionDigits: decimals
      });
    };

    el.setAttribute('data-count-animating', '1');
    el.textContent = format(0) + suffix;

    const tick = (now: number): void => {
      const progress: number = Math.max(0, Math.min(1, (now - startTime) / duration));
      if (progress < 1) {
        const eased: number = 1 - Math.pow(1 - progress, 3);
        el.textContent = format(target * eased) + suffix;
        window.requestAnimationFrame(tick);
      } else {
        el.textContent = raw;
        el.removeAttribute('data-count-animating');
      }
    };

    window.requestAnimationFrame(tick);
  }

  /**
   * True when the web part is running outside a real, deployed SharePoint page:
   * the local debug workbench (gulp serve) or the standalone preview harness.
   * In that case live lists/Graph are unavailable, so demo content is shown.
   */
  private _isPreviewMode(): boolean {
    if (typeof window === 'undefined') {
      return false;
    }

    const { hostname, protocol, search } = window.location;
    return (
      protocol === 'file:' ||
      hostname === 'localhost' ||
      hostname === '127.0.0.1' ||
      search.indexOf('debugManifestsFile') !== -1
    );
  }

  // ------------------------------------------------------------------
  // Data loading
  // ------------------------------------------------------------------

  /** Shared REST read used by every optional list section. */
  private _getListItems<T>(listTitle: string, query: string): Promise<T[]> {
    const { context } = this.props;
    const url: string =
      `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items${query}`;

    return context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          throw new Error(`Unable to load the list "${listTitle}" (${response.status}).`);
        }
        return response.json() as Promise<{ value: T[] }>;
      })
      .then((data: { value: T[] }) => data.value || []);
  }

  private _loadNews(): Promise<void> {
    const { context, newsListTitle } = this.props;

    if (this._isPreviewMode()) {
      this.setState({ news: DEMO_NEWS, isNewsLoading: false, newsError: undefined });
      return Promise.resolve();
    }

    if (!newsListTitle) {
      this.setState({ news: [], isNewsLoading: false, newsError: undefined });
      return Promise.resolve();
    }

    const query: string =
      '?$select=Id,Title,NewsSummary,NewsLink,NewsImage,Modified&$orderby=Modified desc&$top=6';

    this.setState({ isNewsLoading: true, newsError: undefined });

    return context.spHttpClient.get(
      `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(newsListTitle)}')/items${query}`,
      SPHttpClient.configurations.v1
    )
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          throw new Error(`Unable to load the news list "${newsListTitle}" (${response.status}).`);
        }
        return response.json() as Promise<INewsResponse>;
      })
      .then((data: INewsResponse) => {
        this.setState({ news: data.value || [], isNewsLoading: false, newsError: undefined });
      })
      .catch((error: Error) => {
        this.setState({ news: [], isNewsLoading: false, newsError: error.message });
      });
  }

  private _loadCompanyEvents(): Promise<void> {
    const { context, eventsListTitle, eventSource } = this.props;

    if (this._isPreviewMode()) {
      this.setState({
        companyEvents: DEMO_COMPANY_EVENTS,
        isCompanyLoading: false,
        companyError: undefined
      });
      return Promise.resolve();
    }

    if (eventSource === 'calendar' || !eventsListTitle) {
      this.setState({ companyEvents: [], isCompanyLoading: false, companyError: undefined });
      return Promise.resolve();
    }

    const query: string =
      '?$select=Id,Title,EventDate,EventEndDate,EventLocation,EventDescription,RegistrationLink' +
      '&$orderby=EventDate asc&$top=20';

    this.setState({ isCompanyLoading: true, companyError: undefined });

    return context.spHttpClient.get(
      `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(eventsListTitle)}')/items${query}`,
      SPHttpClient.configurations.v1
    )
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          throw new Error(`Unable to load the events list "${eventsListTitle}" (${response.status}).`);
        }
        return response.json() as Promise<ICompanyEventsResponse>;
      })
      .then((data: ICompanyEventsResponse) => {
        this.setState({
          companyEvents: data.value || [],
          isCompanyLoading: false,
          companyError: undefined
        });
      })
      .catch((error: Error) => {
        this.setState({ companyEvents: [], isCompanyLoading: false, companyError: error.message });
      });
  }

  /**
   * Reads the signed-in user's Outlook calendar via Microsoft Graph.
   * Requires the "Calendars.Read" delegated permission to be approved in the
   * SharePoint admin center (API access).
   */
  private _loadCalendarEvents(): Promise<void> {
    const { context, eventSource } = this.props;

    if (this._isPreviewMode()) {
      this.setState({
        calendarEvents: DEMO_CALENDAR_EVENTS,
        isCalendarLoading: false,
        calendarError: undefined
      });
      return Promise.resolve();
    }

    if (eventSource === 'company') {
      this.setState({ calendarEvents: [], isCalendarLoading: false, calendarError: undefined });
      return Promise.resolve();
    }

    const start: Date = new Date();
    const end: Date = new Date();
    end.setDate(end.getDate() + CALENDAR_LOOKAHEAD_DAYS);

    this.setState({ isCalendarLoading: true, calendarError: undefined });

    return context.msGraphClientFactory.getClient('3')
      .then((client: MSGraphClientV3) => {
        const request = client
          .api('/me/calendarView')
          .query({
            startDateTime: start.toISOString(),
            endDateTime: end.toISOString(),
            $select: 'id,subject,start,end,location,bodyPreview,webLink,isAllDay,organizer',
            $top: `${CALENDAR_LOOKAHEAD_DAYS * 2}`
          })
          .header('Prefer', 'outlook.timezone="UTC"');

        return request.get() as Promise<IGraphCalendarResponse>;
      })
      .then((response: IGraphCalendarResponse) => {
        const items: IGraphCalendarEvent[] = response && response.value ? response.value : [];
        this.setState({
          calendarEvents: items.map((item: IGraphCalendarEvent): IIntranetEvent =>
            this._calendarToIntranet(item)),
          isCalendarLoading: false,
          calendarError: undefined
        });
      })
      .catch(() => {
        this.setState({
          calendarEvents: [],
          isCalendarLoading: false,
          calendarError:
            'Calendar access is not approved yet. Ask a SharePoint admin to approve the '
            + '"Calendars.Read" permission, or switch the event source to Company events.'
        });
      });
  }

  /** Quick links: JSON override → Quick Links list → defaults. */
  private _loadQuickLinks(): Promise<void> {
    if (this.props.quickLinksJson || this._isPreviewMode()) {
      this.setState({ quickLinks: this._isPreviewMode() && !this.props.quickLinksJson ? DEMO_QUICK_LINKS : [] });
      return Promise.resolve();
    }

    return this._getListItems<{ Title: string; LinkUrl?: { Url: string }; LinkIcon?: string }>(
      QUICK_LINKS_LIST_TITLE,
      '?$select=Id,Title,LinkUrl,LinkIcon&$top=12'
    )
      .then((rows) => {
        const links: IQuickLink[] = rows
          .filter((row) => !!row && typeof row.Title === 'string' && !!(row.LinkUrl && row.LinkUrl.Url))
          .map((row) => ({
            label: row.Title,
            url: row.LinkUrl ? row.LinkUrl.Url : '',
            icon: row.LinkIcon || undefined
          }));
        this.setState({ quickLinks: links });
      })
      .catch(() => {
        // Missing list → fall back to defaults in _getQuickLinks().
        this.setState({ quickLinks: [] });
      });
  }

  private _loadStats(): Promise<void> {
    const { statsListTitle } = this.props;

    if (this._isPreviewMode()) {
      this.setState({ stats: DEMO_STATS, isStatsLoading: false });
      return Promise.resolve();
    }

    if (!statsListTitle) {
      this.setState({ stats: [], isStatsLoading: false });
      return Promise.resolve();
    }

    this.setState({ isStatsLoading: true });

    return this._getListItems<IStatItem>(
      statsListTitle,
      '?$select=Id,Title,Value,Icon,InHero&$top=20'
    )
      .then((items: IStatItem[]) => {
        this.setState({ stats: items, isStatsLoading: false });
      })
      .catch(() => {
        this.setState({ stats: [], isStatsLoading: false });
      });
  }

  private _loadFeaturedApps(): Promise<void> {
    const { featuredAppsListTitle } = this.props;

    if (this._isPreviewMode()) {
      this.setState({ featuredApps: DEMO_APPS });
      return Promise.resolve();
    }

    if (!featuredAppsListTitle) {
      this.setState({ featuredApps: [] });
      return Promise.resolve();
    }

    return this._getListItems<IFeaturedApp>(
      featuredAppsListTitle,
      '?$select=Id,Title,Description,Icon,Url&$top=8'
    )
      .then((items: IFeaturedApp[]) => {
        this.setState({ featuredApps: items });
      })
      .catch(() => {
        this.setState({ featuredApps: [] });
      });
  }

  private _loadDepartments(): Promise<void> {
    const { departmentsListTitle } = this.props;

    if (this._isPreviewMode()) {
      this.setState({ departments: DEMO_DEPARTMENTS });
      return Promise.resolve();
    }

    if (!departmentsListTitle) {
      this.setState({ departments: [] });
      return Promise.resolve();
    }

    return this._getListItems<IDepartment>(
      departmentsListTitle,
      '?$select=Id,Title,Description,Icon,Head,Url&$top=8'
    )
      .then((items: IDepartment[]) => {
        this.setState({ departments: items });
      })
      .catch(() => {
        this.setState({ departments: [] });
      });
  }

  private _loadClientLogos(): Promise<void> {
    const { clientLogosListTitle } = this.props;

    if (this._isPreviewMode()) {
      this.setState({ clientLogos: DEMO_LOGOS });
      return Promise.resolve();
    }

    if (!clientLogosListTitle) {
      this.setState({ clientLogos: [] });
      return Promise.resolve();
    }

    return this._getListItems<IClientLogo>(
      clientLogosListTitle,
      '?$select=Id,Title,File/ServerRelativeUrl&$expand=File&$top=12'
    )
      .then((items: IClientLogo[]) => {
        this.setState({ clientLogos: items });
      })
      .catch(() => {
        this.setState({ clientLogos: [] });
      });
  }

  /** Reads the crane photos from the Hero Images library and starts the slideshow. */
  private _loadHeroImages(): Promise<void> {
    const { heroImagesListTitle } = this.props;

    this._stopSlideshow();

    if (this._isPreviewMode() || !heroImagesListTitle) {
      this.setState({ heroImages: [], heroIndex: 0 });
      return Promise.resolve();
    }

    return this._getListItems<IClientLogo>(
      heroImagesListTitle,
      '?$select=Id,Title,File/ServerRelativeUrl&$expand=File&$top=12'
    )
      .then((items: IClientLogo[]) => {
        const urls: string[] = items
          .map((item: IClientLogo): string => (item.File && item.File.ServerRelativeUrl) || '')
          .filter((url: string): boolean => /\.(jpe?g|png|webp|gif)$/i.test(url))
          .sort((a: string, b: string): number => (a < b ? -1 : a > b ? 1 : 0))
          .map((url: string): string => encodeURI(this._safeUrl(url)));

        this.setState({ heroImages: urls, heroIndex: 0 }, () => this._startSlideshow());
      })
      .catch(() => {
        // Missing library → the red gradient stays as the hero background.
        this.setState({ heroImages: [], heroIndex: 0 });
      });
  }

  private _startSlideshow(): void {
    this._stopSlideshow();

    const reduced: boolean =
      typeof window.matchMedia === 'function' &&
      window.matchMedia('(prefers-reduced-motion: reduce)').matches;

    if (reduced || this.state.heroImages.length < 2) {
      return;
    }

    this._slideTimer = window.setInterval(() => {
      this.setState((prev: ICompanyIntranetState) => ({
        heroIndex: prev.heroImages.length > 0 ? (prev.heroIndex + 1) % prev.heroImages.length : 0
      }));
    }, 6000);
  }

  private _stopSlideshow(): void {
    if (this._slideTimer !== undefined) {
      window.clearInterval(this._slideTimer);
      this._slideTimer = undefined;
    }
  }

  // ------------------------------------------------------------------
  // Mapping helpers
  // ------------------------------------------------------------------

  /** Graph returns 7-digit fractional seconds and no offset; normalise both. */
  private _graphDate(value: string | undefined): string {
    if (!value) {
      return '';
    }

    let normalized: string = value.replace(/(\.\d{3})\d+/, '$1');

    if (!/[zZ]|[+-]\d{2}:\d{2}$/.test(normalized)) {
      normalized += 'Z';
    }

    return normalized;
  }

  private _calendarToIntranet(event: IGraphCalendarEvent): IIntranetEvent {
    return {
      id: `cal-${event.id}`,
      title: event.subject || 'Untitled meeting',
      start: this._graphDate(event.start && event.start.dateTime),
      end: this._graphDate(event.end && event.end.dateTime),
      location: event.location ? event.location.displayName : undefined,
      description: event.bodyPreview,
      link: event.webLink,
      source: 'calendar',
      organizer: event.organizer && event.organizer.emailAddress
        ? event.organizer.emailAddress.name
        : undefined,
      isAllDay: !!event.isAllDay
    };
  }

  private _companyToIntranet(item: IEventItem): IIntranetEvent {
    return {
      id: `company-${item.Id}`,
      title: item.Title,
      start: item.EventDate,
      end: item.EventEndDate,
      location: item.EventLocation,
      description: item.EventDescription,
      link: item.RegistrationLink ? item.RegistrationLink.Url : undefined,
      source: 'company',
      isAllDay: false
    };
  }

  // ------------------------------------------------------------------
  // Derived data
  // ------------------------------------------------------------------

  private _getQuickLinks(): IQuickLink[] {
    const raw: string = this.props.quickLinksJson;

    if (raw) {
      try {
        const parsed: IQuickLink[] = JSON.parse(raw);
        if (Array.isArray(parsed)) {
          const validLinks: IQuickLink[] = parsed.filter((link: IQuickLink): boolean =>
            !!link && typeof link.label === 'string' && typeof link.url === 'string');
          if (validLinks.length > 0) {
            return validLinks;
          }
        }
      } catch {
        // Fall through to list/defaults.
      }
    }

    if (this.state.quickLinks.length > 0) {
      return this.state.quickLinks;
    }

    return DEFAULT_QUICK_LINKS;
  }

  private _getUpcomingEvents(): IIntranetEvent[] {
    const source: 'calendar' | 'company' | 'both' = this.props.eventSource;
    const merged: IIntranetEvent[] = [];

    if (source === 'calendar' || source === 'both') {
      this.state.calendarEvents.forEach((event: IIntranetEvent) => merged.push(event));
    }

    if (source === 'company' || source === 'both') {
      this.state.companyEvents.forEach((item: IEventItem) =>
        merged.push(this._companyToIntranet(item)));
    }

    const startOfToday: Date = new Date();
    startOfToday.setHours(0, 0, 0, 0);
    const threshold: number = startOfToday.getTime();

    return merged
      .filter((event: IIntranetEvent): boolean => {
        const date: Date = new Date(event.start);
        return !isNaN(date.getTime()) && date.getTime() >= threshold;
      })
      .sort((a: IIntranetEvent, b: IIntranetEvent): number =>
        new Date(a.start).getTime() - new Date(b.start).getTime())
      .slice(0, 6);
  }

  private _getYouTubeId(url: string): string | undefined {
    if (!url) {
      return undefined;
    }

    const patterns: RegExp[] = [
      /(?:youtube\.com\/watch\?v=|youtube\.com\/embed\/|youtu\.be\/|youtube\.com\/shorts\/)([A-Za-z0-9_-]{11})/,
      /^([A-Za-z0-9_-]{11})$/
    ];

    for (const pattern of patterns) {
      const match: RegExpExecArray | null = pattern.exec(url);
      if (match) {
        return match[1];
      }
    }

    return undefined;
  }

  /** "Andrew Smith" → "Andrew"; "Smith, Andrew" → "Andrew". */
  private _getFirstName(name: string): string {
    const clean: string = (name || '').trim();
    if (!clean) {
      return '';
    }
    const given: string = clean.indexOf(',') >= 0 ? clean.split(',')[1].trim() : clean;
    return given.split(/\s+/)[0] || '';
  }

  private _getInitials(name: string): string {
    const parts: string[] = (name || '').trim().split(/\s+/).filter((part: string) => part.length > 0);
    if (parts.length === 0) {
      return '?';
    }
    if (parts.length === 1) {
      return parts[0].substring(0, 2).toUpperCase();
    }
    return (parts[0].charAt(0) + parts[parts.length - 1].charAt(0)).toUpperCase();
  }

  // ------------------------------------------------------------------
  // Formatting
  // ------------------------------------------------------------------

  private _safeUrl(url: string | undefined): string {
    if (!url) {
      return '#';
    }

    if (/^(https?:)?\/\//i.test(url) || /^(mailto|tel):/i.test(url)) {
      return url;
    }

    if (url.charAt(0) === '/') {
      const webUrl: string =
        this.props.context.pageContext.web.serverRelativeUrl.replace(/\/$/, '');

      if (!webUrl || url === webUrl || url.indexOf(`${webUrl}/`) === 0) {
        return url;
      }

      return `${webUrl}${url}`;
    }

    return '#';
  }

  private _formatDate(value: string): string {
    const date: Date = new Date(value);
    return isNaN(date.getTime()) ? '' : date.toLocaleDateString();
  }

  private _formatMonth(value: string): string {
    const date: Date = new Date(value);
    return isNaN(date.getTime())
      ? ''
      : date.toLocaleDateString(undefined, { month: 'short' }).toUpperCase();
  }

  private _formatDay(value: string): string {
    const date: Date = new Date(value);
    return isNaN(date.getTime())
      ? ''
      : date.toLocaleDateString(undefined, { day: '2-digit' });
  }

  private _formatEventTime(event: IIntranetEvent): string {
    const start: Date = new Date(event.start);
    if (isNaN(start.getTime())) {
      return '';
    }

    if (event.isAllDay) {
      return 'All day';
    }

    const timeFormat: Intl.DateTimeFormatOptions = { hour: '2-digit', minute: '2-digit' };
    const startLabel: string = start.toLocaleTimeString(undefined, timeFormat);

    if (!event.end) {
      return startLabel;
    }

    const end: Date = new Date(event.end);
    if (isNaN(end.getTime()) || end.toDateString() !== start.toDateString()) {
      return startLabel;
    }

    return `${startLabel} - ${end.toLocaleTimeString(undefined, timeFormat)}`;
  }


  // ------------------------------------------------------------------
  // Render
  // ------------------------------------------------------------------

  public render(): React.ReactElement<ICompanyIntranetProps> {
    const {
      title,
      subtitle,
      logoUrl,
      youtubeUrl,
      youtubeTitle,
      videoUrl,
      videoPosterUrl,
      videoSubtitle,
      heroLinks,
      userDisplayName,
      environmentMessage,
      eventSource,
      ctaHeading,
      ctaText,
      ctaButtonText,
      ctaUrl,
      hrContact,
      itContact
    } = this.props;

    const quickLinks: IQuickLink[] = this._getQuickLinks();
    const tiles: IQuickLink[] = quickLinks.slice(0, 6);
    const upcomingEvents: IIntranetEvent[] = this._getUpcomingEvents();
    const youtubeId: string | undefined = this._getYouTubeId(youtubeUrl);
    const videoFile: string = videoUrl && this._safeUrl(videoUrl) !== '#' ? this._safeUrl(videoUrl) : '';
    const hasHeroVideo: boolean = !!(videoFile || youtubeId);
    const posterSafe: string = videoPosterUrl ? this._safeUrl(videoPosterUrl) : '';
    const posterUrl: string = posterSafe && posterSafe !== '#'
      ? encodeURI(posterSafe)
      : (youtubeId ? `https://i.ytimg.com/vi/${youtubeId}/hqdefault.jpg` : '');
    const firstName: string = this._getFirstName(userDisplayName);
    const heroEvents: IIntranetEvent[] = upcomingEvents.slice(0, 3);

    const {
      news,
      stats,
      departments,
      clientLogos,
      isNewsLoading,
      isCompanyLoading,
      isCalendarLoading,
      isStatsLoading,
      newsError,
      companyError,
      calendarError,
      heroImages,
      heroIndex,
      isVideoPlaying
    } = this.state;

    const isEventsLoading: boolean = isCompanyLoading || isCalendarLoading;
    const eventsError: string | undefined = eventSource === 'calendar'
      ? calendarError
      : (companyError || calendarError);

    const stripStats: IStatItem[] = stats;

    const showStatsSection: boolean = isStatsLoading || stats.length > 0 || clientLogos.length > 0;
    const navLinks: Array<{ label: string; href: string }> = [
      { label: 'Home', href: '#home' },
      { label: 'News', href: '#news' },
      { label: 'Events', href: '#events' },
      { label: 'Departments', href: '#departments' }
    ];

    return (
      <section className={styles.companyIntranet} ref={this._rootRef}>
        {this._isPreviewMode() && (
          <div className={styles.previewBadge} role="status">
            <Icon iconName="Preview" />
            Preview mode — demo content is shown. Deploy the web part to connect
            live SharePoint lists, your Outlook calendar and the featured video.
          </div>
        )}

        <header className={styles.hero} id="home" data-reveal="1">
          {heroImages.length > 0 && (
            <div className={styles.heroSlides} aria-hidden="true">
              {heroImages.map((src: string, index: number): React.ReactElement => (
                <div
                  key={src}
                  className={`${styles.heroSlide} ${index === heroIndex ? styles.heroSlideActive : ''}`}
                  style={{ backgroundImage: `url("${src}")` }}
                />
              ))}
            </div>
          )}
          <div className={styles.heroShade} aria-hidden="true" />

          <nav className={`${styles.navbar} ${styles.heroNav}`} aria-label="Intranet sections">
            <a className={styles.navBrand} href="#home">
              {logoUrl ? (
                <img className={styles.navLogoImg} src={this._safeUrl(logoUrl)} alt={title} />
              ) : (
                <>
                  <BrandLogo className={styles.navLogo} />
                  <span className={styles.navName}>{(title)}</span>
                </>
              )}
            </a>

            <div className={styles.navLinks}>
              {navLinks.filter((l: { label: string; href: string }): boolean => l.label !== 'Departments').map((link: { label: string; href: string }): React.ReactElement => (
                <a key={link.href} className={styles.navLink} href={link.href}>
                  {link.label}
                </a>
              ))}
              <NavAbout />
              <a className={styles.navLink} href="#departments">Departments</a>
            </div>

            <div className={styles.userChip} title={(userDisplayName)}>
              <span className={styles.userAvatar} aria-hidden="true">
                {this._getInitials(userDisplayName)}
              </span>
              <span className={styles.userName}>{(userDisplayName)}</span>
            </div>
          </nav>

          <div className={styles.heroStage}>
            <h1 className={styles.heroWelcome}>
              {firstName ? `Welcome, ${firstName}!` : 'Welcome!'}
            </h1>

            <div className={styles.heroPanel}>
              <ul className={styles.heroLinks} aria-label="Quick links">
                {heroLinks.slice(0, 8).map((link: IQuickLink): React.ReactElement => (
                  <li key={link.label}>
                    <a className={styles.heroLink} href={this._safeUrl(link.url)}>
                      <Icon iconName={link.icon || 'Link'} />
                      {link.label}
                    </a>
                  </li>
                ))}
              </ul>

              <div className={styles.heroVideo}>
                {hasHeroVideo && isVideoPlaying && videoFile && (
                  <video
                    className={styles.heroVideoEl}
                    src={videoFile}
                    poster={posterUrl || undefined}
                    controls
                    autoPlay
                    playsInline
                  />
                )}
                {hasHeroVideo && isVideoPlaying && !videoFile && youtubeId && (
                  <iframe
                    className={styles.heroVideoEl}
                    src={`https://www.youtube-nocookie.com/embed/${youtubeId}?autoplay=1&rel=0`}
                    title={youtubeTitle || 'Featured video'}
                    allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share"
                    allowFullScreen
                  />
                )}
                {!(hasHeroVideo && isVideoPlaying) && (
                  <button
                    type="button"
                    className={styles.heroVideoPoster}
                    style={posterUrl ? { backgroundImage: `url("${posterUrl}")` } : undefined}
                    disabled={!hasHeroVideo}
                    aria-label={`Play video${youtubeTitle ? `: ${youtubeTitle}` : ''}`}
                    onClick={(): void => this.setState({ isVideoPlaying: true })}
                  >
                    <span className={styles.heroPlay}>
                      <Icon iconName="Play" />
                    </span>
                    <span className={styles.heroVideoCaption}>
                      <strong>{youtubeTitle || (hasHeroVideo ? 'Featured video' : 'Add your featured video')}</strong>
                      {(videoSubtitle || !hasHeroVideo) && (
                        <span>{videoSubtitle || 'Edit the web part and add a YouTube link or video file.'}</span>
                      )}
                    </span>
                  </button>
                )}
              </div>

              <aside className={styles.heroEvents} aria-label="Your upcoming events">
                <div className={styles.heroEventsHead}>
                  <Icon iconName="Calendar" />
                  Coming up for you
                </div>

                {isEventsLoading && [0, 1, 2].map((index: number): React.ReactElement => (
                  <div key={index} className={styles.heroEventSkeleton} aria-hidden="true" />
                ))}

                {!isEventsLoading && heroEvents.length === 0 && (
                  <p className={styles.heroEventsEmpty}>
                    {eventsError || 'No upcoming events in the next 30 days.'}
                  </p>
                )}

                {!isEventsLoading && heroEvents.map((event: IIntranetEvent): React.ReactElement => {
                  const eventLink: string = this._safeUrl(event.link);
                  const external: boolean = eventLink !== '#';
                  const timeLabel: string = this._formatEventTime(event);
                  const meta: string = [timeLabel, event.location].filter((part: string | undefined): boolean => !!part).join(' · ');

                  return (
                    <a
                      key={event.id}
                      className={styles.heroEventCard}
                      href={external ? eventLink : '#events'}
                      target={external ? '_blank' : undefined}
                      rel={external ? 'noreferrer' : undefined}
                    >
                      <span className={styles.heroEventDate}>
                        <span>{this._formatMonth(event.start)}</span>
                        <strong>{this._formatDay(event.start)}</strong>
                      </span>
                      <span className={styles.heroEventBody}>
                        <strong>{event.title}</strong>
                        {meta && <span>{meta}</span>}
                      </span>
                    </a>
                  );
                })}

                <a className={styles.heroEventsMore} href="#events" aria-label="See all events">
                  <Icon iconName="ChevronDown" />
                </a>
              </aside>
            </div>
          </div>
        </header>

        <AppCarousel />

        <section className={styles.section} id="need" data-reveal="1">
          <div className={styles.sectionHeader}>
            <h2 className={styles.sectionTitle}>Everything you need</h2>
            <span className={styles.sectionHint}>One click to the tools you use every day</span>
          </div>
          <div className={styles.tiles}>
            {tiles.map((link: IQuickLink, index: number): React.ReactElement => (
              <a
                key={`${link.label}-${link.url}`}
                className={`${styles.tile} ${index === 0 ? styles.tileHighlighted : ''}`}
                style={{ '--stagger': `${index * 55}ms` } as React.CSSProperties}
                href={this._safeUrl(link.url)}
              >
                <span className={styles.tileIcon}>
                  <Icon iconName={link.icon || 'Link'} />
                </span>
                <span className={styles.tileLabel}>{(link.label)}</span>
                <Icon iconName="ChevronRight" className={styles.tileChevron} />
              </a>
            ))}
          </div>
        </section>

        {showStatsSection && (
          <section className={styles.section} data-reveal="1">
            <div className={styles.sectionHeader}>
              <h2 className={styles.sectionTitle}>{(title)} at a glance</h2>
              <span className={styles.sectionHint}>Numbers our teams create together</span>
            </div>

            {isStatsLoading && (
              <div className={styles.statsStrip} aria-hidden="true">
                {[0, 1, 2, 3].map((index: number): React.ReactElement => (
                  <div key={index} className={`${styles.skeleton} ${styles.statSkeleton}`} />
                ))}
              </div>
            )}

            {!isStatsLoading && stripStats.length > 0 && (
              <div className={styles.statsStrip}>
                {stripStats.slice(0, 5).map((stat: IStatItem): React.ReactElement => (
                  <div key={stat.Id} className={styles.statCard}>
                    <span className={styles.statIcon}>
                      <Icon iconName={stat.Icon || 'FavoriteStar'} />
                    </span>
                    <span className={styles.statValue} data-count={stat.Value}>
                      {(stat.Value)}
                    </span>
                    <span className={styles.statLabel}>{(stat.Title)}</span>
                  </div>
                ))}
              </div>
            )}

            {clientLogos.length > 0 && (
              <div className={styles.logoStrip}>
                <span className={styles.logoStripLabel}>Trusted by</span>
                <div className={styles.logosRow}>
                  {clientLogos.map((logo: IClientLogo): React.ReactElement => {
                    const fileUrl: string = logo.File ? logo.File.ServerRelativeUrl : '';
                    if (fileUrl) {
                      return (
                        <img
                          key={logo.Id}
                          className={styles.logoImg}
                          src={this._safeUrl(fileUrl)}
                          alt={(logo.Title || 'Client logo')}
                        />
                      );
                    }
                    return (
                      <span key={logo.Id} className={styles.logoChip}>
                        <Icon iconName="FavoriteStar" />
                        {(logo.Title || 'Client')}
                      </span>
                    );
                  })}
                </div>
              </div>
            )}
          </section>
        )}

        <div className={styles.mainGrid}>
          <section className={`${styles.section} ${styles.eventsSection}`} id="events" data-reveal="1">
            <div className={styles.sectionHeader}>
              <h2 className={styles.sectionTitle}>Upcoming events</h2>
              <span className={styles.sectionHint}>
                {eventSource === 'calendar'
                  ? 'From your Outlook calendar'
                  : eventSource === 'company'
                    ? 'From the company events list'
                    : 'Your calendar and company events'}
              </span>
            </div>

            {isEventsLoading && (
              <div className={styles.eventList} aria-hidden="true">
                {[0, 1, 2].map((index: number): React.ReactElement => (
                  <div key={index} className={`${styles.eventCard} ${styles.skeletonCard}`}>
                    <div className={`${styles.skeleton} ${styles.skeletonBadge}`} />
                    <div className={styles.eventBody}>
                      <div className={`${styles.skeleton} ${styles.skeletonLineWide}`} />
                      <div className={`${styles.skeleton} ${styles.skeletonLine}`} />
                      <div className={`${styles.skeleton} ${styles.skeletonLineShort}`} />
                    </div>
                  </div>
                ))}
              </div>
            )}

            {!isEventsLoading && eventsError && (
              <p className={styles.status}>{(eventsError)}</p>
            )}

            {!isEventsLoading && !eventsError && upcomingEvents.length === 0 && (
              <p className={styles.status}>
                No upcoming events. Add items to the events list, or check your Outlook calendar.
              </p>
            )}

            {!isEventsLoading && !eventsError && upcomingEvents.length > 0 && (
              <div className={styles.eventList}>
                {upcomingEvents.map((event: IIntranetEvent, index: number): React.ReactElement => {
                  const registrationUrl: string = this._safeUrl(event.link);
                  const isCalendar: boolean = event.source === 'calendar';

                  return (
                    <article
                      key={event.id}
                      className={styles.eventCard}
                      style={{ '--stagger': `${index * 70}ms` } as React.CSSProperties}
                    >
                      <div className={styles.eventDateBadge}>
                        <span className={styles.eventMonth}>{this._formatMonth(event.start)}</span>
                        <span className={styles.eventDay}>{this._formatDay(event.start)}</span>
                      </div>

                      <div className={styles.eventBody}>
                        <div className={styles.eventTopRow}>
                          <span className={isCalendar ? styles.pillCalendar : styles.pillCompany}>
                            {isCalendar ? 'My calendar' : 'Company'}
                          </span>
                        </div>

                        <h3 className={styles.eventTitle}>{(event.title)}</h3>

                        <div className={styles.eventMeta}>
                          <span className={styles.eventMetaItem}>
                            <Icon iconName="Clock" />
                            {this._formatEventTime(event)}
                          </span>
                          {event.location && (
                            <span className={styles.eventMetaItem}>
                              <Icon iconName="Location" />
                              {(event.location)}
                            </span>
                          )}
                        </div>

                        {event.description && (
                          <p className={styles.eventDescription}>{(event.description)}</p>
                        )}

                        {registrationUrl !== '#' && (
                          <a
                            className={styles.eventLink}
                            href={registrationUrl}
                            target="_blank"
                            rel="noreferrer"
                          >
                            {isCalendar ? 'Open invite' : 'Register'}
                            <Icon iconName="ChevronRight" />
                          </a>
                        )}
                      </div>
                    </article>
                  );
                })}
              </div>
            )}
          </section>

          <section className={`${styles.section} ${styles.videoSection}`} data-reveal="1">
            <div className={styles.sectionHeader}>
              <h2 className={styles.sectionTitle}>Featured video</h2>
              <span className={styles.sectionHint}>From our latest update</span>
            </div>

            {youtubeId ? (
              <div className={styles.videoCard}>
                <div className={styles.videoFrame}>
                  <iframe
                    src={`https://www.youtube-nocookie.com/embed/${youtubeId}`}
                    title={youtubeTitle || 'SML Cosmos featured video'}
                    allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share"
                    allowFullScreen
                  />
                </div>
                {youtubeTitle && (
                  <div className={styles.videoCaption}>{(youtubeTitle)}</div>
                )}
              </div>
            ) : (
              <div className={styles.videoPlaceholder}>
                <span className={styles.videoPlaceholderIcon}>
                  <Icon iconName="Play" />
                </span>
                <p>Add a YouTube link in the web part properties to feature a video here.</p>
              </div>
            )}
          </section>
        </div>

        <section className={styles.section} id="news" data-reveal="1">
          <div className={styles.sectionHeader}>
            <h2 className={styles.sectionTitle}>Company news</h2>
            <span className={styles.sectionHint}>Announcements and updates</span>
          </div>

          {isNewsLoading && (
            <div className={styles.newsGrid} aria-hidden="true">
              {[0, 1, 2].map((index: number): React.ReactElement => (
                <div key={index} className={`${styles.newsCard} ${styles.skeletonCard}`}>
                  <div className={`${styles.skeleton} ${styles.skeletonImage}`} />
                  <div className={styles.newsBody}>
                    <div className={`${styles.skeleton} ${styles.skeletonLineWide}`} />
                    <div className={`${styles.skeleton} ${styles.skeletonLine}`} />
                    <div className={`${styles.skeleton} ${styles.skeletonLineShort}`} />
                  </div>
                </div>
              ))}
            </div>
          )}

          {!isNewsLoading && newsError && (
            <p className={styles.status}>
              {(newsError)} Add the list or update the list name in the web part properties.
            </p>
          )}

          {!isNewsLoading && !newsError && news.length === 0 && (
            <p className={styles.status}>
              There is no news yet. Add items to the news list to get started.
            </p>
          )}

          {!isNewsLoading && !newsError && news.length > 0 && (
            <div className={styles.newsGrid}>
              {news.map((item: INewsItem, index: number): React.ReactElement => {
                const itemUrl: string = this._safeUrl(item.NewsLink && item.NewsLink.Url);
                const imageUrl: string = this._safeUrl(item.NewsImage);

                return (
                  <a
                    key={item.Id}
                    className={styles.newsCard}
                    style={{ '--stagger': `${index * 70}ms` } as React.CSSProperties}
                    href={itemUrl}
                  >
                    <div className={styles.newsMedia}>
                      {imageUrl !== '#' ? (
                        <img className={styles.newsImage} src={imageUrl} alt="" />
                      ) : (
                        <div className={styles.newsImagePlaceholder}>
                          {(item.Title ? item.Title.charAt(0) : 'S')}
                        </div>
                      )}
                    </div>
                    <div className={styles.newsBody}>
                      <span className={styles.newsDate}>{this._formatDate(item.Modified)}</span>
                      <h3 className={styles.newsTitle}>{(item.Title)}</h3>
                      {item.NewsSummary && (
                        <p className={styles.newsSummary}>{(item.NewsSummary)}</p>
                      )}
                      <span className={styles.newsCta}>
                        Read more
                        <Icon iconName="ChevronRight" />
                      </span>
                    </div>
                  </a>
                );
              })}
            </div>
          )}
        </section>

        <LeadershipSection
          context={this.props.context}
          listTitle={this.props.leadersListTitle}
          isPreview={this._isPreviewMode()}
        />

        <CelebrationsSection
          context={this.props.context}
          listTitle={this.props.celebrationsListTitle}
          isPreview={this._isPreviewMode()}
        />

        {departments.length > 0 && (
          <section className={styles.section} id="departments" data-reveal="1">
            <div className={styles.sectionHeader}>
              <h2 className={styles.sectionTitle}>Departments</h2>
              <span className={styles.sectionHint}>Who to reach, and how</span>
            </div>
            <div className={styles.deptGrid}>
              {departments.slice(0, 3).map((dept: IDepartment): React.ReactElement => {
                const deptUrl: string = this._safeUrl(dept.Url && dept.Url.Url);
                const cardContent: React.ReactElement = (
                  <>
                    <span className={styles.deptIcon}>
                      <Icon iconName={dept.Icon || 'People'} />
                    </span>
                    <h3 className={styles.deptTitle}>{(dept.Title)}</h3>
                    {dept.Description && (
                      <p className={styles.deptDesc}>{(dept.Description)}</p>
                    )}
                    {dept.Head && (
                      <span className={styles.deptHead}>
                        <Icon iconName="Contact" />
                        {(dept.Head)}
                      </span>
                    )}
                  </>
                );

                return deptUrl !== '#' ? (
                  <a key={dept.Id} className={styles.deptCard} href={deptUrl} target="_blank" rel="noreferrer">
                    {cardContent}
                  </a>
                ) : (
                  <div key={dept.Id} className={styles.deptCard}>
                    {cardContent}
                  </div>
                );
              })}
            </div>
          </section>
        )}

        {ctaHeading && (
          <section className={styles.ctaSection} data-reveal="1">
            <div className={styles.ctaBanner}>
              <div className={styles.heroClouds} aria-hidden="true">
                <span className={`${styles.cloud} ${styles.cloud1}`} />
                <span className={`${styles.cloud} ${styles.cloud3}`} />
              </div>
              <div className={styles.ctaContent}>
                <div>
                  <h2 className={styles.ctaHeading}>{(ctaHeading)}</h2>
                  {ctaText && <p className={styles.ctaText}>{(ctaText)}</p>}
                </div>
                {ctaUrl && ctaButtonText && (
                  <a className={styles.ctaButton} href={this._safeUrl(ctaUrl)}>
                    {(ctaButtonText)}
                    <Icon iconName="ChevronRight" />
                  </a>
                )}
              </div>
            </div>
          </section>
        )}

        <footer className={styles.footer} data-reveal="1">
          <div className={styles.footerTop}>
            <div className={styles.footerBrand}>
              {logoUrl ? (
                <img className={styles.navLogoImg} src={this._safeUrl(logoUrl)} alt="" />
              ) : (
                <BrandLogo className={styles.navLogo} />
              )}
              <div className={styles.footerBrandText}>
                <span className={styles.footerBrandName}>{(title)}</span>
                <span className={styles.footerBrandTag}>{(subtitle)}</span>
              </div>
            </div>

            {(hrContact || itContact) && (
              <div className={styles.footerContacts}>
                {hrContact && (
                  <span className={styles.footerContact}>
                    <Icon iconName="Mail" />
                    {(hrContact)}
                  </span>
                )}
                {itContact && (
                  <span className={styles.footerContact}>
                    <Icon iconName="Phone" />
                    {(itContact)}
                  </span>
                )}
              </div>
            )}

            <nav className={styles.footerLinks} aria-label="Footer links">
              {quickLinks.slice(0, 5).map((link: IQuickLink): React.ReactElement => (
                <a key={`${link.label}-${link.url}`} className={styles.footerLink} href={this._safeUrl(link.url)}>
                  {(link.label)}
                </a>
              ))}
            </nav>
          </div>

          <div className={styles.footerBottom}>
            <span>© {new Date().getFullYear()} {(title)}</span>
            <span>{(environmentMessage)}</span>
          </div>

          <div className={styles.watermark} aria-hidden="true">SML COSMOS</div>
        </footer>
      </section>
    );
  }
}
