import * as React from 'react';
import styles from './CompanyIntranet.module.scss';
import type { ICompanyIntranetProps } from './ICompanyIntranetProps';
import type { ICelebrationItem } from './IntranetModels';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Icon } from '@fluentui/react/lib/Icon';

/** How many days ahead to look, and how many people to show per card. */
const LOOK_AHEAD_DAYS = 30;
const MAX_UPCOMING = 4;
const MS_PER_DAY = 86400000;

type CelebrationKind = 'birthday' | 'anniversary';

interface ICelebration {
  key: string;
  kind: CelebrationKind;
  years?: number;
  name: string;
  detail: string;
  email?: string;
  photoUrl?: string;
  daysAway: number;
  date: Date;
}

interface ICelebrationGroups {
  birthdays: ICelebration[];
  anniversaries: ICelebration[];
}

interface IDayParts {
  y: number;
  m: number;
  d: number;
}

/**
 * Date-only SharePoint columns come back as midnight in the site's time zone, converted to UTC
 * (e.g. 2020-03-13T18:30:00Z for 14 March in India). Adding 12 hours and reading the UTC date gives
 * the intended calendar day for sites between UTC-12 and UTC+11.
 */
function parseDay(value: string | undefined): IDayParts | undefined {
  if (!value) {
    return undefined;
  }
  const time: number = new Date(value).getTime();
  if (isNaN(time)) {
    return undefined;
  }
  const shifted: Date = new Date(time + 12 * 3600 * 1000);
  return { y: shifted.getUTCFullYear(), m: shifted.getUTCMonth() + 1, d: shifted.getUTCDate() };
}

function isLeapYear(y: number): boolean {
  return (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
}

/** Next time this month/day comes round (today counts). 29 February falls on 28 February in other years. */
function nextOccurrence(parts: IDayParts, today: Date): { date: Date; year: number; daysAway: number } {
  const build = (year: number): Date => {
    const day: number = parts.m === 2 && parts.d === 29 && !isLeapYear(year) ? 28 : parts.d;
    return new Date(year, parts.m - 1, day);
  };
  let year: number = today.getFullYear();
  let date: Date = build(year);
  if (date.getTime() < today.getTime()) {
    year += 1;
    date = build(year);
  }
  return { date, year, daysAway: Math.round((date.getTime() - today.getTime()) / MS_PER_DAY) };
}

type IRawRow = Record<string, unknown>;

/** First non-empty value among the possible internal names (columns made by hand get names like Join_x0020_date). */
function pickValue(row: IRawRow, names: string[]): unknown {
  for (const name of names) {
    const value: unknown = row[name];
    if (value) {
      return value;
    }
  }
  return undefined;
}

/** Text from a plain-text column, or the address from a Hyperlink column. */
function asString(value: unknown): string {
  if (typeof value === 'string') {
    return value;
  }
  if (value && typeof value === 'object' && typeof (value as { Url?: unknown }).Url === 'string') {
    return (value as { Url: string }).Url;
  }
  return '';
}

/** A "Copy link" sharing address (…/:i:/r/sites/x/file.jpg?d=…) is a web page, not an image; turn it into the direct file address. */
function toDirectImageUrl(url: string): string {
  const match: RegExpMatchArray | undefined = url.match(/^(https?:\/\/[^/]+)\/:[a-z]:\/[a-z](\/[^?]*)/i) || undefined;
  return match ? `${match[1]}${match[2]}` : url;
}

function mapRow(row: IRawRow): ICelebrationItem {
  const photo: string = asString(pickValue(row, ['PersonPhoto', 'Photo', 'Picture', 'Image']));
  return {
    Id: Number(row.Id || row.ID || 0),
    Title: asString(row.Title),
    PersonBirthday: asString(pickValue(row, ['PersonBirthday', 'Birthday', 'BirthDate', 'DateOfBirth', 'DOB'])),
    PersonJoinDate: asString(pickValue(row, ['PersonJoinDate', 'JoinDate', 'Join_x0020_date', 'Join_x0020_Date', 'DateOfJoining'])),
    PersonEmail: asString(pickValue(row, ['PersonEmail', 'Email', 'E_x002d_mail'])),
    PersonDepartment: asString(pickValue(row, ['PersonDepartment', 'Department'])),
    PersonPhoto: photo ? { Url: toDirectImageUrl(photo) } : undefined
  };
}

function initialsOf(name: string): string {
  return name.split(/\s+/).filter(Boolean).slice(0, 2).map((w: string): string => w.charAt(0).toUpperCase()).join('');
}

function firstNameOf(name: string): string {
  return name.split(/\s+/)[0] || name;
}

/** "Birthday today", "Work anniversary tomorrow", "Birthday on 24 Nov". */
function chipText(c: ICelebration): string {
  const label: string = c.kind === 'birthday' ? 'Birthday' : 'Work anniversary';
  if (c.daysAway === 0) {
    return `${label} today`;
  }
  if (c.daysAway === 1) {
    return `${label} tomorrow`;
  }
  return `${label} on ${c.date.toLocaleDateString(undefined, { day: 'numeric', month: 'short' })}`;
}

function buildGroups(items: ICelebrationItem[], today: Date): ICelebrationGroups {
  const birthdays: ICelebration[] = [];
  const anniversaries: ICelebration[] = [];

  items.forEach((row: ICelebrationItem): void => {
    const name: string = (row.Title || '').trim();
    if (!name) {
      return;
    }
    const photoUrl: string | undefined = row.PersonPhoto && row.PersonPhoto.Url ? row.PersonPhoto.Url : undefined;
    const email: string | undefined = (row.PersonEmail || '').trim() || undefined;
    const department: string = (row.PersonDepartment || '').trim();

    const birthday: IDayParts | undefined = parseDay(row.PersonBirthday);
    if (birthday) {
      const next = nextOccurrence(birthday, today);
      if (next.daysAway <= LOOK_AHEAD_DAYS) {
        birthdays.push({ key: `b-${row.Id}`, kind: 'birthday', name, detail: department, email, photoUrl, daysAway: next.daysAway, date: next.date });
      }
    }

    const joined: IDayParts | undefined = parseDay(row.PersonJoinDate);
    if (joined) {
      const next = nextOccurrence(joined, today);
      const years: number = next.year - joined.y;
      if (years >= 1 && next.daysAway <= LOOK_AHEAD_DAYS) {
        anniversaries.push({
          key: `a-${row.Id}`,
          kind: 'anniversary',
          years,
          name,
          detail: `${years} ${years === 1 ? 'year' : 'years'} with SML Cosmos`,
          email,
          photoUrl,
          daysAway: next.daysAway,
          date: next.date
        });
      }
    }
  });

  const byDate = (a: ICelebration, b: ICelebration): number =>
    a.daysAway - b.daysAway || a.name.localeCompare(b.name);
  return { birthdays: birthdays.sort(byDate), anniversaries: anniversaries.sort(byDate) };
}

/** Sample rows for the offline preview, dated relative to today so the widget always has something to show. */
function demoItems(today: Date): ICelebrationItem[] {
  const iso = (offsetDays: number, year: number): string => {
    const d: Date = new Date(today.getFullYear(), today.getMonth(), today.getDate() + offsetDays);
    return new Date(Date.UTC(year, d.getMonth(), d.getDate())).toISOString();
  };
  return [
    { Id: 1, Title: 'Aarav Mehta', PersonBirthday: iso(0, 1992), PersonDepartment: 'Projects & Engineering' },
    { Id: 2, Title: 'Priya Nair', PersonBirthday: iso(2, 1989), PersonDepartment: 'Safety & Compliance' },
    { Id: 3, Title: 'Rohan Deshmukh', PersonBirthday: iso(9, 1995), PersonDepartment: 'Operations' },
    { Id: 4, Title: 'Sneha Kulkarni', PersonJoinDate: iso(0, today.getFullYear() - 5) },
    { Id: 5, Title: 'Vikram Patil', PersonJoinDate: iso(4, today.getFullYear() - 10) },
    { Id: 6, Title: 'Anjali Shah', PersonJoinDate: iso(12, today.getFullYear() - 2) }
  ];
}

/** Confetti pieces: position, size, colour and timing. They fall slowly; with reduced motion they sit still. */
const CONFETTI: Array<{ left: string; top: string; w: number; h: number; round?: boolean; color: string; delay: string; time: string }> = [
  { left: '4%', top: '18%', w: 9, h: 16, color: '#86efac', delay: '0s', time: '9s' },
  { left: '11%', top: '62%', w: 8, h: 8, round: true, color: '#fecaca', delay: '-4s', time: '11s' },
  { left: '19%', top: '34%', w: 12, h: 6, color: '#ffffff', delay: '-2s', time: '10s' },
  { left: '27%', top: '78%', w: 9, h: 16, color: '#fecaca', delay: '-6s', time: '12s' },
  { left: '36%', top: '12%', w: 8, h: 8, round: true, color: '#86efac', delay: '-8s', time: '9.5s' },
  { left: '44%', top: '54%', w: 12, h: 6, color: '#ffffff', delay: '-1s', time: '13s' },
  { left: '53%', top: '26%', w: 9, h: 16, color: '#86efac', delay: '-5s', time: '10.5s' },
  { left: '61%', top: '70%', w: 8, h: 8, round: true, color: '#fecaca', delay: '-9s', time: '11.5s' },
  { left: '69%', top: '16%', w: 12, h: 6, color: '#ffffff', delay: '-3s', time: '9s' },
  { left: '77%', top: '48%', w: 9, h: 16, color: '#fecaca', delay: '-7s', time: '12.5s' },
  { left: '85%', top: '82%', w: 8, h: 8, round: true, color: '#86efac', delay: '-10s', time: '10s' },
  { left: '93%', top: '30%', w: 12, h: 6, color: '#ffffff', delay: '-2.5s', time: '11s' }
];

interface ICelebrationsSectionProps {
  context: ICompanyIntranetProps['context'];
  listTitle: string;
  /** True on the local/offline preview, where there is no SharePoint to read from. */
  isPreview: boolean;
}

export function CelebrationsSection({ context, listTitle, isPreview }: ICelebrationsSectionProps): React.ReactElement {
  const [groups, setGroups] = React.useState<ICelebrationGroups | undefined>(undefined);
  const [brokenPhotos, setBrokenPhotos] = React.useState<Record<string, boolean>>({});

  React.useEffect((): (() => void) => {
    let cancelled = false;
    const now: Date = new Date();
    const today: Date = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const apply = (items: ICelebrationItem[]): void => {
      if (!cancelled) {
        setGroups(buildGroups(items, today));
      }
    };

    if (isPreview) {
      apply(demoItems(today));
      return (): void => { cancelled = true; };
    }
    if (!listTitle) {
      apply([]);
      return (): void => { cancelled = true; };
    }

    const url: string =
      `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items` +
      '?$top=5000';

    context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          throw new Error(`Unable to load the list "${listTitle}" (${response.status}).`);
        }
        return response.json() as Promise<{ value: IRawRow[] }>;
      })
      .then((data: { value: IRawRow[] }) => apply((data.value || []).map(mapRow)))
      .catch(() => apply([]));

    return (): void => { cancelled = true; };
  }, [context, listTitle, isPreview]);

  // Nothing to celebrate (or list not set up yet): the section stays out of the way.
  if (!groups || (groups.birthdays.length === 0 && groups.anniversaries.length === 0)) {
    return <></>;
  }

  const everyone: ICelebration[] = groups.birthdays
    .concat(groups.anniversaries)
    .sort((x: ICelebration, y: ICelebration): number => x.daysAway - y.daysAway || x.name.localeCompare(y.name));
  const today: ICelebration[] = everyone.filter((c: ICelebration): boolean => c.daysAway === 0);
  // Today's people get the stage; on a quiet day it shows whoever is next.
  const spotlight: ICelebration[] = today.length > 0 ? today : [everyone[0]];
  const lead: ICelebration = spotlight[0];
  const others: ICelebration[] = spotlight.slice(1);
  const spotlightKeys: Record<string, boolean> = {};
  spotlight.forEach((c: ICelebration): void => { spotlightKeys[c.key] = true; });
  const upcoming: ICelebration[] = everyone.filter((c: ICelebration): boolean => !spotlightKeys[c.key]);
  const shown: ICelebration[] = upcoming.slice(0, MAX_UPCOMING);
  const hiddenCount: number = upcoming.length - shown.length;

  const first: string = firstNameOf(lead.name);
  const years: number = lead.years || 0;
  let headline: string;
  if (lead.kind === 'birthday') {
    headline = lead.daysAway === 0 ? `Happy birthday, ${first}!` : `${first}'s birthday is ${lead.daysAway === 1 ? 'tomorrow' : 'coming up'}`;
  } else {
    headline = lead.daysAway === 0 ? `Congratulations, ${first}!` : `${first}'s anniversary is ${lead.daysAway === 1 ? 'tomorrow' : 'coming up'}`;
  }

  const renderAvatar = (c: ICelebration, className: string): React.ReactElement => (
    <span className={`${styles.celebAvatar} ${c.kind === 'anniversary' ? styles.celebAvatarGreen : ''} ${className}`} aria-hidden="true">
      {c.photoUrl && !brokenPhotos[c.key] ? (
        <img
          src={c.photoUrl}
          alt=""
          onError={(): void => setBrokenPhotos((prev: Record<string, boolean>) => ({ ...prev, [c.key]: true }))}
        />
      ) : (
        initialsOf(c.name)
      )}
    </span>
  );

  const mailHref = (c: ICelebration): string =>
    `mailto:${c.email}?subject=${encodeURIComponent(c.kind === 'birthday' ? 'Happy birthday! 🎂' : 'Happy work anniversary! 🎉')}`;

  return (
    <section className={`${styles.section} ${styles.revealed}`} id="celebrations" data-reveal="1">
      <div className={styles.sectionHeader}>
        <h2 className={styles.sectionTitle}>Celebrations</h2>
        <span className={styles.sectionHint}>Birthdays and work anniversaries in the next {LOOK_AHEAD_DAYS} days</span>
      </div>

      <div className={`${styles.celebBanner} ${shown.length > 0 ? styles.celebBannerSplit : ''}`}>
        <span className={styles.celebConfetti} aria-hidden="true">
          {CONFETTI.map((p, i: number): React.ReactElement => (
            <i
              key={i}
              style={{
                left: p.left,
                top: p.top,
                width: `${p.w}px`,
                height: `${p.h}px`,
                borderRadius: p.round ? '50%' : '3px',
                background: p.color,
                animationDelay: p.delay,
                animationDuration: p.time
              }}
            />
          ))}
        </span>

        <div className={styles.celebMain}>
          <div className={styles.celebPortrait}>
            <span className={styles.celebRing}>{renderAvatar(lead, styles.celebAvatarXL)}</span>
            <span className={styles.celebBadge} aria-hidden="true">
              <Icon iconName={lead.kind === 'birthday' ? 'Cake' : 'Ribbon'} />
            </span>
          </div>

          <div className={styles.celebCopy}>
            <span className={styles.celebChip}>{chipText(lead)}</span>
            <h3 className={styles.celebHeadline}>{headline}</h3>
            <p className={styles.celebSub}>
              {lead.kind === 'anniversary'
                ? `${years} ${years === 1 ? 'year' : 'years'} with SML Cosmos${lead.detail && lead.detail.indexOf('year') === -1 ? ` · ${lead.detail}` : ''}`
                : (lead.detail || 'SML Cosmos')}
            </p>
            <div className={styles.celebActions}>
              {lead.email && (
                <a className={styles.celebCta} href={mailHref(lead)}>
                  <Icon iconName="Mail" />
                  Send wishes
                </a>
              )}
              {others.length > 0 && (
                <span className={styles.celebAlso}>
                  <span className={styles.celebStack} aria-hidden="true">
                    {others.slice(0, 3).map((c: ICelebration): React.ReactElement => (
                      <React.Fragment key={c.key}>{renderAvatar(c, styles.celebAvatarXS)}</React.Fragment>
                    ))}
                  </span>
                  Also today: {others.map((c: ICelebration): string => firstNameOf(c.name)).join(', ')}
                </span>
              )}
            </div>
          </div>
        </div>

        {shown.length > 0 && (
          <aside className={styles.celebPanel} aria-label="Coming up">
            <h3 className={styles.celebPanelTitle}>Coming up</h3>
            <ol className={styles.celebList}>
              {shown.map((c: ICelebration): React.ReactElement => (
                <li key={c.key} className={styles.celebRow}>
                  <span className={styles.celebDate}>
                    <b>{c.date.getDate()}</b>
                    {c.date.toLocaleDateString(undefined, { month: 'short' })}
                  </span>
                  <div className={styles.celebText}>
                    <p className={styles.celebName}>{c.name}</p>
                    <p className={styles.celebMeta}>
                      {c.kind === 'birthday' ? 'Birthday' : 'Work anniversary'}
                      {c.detail && c.kind === 'birthday' ? ` · ${c.detail}` : ''}
                      {c.kind === 'anniversary' && c.years ? ` · ${c.years} ${c.years === 1 ? 'yr' : 'yrs'}` : ''}
                    </p>
                  </div>
                  {c.email && (
                    <a className={styles.celebWish} href={mailHref(c)} aria-label={`Send ${c.name} a message`} title="Send wishes">
                      <Icon iconName="Mail" />
                    </a>
                  )}
                </li>
              ))}
            </ol>
            {hiddenCount > 0 && <p className={styles.celebMore}>and {hiddenCount} more in the next {LOOK_AHEAD_DAYS} days</p>}
          </aside>
        )}
      </div>
    </section>
  );
}

export default CelebrationsSection;