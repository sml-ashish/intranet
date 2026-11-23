import * as React from 'react';
import styles from './CompanyIntranet.module.scss';
import type { ICompanyIntranetProps } from './ICompanyIntranetProps';
import { asString, pickValue, toDirectImageUrl } from './ListHelpers';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Icon } from '@fluentui/react/lib/Icon';

/** A leadership team member. Set photoUrl (e.g. /SiteAssets/leaders/name.jpg) to show a real photo. */
export interface ILeader {
  name: string;
  role: string;
  photoUrl?: string;
  linkedinUrl?: string;
}

const LEADERS: ILeader[] = [
  { name: 'Leader Name', role: 'Managing Director' },
  { name: 'Leader Name', role: 'Chief Operating Officer' },
  { name: 'Leader Name', role: 'Chief Financial Officer' },
  { name: 'Leader Name', role: 'Head of Projects & Engineering' },
  { name: 'Leader Name', role: 'Head of Safety & Compliance' },
  { name: 'Leader Name', role: 'Head of People & Culture' }
];

function getInitialsOf(name: string): string {
  return name.split(/\s+/).filter(Boolean).slice(0, 2).map((w: string): string => w.charAt(0).toUpperCase()).join('');
}

/** How long each leader stays open before the next one takes over. */
const LEADER_AUTOPLAY_MS = 8000;

interface ILeadershipSectionProps {
  context: ICompanyIntranetProps['context'];
  listTitle: string;
  /** True on the local/offline preview, where there is no SharePoint to read from. */
  isPreview: boolean;
}

function twoDigits(n: number): string {
  return n < 10 ? `0${n}` : `${n}`;
}

export function LeadershipSection({ context, listTitle, isPreview }: ILeadershipSectionProps): React.ReactElement {
  // undefined = still loading (nothing is drawn, so placeholder names never flash)
  const [leaders, setLeaders] = React.useState<ILeader[] | undefined>(undefined);
  const [active, setActive] = React.useState<number>(0);
  const [hovered, setHovered] = React.useState<boolean>(false);
  const [focused, setFocused] = React.useState<boolean>(false);
  const [reducedMotion, setReducedMotion] = React.useState<boolean>(false);
  const [brokenPhotos, setBrokenPhotos] = React.useState<Record<string, boolean>>({});

  const list: ILeader[] = leaders || [];
  const count: number = list.length;

  // ---- Load the leaders from the SharePoint list (HR edits the list, no code change) ----
  React.useEffect((): (() => void) => {
    let cancelled = false;
    const applyLeaders = (items: ILeader[]): void => {
      if (cancelled) {
        return;
      }
      setLeaders(items);
      setActive(0);
      setBrokenPhotos({});
    };

    if (isPreview || !listTitle) {
      applyLeaders(LEADERS);
      return (): void => { cancelled = true; };
    }

    // No $select: columns made by hand have names like Sort_x0020_order, so read whatever is there.
    const url: string =
      `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items?$top=100`;

    context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          throw new Error(`Unable to load the list "${listTitle}" (${response.status}).`);
        }
        return response.json() as Promise<{ value: Array<Record<string, unknown>> }>;
      })
      .then((data: { value: Array<Record<string, unknown>> }) => {
        const rows: Array<{ order: number; id: number; leader: ILeader }> = (data.value || [])
          .map((row: Record<string, unknown>) => {
            const photo: string = asString(pickValue(row, ['LeaderPhoto', 'Photo', 'Picture', 'Image']));
            const linkedin: string = asString(pickValue(row, ['LeaderLinkedIn', 'LinkedIn', 'Linkedin']));
            const order: number = Number(pickValue(row, ['SortOrder', 'Sort_x0020_order', 'Sort_x0020_Order', 'Order']));
            return {
              order: isNaN(order) ? 9999 : order,
              id: Number(row.Id || row.ID || 0),
              leader: {
                name: asString(row.Title).trim(),
                role: asString(pickValue(row, ['Role', 'JobTitle', 'Job_x0020_title', 'Designation'])).trim(),
                photoUrl: photo ? toDirectImageUrl(photo) : undefined,
                linkedinUrl: linkedin || undefined
              }
            };
          })
          .filter((r: { leader: ILeader }): boolean => !!r.leader.name)
          .sort((x: { order: number; id: number }, y: { order: number; id: number }): number => x.order - y.order || x.id - y.id);
        const items: ILeader[] = rows.map((r: { leader: ILeader }): ILeader => r.leader);
        // An empty/missing list falls back to the placeholder cards so the section never disappears.
        applyLeaders(items.length > 0 ? items : LEADERS);
      })
      .catch(() => applyLeaders(LEADERS));

    return (): void => { cancelled = true; };
  }, [context, listTitle, isPreview]);

  // ---- Respect the operating-system "reduce motion" setting ----
  React.useEffect((): (() => void) | void => {
    if (typeof window === 'undefined' || typeof window.matchMedia !== 'function') {
      return;
    }
    const query: MediaQueryList = window.matchMedia('(prefers-reduced-motion: reduce)');
    setReducedMotion(query.matches);
    const onChange = (e: MediaQueryListEvent): void => setReducedMotion(e.matches);
    if (query.addEventListener) {
      query.addEventListener('change', onChange);
      return (): void => query.removeEventListener('change', onChange);
    }
    return undefined;
  }, []);

  const goTo = (i: number): void => {
    if (count > 0) {
      setActive(((i % count) + count) % count);
    }
  };

  // Autoplay is driven by the progress line on the open card: when it finishes, the next card opens.
  // Hovering or focusing the section freezes that line, so nothing moves while someone is reading.
  const autoplay: boolean = count > 1 && !reducedMotion;
  const paused: boolean = hovered || focused;

  return (
    <section
      className={styles.section}
      id="about"
      data-reveal="1"
      onMouseEnter={(): void => setHovered(true)}
      onMouseLeave={(): void => setHovered(false)}
      onFocus={(): void => setFocused(true)}
      onBlur={(e: React.FocusEvent<HTMLElement>): void => {
        if (!e.currentTarget.contains(e.relatedTarget as Node)) {
          setFocused(false);
        }
      }}
    >
      <div className={styles.sectionHeader} id="leadership">
        <h2 className={styles.sectionTitle}>Meet our leadership</h2>
        <span className={styles.sectionHint}>The people guiding SML Cosmos</span>
      </div>

      <div
        className={`${styles.lxRow} ${paused ? styles.lxPaused : ''}`}
        role="group"
        aria-label="Leadership team"
        style={{ minHeight: leaders === undefined ? '540px' : undefined }}
      >
        {list.map((l: ILeader, i: number): React.ReactElement => {
          const on: boolean = i === active;
          const photoKey: string = `${l.name}-${i}`;
          return (
            <div key={photoKey} className={`${styles.lxCard} ${on ? styles.lxCardOn : ''}`}>
              {l.photoUrl && !brokenPhotos[photoKey] ? (
                <img
                  className={styles.lxPhoto}
                  src={l.photoUrl}
                  alt={on ? l.name : ''}
                  onError={(): void => setBrokenPhotos((prev: Record<string, boolean>) => ({ ...prev, [photoKey]: true }))}
                />
              ) : (
                <span className={styles.lxInitials} aria-hidden="true">{getInitialsOf(l.name)}</span>
              )}
              <span className={styles.lxShade} aria-hidden="true" />
              <span className={styles.lxNum} aria-hidden="true">{twoDigits(i + 1)}</span>

              {on ? (
                <div className={styles.lxBody} key={`body-${i}`}>
                  <span className={styles.lxRole}>{l.role || 'SML Cosmos'}</span>
                  <h3 className={styles.lxName}>{l.name}</h3>
                  {l.linkedinUrl && (
                    <a className={styles.lxLink} href={l.linkedinUrl} target="_blank" rel="noopener noreferrer" aria-label={`${l.name} on LinkedIn`}>
                      <Icon iconName="LinkedInLogo" />
                      <span>LinkedIn</span>
                    </a>
                  )}
                </div>
              ) : (
                <>
                  <span className={styles.lxSide} aria-hidden="true"><span>{l.name}</span></span>
                  <button
                    type="button"
                    className={styles.lxPick}
                    aria-label={`Show ${l.name}, ${l.role}`}
                    onClick={(): void => goTo(i)}
                  />
                </>
              )}

              {on && count > 1 && (
                <span className={styles.lxBar} aria-hidden="true">
                  <span
                    key={active}
                    className={`${styles.lxFill} ${autoplay ? styles.lxFillRun : styles.lxFillStill}`}
                    style={{ animationDuration: `${LEADER_AUTOPLAY_MS}ms` }}
                    onAnimationEnd={(e: React.AnimationEvent): void => {
                      if (e.target === e.currentTarget) {
                        goTo(active + 1);
                      }
                    }}
                  />
                </span>
              )}
            </div>
          );
        })}
      </div>
    </section>
  );
}

export default LeadershipSection;