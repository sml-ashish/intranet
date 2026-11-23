import * as React from 'react';

interface IBrandLogoProps {
  className?: string;
}

/**
 * SML Cosmos mark. Rendered inline so the web part has no external image
 * dependency and works before a custom logo is configured.
 */
export default function BrandLogo({ className }: IBrandLogoProps): React.ReactElement<IBrandLogoProps> {
  return (
    <svg
      className={className}
      viewBox="0 0 64 64"
      role="img"
      aria-label="SML Cosmos"
      xmlns="http://www.w3.org/2000/svg"
    >
      <defs>
        <linearGradient id="smlCosmosBg" x1="0" y1="0" x2="1" y2="1">
          <stop offset="0%" stopColor="#991b1b" />
          <stop offset="55%" stopColor="#dc2626" />
          <stop offset="100%" stopColor="#86efac" />
        </linearGradient>
        <linearGradient id="smlCosmosRing" x1="0" y1="0" x2="1" y2="1">
          <stop offset="0%" stopColor="#dcfce7" />
          <stop offset="100%" stopColor="#fecaca" />
        </linearGradient>
      </defs>

      <rect width="64" height="64" rx="16" fill="url(#smlCosmosBg)" />

      <ellipse
        cx="32"
        cy="32"
        rx="22"
        ry="9"
        fill="none"
        stroke="url(#smlCosmosRing)"
        strokeWidth="2.4"
        opacity="0.85"
        transform="rotate(-28 32 32)"
      />

      <circle cx="32" cy="32" r="7.5" fill="#ffffff" />
      <circle cx="48.5" cy="22.5" r="3.2" fill="#dcfce7" />
      <circle cx="16.5" cy="41.5" r="2.2" fill="#ffffff" opacity="0.75" />
    </svg>
  );
}
