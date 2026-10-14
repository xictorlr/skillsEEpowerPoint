/** Iconos SVG 16x16 con trazo currentColor. Las barras rojas marcan la referencia. */
const S =
  'xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.2" stroke-linecap="round" aria-hidden="true" focusable="false"';
const REF = 'stroke="#d33"';

const box = (x: number, y: number, w: number, h: number) => `<rect x="${x}" y="${y}" width="${w}" height="${h}" fill="currentColor" fill-opacity=".25"/>`;

export const icons: Record<string, string> = {
  alignLeft: `<svg ${S}><line x1="2" y1="1" x2="2" y2="15" ${REF}/>${box(3, 3, 9, 3)}${box(3, 9, 5, 3)}</svg>`,
  alignCenterH: `<svg ${S}><line x1="8" y1="1" x2="8" y2="15" ${REF}/>${box(3, 3, 10, 3)}${box(5.5, 9, 5, 3)}</svg>`,
  alignRight: `<svg ${S}><line x1="14" y1="1" x2="14" y2="15" ${REF}/>${box(4, 3, 9, 3)}${box(8, 9, 5, 3)}</svg>`,
  alignTop: `<svg ${S}><line x1="1" y1="2" x2="15" y2="2" ${REF}/>${box(3, 3, 3, 9)}${box(9, 3, 3, 5)}</svg>`,
  alignCenterV: `<svg ${S}><line x1="1" y1="8" x2="15" y2="8" ${REF}/>${box(3, 3, 3, 10)}${box(9, 5.5, 3, 5)}</svg>`,
  alignBottom: `<svg ${S}><line x1="1" y1="14" x2="15" y2="14" ${REF}/>${box(3, 4, 3, 9)}${box(9, 8, 3, 5)}</svg>`,
  distributeH: `<svg ${S}>${box(1, 4, 3, 8)}${box(6.5, 4, 3, 8)}${box(12, 4, 3, 8)}<path d="M4.5 8h1.5M10 8h1.5"/></svg>`,
  distributeV: `<svg ${S}>${box(4, 1, 8, 3)}${box(4, 6.5, 8, 3)}${box(4, 12, 8, 3)}<path d="M8 4.5v1.5M8 10v1.5"/></svg>`,
  swap: `<svg ${S}>${box(1, 2, 6, 5)}${box(9, 9, 6, 5)}<path d="M8 4.5h5l-1.5-1.5M13 4.5l-1.5 1.5M8 11.5H3l1.5-1.5M3 11.5l1.5 1.5"/></svg>`,
  matrix: `<svg ${S}>${box(1.5, 1.5, 5, 5)}${box(9.5, 1.5, 5, 5)}${box(1.5, 9.5, 5, 5)}${box(9.5, 9.5, 5, 5)}</svg>`,
  slideCenterH: `<svg ${S}><rect x="1" y="2" width="14" height="12" rx="1"/><line x1="8" y1="2" x2="8" y2="14" stroke-dasharray="1.5 1.5" ${REF}/>${box(5, 6, 6, 4)}</svg>`,
  slideCenterV: `<svg ${S}><rect x="1" y="2" width="14" height="12" rx="1"/><line x1="1" y1="8" x2="15" y2="8" stroke-dasharray="1.5 1.5" ${REF}/>${box(4, 5, 8, 6)}</svg>`,
  slideCenter: `<svg ${S}><rect x="1" y="2" width="14" height="12" rx="1"/><line x1="8" y1="2" x2="8" y2="14" stroke-dasharray="1.5 1.5" ${REF}/><line x1="1" y1="8" x2="15" y2="8" stroke-dasharray="1.5 1.5" ${REF}/>${box(5, 5.5, 6, 5)}</svg>`,
  matchWidth: `<svg ${S}>${box(2, 2, 12, 4)}${box(2, 9, 12, 4)}<path d="M2 15h12" ${REF}/></svg>`,
  matchHeight: `<svg ${S}>${box(2, 2, 4, 12)}${box(9, 2, 4, 12)}<path d="M15 2v12" ${REF}/></svg>`,
  matchBoth: `<svg ${S}>${box(1, 1, 6, 6)}${box(9, 9, 6, 6)}<path d="M9 1h6v6" ${REF}/></svg>`,
  matchWidest: `<svg ${S}>${box(1, 2, 14, 4)}${box(1, 9, 8, 4)}<path d="M9 11h6" stroke-dasharray="1 1"/></svg>`,
  matchTallest: `<svg ${S}>${box(2, 1, 4, 14)}${box(9, 1, 4, 8)}<path d="M11 9v6" stroke-dasharray="1 1"/></svg>`,
  matchNarrowest: `<svg ${S}>${box(1, 2, 6, 4)}${box(1, 9, 14, 4)}<path d="M7 11h8" stroke-dasharray="1 1" ${REF}/></svg>`,
  matchShortest: `<svg ${S}>${box(2, 1, 4, 6)}${box(9, 1, 4, 14)}<path d="M11 7v8" stroke-dasharray="1 1" ${REF}/></svg>`,
  stretchLeft: `<svg ${S}><line x1="2" y1="1" x2="2" y2="15" ${REF}/>${box(7, 5, 7, 6)}<path d="M7 8H3M4.5 6.5 3 8l1.5 1.5"/></svg>`,
  stretchRight: `<svg ${S}><line x1="14" y1="1" x2="14" y2="15" ${REF}/>${box(2, 5, 7, 6)}<path d="M9 8h4M11.5 6.5 13 8l-1.5 1.5"/></svg>`,
  stretchTop: `<svg ${S}><line x1="1" y1="2" x2="15" y2="2" ${REF}/>${box(5, 7, 6, 7)}<path d="M8 7V3M6.5 4.5 8 3l1.5 1.5"/></svg>`,
  stretchBottom: `<svg ${S}><line x1="1" y1="14" x2="15" y2="14" ${REF}/>${box(5, 2, 6, 7)}<path d="M8 9v4M6.5 11.5 8 13l1.5-1.5"/></svg>`,
  copyFormat: `<svg ${S}>${box(2, 2, 7, 9)}<path d="M9 4h4v10H5v-2" /><path d="M4 5.5 6 7l2.5-3" ${REF}/></svg>`,
  pasteFormat: `<svg ${S}><path d="M3 3h10v11H3z"/><path d="M6 3V1.5h4V3"/><path d="M5 8h6M5 11h6" ${REF}/></svg>`,
  pasteFill: `<svg ${S}><path d="M3 3h10v11H3z"/>${box(5, 6, 6, 6)}</svg>`,
  pasteLine: `<svg ${S}><path d="M3 3h10v11H3z"/><rect x="5" y="6" width="6" height="6" ${REF}/></svg>`,
  pasteFont: `<svg ${S}><path d="M3 3h10v11H3z"/><path d="M5.5 12 8 5.5 10.5 12M6.3 10h3.4" ${REF}/></svg>`,
  pasteSize: `<svg ${S}><path d="M3 3h10v11H3z"/><path d="M5.5 11.5v-6h6" ${REF}/><path d="M5.5 5.5 7 7M5.5 5.5 7 4"/></svg>`,
  noFill: `<svg ${S}><rect x="3" y="3" width="10" height="10"/><path d="M3 13 13 3" ${REF}/></svg>`,
  noLine: `<svg ${S}><rect x="3" y="3" width="10" height="10" stroke-dasharray="2 2"/><path d="M3 13 13 3" ${REF}/></svg>`,
  transparencyUp: `<svg ${S}>${box(2, 4, 8, 8)}<rect x="6" y="2" width="8" height="8" fill="currentColor" fill-opacity=".1"/><path d="M11 13v3M9.5 14.5h3" ${REF}/></svg>`,
  transparencyDown: `<svg ${S}>${box(2, 4, 8, 8)}<rect x="6" y="2" width="8" height="8" fill="currentColor" fill-opacity=".6"/><path d="M9.5 14.5h3" ${REF}/></svg>`,
  targetFill: `<svg ${S}>${box(3, 3, 10, 10)}</svg>`,
  targetLine: `<svg ${S}><rect x="3" y="3" width="10" height="10" stroke-width="2"/></svg>`,
  targetFont: `<svg ${S}><path d="M4 13 8 3l4 10M5.5 9.5h5"/></svg>`,
  setColor: `<svg ${S}><circle cx="8" cy="8" r="5"/></svg>`,
};

export function iconSvg(id: string): string {
  return icons[id] ?? `<svg ${S}><rect x="2" y="2" width="12" height="12"/></svg>`;
}
