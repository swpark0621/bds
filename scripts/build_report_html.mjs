import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = dirname(dirname(fileURLToPath(import.meta.url)));

const reports = [
  {
    source: '서계동_재개발_조사보고서.md',
    output: '서계동_재개발_조사보고서.html',
    title: '용산 서계동 재개발 조사 보고서',
    shortTitle: '서계동',
    description: '서계 통합구역의 공식 단계, 사업성, 실거래, 권리 리스크를 정리한 HTML 보고서',
    accent: '#2563eb',
  },
  {
    source: '청파동_재개발_조사보고서.md',
    output: '청파동_재개발_조사보고서.html',
    title: '용산 청파동 재개발 조사 보고서',
    shortTitle: '청파동',
    description: '청파2구역, 청파제1구역, 청파동1가 46번지 일대의 차이를 분리한 HTML 보고서',
    accent: '#0f766e',
  },
  {
    source: '후암동_재개발_조사보고서.md',
    output: '후암동_재개발_조사보고서.html',
    title: '용산 후암동 재개발·재건축 조사 보고서',
    shortTitle: '후암동',
    description: '동후암1구역, 동후암3구역, 후암동제1구역 참고군을 비교한 HTML 보고서',
    accent: '#b7791f',
  },
  {
    source: '서계통합_청파2_비교_실거래_프리미엄_보고서.md',
    output: '서계통합_청파2_비교_실거래_프리미엄_보고서.html',
    title: '용산 서계 통합·청파2 비교 및 실거래 프리미엄 보고서',
    shortTitle: '서계·청파2',
    description: '조합설립인가 이후의 사업 비교와 실거래 기반 기대감 프리미엄을 같은 기준으로 정리한 HTML 보고서',
    accent: '#9a3412',
  },
];

function escapeHtml(value) {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function inlineMarkdown(value) {
  let html = escapeHtml(value);
  html = html.replace(/`([^`]+)`/g, '<code>$1</code>');
  html = html.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
  html = html.replace(
    /(https?:\/\/[^\s<)]+)([).,]?)?/g,
    (_match, url, tail = '') => `<a href="${url}" target="_blank" rel="noopener noreferrer">${url}</a>${tail}`,
  );
  return html;
}

function normalizeHeading(value) {
  return value.replace(/^#+\s+/, '').trim();
}

function parseTableRow(line) {
  return line
    .trim()
    .replace(/^\|/, '')
    .replace(/\|$/, '')
    .split('|')
    .map((cell) => cell.trim());
}

function isTableStart(lines, index) {
  const current = lines[index]?.trim() ?? '';
  const next = lines[index + 1]?.trim() ?? '';
  return current.startsWith('|') && current.endsWith('|') && /^\|?[\s:-]+\|[\s|:-]*$/.test(next);
}

function renderMarkdown(markdown) {
  const lines = markdown.replace(/\r\n/g, '\n').split('\n');
  const toc = [];
  const html = [];
  let index = 0;
  let sectionId = 0;

  while (index < lines.length) {
    const line = lines[index];
    const trimmed = line.trim();

    if (!trimmed) {
      index += 1;
      continue;
    }

    if (/^---+$/.test(trimmed)) {
      html.push('<hr>');
      index += 1;
      continue;
    }

    const heading = trimmed.match(/^(#{1,4})\s+(.+)$/);
    if (heading) {
      const level = heading[1].length;
      const text = normalizeHeading(trimmed);
      const id = `section-${++sectionId}`;
      if (level >= 2 && level <= 3) toc.push({ id, level, text });
      html.push(`<h${level} id="${id}">${inlineMarkdown(text)}</h${level}>`);
      index += 1;
      continue;
    }

    if (trimmed.startsWith('>')) {
      const quote = [];
      while (index < lines.length && lines[index].trim().startsWith('>')) {
        quote.push(lines[index].trim().replace(/^>\s?/, ''));
        index += 1;
      }
      html.push(`<blockquote>${quote.map(inlineMarkdown).join('<br>')}</blockquote>`);
      continue;
    }

    if (isTableStart(lines, index)) {
      const rows = [];
      while (index < lines.length && lines[index].trim().startsWith('|') && lines[index].trim().endsWith('|')) {
        rows.push(lines[index]);
        index += 1;
      }

      const headers = parseTableRow(rows[0]);
      const bodyRows = rows.slice(2).map(parseTableRow);
      html.push('<div class="table-scroll"><table>');
      html.push('<thead><tr>');
      headers.forEach((cell) => html.push(`<th>${inlineMarkdown(cell)}</th>`));
      html.push('</tr></thead>');
      html.push('<tbody>');
      bodyRows.forEach((row) => {
        html.push('<tr>');
        row.forEach((cell) => html.push(`<td>${inlineMarkdown(cell)}</td>`));
        html.push('</tr>');
      });
      html.push('</tbody></table></div>');
      continue;
    }

    if (/^-\s+/.test(trimmed)) {
      html.push('<ul>');
      while (index < lines.length && /^-\s+/.test(lines[index].trim())) {
        html.push(`<li>${inlineMarkdown(lines[index].trim().replace(/^-\s+/, ''))}</li>`);
        index += 1;
      }
      html.push('</ul>');
      continue;
    }

    if (/^\d+\.\s+/.test(trimmed)) {
      html.push('<ol>');
      while (index < lines.length && /^\d+\.\s+/.test(lines[index].trim())) {
        html.push(`<li>${inlineMarkdown(lines[index].trim().replace(/^\d+\.\s+/, ''))}</li>`);
        index += 1;
      }
      html.push('</ol>');
      continue;
    }

    const paragraph = [];
    while (
      index < lines.length &&
      lines[index].trim() &&
      !/^---+$/.test(lines[index].trim()) &&
      !/^(#{1,4})\s+/.test(lines[index].trim()) &&
      !lines[index].trim().startsWith('>') &&
      !isTableStart(lines, index) &&
      !/^-\s+/.test(lines[index].trim()) &&
      !/^\d+\.\s+/.test(lines[index].trim())
    ) {
      paragraph.push(lines[index].trim());
      index += 1;
    }
    html.push(`<p>${inlineMarkdown(paragraph.join(' '))}</p>`);
  }

  return { body: html.join('\n'), toc };
}

function renderPage(report, markdown) {
  const { body, toc } = renderMarkdown(markdown);
  const reportLinks = reports
    .map((item) => {
      const current = item.output === report.output ? ' aria-current="page"' : '';
      return `<a href="${item.output}"${current}>${item.shortTitle}</a>`;
    })
    .join('\n');
  const tocLinks = toc
    .map((item) => `<a class="level-${item.level}" href="#${item.id}">${inlineMarkdown(item.text)}</a>`)
    .join('\n');

  return `<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>${escapeHtml(report.title)}</title>
  <style>
    :root {
      --bg: #f6f7f8;
      --surface: #ffffff;
      --ink: #20262d;
      --muted: #69737d;
      --line: #d8dde3;
      --accent: ${report.accent};
      --accent-soft: color-mix(in srgb, var(--accent) 12%, white);
      --shadow: 0 12px 30px rgba(32, 38, 45, 0.08);
    }

    * { box-sizing: border-box; }

    html { scroll-behavior: smooth; }

    body {
      margin: 0;
      background: var(--bg);
      color: var(--ink);
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Noto Sans KR", "Apple SD Gothic Neo", Arial, sans-serif;
      font-size: 16px;
      line-height: 1.7;
      letter-spacing: 0;
    }

    a { color: var(--accent); text-decoration: none; }
    a:hover { text-decoration: underline; }

    .layout {
      display: grid;
      grid-template-columns: 280px minmax(0, 1fr);
      min-height: 100vh;
    }

    aside {
      position: sticky;
      top: 0;
      height: 100vh;
      padding: 22px 18px;
      background: #fff;
      border-right: 1px solid var(--line);
      overflow: auto;
    }

    .brand {
      display: grid;
      gap: 5px;
      margin-bottom: 16px;
    }

    .brand strong {
      font-size: 18px;
      line-height: 1.3;
    }

    .brand span {
      color: var(--muted);
      font-size: 13px;
    }

    .quick-links,
    .toc {
      display: grid;
      gap: 7px;
      margin-top: 16px;
      padding-top: 16px;
      border-top: 1px solid var(--line);
    }

    .quick-links a,
    .toc a {
      display: block;
      border-radius: 8px;
      padding: 8px 10px;
      color: #34404a;
      font-size: 14px;
    }

    .quick-links a {
      border: 1px solid var(--line);
      background: #fbfbfc;
    }

    .quick-links a[aria-current="page"] {
      border-color: color-mix(in srgb, var(--accent) 46%, var(--line));
      background: var(--accent-soft);
      color: var(--accent);
      font-weight: 700;
    }

    .toc a:hover,
    .quick-links a:hover {
      background: #eef2f5;
      text-decoration: none;
    }

    .toc .level-3 {
      padding-left: 22px;
      font-size: 13px;
      color: var(--muted);
    }

    main {
      min-width: 0;
      padding: 30px;
    }

    .hero {
      max-width: 1120px;
      margin: 0 auto 18px;
      display: grid;
      gap: 14px;
    }

    .eyebrow {
      color: var(--accent);
      font-weight: 760;
      font-size: 13px;
    }

    .hero h1 {
      margin: 0;
      font-size: clamp(28px, 4vw, 44px);
      line-height: 1.18;
      letter-spacing: 0;
    }

    .hero p {
      margin: 0;
      color: var(--muted);
      max-width: 840px;
    }

    .actions {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
    }

    .button {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-height: 40px;
      border-radius: 8px;
      border: 1px solid var(--line);
      background: var(--surface);
      color: #2c3640;
      padding: 0 12px;
      font-size: 14px;
      box-shadow: 0 2px 8px rgba(32, 38, 45, 0.05);
    }

    .button.primary {
      border-color: var(--accent);
      background: var(--accent);
      color: #fff;
    }

    .button:hover {
      text-decoration: none;
      background: #f9fafb;
    }

    .button.primary:hover {
      background: color-mix(in srgb, var(--accent) 88%, black);
    }

    article {
      max-width: 1120px;
      margin: 0 auto;
      background: var(--surface);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
      padding: 28px;
    }

    article h1 {
      margin: 0 0 14px;
      font-size: clamp(24px, 3vw, 34px);
      line-height: 1.25;
      letter-spacing: 0;
    }

    article h2 {
      margin: 34px 0 12px;
      padding-top: 10px;
      border-top: 1px solid var(--line);
      font-size: 24px;
      line-height: 1.35;
      letter-spacing: 0;
    }

    article h3 {
      margin: 26px 0 10px;
      font-size: 19px;
      line-height: 1.4;
      letter-spacing: 0;
    }

    article p,
    article ul,
    article ol,
    article blockquote {
      margin: 0 0 15px;
    }

    article ul,
    article ol {
      padding-left: 22px;
    }

    article li + li {
      margin-top: 6px;
    }

    blockquote {
      border-left: 4px solid var(--accent);
      background: var(--accent-soft);
      padding: 13px 16px;
      border-radius: 0 8px 8px 0;
      color: #31404d;
    }

    hr {
      border: 0;
      border-top: 1px solid var(--line);
      margin: 26px 0;
    }

    strong { font-weight: 780; }

    code {
      border-radius: 6px;
      background: #eef1f4;
      padding: 2px 5px;
      font-family: "SFMono-Regular", Consolas, "Liberation Mono", monospace;
      font-size: 0.92em;
    }

    .table-scroll {
      width: 100%;
      overflow-x: auto;
      margin: 14px 0 18px;
      border: 1px solid var(--line);
      border-radius: 8px;
    }

    table {
      width: 100%;
      min-width: 720px;
      border-collapse: collapse;
      background: #fff;
    }

    th,
    td {
      padding: 10px 12px;
      border-bottom: 1px solid var(--line);
      border-right: 1px solid var(--line);
      text-align: left;
      vertical-align: top;
      font-size: 14px;
    }

    th:last-child,
    td:last-child {
      border-right: 0;
    }

    tr:last-child td {
      border-bottom: 0;
    }

    th {
      background: #f2f5f7;
      font-weight: 740;
      color: #28323b;
    }

    @media (max-width: 900px) {
      .layout {
        display: block;
      }

      aside {
        position: static;
        height: auto;
        border-right: 0;
        border-bottom: 1px solid var(--line);
      }

      .toc {
        max-height: 260px;
        overflow: auto;
      }

      main {
        padding: 18px;
      }

      article {
        padding: 20px;
      }
    }

    @media (max-width: 560px) {
      body { font-size: 15px; }

      main {
        padding: 14px;
      }

      article {
        padding: 16px;
      }

      .hero h1 {
        font-size: 29px;
      }

      article h2 {
        font-size: 21px;
      }

      table {
        min-width: 620px;
      }

      th,
      td {
        padding: 9px;
      }
    }

    @media print {
      aside,
      .actions {
        display: none !important;
      }

      .layout {
        display: block;
      }

      main {
        padding: 0;
      }

      article {
        box-shadow: none;
        border: 0;
      }
    }
  </style>
</head>
<body>
  <div class="layout">
    <aside>
      <div class="brand">
        <strong>용산 재개발 보고서</strong>
        <span>서계 · 청파 · 후암 비교 자료</span>
      </div>
      <div class="quick-links" aria-label="보고서 이동">
        <a href="서계동_재개발_스터디_대시보드.html">대시보드</a>
        ${reportLinks}
      </div>
      <nav class="toc" aria-label="문서 목차">
        ${tocLinks}
      </nav>
    </aside>
    <main>
      <header class="hero">
        <div class="eyebrow">${escapeHtml(report.shortTitle)} HTML 보고서</div>
        <h1>${escapeHtml(report.title)}</h1>
        <p>${escapeHtml(report.description)}</p>
        <div class="actions">
          <a class="button primary" href="서계동_재개발_스터디_대시보드.html">대시보드</a>
          <a class="button" href="${report.source}">마크다운 원문</a>
          <button class="button" type="button" onclick="window.print()">인쇄</button>
        </div>
      </header>
      <article>
${body}
      </article>
    </main>
  </div>
</body>
</html>
`;
}

reports.forEach((report) => {
  const markdown = readFileSync(join(root, report.source), 'utf8');
  const html = renderPage(report, markdown);
  writeFileSync(join(root, report.output), html, 'utf8');
  console.log(`wrote ${report.output}`);
});
