/**
 * Inline Markdown → Google Sheets textFormatRuns
 *
 * The Sheets values API can only write plain text or formulas — it cannot set
 * `textFormatRuns`, so plain-text URLs are never linkified and `=HYPERLINK()`
 * only ever produces one link per cell. Multiple distinct clickable links (or
 * any other rich text) inside a single cell require `textFormatRuns`, which is
 * only settable via `spreadsheets.batchUpdate` → `updateCells`. This module
 * parses a small, deliberately restricted subset of inline markdown into plain
 * text plus the `textFormatRuns` needed to render it.
 *
 * Supported inline constructs:
 *   - `[label](url)` and bare `http(s)://` URLs -> `format.link.uri`
 *   - `**bold**` / `__bold__`                    -> `format.bold`
 *   - `*italic*` / `_italic_`                     -> `format.italic`
 *   - `~~strikethrough~~`                         -> `format.strikethrough`
 *   - `\X` escapes any of the delimiter characters above to a literal `X`
 *
 * Everything else (headers, lists, tables, backticks, unmatched delimiters,
 * ...) is passed through as literal text. This module never throws on
 * malformed markdown — graceful degradation to literal text is the contract.
 *
 * A note on indices: Google Sheets `textFormatRuns[].startIndex` is defined in
 * UTF-16 code units. JavaScript string indexing (`str[i]`, `str.length`,
 * `str.slice()`) already operates on UTF-16 code units, so plain indexing is
 * exactly correct here. Do NOT iterate with `[...str]` or any other code-point
 * aware iteration — that would silently misalign indices for any text
 * containing astral characters (e.g. emoji), which are represented as
 * surrogate pairs (2 code units).
 */

import type { sheets_v4 } from 'googleapis';

// Default styling applied to link runs so they read as clickable links (Sheets'
// own link color), matching what Sheets applies automatically for =HYPERLINK().
const LINK_BLUE: sheets_v4.Schema$Color = { red: 0.06666667, green: 0.33333334, blue: 0.8 };

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/** A formatting span over `text`, in UTF-16 code unit coordinates (half-open: [start, end)). */
export interface MarkdownSpan {
  start: number;
  end: number;
  format: sheets_v4.Schema$TextFormat;
}

export interface ParsedMarkdown {
  /** Plain text with all markdown syntax stripped (or literalized). */
  text: string;
  /** Formatting spans over `text`. May overlap/nest (e.g. a bold span wrapping a link span). */
  spans: MarkdownSpan[];
  /** Number of link spans found (spans whose format has `link` set). */
  linkCount: number;
}

// ---------------------------------------------------------------------------
// Tokenizer
// ---------------------------------------------------------------------------

type EmphasisKind = 'bold' | 'italic' | 'strikethrough';

type Token = { type: 'text'; value: string } | { type: 'link'; label: string; uri: string } | { type: 'delim'; kind: EmphasisKind; raw: string; canOpen: boolean; canClose: boolean };

// Matched against a suffix of the input (anchored with ^), so `.exec` always tests position 0.
const MD_LINK = /^\[([^\]]*?)\]\(\s*(<[^>]+>|[^)\s]+)(?:\s+"[^"]*")?\s*\)/;
const BARE_URL = /^https?:\/\/[^\s<>"'()[\]]+/;

// Characters that may be backslash-escaped to a literal (the delimiter chars this
// parser recognizes as syntax, plus the backslash itself).
const ESCAPABLE = new Set(['\\', '*', '_', '~', '[', ']', '(', ')']);

/** Trailing punctuation that is almost certainly prose, not part of a bare URL. */
function trimTrailingPunctuation(url: string): string {
  return url.replace(/[.,;:!?]+$/, '');
}

const isWhitespaceOrEdge = (ch: string | undefined): boolean => ch === undefined || /\s/.test(ch);
const isWordChar = (ch: string | undefined): boolean => ch !== undefined && /[A-Za-z0-9]/.test(ch);

/**
 * Simplified CommonMark flanking rules — without these, whitespace-flanked
 * delimiters pair up ("3 * 4 * 5" would italicize " 4 ") and intraword
 * underscores become emphasis ("user_id and account_id" would italicize
 * "id and account" and delete the underscores). Both are silent data
 * corruption for ordinary spreadsheet content, so:
 *   - a delimiter can OPEN emphasis only if not followed by whitespace/end
 *   - a delimiter can CLOSE emphasis only if not preceded by whitespace/start
 *   - `_`/`__` additionally can never open after, or close before, a word
 *     character (intraword underscores are literal, per CommonMark)
 * Intraword `*` is intentionally still allowed (`a*b*c` emphasizes `b`),
 * matching CommonMark.
 */
function delimiterFlanking(_kind: EmphasisKind, raw: string, prev: string | undefined, next: string | undefined): { canOpen: boolean; canClose: boolean } {
  const underscore = raw.startsWith('_');
  const canOpen = !isWhitespaceOrEdge(next) && !(underscore && isWordChar(prev));
  const canClose = !isWhitespaceOrEdge(prev) && !(underscore && isWordChar(next));
  return { canOpen, canClose };
}

/**
 * Scans the raw markdown left-to-right into a flat token stream: literal text
 * runs, resolved links (markdown-link or bare-URL), and emphasis delimiters.
 * Emphasis delimiters are emitted as opaque tokens here — whether a given
 * delimiter is an opener, a closer, or unmatched (and thus literal) is decided
 * in a second pass, since that requires seeing the whole stream.
 */
function tokenize(markdown: string): Token[] {
  const tokens: Token[] = [];
  let buffer = '';
  const flush = () => {
    if (buffer) {
      tokens.push({ type: 'text', value: buffer });
      buffer = '';
    }
  };

  let i = 0;
  while (i < markdown.length) {
    const ch = markdown[i] as string;
    const next = markdown[i + 1];

    // Backslash escapes: \* \_ \~ \[ \] \( \) \\  -> literal char, never treated as syntax.
    if (ch === '\\' && next !== undefined && ESCAPABLE.has(next)) {
      buffer += next;
      i += 2;
      continue;
    }

    // Markdown links: [label](url) — checked before bare URLs so the label wins.
    const linkMatch = MD_LINK.exec(markdown.slice(i));
    if (linkMatch) {
      flush();
      const rawLabel = (linkMatch[1] ?? '').trim();
      const uri = (linkMatch[2] ?? '').replace(/^<|>$/g, '');
      tokens.push({ type: 'link', label: rawLabel.length > 0 ? rawLabel : uri, uri });
      i += linkMatch[0].length;
      continue;
    }

    // Bare http(s) URLs — the URL itself becomes the label.
    if (ch === 'h') {
      const bareMatch = BARE_URL.exec(markdown.slice(i));
      if (bareMatch) {
        flush();
        const raw = bareMatch[0];
        const uri = trimTrailingPunctuation(raw);
        tokens.push({ type: 'link', label: uri, uri });
        // Trailing punctuation trimmed off the URL is not part of the link — keep it as literal text.
        if (uri.length < raw.length) buffer += raw.slice(uri.length);
        i += raw.length;
        continue;
      }
    }

    // Emphasis delimiters — longest first: ** and __ (bold) and ~~ (strikethrough) before * and _ (italic).
    // The delimiter's width is fixed BEFORE the flanking check: a ** that fails
    // flanking is two literal asterisks, never reinterpreted as two * delimiters.
    const two = markdown.slice(i, i + 2);
    const kind: EmphasisKind | undefined = two === '**' || two === '__' ? 'bold' : two === '~~' ? 'strikethrough' : ch === '*' || ch === '_' ? 'italic' : undefined;
    if (kind !== undefined) {
      const raw = kind === 'italic' ? ch : two;
      const { canOpen, canClose } = delimiterFlanking(kind, raw, markdown[i - 1], markdown[i + raw.length]);
      if (canOpen || canClose) {
        flush();
        tokens.push({ type: 'delim', kind, raw, canOpen, canClose });
      } else {
        // Fails flanking on both sides — literal text (e.g. " * " or intraword "_").
        buffer += raw;
      }
      i += raw.length;
      continue;
    }

    buffer += ch;
    i += 1;
  }
  flush();
  return tokens;
}

// ---------------------------------------------------------------------------
// Delimiter matching (unmatched delimiters degrade to literal text)
// ---------------------------------------------------------------------------

/** Matches emphasis delimiter tokens by kind using a simple open/close stack (assumes well-nested markdown). */
function matchDelimiters(tokens: Token[]): Map<number, number> {
  const stack: { kind: EmphasisKind; tokenIndex: number }[] = [];
  const openerToCloser = new Map<number, number>();

  tokens.forEach((token, index) => {
    if (token.type !== 'delim') return;
    const top = stack[stack.length - 1];
    if (token.canClose && top && top.kind === token.kind) {
      stack.pop();
      openerToCloser.set(top.tokenIndex, index);
    } else if (token.canOpen) {
      stack.push({ kind: token.kind, tokenIndex: index });
    }
    // Can neither close the current context nor open — stays unmatched (literal).
  });

  // Anything left on the stack never closed — leave unmatched (rendered literally by the caller).
  return openerToCloser;
}

const FORMAT_FOR_KIND: Record<EmphasisKind, sheets_v4.Schema$TextFormat> = {
  bold: { bold: true },
  italic: { italic: true },
  strikethrough: { strikethrough: true },
};

/** Walks the token stream, resolving matched delimiter pairs into spans and unmatched ones into literal text. */
function buildTextAndSpans(tokens: Token[]): { text: string; spans: MarkdownSpan[] } {
  const openerToCloser = matchDelimiters(tokens);
  const closerToOpener = new Map<number, number>();
  for (const [opener, closer] of openerToCloser) closerToOpener.set(closer, opener);

  let text = '';
  const spans: MarkdownSpan[] = [];
  const openerStart = new Map<number, number>(); // opener token index -> plain-text start offset

  tokens.forEach((token, index) => {
    if (token.type === 'text') {
      text += token.value;
      return;
    }

    if (token.type === 'link') {
      const start = text.length;
      text += token.label;
      spans.push({ start, end: text.length, format: { link: { uri: token.uri } } });
      return;
    }

    // Emphasis delimiter token.
    if (openerToCloser.has(index)) {
      // Matched opener — record where its span begins; the delimiter itself is consumed (not emitted).
      openerStart.set(index, text.length);
      return;
    }
    const openerIndex = closerToOpener.get(index);
    if (openerIndex !== undefined) {
      // Matched closer — emit the span from the recorded opener position to here.
      const start = openerStart.get(openerIndex);
      if (start === undefined) {
        throw new Error('markdown-runs: internal error — emphasis opener position was not recorded');
      }
      spans.push({ start, end: text.length, format: FORMAT_FOR_KIND[token.kind] });
      return;
    }
    // Unmatched delimiter — no syntactic partner, so it degrades to its literal characters.
    text += token.raw;
  });

  return { text, spans };
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/** Parses inline markdown into plain text plus formatting spans (see module docs for the supported subset). */
export function parseInlineMarkdown(markdown: string): ParsedMarkdown {
  const tokens = tokenize(markdown);
  const { text, spans } = buildTextAndSpans(tokens);
  const linkCount = spans.filter((span) => span.format.link !== undefined).length;
  return { text, spans, linkCount };
}

/**
 * Builds `textFormatRuns` from parsed markdown spans.
 *
 * Three Sheets API rules drive this, and are easy to get wrong:
 *   1. A run's format persists to the end of the string unless a later run
 *      resets it with an empty `format: {}` — runs are NOT scoped to a range.
 *   2. A run must never be emitted at `startIndex === text.length`: a span
 *      that reaches the end of the string needs no "turn it off" run, because
 *      there's no text left for stray formatting to leak into.
 *   3. Runs must be listed in strictly increasing `startIndex` order, with no
 *      two runs sharing an index.
 *
 * The approach: collect every index where the set of "active" spans changes
 * (every span start, and every span end that isn't end-of-string), merge the
 * formats of all spans active at each such index into one run, then ensure
 * there is a run at index 0 — either a span already starts there, or an
 * explicit empty-format baseline is prepended so the cell doesn't inherit any
 * pre-existing rich text formatting.
 */
export function buildTextFormatRuns(parsed: ParsedMarkdown, styleLinks = true): sheets_v4.Schema$TextFormatRun[] {
  const { text, spans } = parsed;
  if (text.length === 0) return [];

  const boundaries = new Set<number>();
  for (const span of spans) {
    boundaries.add(span.start);
    if (span.end < text.length) boundaries.add(span.end);
  }
  const sortedBoundaries = [...boundaries].sort((a, b) => a - b);

  const runs: sheets_v4.Schema$TextFormatRun[] = sortedBoundaries.map((startIndex) => {
    const format: sheets_v4.Schema$TextFormat = {};
    for (const span of spans) {
      if (span.start <= startIndex && startIndex < span.end) {
        Object.assign(format, span.format);
      }
    }
    if (styleLinks && format.link) {
      format.underline = true;
      format.foregroundColorStyle = { rgbColor: LINK_BLUE };
    }
    return { startIndex, format };
  });

  if (runs.length === 0 || runs[0]?.startIndex !== 0) {
    runs.unshift({ startIndex: 0, format: {} });
  }

  return runs;
}
