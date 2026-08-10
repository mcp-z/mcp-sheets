import assert from 'assert';
import { buildTextFormatRuns, type MarkdownSpan, type ParsedMarkdown, parseInlineMarkdown } from '../../../src/spreadsheet/markdown-runs.ts';

const LINK_BLUE = { red: 0.06666667, green: 0.33333334, blue: 0.8 };

/** Shared invariants every run list must satisfy, per the Sheets textFormatRuns contract. */
function assertRunInvariants(runs: ReturnType<typeof buildTextFormatRuns>, text: string): void {
  let previousStartIndex = -1;
  for (const run of runs) {
    const startIndex = run.startIndex;
    assert.ok(startIndex !== null && startIndex !== undefined, 'run must have a startIndex');
    assert.ok((startIndex as number) > previousStartIndex, `runs must be strictly increasing (${startIndex} > ${previousStartIndex})`);
    assert.notStrictEqual(startIndex, text.length, 'no run may be emitted at startIndex === text.length');
    previousStartIndex = startIndex as number;
  }
}

describe('markdown-runs', () => {
  describe('parseInlineMarkdown', () => {
    it('parses two markdown links separated by plain text', () => {
      const parsed = parseInlineMarkdown('See [Alpha](https://a.example) and [Beta](https://b.example) for details');
      assert.strictEqual(parsed.text, 'See Alpha and Beta for details');
      assert.strictEqual(parsed.linkCount, 2);
      const [first, second] = parsed.spans;
      assert.deepStrictEqual(first, { start: 4, end: 9, format: { link: { uri: 'https://a.example' } } });
      assert.deepStrictEqual(second, { start: 14, end: 18, format: { link: { uri: 'https://b.example' } } });
    });

    it('trims trailing prose punctuation off bare URLs, keeping it as literal text', () => {
      const parsed = parseInlineMarkdown('see https://x.com/a.');
      assert.strictEqual(parsed.text, 'see https://x.com/a.');
      assert.strictEqual(parsed.linkCount, 1);
      const link = parsed.spans[0] as MarkdownSpan;
      assert.strictEqual(link.format.link?.uri, 'https://x.com/a');
      // The trailing "." is literal text outside the link span.
      assert.strictEqual(parsed.text.slice(link.start, link.end), 'https://x.com/a');
      assert.strictEqual(parsed.text.slice(link.end), '.');
    });

    it('parses bold with ** and __', () => {
      const asterisks = parseInlineMarkdown('**bold**');
      assert.strictEqual(asterisks.text, 'bold');
      assert.deepStrictEqual(asterisks.spans, [{ start: 0, end: 4, format: { bold: true } }]);

      const underscores = parseInlineMarkdown('__bold__');
      assert.strictEqual(underscores.text, 'bold');
      assert.deepStrictEqual(underscores.spans, [{ start: 0, end: 4, format: { bold: true } }]);
    });

    it('parses italic with * and _', () => {
      const asterisk = parseInlineMarkdown('*italic*');
      assert.strictEqual(asterisk.text, 'italic');
      assert.deepStrictEqual(asterisk.spans, [{ start: 0, end: 6, format: { italic: true } }]);

      const underscore = parseInlineMarkdown('_italic_');
      assert.strictEqual(underscore.text, 'italic');
      assert.deepStrictEqual(underscore.spans, [{ start: 0, end: 6, format: { italic: true } }]);
    });

    it('parses strikethrough with ~~', () => {
      const parsed = parseInlineMarkdown('~~gone~~');
      assert.strictEqual(parsed.text, 'gone');
      assert.deepStrictEqual(parsed.spans, [{ start: 0, end: 4, format: { strikethrough: true } }]);
    });

    it('merges bold wrapping a link into a single span covering both formats', () => {
      const parsed = parseInlineMarkdown('**[Acme](https://acme.example)**');
      assert.strictEqual(parsed.text, 'Acme');
      assert.strictEqual(parsed.linkCount, 1);
      assert.strictEqual(parsed.spans.length, 2);
      const linkSpan = parsed.spans.find((span) => span.format.link);
      const boldSpan = parsed.spans.find((span) => span.format.bold);
      assert.deepStrictEqual(linkSpan, { start: 0, end: 4, format: { link: { uri: 'https://acme.example' } } });
      assert.deepStrictEqual(boldSpan, { start: 0, end: 4, format: { bold: true } });
    });

    it('treats backslash-escaped delimiters as literal characters with no span', () => {
      const parsed = parseInlineMarkdown('\\*not bold\\*');
      assert.strictEqual(parsed.text, '*not bold*');
      assert.strictEqual(parsed.spans.length, 0);
    });

    it('passes through unsupported block syntax as literal text', () => {
      const header = parseInlineMarkdown('# Header');
      assert.strictEqual(header.text, '# Header');
      assert.strictEqual(header.spans.length, 0);

      const backticks = parseInlineMarkdown('`code span`');
      assert.strictEqual(backticks.text, '`code span`');
      assert.strictEqual(backticks.spans.length, 0);
    });

    it('passes through unmatched delimiters as literal text', () => {
      const parsed = parseInlineMarkdown('a * b');
      assert.strictEqual(parsed.text, 'a * b');
      assert.strictEqual(parsed.spans.length, 0);
    });

    it('does not pair whitespace-flanked delimiters (flanking rules)', () => {
      // Without flanking rules these would italicize " 4 " and swallow the asterisks.
      const arithmetic = parseInlineMarkdown('3 * 4 * 5');
      assert.strictEqual(arithmetic.text, '3 * 4 * 5');
      assert.strictEqual(arithmetic.spans.length, 0);

      const doubles = parseInlineMarkdown('x ** y ** z');
      assert.strictEqual(doubles.text, 'x ** y ** z');
      assert.strictEqual(doubles.spans.length, 0);
    });

    it('treats intraword underscores as literal, never emphasis', () => {
      // Identifier-style content is common in spreadsheets; underscores must survive.
      const parsed = parseInlineMarkdown('user_id and account_id');
      assert.strictEqual(parsed.text, 'user_id and account_id');
      assert.strictEqual(parsed.spans.length, 0);
    });

    it('still allows intraword asterisk emphasis, per CommonMark', () => {
      const parsed = parseInlineMarkdown('a*b*c');
      assert.strictEqual(parsed.text, 'abc');
      assert.deepStrictEqual(parsed.spans, [{ start: 1, end: 2, format: { italic: true } }]);
    });

    it('accounts for UTF-16 code units when an astral character (emoji) precedes a link', () => {
      // "😀" is a surrogate pair -> 2 UTF-16 code units, not 1.
      const parsed = parseInlineMarkdown('😀 [go](https://example.com)');
      assert.strictEqual(parsed.text, '😀 go');
      assert.strictEqual(parsed.linkCount, 1);
      const link = parsed.spans[0] as MarkdownSpan;
      assert.strictEqual(link.start, '😀 '.length); // 3 code units (2 surrogate + 1 space), not 2
      assert.strictEqual(parsed.text.slice(link.start, link.end), 'go');
    });
  });

  describe('buildTextFormatRuns', () => {
    it('emits link runs with a reset between and after when the second link is not at end-of-string', () => {
      const parsed = parseInlineMarkdown('See [Alpha](https://a.example) and [Beta](https://b.example) for details');
      const runs = buildTextFormatRuns(parsed);
      assertRunInvariants(runs, parsed.text);

      // baseline, link(Alpha), reset, link(Beta), reset
      assert.strictEqual(runs.length, 5);
      assert.strictEqual(runs[0]?.startIndex, 0);
      assert.deepStrictEqual(runs[0]?.format, {});

      assert.strictEqual(runs[1]?.startIndex, 4);
      assert.strictEqual(runs[1]?.format?.link?.uri, 'https://a.example');
      assert.strictEqual(runs[1]?.format?.underline, true);
      assert.deepStrictEqual(runs[1]?.format?.foregroundColorStyle, { rgbColor: LINK_BLUE });

      assert.strictEqual(runs[2]?.startIndex, 9);
      assert.deepStrictEqual(runs[2]?.format, {});

      assert.strictEqual(runs[3]?.startIndex, 14);
      assert.strictEqual(runs[3]?.format?.link?.uri, 'https://b.example');

      assert.strictEqual(runs[4]?.startIndex, 18);
      assert.deepStrictEqual(runs[4]?.format, {});
    });

    it('omits the trailing reset run when the link reaches the end of the string', () => {
      const parsed = parseInlineMarkdown('See [Alpha](https://a.example)');
      const runs = buildTextFormatRuns(parsed);
      assertRunInvariants(runs, parsed.text);

      assert.strictEqual(runs.length, 2);
      assert.strictEqual(runs[0]?.startIndex, 0);
      assert.deepStrictEqual(runs[0]?.format, {});
      assert.strictEqual(runs[1]?.startIndex, 4);
      assert.strictEqual(runs[1]?.format?.link?.uri, 'https://a.example');
    });

    it('does not emit a separate plain baseline run when a link starts at index 0', () => {
      const parsed = parseInlineMarkdown('[Alpha](https://a.example) tail');
      const runs = buildTextFormatRuns(parsed);
      assertRunInvariants(runs, parsed.text);

      assert.strictEqual(runs.length, 2);
      assert.strictEqual(runs[0]?.startIndex, 0);
      assert.strictEqual(runs[0]?.format?.link?.uri, 'https://a.example');
      assert.strictEqual(runs[1]?.startIndex, 5);
      assert.deepStrictEqual(runs[1]?.format, {});
    });

    it('handles two adjacent links with strictly increasing indices and no separate reset between them', () => {
      const parsed = parseInlineMarkdown('[a](https://a.example)[b](https://b.example)');
      const runs = buildTextFormatRuns(parsed);
      assertRunInvariants(runs, parsed.text);

      assert.strictEqual(parsed.text, 'ab');
      assert.strictEqual(runs.length, 2);
      assert.strictEqual(runs[0]?.startIndex, 0);
      assert.strictEqual(runs[0]?.format?.link?.uri, 'https://a.example');
      assert.strictEqual(runs[1]?.startIndex, 1);
      assert.strictEqual(runs[1]?.format?.link?.uri, 'https://b.example');
    });

    it('produces a single link run (no style) with styleLinks=false', () => {
      const parsed = parseInlineMarkdown('[a](https://a.example)');
      const runs = buildTextFormatRuns(parsed, false);
      assertRunInvariants(runs, parsed.text);
      assert.deepStrictEqual(runs, [{ startIndex: 0, format: { link: { uri: 'https://a.example' } } }]);
    });

    it('emits a single empty-format baseline run for plain text with no spans', () => {
      const parsed: ParsedMarkdown = { text: '# Header', spans: [], linkCount: 0 };
      const runs = buildTextFormatRuns(parsed);
      assertRunInvariants(runs, parsed.text);
      assert.deepStrictEqual(runs, [{ startIndex: 0, format: {} }]);
    });

    it('returns no runs for empty text', () => {
      const parsed: ParsedMarkdown = { text: '', spans: [], linkCount: 0 };
      assert.deepStrictEqual(buildTextFormatRuns(parsed), []);
    });

    it('satisfies run invariants across a mixed-formatting case', () => {
      const parsed = parseInlineMarkdown('😀 **[go](https://example.com)** and *also* ~~this~~ plain https://tail.example/x!');
      const runs = buildTextFormatRuns(parsed);
      assertRunInvariants(runs, parsed.text);
      // Sanity: the last URL is at end-of-string minus trimmed punctuation, so its run must exist
      // but no run should exist at text.length.
      assert.ok(runs.some((run) => (run.format?.link?.uri ?? '') === 'https://tail.example/x'));
    });
  });
});
