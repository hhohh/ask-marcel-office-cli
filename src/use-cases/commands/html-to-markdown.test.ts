import { describe, expect, it } from 'bun:test';
import { htmlToMarkdown } from './html-to-markdown.ts';

describe('htmlToMarkdown — convert Graph-returned HTML (Office docs, OneNote, Outlook bodies) into clean markdown', () => {
  it('renders headings as ATX (# H1) not setext underline', () => {
    const md = htmlToMarkdown('<h1>Title</h1><h2>Subtitle</h2>');
    expect(md).toContain('# Title');
    expect(md).toContain('## Subtitle');
  });

  it('emits fenced code blocks rather than indented ones', () => {
    const md = htmlToMarkdown('<pre><code>const x = 1;</code></pre>');
    expect(md).toContain('```');
    expect(md).toContain('const x = 1;');
  });

  it('preserves inline links with their text', () => {
    const md = htmlToMarkdown('<p>See <a href="https://example.com/docs">our docs</a>.</p>');
    expect(md).toContain('[our docs](https://example.com/docs)');
  });

  it('preserves images as ![alt](src) with the data URI intact', () => {
    const md = htmlToMarkdown('<p><img src="data:image/png;base64,iVBORw0=" alt="diagram"></p>');
    expect(md).toContain('![diagram](data:image/png;base64,iVBORw0=)');
  });

  it('renders unordered lists with `-` markers', () => {
    const md = htmlToMarkdown('<ul><li>one</li><li>two</li></ul>');
    expect(md).toContain('-   one');
    expect(md).toContain('-   two');
  });

  it('strips MSO conditional comments that pollute Outlook HTML', () => {
    const html = '<p>before<!--[if !mso]> noisy MSO bracket <![endif]-->after</p>';
    const md = htmlToMarkdown(html);
    expect(md).toContain('before');
    expect(md).toContain('after');
    expect(md).not.toContain('mso');
    expect(md).not.toContain('endif');
  });

  it('strips ordinary HTML comments', () => {
    const md = htmlToMarkdown('<p>visible<!-- secret note --></p>');
    expect(md).toContain('visible');
    expect(md).not.toContain('secret');
  });

  it('drops <script> and <style> blocks (turndown defaults handle these)', () => {
    const md = htmlToMarkdown('<p>kept</p><script>alert(1)</script><style>p{color:red}</style>');
    expect(md).toContain('kept');
    expect(md).not.toContain('alert');
    expect(md).not.toContain('color:red');
  });
});
