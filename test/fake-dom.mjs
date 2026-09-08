// A document small enough to write a fixture in, and no bigger: what
// collectRenderedMessages asks of the document it is given, and nothing else.
// The point of it is the selectors — a fixture written with the ids and
// data-tid attributes Teams renders stops matching the moment the collector
// looks for different ones, which is the failure that otherwise only shows up
// as a run that quietly reads no messages.
//
// It is a stand-in, not a browser: innerText is textContent, so it says nothing
// about the whitespace collapsing and hidden-element handling a real innerText
// does. Fixture text is therefore written the way a browser would already have
// rendered it, and what the two do differently stays a question for a live run.

// Only the selector shapes the collector uses: an optional tag name and at most
// one attribute, either present or equal to a value. Anything else throws
// rather than quietly matching nothing, so a selector this cannot answer for is
// a failing test and not a passing one.
const SELECTOR = /^([a-z]+)?(?:\[([\w-]+)(?:="([^"]*)")?\])?$/;

function matcher(selector) {
  const parsed = SELECTOR.exec(selector);
  if (!parsed || selector === '') throw new Error(`fake-dom cannot match the selector "${selector}".`);
  const [, tag, attr, value] = parsed;
  return node => (!tag || node.tag === tag)
    && (!attr || (node.attrs[attr] !== undefined && (value === undefined || node.attrs[attr] === value)));
}

class Element {
  constructor(tag, attrs, children) {
    this.tag = tag;
    this.attrs = attrs;
    this.children = children;
    this.parent = null;
    for (const child of children) if (typeof child !== 'string') child.parent = this;
  }

  getAttribute(name) {
    return this.attrs[name] ?? null;
  }

  get textContent() {
    return this.children
      .map(child => (typeof child === 'string' ? child : child.textContent))
      .join('');
  }

  get innerText() {
    return this.textContent;
  }

  querySelectorAll(selector) {
    return this.descendants().filter(matcher(selector));
  }

  querySelector(selector) {
    return this.querySelectorAll(selector)[0] ?? null;
  }

  closest(selector) {
    const matches = matcher(selector);
    for (let node = this; node; node = node.parent) if (matches(node)) return node;
    return null;
  }

  // Document order, so that "the first match" means the same here as in a page.
  descendants() {
    const found = [];
    for (const child of this.children) {
      if (typeof child === 'string') continue;
      found.push(child, ...child.descendants());
    }
    return found;
  }
}

export function el(tag, attrs = {}, children = []) {
  return new Element(tag, attrs, children);
}

// The root the collector is handed. getElementById lives here rather than on
// Element because that is where a page has it too.
export function fakeDocument(children) {
  const root = el('html', {}, children);
  root.getElementById = id => root.querySelectorAll('[id]').find(node => node.attrs.id === id) ?? null;
  return root;
}
