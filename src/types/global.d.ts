declare namespace NodeJS {
  interface Timeout {
    _idleTimeout: number;
    _idlePrev: Timeout | null;
    _idleNext: Timeout | null;
    _idleStart: number;
    _onTimeout: () => void;
    _timerArgs: unknown[];
    _repeat: number | null;
  }
}

interface CSSStyleDeclaration {
  display: string;
  backgroundColor: string;
  borderColor: string;
  cssText: string;
  right: string;
  top: string;
  [key: string]: string | any;
}

interface Element {
  style: CSSStyleDeclaration;
  className: string;
  textContent: string | null;
  addEventListener(type: string, listener: (event: Event) => void): void;
  remove(): void;
  getBoundingClientRect(): DOMRect;
}

interface HTMLElement extends Element {
}

interface HTMLDivElement extends HTMLElement {
  align: string;
  addEventListener(type: string, listener: (event: Event) => void): void;
  focus(): void;
  querySelector(selectors: string): Element | null;
  tabIndex: number;
}

interface HTMLInputElement extends HTMLElement {
  value: string;
  type: string;
}

interface HTMLSelectElement extends HTMLElement {
  value: string;
  innerHTML: string;
}

interface HTMLTextAreaElement extends HTMLElement {
  value: string;
  rows: number;
}

interface Document {
  createElement(tagName: string): HTMLElement;
  getElementById(id: string): HTMLElement | null;
  body: HTMLElement;
  documentElement: HTMLElement;
}

declare var document: Document;

interface DOMRect {
  top: number;
  right: number;
  bottom: number;
  left: number;
  width: number;
  height: number;
}

interface Event {
  target: EventTarget | null;
}

interface EventTarget {
  value?: string;
}

interface Console {
  log(message?: any, ...optionalParams: any[]): void;
  error(message?: any, ...optionalParams: any[]): void;
}

declare var console: Console;

declare module "*.css" {
  const content: { [className: string]: string };
  export default content;
}

declare module "*.png" {
  const content: string;
  export default content;
}

declare module "*.jpg" {
  const content: string;
  export default content;
}

declare module "*.svg" {
  const content: string;
  export default content;
}