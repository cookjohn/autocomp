declare namespace Word {
  function run<T>(callback: (context: RequestContext) => Promise<T>): Promise<T>;

  interface ClientObject {
    context: RequestContext;
  }

  interface TrackedObjects {
    add(object: ClientObject): void;
    remove(object: ClientObject): void;
  }

  interface RequestContext {
    document: Document;
    sync(): Promise<void>;
    trackedObjects: TrackedObjects;
  }

  interface Document {
    body: Body;
    getSelection(): Range;
  }

  interface Body {
    text: string;
    load(properties: string): void;
  }

  enum SelectionMode {
    select = "Select",
    start = "Start",
    end = "End"
  }

  interface Range {
    text: string;
    paragraphs: ParagraphCollection;
    font: Font;
    load(properties: string): void;
    getRange(location: RangeLocation): Range;
    insertText(text: string, location: InsertLocation): Range;
    delete(): void;
    collapse(selectionMode: SelectionMode): void;
  }

  interface Font {
    color: string;
    italic: boolean;
  }

  interface ParagraphCollection {
    getFirst(): Paragraph;
  }

  interface Paragraph {
    text: string;
    load(properties: string): void;
    getRange(): Range;
  }

  enum RangeLocation {
    before = "Before",
    after = "After",
    start = "Start",
    end = "End",
    whole = "Whole",
  }

  enum InsertLocation {
    before = "Before",
    after = "After",
    start = "Start",
    end = "End",
    replace = "Replace",
  }
}

declare namespace Office {
  enum EventType {
    DocumentSelectionChanged = "DocumentSelectionChanged",
  }

  enum HostType {
    Word = "Word",
  }

  interface InitializationInfo {
    host: HostType;
    platform: string;
  }

  function onReady(callback: (info: InitializationInfo) => void): Promise<void>;

  const context: Context;

  interface Context {
    document: Document;
  }

  interface Document {
    addHandlerAsync(eventType: EventType, handler: () => void): Promise<void>;
    removeHandlerAsync(eventType: EventType, options: { handler: () => void }): Promise<void>;
  }
}