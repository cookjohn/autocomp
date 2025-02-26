/* global Word, Office */

interface ContextConfig {
  contextRange: "paragraph" | "document" | "custom";
  maxContextLength: number;
  customParagraphs?: number;
}

export class AutoCompleteContext {
  private config: ContextConfig;

  constructor(config: ContextConfig) {
    this.config = {
      ...config,
      customParagraphs: config.customParagraphs || 3,
    };
  }

  /**
   * 获取当前上下文
   */
  public async getContext(context: Word.RequestContext): Promise<string> {
    try {
      switch (this.config.contextRange) {
        case "paragraph":
          return await this.getCurrentParagraph(context);
        case "document":
          return await this.getDocumentContext(context);
        case "custom":
          return await this.getCustomParagraphs(context);
        default:
          return await this.getCurrentParagraph(context);
      }
    } catch (error) {
      console.error("Failed to get context:", error);
      return "";
    }
  }

  /**
   * 获取当前段落内容
   */
  private async getCurrentParagraph(context: Word.RequestContext): Promise<string> {
    const selection = context.document.getSelection();
    const paragraph = selection.paragraphs.getFirst();
    paragraph.load("text");
    await context.sync();

    return this.truncateContext(paragraph.text);
  }

  /**
   * 获取整个文档内容
   */
  private async getDocumentContext(context: Word.RequestContext): Promise<string> {
    const body = context.document.body;
    body.load("text");
    await context.sync();

    return this.truncateContext(body.text);
  }

  /**
   * 获取指定数量段落的内容
   */
  private async getCustomParagraphs(context: Word.RequestContext): Promise<string> {
    const selection = context.document.getSelection();
    const currentParagraph = selection.paragraphs.getFirst();
    
    // 获取当前段落的范围
    const range = currentParagraph.getRange();
    range.load("text");
    await context.sync();

    // 向前获取段落
    let beforeParagraphs = "";
    let beforeRange = range;
    for (let i = 0; i < Math.floor(this.config.customParagraphs / 2); i++) {
      const prevRange = beforeRange.getRange(Word.RangeLocation.before);
      prevRange.load("text");
      await context.sync();
      beforeParagraphs = prevRange.text + beforeParagraphs;
      beforeRange = prevRange;
    }

    // 向后获取段落
    let afterParagraphs = "";
    let afterRange = range;
    for (let i = 0; i < Math.floor(this.config.customParagraphs / 2); i++) {
      const nextRange = afterRange.getRange(Word.RangeLocation.after);
      nextRange.load("text");
      await context.sync();
      afterParagraphs += nextRange.text;
      afterRange = nextRange;
    }

    // 组合上下文
    const fullContext = beforeParagraphs + range.text + afterParagraphs;
    return this.truncateContext(fullContext);
  }

  /**
   * 截断上下文到指定长度
   */
  private truncateContext(text: string): string {
    if (text.length <= this.config.maxContextLength) {
      return text;
    }

    // 保留后半部分，因为这是最接近当前编辑位置的内容
    return text.slice(-this.config.maxContextLength);
  }
}