/* global Word, console */

interface ContextConfig {
  contextRange: "paragraph" | "document" | "custom" | "smart";
  maxContextLength: number;
  customParagraphs?: number;
  smartOptions?: {
    includeTitles?: boolean;
    includeStructure?: boolean;
    weightCurrentParagraph?: number;
  };
}

// 用于缓存上下文的接口
interface ContextCache {
  content: string;
  timestamp: number;
  paragraphText?: string;
}

export class AutoCompleteContext {
  private config: ContextConfig;
  private cache: ContextCache | null = null;
  private cacheTTL = 5000; // 缓存有效期，默认5秒

  constructor(config: ContextConfig) {
    this.config = {
      ...config,
      customParagraphs: config.customParagraphs || 3,
      smartOptions: {
        includeTitles: config.smartOptions?.includeTitles ?? true,
        includeStructure: config.smartOptions?.includeStructure ?? true,
        weightCurrentParagraph: config.smartOptions?.weightCurrentParagraph ?? 2.0,
      },
    };
  }

  /**
   * 获取当前上下文
   */
  public async getContext(context: Word.RequestContext): Promise<string> {
    console.log("getContext方法被调用");
    try {
      // 检查缓存是否有效
      console.log("开始检查缓存是否有效");
      const cacheValid = await this.isCacheValid(context);
      console.log("缓存是否有效:", cacheValid);

      if (cacheValid) {
        console.log("使用缓存的上下文");
        return this.cache!.content;
      }

      console.log("缓存无效，获取新的上下文");
      let contextContent: string;
      switch (this.config.contextRange) {
        case "paragraph":
          console.log("获取当前段落上下文");
          contextContent = await this.getCurrentParagraph(context);
          break;
        case "document":
          console.log("获取整个文档上下文");
          contextContent = await this.getDocumentContext(context);
          break;
        case "custom":
          console.log("获取自定义段落上下文");
          contextContent = await this.getCustomParagraphs(context);
          break;
        case "smart":
          console.log("获取智能上下文");
          contextContent = await this.getSmartContext(context);
          break;
        default:
          console.log("使用默认的当前段落上下文");
          contextContent = await this.getCurrentParagraph(context);
      }
      console.log("获取上下文成功，长度:", contextContent.length);

      // 更新缓存
      console.log("开始更新缓存");
      await this.updateCache(context, contextContent);
      console.log("缓存更新成功");

      return contextContent;
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
   * 获取指定数量段落的内容，并考虑光标位置
   */
  private async getCustomParagraphs(context: Word.RequestContext): Promise<string> {
    const selection = context.document.getSelection();
    const currentParagraph = selection.paragraphs.getFirst();

    // 获取当前段落的范围
    const range = currentParagraph.getRange();
    range.load("text");
    await context.sync();

    // 估计光标在当前段落中的相对位置
    // 由于无法直接获取光标位置，我们假设它在段落的中间
    const currentText = range.text;
    const relativePosition = Math.floor(currentText.length / 2);

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

    // 分割当前段落为估计的光标前后两部分
    const beforeCursor = currentText.substring(0, relativePosition);
    const afterCursor = currentText.substring(relativePosition);

    // 组合上下文，确保光标位置周围的内容被优先保留
    const fullContext = beforeParagraphs + beforeCursor + afterCursor + afterParagraphs;
    return this.intelligentTruncate(fullContext, beforeParagraphs.length + beforeCursor.length);
  }

  /**
   * 获取智能上下文，考虑文档结构和语义相关性
   */
  private async getSmartContext(context: Word.RequestContext): Promise<string> {
    // 获取基本上下文
    const customContext = await this.getCustomParagraphs(context);

    // 如果不需要额外的结构信息，直接返回自定义段落上下文
    if (!this.config.smartOptions?.includeTitles && !this.config.smartOptions?.includeStructure) {
      return customContext;
    }

    let structureInfo = "";

    // 获取文档标题和结构信息
    if (this.config.smartOptions?.includeTitles || this.config.smartOptions?.includeStructure) {
      const body = context.document.body;
      // 注意：这里假设body有paragraphs属性，实际可能需要调整
      body.load("paragraphs");
      await context.sync();

      // 这里假设paragraphs有items属性，实际可能需要调整
      const paragraphs = body.paragraphs;
      paragraphs.load("text");
      await context.sync();

      // 提取标题和结构信息
      const titles: string[] = [];
      const structure: string[] = [];

      // 这里假设paragraphs.items是一个数组，实际可能需要调整
      for (let i = 0; i < paragraphs.items.length; i++) {
        const paragraph = paragraphs.items[i];
        // 由于无法直接获取段落样式，我们通过文本特征来猜测是否是标题
        const text = paragraph.text.trim();

        // 简单的启发式方法：短且以冒号结尾的可能是标题，或者全大写的短文本
        const isLikelyTitle =
          (text.length < 50 && text.endsWith(":")) ||
          (text.length < 30 && text === text.toUpperCase()) ||
          (text.length < 100 && !text.includes("."));

        if (isLikelyTitle) {
          titles.push(text);

          if (this.config.smartOptions?.includeStructure) {
            structure.push(`Heading: ${text}`);
          }
        } else if (this.config.smartOptions?.includeStructure && text.length > 0) {
          // 添加普通段落的简要信息到结构中
          structure.push(`Paragraph: ${text.substring(0, 50)}${text.length > 50 ? "..." : ""}`);
        }
      }

      // 构建结构信息字符串
      if (this.config.smartOptions?.includeTitles && titles.length > 0) {
        structureInfo += "Document Titles:\n" + titles.join("\n") + "\n\n";
      }

      if (this.config.smartOptions?.includeStructure && structure.length > 0) {
        structureInfo += "Document Structure:\n" + structure.join("\n") + "\n\n";
      }
    }

    // 组合结构信息和自定义段落上下文
    const combinedContext = structureInfo + "Current Content:\n" + customContext;

    // 确保不超过最大上下文长度
    return this.truncateContext(combinedContext);
  }

  /**
   * 智能截断上下文，保持句子和段落的完整性
   */
  private intelligentTruncate(text: string, cursorPosition: number): string {
    if (text.length <= this.config.maxContextLength) {
      return text;
    }

    // 计算应该保留的前后文本长度
    const totalLength = this.config.maxContextLength;
    const halfLength = Math.floor(totalLength / 2);

    // 从光标位置向前后扩展
    let startPos = Math.max(0, cursorPosition - halfLength);
    let endPos = Math.min(text.length, cursorPosition + halfLength);

    // 调整以确保不超过总长度
    if (endPos - startPos > totalLength) {
      if (startPos === 0) {
        endPos = totalLength;
      } else if (endPos === text.length) {
        startPos = text.length - totalLength;
      }
    }

    // 尝试调整到句子边界
    startPos = this.findSentenceBoundary(text, startPos, false);
    endPos = this.findSentenceBoundary(text, endPos, true);

    // 如果调整后长度超过限制，进行二次调整
    if (endPos - startPos > totalLength) {
      // 优先保留光标后的内容，因为这通常是用户正在编辑的部分
      const excessLength = endPos - startPos - totalLength;
      startPos += excessLength;

      // 再次尝试调整到句子边界
      startPos = this.findSentenceBoundary(text, startPos, false);
    }

    return text.substring(startPos, endPos);
  }

  /**
   * 查找最近的句子边界
   */
  private findSentenceBoundary(text: string, position: number, forward: boolean): number {
    const sentenceEndMarkers = [".", "!", "?", "\n"];
    const maxLookup = 100; // 最大查找范围

    if (forward) {
      // 向前查找句子结束
      for (let i = 0; i < maxLookup; i++) {
        const pos = position + i;
        if (pos >= text.length) return text.length;

        if (sentenceEndMarkers.includes(text[pos])) {
          return pos + 1; // 包含句子结束符
        }
      }
      return Math.min(position + maxLookup, text.length);
    } else {
      // 向后查找句子开始
      for (let i = 0; i < maxLookup; i++) {
        const pos = position - i;
        if (pos <= 0) return 0;

        if (sentenceEndMarkers.includes(text[pos - 1])) {
          return pos; // 从句子开始处开始
        }
      }
      return Math.max(position - maxLookup, 0);
    }
  }

  /**
   * 截断上下文到指定长度，保持语义完整性
   */
  private truncateContext(text: string): string {
    if (text.length <= this.config.maxContextLength) {
      return text;
    }

    // 如果文本太长，优先保留后半部分，但尝试在句子边界处截断
    const startPos = text.length - this.config.maxContextLength;
    const adjustedStartPos = this.findSentenceBoundary(text, startPos, false);

    return text.substring(adjustedStartPos);
  }

  /**
   * 检查缓存是否有效
   */
  private async isCacheValid(context: Word.RequestContext): Promise<boolean> {
    console.log("isCacheValid方法被调用");
    if (!this.cache) {
      console.log("缓存不存在");
      return false;
    }

    // 检查缓存是否过期
    const now = Date.now();
    const cacheAge = now - this.cache.timestamp;
    console.log("缓存年龄:", cacheAge, "ms, TTL:", this.cacheTTL, "ms");
    if (cacheAge > this.cacheTTL) {
      console.log("缓存已过期");
      return false;
    }

    try {
      // 检查当前段落是否改变
      const selection = context.document.getSelection();
      const paragraph = selection.paragraphs.getFirst();
      paragraph.load("text");
      await context.sync();

      console.log("当前段落文本:", paragraph.text);
      console.log("缓存的段落文本:", this.cache.paragraphText);

      // 只检查段落文本是否改变
      if (paragraph.text !== this.cache.paragraphText) {
        console.log("段落文本已改变，缓存无效");
        return false;
      }

      console.log("段落文本未改变，缓存有效");
      return true;
    } catch (error) {
      console.error("缓存验证失败:", error);
      return false;
    }
  }

  /**
   * 更新上下文缓存
   */
  private async updateCache(context: Word.RequestContext, content: string): Promise<void> {
    console.log("updateCache方法被调用");
    try {
      const selection = context.document.getSelection();
      const paragraph = selection.paragraphs.getFirst();
      paragraph.load("text");
      await context.sync();

      const timestamp = Date.now();
      console.log("更新缓存:", {
        contentLength: content.length,
        timestamp: timestamp,
        paragraphText: paragraph.text,
      });

      this.cache = {
        content,
        timestamp: timestamp,
        paragraphText: paragraph.text,
      };
      console.log("缓存更新成功");
    } catch (error) {
      console.error("更新缓存失败:", error);
    }
  }

  /**
   * 清除缓存
   */
  public clearCache(): void {
    this.cache = null;
  }

  /**
   * 设置缓存TTL（毫秒）
   */
  public setCacheTTL(ttl: number): void {
    this.cacheTTL = ttl;
  }
}
