/* global document, Office, Word, console, HTMLElement */

interface SuggestionOptions {
  position?: {
    top: number;
    left: number;
  };
  style?: {
    maxWidth: string;
    maxHeight: string;
    fontSize: string;
  };
}

export class SuggestionManager {
  private currentSuggestion: string | null = null;
  private suggestionDiv: HTMLElement | null = null;
  private defaultOptions: SuggestionOptions = {
    style: {
      maxWidth: "300px",
      maxHeight: "150px",
      fontSize: "14px",
    },
  };

  constructor(options: Partial<SuggestionOptions> = {}) {
    this.defaultOptions = { ...this.defaultOptions, ...options };
    this.initializeSuggestionUI();
  }

  private initializeSuggestionUI(): void {
    if (this.suggestionDiv) {
      document.body.removeChild(this.suggestionDiv);
    }

    // 创建建议框DOM
    const div = document.createElement("div");
    div.id = "auto-complete-suggestion";
    div.style.cssText = `
      position: fixed;
      top: 20px;
      right: 20px;
      background: white;
      border: 1px solid #e0e0e0;
      border-radius: 6px;
      padding: 16px;
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
      z-index: 1000;
      overflow-y: auto;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial;
      cursor: pointer;
      display: none;
      min-width: 200px;
      transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
      ${this.defaultOptions.style?.maxWidth ? `max-width: ${this.defaultOptions.style.maxWidth};` : ""}
      ${this.defaultOptions.style?.maxHeight ? `max-height: ${this.defaultOptions.style.maxHeight};` : ""}
      ${this.defaultOptions.style?.fontSize ? `font-size: ${this.defaultOptions.style.fontSize};` : ""}
    `;

    // 创建标题
    const title = document.createElement("div");
    title.style.cssText = `
      font-weight: 600;
      margin-bottom: 12px;
      padding-bottom: 12px;
      border-bottom: 1px solid #e0e0e0;
      color: #333;
      font-size: 15px;
    `;
    title.textContent = "自动补全建议";
    div.appendChild(title);

    // 创建内容容器
    const content = document.createElement("div");
    content.id = "suggestion-content";
    content.style.cssText = `
      white-space: pre-wrap;
      word-break: break-word;
      line-height: 1.5;
      color: #333;
    `;
    div.appendChild(content);

    // 添加点击事件
    div.addEventListener("click", () => {
      void this.applySuggestion();
    });

    // 添加焦点相关事件
    div.addEventListener("mouseenter", () => {
      div.style.backgroundColor = "#f8f9fa";
      div.style.borderColor = "#0078d4";
      div.style.boxShadow = "0 6px 16px rgba(0, 0, 0, 0.15)";
      div.style.transform = "translateY(-2px)";
    });

    div.addEventListener("mouseleave", () => {
      div.style.backgroundColor = "white";
      div.style.borderColor = "#e0e0e0";
      div.style.boxShadow = "0 4px 12px rgba(0, 0, 0, 0.1)";
      div.style.transform = "none";
    });

    // 设置tabIndex使div可以获取焦点
    div.tabIndex = 0;

    // 添加到文档
    document.body.appendChild(div);
    this.suggestionDiv = div;
  }

  public async showSuggestion(suggestion: string): Promise<void> {
    if (!this.suggestionDiv) return;

    // 设置新的建议内容
    this.currentSuggestion = suggestion;
    const contentDiv = this.suggestionDiv.querySelector("#suggestion-content");
    if (contentDiv) {
      contentDiv.textContent = suggestion;
    }
    this.suggestionDiv.style.display = "block";
    
    // 自动获取焦点
    this.suggestionDiv.focus();
  }

  public async applySuggestion(): Promise<void> {
    const suggestion = this.currentSuggestion;
    if (!suggestion) return;

    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.insertText(suggestion, Word.InsertLocation.end);
        await context.sync();
        this.hideSuggestion();
      });
    } catch (error) {
      console.error("应用补全建议失败:", error);
    }
  }

  private hideSuggestion(): void {
    if (this.suggestionDiv) {
      this.suggestionDiv.style.display = "none";
      this.currentSuggestion = null;
    }
  }

  public dispose(): void {
    if (this.suggestionDiv) {
      document.body.removeChild(this.suggestionDiv);
      this.suggestionDiv = null;
    }
    this.currentSuggestion = null;
  }
}