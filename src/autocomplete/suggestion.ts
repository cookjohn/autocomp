/* global document, Word, console, HTMLElement */

interface SuggestionOptions {
  position?: {
    top: number;
    left: number;
  };
  style?: {
    maxWidth: string;
    maxHeight?: string;
    fontSize: string;
  };
  displayMode?: "sidebar" | "inline";
}

export class SuggestionManager {
  private currentSuggestion: string | null = null;
  private suggestionDiv: HTMLElement | null = null;
  private defaultOptions: SuggestionOptions = {
    style: {
      maxWidth: "100%",
      fontSize: "14px",
    },
    displayMode: "sidebar",
  };

  constructor(options: Partial<SuggestionOptions> = {}) {
    this.defaultOptions = { ...this.defaultOptions, ...options };
    if (this.defaultOptions.displayMode === "sidebar") {
      this.initializeSuggestionUI();
    }
  }

  private initializeSuggestionUI(): void {
    if (this.suggestionDiv) {
      document.body.removeChild(this.suggestionDiv);
    }

    // 创建建议框DOM
    const div = document.createElement("div");
    div.id = "auto-complete-suggestion";
    div.style.cssText = `
      width: 100%;
      box-sizing: border-box;
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

    // 添加键盘事件监听
    div.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === "Tab") {
        event.preventDefault(); // 阻止默认行为
        void this.applySuggestion();
      }
    });

    // 添加到建议框容器
    const container = document.getElementById("suggestion-container");
    if (container) {
      container.appendChild(div);
      div.style.position = "static";
      div.style.margin = "0 0 20px 0";
    }
    this.suggestionDiv = div;
  }

  public async showSuggestion(suggestion: string): Promise<void> {
    if (this.defaultOptions.displayMode === "sidebar") {
      // 清除之前的建议
      await this.clearSuggestion();
      this.currentSuggestion = suggestion;

      // 侧边栏模式
      if (!this.suggestionDiv) return;

      const contentDiv = this.suggestionDiv.querySelector("#suggestion-content");
      if (contentDiv) {
        contentDiv.textContent = suggestion;
      }
      this.suggestionDiv.style.display = "block";
      this.suggestionDiv.focus();
    } else {
      // 内联模式
      try {
        await Word.run(async (context) => {
          // 清除之前的建议
          const selection = context.document.getSelection();
          const afterRange = selection.getRange(Word.RangeLocation.after);
          afterRange.load("text");
          await context.sync();

          if (afterRange.text === this.currentSuggestion) {
            afterRange.delete();
            await context.sync();
          }

          // 插入新建议
          const range = selection.insertText(suggestion, Word.InsertLocation.after);
          range.font.italic = true;
          range.font.color = "gray";
          await context.sync();
        });

        // 保存当前建议
        this.currentSuggestion = suggestion;
      } catch (error) {
        console.error("Failed to show inline suggestion:", error);
      }
    }
  }

  public async applySuggestion(): Promise<void> {
    const suggestion = this.currentSuggestion;
    if (!suggestion) return;

    try {
      if (this.defaultOptions.displayMode === "sidebar") {
        // 侧边栏模式：插入文本
        await Word.run(async (context) => {
          const selection = context.document.getSelection();
          selection.insertText(suggestion, Word.InsertLocation.end);
          await context.sync();
        });
      } else {
        // 内联模式：修改建议文本样式并移动光标
        await Word.run(async (context) => {
          const selection = context.document.getSelection();
          const afterRange = selection.getRange(Word.RangeLocation.after);
          afterRange.load("text");
          await context.sync();

          if (afterRange.text === suggestion) {
            // 修改样式
            afterRange.font.italic = false;
            afterRange.font.color = "black";

            // 移动光标到建议文本末尾
            const endRange = afterRange.getRange(Word.RangeLocation.end);
            endRange.select();
            await context.sync();
          }
        });
      }
    } catch (error) {
      console.error("Failed to apply suggestion:", error);
    } finally {
      await this.clearSuggestion();
    }
  }

  private async clearSuggestion(): Promise<void> {
    if (this.defaultOptions.displayMode === "sidebar") {
      if (this.suggestionDiv) {
        this.suggestionDiv.style.display = "none";
      }
    } else if (this.currentSuggestion) {
      try {
        await Word.run(async (context) => {
          const selection = context.document.getSelection();
          const range = selection.getRange(Word.RangeLocation.after);
          range.load("text");
          await context.sync();

          if (range.text === this.currentSuggestion) {
            range.delete();
            await context.sync();
          }
        });
      } catch (error) {
        console.error("Failed to clear inline suggestion:", error);
      }
    }
    this.currentSuggestion = null;
  }

  public hasSuggestion(): boolean {
    return this.currentSuggestion !== null;
  }

  public dispose(): void {
    if (this.suggestionDiv) {
      const container = document.getElementById("suggestion-container");
      if (container && container.contains(this.suggestionDiv)) {
        container.removeChild(this.suggestionDiv);
      }
      this.suggestionDiv = null;
    }
    void this.clearSuggestion();
  }

  public updateDisplayMode(mode: "sidebar" | "inline"): void {
    if (mode === this.defaultOptions.displayMode) return;

    void this.clearSuggestion();
    this.defaultOptions.displayMode = mode;

    if (mode === "sidebar") {
      this.initializeSuggestionUI();
    } else {
      if (this.suggestionDiv) {
        const container = document.getElementById("suggestion-container");
        if (container && container.contains(this.suggestionDiv)) {
          container.removeChild(this.suggestionDiv);
        }
        this.suggestionDiv = null;
      }
    }
  }
}
