/* global Office, Word, console, clearTimeout, window */

import { debounce } from "../utils/debounce";
import { AutoCompleteContext } from "./context";
import { LLMService } from "./api";
import { SuggestionManager } from "./suggestion";

interface ProviderConfig {
  apiKey: string;
  model: string;
  endpoint?: string;
}

export interface AutoCompleteConfig {
  triggerMode: "auto" | "manual";
  contextRange: "paragraph" | "document" | "custom";
  maxContextLength: number;
  customParagraphs?: number; // 自定义段落数
  debounceMs: number;
  triggerDelayMs: number; // 触发延迟时间
  suggestionPosition: "sidebar" | "inline"; // 建议显示位置
  apiConfig: {
    provider: "openai" | "anthropic" | "openroute" | "gemini" | "custom" | "doubao" | "deepseek" | "vertex";
    maxTokens: number;
    temperature: number;
    systemPrompt?: string; // 自定义系统提示词
    providerConfigs: {
      openai: ProviderConfig;
      anthropic: ProviderConfig;
      openroute: ProviderConfig;
      gemini: ProviderConfig;
      doubao: ProviderConfig;
      deepseek: ProviderConfig;
      vertex: ProviderConfig;
      custom: ProviderConfig;
    };
  };
}

export class AutoCompleteEngine {
  private config: AutoCompleteConfig;
  private context: AutoCompleteContext;
  private llmService: LLMService;
  private suggestionManager: SuggestionManager;
  private isProcessing = false;
  private debouncedRequestCompletion: (context: string) => Promise<string | null>;
  private inputDebounceTimer: number | null = null; // 用于跟踪输入防抖定时器

  constructor(config: AutoCompleteConfig) {
    this.config = {
      ...config,
      customParagraphs: config.customParagraphs || 3, // 默认3段
      triggerDelayMs: config.triggerDelayMs || 2000, // 默认2秒
      suggestionPosition: config.suggestionPosition || "sidebar", // 默认侧边栏悬浮窗
      apiConfig: {
        ...config.apiConfig,
        systemPrompt:
          config.apiConfig.systemPrompt ||
          "你是一个专业文档助手，负责根据上下文生成连贯、自然的后续内容。生成内容应与前文风格一致，并保持逻辑连贯性。请直接输出续写内容，无需解释或添加其他信息。",
      },
    };

    this.context = new AutoCompleteContext({
      contextRange: config.contextRange,
      maxContextLength: config.maxContextLength,
      customParagraphs: config.customParagraphs,
    });

    this.llmService = new LLMService(this.config.apiConfig);
    this.suggestionManager = new SuggestionManager({
      displayMode: this.config.suggestionPosition,
    });

    // 防抖包装的补全请求
    this.debouncedRequestCompletion = debounce(
      (context: string) => this.llmService.complete(context),
      config.debounceMs
    );
  }

  /**
   * 初始化自动补全引擎
   */
  public async initialize(): Promise<void> {
    try {
      // 注册文档变更事件
      await Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        this.handleDocumentChange.bind(this)
      );
    } catch (error) {
      console.error("Failed to initialize auto-complete:", error);
    }
  }

  // 用于跟踪上一次内容
  private lastParagraphText: string | null = null;

  /**
   * 处理文档变更事件
   */
  private async handleDocumentChange(): Promise<void> {
    console.log("文档变更事件触发");
    try {
      // 检查是否是Tab键导致的变化
      let tabKeyHandled = false;
      if (this.suggestionManager.hasSuggestion()) {
        console.log("有建议存在，检查是否是Tab键");
        await Word.run(async (context) => {
          const selection = context.document.getSelection();
          selection.load("text");
          await context.sync();

          // 如果选区文本是制表符，删除制表符并应用建议
          if (selection.text === "\t") {
            console.log("检测到制表符，应用建议");

            // 删除制表符
            selection.delete();
            await context.sync();

            // 应用建议
            await this.suggestionManager.applySuggestion();
            tabKeyHandled = true;
          }
        });
      }

      // 如果处理了Tab键，不需要继续处理
      if (tabKeyHandled) {
        console.log("Tab键已处理，不检查内容变化");
        return;
      }

      // 自动模式下检查内容变化并延迟触发补全
      if (this.config.triggerMode === "auto" && !this.isProcessing) {
        console.log("自动模式且未处理中，检查内容变化");
        await Word.run(async (context) => {
          const selection = context.document.getSelection();
          const paragraph = selection.paragraphs.getFirst();
          paragraph.load("text");
          await context.sync();

          const paragraphText = paragraph.text;
          console.log("当前段落文本:", paragraphText);
          console.log("上一次段落文本:", this.lastParagraphText);

          // 检查内容是否变化
          let shouldTrigger = true;

          // 如果段落文本没有变化（内容相同）
          if (this.lastParagraphText === paragraphText) {
            console.log("内容未变化，不触发补全");
            shouldTrigger = false;
          } else {
            // 内容发生变化，触发请求
            console.log("内容已变化，触发补全");
          }

          // 更新上一次段落文本
          this.lastParagraphText = paragraphText;
          console.log("更新上一次段落文本:", this.lastParagraphText);

          // 如果需要触发补全，使用防抖方式延迟执行
          if (shouldTrigger) {
            console.log("需要触发补全，使用防抖延迟执行");

            // 清除之前的定时器（如果存在）
            if (this.inputDebounceTimer !== null) {
              console.log("清除之前的输入防抖定时器");
              clearTimeout(this.inputDebounceTimer);
              this.inputDebounceTimer = null;
            }

            // 设置新的定时器
            console.log(`设置新的输入防抖定时器，等待${this.config.triggerDelayMs}ms后执行`);
            this.inputDebounceTimer = window.setTimeout(() => {
              console.log("输入防抖定时器触发，开始补全");
              this.inputDebounceTimer = null;
              void this.triggerCompletion();
            }, this.config.triggerDelayMs);
          }
        });
      }
    } catch (error) {
      console.error("处理文档变更事件失败:", error);
    }
  }

  /**
   * 触发补全请求
   */
  public async triggerCompletion(): Promise<void> {
    console.log("触发补全方法被调用");
    if (this.isProcessing) {
      console.log("已经在处理中，不重复触发");
      return;
    }

    try {
      console.log("设置isProcessing = true");
      this.isProcessing = true;

      // 如果有建议存在，先清除
      if (this.suggestionManager.hasSuggestion()) {
        console.log("有建议存在，先清除");
        // 我们不直接调用clearSuggestion，因为showSuggestion会自动清除
      }

      await Word.run(async (context) => {
        console.log("开始获取上下文");
        // 获取当前上下文
        const contextContent = await this.context.getContext(context);
        console.log("获取上下文成功，长度:", contextContent.length);

        console.log("开始请求LLM补全");
        // 请求LLM补全
        const completion = await this.debouncedRequestCompletion(contextContent);
        console.log("LLM补全请求结果:", completion ? "成功" : "失败");

        // 显示补全建议
        if (completion) {
          console.log("开始显示补全建议");
          await this.suggestionManager.showSuggestion(completion);
          console.log("显示补全建议成功");
        }

        await context.sync();
      });
    } catch (error) {
      console.error("Completion request failed:", error);
    } finally {
      console.log("设置isProcessing = false");
      this.isProcessing = false;
    }
  }

  /**
   * 应用当前补全建议
   */
  public async applySuggestion(): Promise<void> {
    await this.suggestionManager.applySuggestion();
  }

  /**
   * 更新配置
   */
  public updateConfig(newConfig: Partial<AutoCompleteConfig>): void {
    // 合并配置
    this.config = {
      ...this.config,
      ...newConfig,
      apiConfig: {
        ...this.config.apiConfig,
        ...newConfig.apiConfig,
      },
    };

    // 确保自定义段落数有效
    if (this.config.contextRange === "custom") {
      this.config.customParagraphs = Math.max(1, Math.min(10, this.config.customParagraphs || 3));
    }

    // 更新各组件配置
    this.context = new AutoCompleteContext({
      contextRange: this.config.contextRange,
      maxContextLength: this.config.maxContextLength,
      customParagraphs: this.config.customParagraphs,
    });

    this.llmService = new LLMService({
      ...this.config.apiConfig,
      systemPrompt: this.config.apiConfig.systemPrompt,
    });

    // 更新建议显示位置
    this.suggestionManager.updateDisplayMode(this.config.suggestionPosition);

    // 更新防抖函数
    this.debouncedRequestCompletion = debounce(
      (context: string) => this.llmService.complete(context),
      this.config.debounceMs
    );
  }

  /**
   * 清理资源
   */
  public dispose(): void {
    // 移除事件监听
    void Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, {
      handler: this.handleDocumentChange.bind(this),
    });
    this.suggestionManager.dispose();
  }
}
