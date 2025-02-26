import { debounce } from "../utils/debounce";
import { AutoCompleteContext } from "./context";
import { LLMService } from "./api";
import { SuggestionManager } from "./suggestion";

export interface AutoCompleteConfig {
  triggerMode: "auto" | "manual";
  contextRange: "paragraph" | "document" | "custom";
  maxContextLength: number;
  customParagraphs?: number;  // 自定义段落数
  debounceMs: number;
  triggerDelayMs: number;  // 触发延迟时间
  suggestionPosition: "sidebar" | "inline";  // 建议显示位置
  apiConfig: {
    provider: "openai" | "anthropic" | "openroute" | "gemini" | "custom";
    apiKey: string;
    endpoint?: string;
    maxTokens: number;
    temperature: number;
    model?: string;
    systemPrompt?: string;  // 自定义系统提示词
  };
}

export class AutoCompleteEngine {
  private config: AutoCompleteConfig;
  private context: AutoCompleteContext;
  private llmService: LLMService;
  private suggestionManager: SuggestionManager;
  private isProcessing = false;
  private debouncedRequestCompletion: (context: string) => Promise<string | null>;

  constructor(config: AutoCompleteConfig) {
    this.config = {
      ...config,
      customParagraphs: config.customParagraphs || 3,  // 默认3段
      triggerDelayMs: config.triggerDelayMs || 2000,  // 默认2秒
      suggestionPosition: config.suggestionPosition || "sidebar",  // 默认侧边栏悬浮窗
      apiConfig: {
        ...config.apiConfig,
        systemPrompt: config.apiConfig.systemPrompt ||
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
      displayMode: this.config.suggestionPosition
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

  /**
   * 处理文档变更事件
   */
  private async handleDocumentChange(): Promise<void> {
    try {
      // 自动模式下延迟触发补全
      if (this.config.triggerMode === "auto" && !this.isProcessing) {
        setTimeout(() => {
          void this.triggerCompletion();
        }, this.config.triggerDelayMs);
      }
      
      // 检查是否是Tab键导致的变化
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();
        
        // 如果选区文本是制表符，且有建议显示，删除制表符并应用建议
        if (selection.text === "\t") {
          console.log("检测到制表符，可能是Tab键导致的变化");
          
          // 删除制表符
          selection.delete();
          await context.sync();
          
          // 应用建议
          await this.suggestionManager.applySuggestion();
        }
      });
    } catch (error) {
      console.error("处理文档变更事件失败:", error);
    }
  }

  /**
   * 触发补全请求
   */
  public async triggerCompletion(): Promise<void> {
    if (this.isProcessing) return;

    try {
      this.isProcessing = true;

      await Word.run(async (context) => {
        // 获取当前上下文
        const contextContent = await this.context.getContext(context);
        
        // 请求LLM补全
        const completion = await this.debouncedRequestCompletion(contextContent);
        
        // 显示补全建议
        if (completion) {
          await this.suggestionManager.showSuggestion(completion);
        }

        await context.sync();
      });
    } catch (error) {
      console.error("Completion request failed:", error);
    } finally {
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
    void Office.context.document.removeHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      {
        handler: this.handleDocumentChange.bind(this),
      }
    );
    this.suggestionManager.dispose();
  }
}