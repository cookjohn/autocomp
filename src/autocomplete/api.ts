/* global console, XMLHttpRequest */

import { showMessage } from "../taskpane/taskpane";

type RequestMode = "cors" | "no-cors" | "same-origin" | "navigate";

interface RequestInit {
  method?: string;
  headers?: Record<string, string>;
  body?: string | undefined;
  mode?: RequestMode;
}

interface Response {
  ok: boolean;
  status: number;
  statusText: string;
  json(): Promise<any>;
  text(): Promise<string>;
}

// 自定义fetch函数，使用XMLHttpRequest实现
async function customFetch(url: string, init?: RequestInit): Promise<Response> {
  // 否则使用XMLHttpRequest
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.open(init?.method || "GET", url);

    // 设置请求头
    if (init?.headers) {
      Object.entries(init.headers).forEach(([key, value]) => {
        xhr.setRequestHeader(key, value);
      });
    }

    xhr.onload = () => {
      const response = {
        ok: xhr.status >= 200 && xhr.status < 300,
        status: xhr.status,
        statusText: xhr.statusText,
        json: async () => JSON.parse(xhr.responseText),
        text: async () => xhr.responseText,
      };
      resolve(response as Response);
    };

    xhr.onerror = () => {
      reject(new Error("Network request failed"));
    };

    xhr.send(init?.body);
  });
}

interface ProviderConfig {
  apiKey: string;
  model: string;
  endpoint?: string;
}

interface LLMConfig {
  provider: "openai" | "anthropic" | "openroute" | "gemini" | "custom" | "doubao" | "deepseek" | "vertex";
  maxTokens: number;
  temperature: number;
  systemPrompt?: string;
  providerConfigs: {
    [key in
      | "openai"
      | "anthropic"
      | "openroute"
      | "gemini"
      | "custom"
      | "doubao"
      | "deepseek"
      | "vertex"]: ProviderConfig;
  };
}

interface LLMModel {
  id: string;
  name: string;
  description: string;
  context_length: number;
  pricing?: {
    prompt: string;
    completion: string;
  };
}

export class LLMService {
  private config: LLMConfig;
  private defaultPrompt =
    "You are a professional document assistant. Please continue the text based on the context, maintaining consistency in style and logic. Provide only the continuation without explanations.";

  constructor(config: LLMConfig) {
    this.config = {
      ...config,
      systemPrompt: config.systemPrompt || this.defaultPrompt,
    };
  }

  /**
   * 获取可用模型列表
   */
  public async getAvailableModels(): Promise<LLMModel[]> {
    try {
      switch (this.config.provider) {
        case "openai":
          return await this.getOpenAIModels();
        case "anthropic":
          return await this.getAnthropicModels();
        case "openroute":
          return await this.getOpenRouteModels();
        case "gemini":
          return await this.getGeminiModels();
        case "doubao":
          return await this.getDoubaoModels();
        case "deepseek":
          return await this.getDeepseekModels();
        case "vertex":
          return await this.getVertexModels();
        default:
          throw new Error(`不支持的模型提供商: ${this.config.provider}`);
      }
    } catch (error: any) {
      console.error("获取模型列表失败:", error);
      let errorMessage = "获取模型列表失败";

      // 处理常见的错误类型
      if (error.message.includes("API key")) {
        errorMessage = `API密钥无效或未提供 (${this.config.provider})`;
      } else if (error.message.includes("network") || error.message.includes("Failed to fetch")) {
        errorMessage = `网络连接失败，请检查网络状态或代理设置 (${this.config.provider})`;
      } else if (error.message.includes("timeout")) {
        errorMessage = `请求超时，请稍后重试 (${this.config.provider})`;
      } else if (error.message.includes("rate limit") || error.message.includes("429")) {
        errorMessage = `已达到API请求限制，请稍后重试 (${this.config.provider})`;
      } else if (error.message.includes("API error")) {
        const statusMatch = error.message.match(/(\d{3})/);
        const status = statusMatch ? statusMatch[1] : "";
        switch (status) {
          case "401":
            errorMessage = `认证失败，请检查API密钥是否正确 (${this.config.provider})`;
            break;
          case "403":
            errorMessage = `没有访问权限，请检查API密钥权限 (${this.config.provider})`;
            break;
          case "404":
            errorMessage = `请求的资源不存在 (${this.config.provider})`;
            break;
          case "500":
          case "502":
          case "503":
            errorMessage = `服务器错误，请稍后重试 (${this.config.provider})`;
            break;
          default:
            errorMessage = `${error.message} (${this.config.provider})`;
        }
      } else {
        errorMessage = `${error.message} (${this.config.provider})`;
      }

      // 显示错误信息
      showMessage(errorMessage, "error");
      throw new Error(errorMessage);
    }
  }

  private async getOpenAIModels(): Promise<LLMModel[]> {
    const openaiConfig = this.config.providerConfigs.openai;
    const response = await customFetch("https://api.openai.com/v1/models", {
      method: "GET",
      headers: {
        Authorization: `Bearer ${openaiConfig.apiKey}`,
      },
    });

    if (!response.ok) {
      throw new Error(`OpenAI API error: ${response.statusText}`);
    }

    const data = await response.json();
    return data.data
      .filter((model: any) => model.id.startsWith("gpt-") && !model.id.includes("instruct"))
      .map((model: any) => ({
        id: model.id,
        name: model.id,
        description: "OpenAI GPT Model",
        context_length: model.id.includes("32k") ? 32768 : model.id.includes("16k") ? 16384 : 4096,
      }))
      .sort((a: LLMModel, b: LLMModel) => b.context_length - a.context_length);
  }

  private async getAnthropicModels(): Promise<LLMModel[]> {
    const anthropicConfig = this.config.providerConfigs.anthropic;
    const response = await customFetch("https://api.anthropic.com/v1/models", {
      method: "GET",
      headers: {
        "x-api-key": anthropicConfig.apiKey,
      },
    });

    if (!response.ok) {
      throw new Error(`Anthropic API error: ${response.statusText}`);
    }

    const data = await response.json();
    return data.models.map((model: any) => ({
      id: model.id,
      name: model.name || model.id,
      description: "Anthropic Claude Model",
      context_length: model.context_window || 100000,
    }));
  }

  private async getOpenRouteModels(): Promise<LLMModel[]> {
    const openrouteConfig = this.config.providerConfigs.openroute;
    const response = await customFetch("https://openrouter.ai/api/v1/models", {
      method: "GET",
      headers: {
        Authorization: `Bearer ${openrouteConfig.apiKey}`,
        "HTTP-Referer": "https://github.com/cookjohn/autocomp",
        "X-Title": "Word LLM AutoComplete",
      },
    });

    if (!response.ok) {
      throw new Error(`OpenRoute API error: ${response.statusText}`);
    }

    const data = await response.json();
    return data.data
      .map((model: any) => ({
        id: model.id,
        name: `${model.name} (${model.context_length}tokens)`,
        description: model.description || "",
        context_length: model.context_length,
        pricing: {
          prompt: model.pricing?.prompt || "Unknown",
          completion: model.pricing?.completion || "Unknown",
        },
      }))
      .sort((a: LLMModel, b: LLMModel) => b.context_length - a.context_length);
  }

  private async getGeminiModels(): Promise<LLMModel[]> {
    const geminiConfig = this.config.providerConfigs.gemini;
    const response = await customFetch(
      `https://generativelanguage.googleapis.com/v1/models?key=${geminiConfig.apiKey}`,
      {
        method: "GET",
      }
    );

    if (!response.ok) {
      throw new Error(`Gemini API error: ${response.statusText}`);
    }

    const data = await response.json();
    return data.models
      .filter((model: any) => model.name.includes("gemini"))
      .map((model: any) => ({
        id: model.name,
        name: model.displayName || model.name,
        description: model.description || "Google Gemini Model",
        context_length: model.name.includes("pro") ? 32768 : 16384,
      }))
      .sort((a: LLMModel, b: LLMModel) => b.context_length - a.context_length);
  }

  public async complete(context: string): Promise<string | null> {
    try {
      switch (this.config.provider) {
        case "openai":
          return await this.completeWithOpenAI(context);
        case "anthropic":
          return await this.completeWithAnthropic(context);
        case "openroute":
          return await this.completeWithOpenRoute(context);
        case "gemini":
          return await this.completeWithGemini(context);
        case "doubao":
          return await this.completeWithDoubao(context);
        case "deepseek":
          return await this.completeWithDeepseek(context);
        case "custom":
          return await this.completeWithCustomAPI(context);
        default:
          throw new Error(`Unsupported LLM provider: ${this.config.provider}`);
      }
    } catch (error) {
      console.error("LLM API request failed:", error);
      return null;
    }
  }

  private async completeWithOpenAI(context: string): Promise<string> {
    const openaiConfig = this.config.providerConfigs.openai;
    if (!openaiConfig.model) {
      throw new Error("OpenAI requires a model selection");
    }

    const response = await customFetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${openaiConfig.apiKey}`,
      },
      body: JSON.stringify({
        model: openaiConfig.model,
        messages: [
          {
            role: "system",
            content: this.config.systemPrompt,
          },
          {
            role: "user",
            content: `Current content:\n${context}\n\nContinue:`,
          },
        ],
        max_tokens: this.config.maxTokens,
        temperature: this.config.temperature,
        top_p: 0.95,
        frequency_penalty: 0.5,
        presence_penalty: 0.5,
        stream: false,
      }),
    });

    if (!response.ok) {
      throw new Error(`OpenAI API error: ${response.statusText}`);
    }

    const data = await response.json();
    return data.choices[0].message.content.trim();
  }

  private async completeWithOpenRoute(context: string): Promise<string> {
    const openrouteConfig = this.config.providerConfigs.openroute;
    if (!openrouteConfig.model) {
      throw new Error("OpenRoute requires a model selection");
    }

    const response = await customFetch("https://openrouter.ai/api/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${openrouteConfig.apiKey}`,
        "HTTP-Referer": "https://github.com/cookjohn/autocomp",
        "X-Title": "Word LLM AutoComplete",
      },
      body: JSON.stringify({
        model: openrouteConfig.model,
        messages: [
          {
            role: "system",
            content: this.config.systemPrompt,
          },
          {
            role: "user",
            content: `Current content:\n${context}\n\nContinue:`,
          },
        ],
        max_tokens: this.config.maxTokens,
        temperature: this.config.temperature,
        top_p: 0.95,
        frequency_penalty: 0.5,
        presence_penalty: 0.5,
      }),
    });

    if (!response.ok) {
      throw new Error(`OpenRoute API error: ${response.statusText}`);
    }

    const data = await response.json();
    return data.choices[0].message.content.trim();
  }

  private async completeWithGemini(context: string): Promise<string> {
    const geminiConfig = this.config.providerConfigs.gemini;
    if (!geminiConfig.model) {
      throw new Error("Gemini requires a model selection");
    }

    const url = `https://generativelanguage.googleapis.com/v1/${geminiConfig.model}:generateContent?key=${geminiConfig.apiKey}`;
    const response = await customFetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        contents: [
          {
            role: "user",
            parts: [{ text: `${this.config.systemPrompt}\n\nCurrent content:\n${context}\n\nContinue:` }],
          },
        ],
        generationConfig: {
          temperature: this.config.temperature,
          maxOutputTokens: this.config.maxTokens,
          topP: 0.95,
        },
        safetySettings: [
          { category: "HARM_CATEGORY_HARASSMENT", threshold: "BLOCK_NONE" },
          { category: "HARM_CATEGORY_HATE_SPEECH", threshold: "BLOCK_NONE" },
          { category: "HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold: "BLOCK_NONE" },
          { category: "HARM_CATEGORY_DANGEROUS_CONTENT", threshold: "BLOCK_NONE" },
        ],
      }),
    });

    if (!response.ok) {
      throw new Error(`Gemini API error: ${response.statusText}`);
    }

    const data = await response.json();
    return data.candidates[0].content.parts[0].text.trim();
  }

  private async completeWithDoubao(context: string): Promise<string> {
    const doubaoConfig = this.config.providerConfigs.doubao;
    if (!doubaoConfig.model) {
      throw new Error("Doubao requires a model selection");
    }
    console.log(doubaoConfig.apiKey);
    const response = await customFetch("https://ark.cn-beijing.volces.com/api/v3/chat/completions", {
      method: "POST",
      mode: "no-cors",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${doubaoConfig.apiKey}`,
      },
      body: JSON.stringify({
        model: doubaoConfig.model,
        messages: [
          {
            role: "system",
            content: this.config.systemPrompt,
          },
          {
            role: "user",
            content: `Current content:\n${context}\n\nContinue:`,
          },
        ],
      }),
    });
    if (!response.ok) {
      throw new Error(`Doubao API error: ${response.statusText}`);
    }

    const data = await response.json();
    return data.choices[0].message.content.trim();
  }

  private async completeWithDeepseek(context: string): Promise<string> {
    const deepseekConfig = this.config.providerConfigs.deepseek;
    if (!deepseekConfig.model) {
      throw new Error("Deepseek requires a model selection");
    }

    const response = await customFetch("https://api.deepseek.com/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${deepseekConfig.apiKey}`,
      },
      body: JSON.stringify({
        model: deepseekConfig.model,
        messages: [
          {
            role: "system",
            content: this.config.systemPrompt,
          },
          {
            role: "user",
            content: `Current content:\n${context}\n\nContinue:`,
          },
        ],
        max_tokens: this.config.maxTokens,
        temperature: this.config.temperature,
        top_p: 0.95,
        frequency_penalty: 0.5,
        presence_penalty: 0.5,
      }),
    });

    if (!response.ok) {
      throw new Error(`Deepseek API error: ${response.statusText}`);
    }

    const data = await response.json();
    return data.choices[0].message.content.trim();
  }

  private async getDoubaoModels(): Promise<LLMModel[]> {
    // 豆包提供固定的模型列表
    const models = [
      {
        id: "doubao-1-5-pro-32k-250115",
        name: "豆包Pro 32K",
        description: "Doubao Pro Model (32K)",
        context_length: 32768,
      },
      {
        id: "doubao-1.5-pro-256k-250115",
        name: "豆包Pro 256K",
        description: "Doubao Pro Model (256K)",
        context_length: 262144,
      },
      {
        id: "doubao-1.5-lite-32k-250115",
        name: "豆包Lite 32K",
        description: "Doubao Lite Model (32K)",
        context_length: 32768,
      },
      {
        id: "deepseek-r1-250120",
        name: "Deepseek R1",
        description: "Deepseek R1 Model",
        context_length: 32768,
      },
      {
        id: "deepseek-r1-distill-qwen-32b-250120",
        name: "Deepseek R1 Qwen 32B",
        description: "Deepseek R1 Qwen 32B Model",
        context_length: 32768,
      },
      {
        id: "deepseek-r1-distill-qwen-7b-250120",
        name: "Deepseek R1 Qwen 7B",
        description: "Deepseek R1 Qwen 7B Model",
        context_length: 32768,
      },
      {
        id: "deepseek-v3-241226",
        name: "Deepseek V3",
        description: "Deepseek V3 Model",
        context_length: 32768,
      },
    ];
    return models;
  }

  private async getDeepseekModels(): Promise<LLMModel[]> {
    const deepseekConfig = this.config.providerConfigs.deepseek;
    const response = await customFetch("https://api.deepseek.com/models", {
      method: "GET",
      headers: {
        Authorization: `Bearer ${deepseekConfig.apiKey}`,
      },
    });

    if (!response.ok) {
      throw new Error(`Deepseek API error: ${response.statusText}`);
    }

    const data = await response.json();
    return data.data
      .filter((model: any) => model.object === "model")
      .map((model: any) => ({
        id: model.id,
        name: model.id,
        description: `Deepseek AI Model (${model.owned_by})`,
        context_length: 32768, // 默认上下文长度
      }))
      .sort((a: LLMModel, b: LLMModel) => a.name.localeCompare(b.name));
  }

  private async getVertexModels(): Promise<LLMModel[]> {
    const vertexConfig = this.config.providerConfigs.vertex;
    const response = await customFetch("https://us-central1-aiplatform.googleapis.com/v1/models", {
      method: "GET",
      headers: {
        Authorization: `Bearer ${vertexConfig.apiKey}`,
      },
    });

    if (!response.ok) {
      throw new Error(`Vertex API error: ${response.statusText}`);
    }

    const data = await response.json();
    return data.models
      .filter((model: any) => model.name.includes("text-"))
      .map((model: any) => ({
        id: model.name,
        name: model.displayName || model.name,
        description: model.description || "Google Vertex AI Model",
        context_length: 32768,
      }))
      .sort((a: LLMModel, b: LLMModel) => b.context_length - a.context_length);
  }

  private async completeWithVertex(context: string): Promise<string> {
    const vertexConfig = this.config.providerConfigs.vertex;
    if (!vertexConfig.model) {
      throw new Error("Vertex requires a model selection");
    }

    const response = await customFetch(
      `https://us-central1-aiplatform.googleapis.com/v1/${vertexConfig.model}:predict`,
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${vertexConfig.apiKey}`,
        },
        body: JSON.stringify({
          instances: [
            {
              prompt: `${this.config.systemPrompt}\n\nCurrent content:\n${context}\n\nContinue:`,
            },
          ],
          parameters: {
            temperature: this.config.temperature,
            maxOutputTokens: this.config.maxTokens,
            topP: 0.95,
          },
        }),
      }
    );

    if (!response.ok) {
      throw new Error(`Vertex API error: ${response.statusText}`);
    }

    const data = await response.json();
    return data.predictions[0].content.trim();
  }

  private async completeWithAnthropic(context: string): Promise<string> {
    const anthropicConfig = this.config.providerConfigs.anthropic;
    if (!anthropicConfig.model) {
      throw new Error("Anthropic requires a model selection");
    }

    const response = await customFetch("https://api.anthropic.com/v1/complete", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "X-API-Key": anthropicConfig.apiKey,
      },
      body: JSON.stringify({
        model: anthropicConfig.model,
        prompt: `${this.config.systemPrompt}\n\nCurrent content:\n${context}\n\nContinue:`,
        max_tokens_to_sample: this.config.maxTokens,
        temperature: this.config.temperature,
        top_p: 0.95,
      }),
    });

    if (!response.ok) {
      throw new Error(`Anthropic API error: ${response.statusText}`);
    }

    const data = await response.json();
    return data.completion.trim();
  }

  private async completeWithCustomAPI(context: string): Promise<string> {
    const customConfig = this.config.providerConfigs.custom;
    if (!customConfig.endpoint) {
      throw new Error("Custom API requires an endpoint");
    }

    const response = await customFetch(customConfig.endpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${customConfig.apiKey}`,
      },
      body: JSON.stringify({
        prompt: `${this.config.systemPrompt}\n\nCurrent content:\n${context}\n\nContinue:`,
        max_tokens: this.config.maxTokens,
        temperature: this.config.temperature,
      }),
    });

    if (!response.ok) {
      throw new Error(`Custom API error: ${response.statusText}`);
    }

    const data = await response.json();
    return data.completion || data.text || "";
  }
}
