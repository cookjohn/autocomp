interface ProviderConfig {
  apiKey: string;
  model: string;
  endpoint?: string;
}

interface LLMConfig {
  provider: "openai" | "anthropic" | "openroute" | "gemini" | "custom";
  maxTokens: number;
  temperature: number;
  systemPrompt?: string;
  providerConfigs: {
    [key in "openai" | "anthropic" | "openroute" | "gemini" | "custom"]: ProviderConfig;
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
        default:
          throw new Error(`Unsupported provider for model list: ${this.config.provider}`);
      }
    } catch (error) {
      console.error("Failed to fetch models:", error);
      throw error;
    }
  }

  private async getOpenAIModels(): Promise<LLMModel[]> {
    const openaiConfig = this.config.providerConfigs.openai;
    const response = await fetch("https://api.openai.com/v1/models", {
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
    const response = await fetch("https://api.anthropic.com/v1/models", {
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
    const response = await fetch("https://openrouter.ai/api/v1/models", {
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
    const response = await fetch(`https://generativelanguage.googleapis.com/v1/models?key=${geminiConfig.apiKey}`, {
      method: "GET",
    });

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

    const response = await fetch("https://api.openai.com/v1/chat/completions", {
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

    const response = await fetch("https://openrouter.ai/api/v1/chat/completions", {
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
    const response = await fetch(url, {
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

  private async completeWithAnthropic(context: string): Promise<string> {
    const anthropicConfig = this.config.providerConfigs.anthropic;
    if (!anthropicConfig.model) {
      throw new Error("Anthropic requires a model selection");
    }

    const response = await fetch("https://api.anthropic.com/v1/complete", {
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

    const response = await fetch(customConfig.endpoint, {
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
