/* global document, Office, console, localStorage, setTimeout */
/* global HTMLElement, HTMLInputElement, HTMLSelectElement, HTMLTextAreaElement */

import { AutoCompleteEngine, AutoCompleteConfig } from "../autocomplete/engine";
import { LLMService } from "../autocomplete/api";
import { setAutoCompleteEngine } from "../commands/commands";
import { debounce } from "../utils/debounce";

type Provider = "openai" | "anthropic" | "openroute" | "gemini" | "custom";
type ContextRange = "paragraph" | "document" | "custom";

let autoCompleteEngine: AutoCompleteEngine | null = null;
let llmService: LLMService | null = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    const sideloadMsg = document.getElementById("sideload-msg");
    const appBody = document.getElementById("app-body");

    if (sideloadMsg && appBody) {
      sideloadMsg.style.display = "none";
      appBody.style.display = "flex";
    }

    // 初始化表单和绑定事件
    void (async () => {
      await initializeForm();

      document.getElementById("start-auto")?.addEventListener("click", startAutoComplete);
      document.getElementById("stop-auto")?.addEventListener("click", stopAutoComplete);
      document.getElementById("trigger-completion")?.addEventListener("click", triggerCompletion);
      document.getElementById("save-config")?.addEventListener("click", saveConfig);

      // 绑定provider变更事件
      document.getElementById("provider")?.addEventListener("change", () => {
        void updateModelList();
      });

      // 绑定API密钥变更事件
      const apiKeyInput = document.getElementById("api-key");
      const updateModelListOnInput = debounce(async () => {
        await updateModelList();
      }, 500); // 500ms防抖

      apiKeyInput?.addEventListener("input", () => {
        void updateModelListOnInput();
      });

      // 绑定模型选择事件
      document.getElementById("model")?.addEventListener("change", () => {
        const config = ConfigManager.getFormConfig();
        ConfigManager.saveConfig(config);
      });

      // 绑定配置折叠事件
      document.getElementById("toggle-config")?.addEventListener("click", toggleConfig);
    })();
  }
});

// 配置管理
const ConfigManager = {
  // 获取当前表单配置
  getFormConfig(): AutoCompleteConfig {
    return {
      triggerMode: (document.getElementById("trigger-mode") as HTMLSelectElement).value as "auto" | "manual",
      contextRange: (document.getElementById("context-range") as HTMLSelectElement).value as ContextRange,
      maxContextLength: parseInt((document.getElementById("max-context") as HTMLInputElement).value) || 2000,
      customParagraphs: parseInt((document.getElementById("custom-paragraphs") as HTMLInputElement).value) || 3,
      debounceMs: parseInt((document.getElementById("debounce") as HTMLInputElement).value) || 1000,
      triggerDelayMs: parseInt((document.getElementById("trigger-delay") as HTMLInputElement).value) || 2000,
      suggestionPosition: (document.getElementById("suggestion-position") as HTMLSelectElement).value as
        | "sidebar"
        | "inline",
      apiConfig: {
        provider: (document.getElementById("provider") as HTMLSelectElement).value as Provider,
        apiKey: (document.getElementById("api-key") as HTMLInputElement).value,
        model: (document.getElementById("model") as HTMLSelectElement).value,
        maxTokens: parseInt((document.getElementById("max-tokens") as HTMLInputElement).value) || 150,
        temperature: parseFloat((document.getElementById("temperature") as HTMLInputElement).value) || 0.7,
        systemPrompt: (document.getElementById("system-prompt") as HTMLTextAreaElement).value,
      },
    };
  },

  // 获取保存的配置
  getSavedConfig(): AutoCompleteConfig | null {
    const saved = localStorage.getItem("autoCompleteConfig");
    return saved ? JSON.parse(saved) : null;
  },

  // 保存配置
  saveConfig(config: AutoCompleteConfig): void {
    localStorage.setItem("autoCompleteConfig", JSON.stringify(config));
  },

  // 更新配置
  updateConfig(config: Partial<AutoCompleteConfig>): void {
    const currentConfig = this.getSavedConfig() || this.getDefaultConfig();
    this.saveConfig({
      ...currentConfig,
      ...config,
      apiConfig: {
        ...currentConfig.apiConfig,
        ...config.apiConfig,
      },
    });
  },

  // 获取默认配置
  getDefaultConfig(): AutoCompleteConfig {
    return {
      triggerMode: "manual",
      contextRange: "paragraph",
      maxContextLength: 2000,
      customParagraphs: 3,
      debounceMs: 1000,
      triggerDelayMs: 2000,
      suggestionPosition: "sidebar",
      apiConfig: {
        provider: "openai",
        apiKey: "",
        maxTokens: 150,
        temperature: 0.7,
        systemPrompt: "",
      },
    };
  },
};

async function updateModelList(): Promise<void> {
  try {
    const config = ConfigManager.getFormConfig();
    const modelElement = document.getElementById("model") as HTMLSelectElement;
    modelElement.innerHTML = "<option value=''>加载中...</option>";

    if (!config.apiConfig.apiKey) {
      modelElement.innerHTML = "<option value=''>请先填写API密钥</option>";
      return;
    }

    llmService = new LLMService(config.apiConfig);
    const models = await llmService.getAvailableModels();
    if (!models.length) {
      modelElement.innerHTML = "<option value=''>未找到可用模型</option>";
      return;
    }

    const sortedModels = models.sort((a, b) => a.name.localeCompare(b.name));
    const savedConfig = ConfigManager.getSavedConfig();
    const savedModel = savedConfig?.apiConfig.model;

    // 生成模型列表选项
    const options = sortedModels.map(
      (model) => `<option value="${model.id}">${model.name} (${model.context_length}tokens)</option>`
    );

    // 如果有上次使用的模型且在列表中，添加选中状态
    if (savedModel && sortedModels.some((model) => model.id === savedModel)) {
      modelElement.innerHTML = options.join("");
      modelElement.value = savedModel;
      // 保存当前配置
      config.apiConfig.model = savedModel;
      ConfigManager.updateConfig(config);
    } else if (savedModel) {
      // 如果上次使用的模型不在列表中，显示提示
      modelElement.innerHTML =
        `<option value="">上次使用的模型 ${savedModel} 不可用，请重新选择</option>` + options.join("");
      showMessage("模型列表已更新，请重新选择模型", "error");
    } else {
      // 首次使用或没有保存的模型
      modelElement.innerHTML = `<option value="">请选择模型</option>` + options.join("");
    }
  } catch (error) {
    console.error("Failed to fetch models:", error);
    const modelElement = document.getElementById("model") as HTMLSelectElement;
    modelElement.innerHTML = "<option value=''>获取模型列表失败</option>";
  }
}

async function loadSavedConfig(): Promise<void> {
  const config = ConfigManager.getSavedConfig() || ConfigManager.getDefaultConfig();

  // 设置表单值
  const formConfig = {
    provider: config.apiConfig.provider,
    "trigger-mode": config.triggerMode,
    "context-range": config.contextRange,
    "custom-paragraphs": config.customParagraphs?.toString(),
    "max-context": config.maxContextLength?.toString(),
    debounce: config.debounceMs?.toString(),
    "trigger-delay": config.triggerDelayMs?.toString(),
    "suggestion-position": config.suggestionPosition,
    "api-key": config.apiConfig.apiKey,
    "max-tokens": config.apiConfig.maxTokens?.toString(),
    temperature: config.apiConfig.temperature?.toString(),
    "system-prompt": config.apiConfig.systemPrompt,
  };

  Object.entries(formConfig).forEach(([id, value]) => {
    const element = document.getElementById(id);
    if (element) {
      (element as HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement).value = value || "";
    }
  });
}

async function initializeForm(): Promise<void> {
  try {
    await loadSavedConfig();

    // 更新模型列表
    await updateModelList();
  } catch (error) {
    console.error("Failed to initialize form:", error);
  }
}

function getConfig(): AutoCompleteConfig {
  return {
    triggerMode: (document.getElementById("trigger-mode") as HTMLSelectElement).value as "auto" | "manual",
    contextRange: (document.getElementById("context-range") as HTMLSelectElement).value as ContextRange,
    maxContextLength: parseInt((document.getElementById("max-context") as HTMLInputElement).value) || 2000,
    customParagraphs: parseInt((document.getElementById("custom-paragraphs") as HTMLInputElement).value) || 3,
    debounceMs: parseInt((document.getElementById("debounce") as HTMLInputElement).value) || 1000,
    triggerDelayMs: parseInt((document.getElementById("trigger-delay") as HTMLInputElement).value) || 2000,
    suggestionPosition: (document.getElementById("suggestion-position") as HTMLSelectElement).value as
      | "sidebar"
      | "inline",
    apiConfig: {
      provider: (document.getElementById("provider") as HTMLSelectElement).value as Provider,
      apiKey: (document.getElementById("api-key") as HTMLInputElement).value,
      model: (document.getElementById("model") as HTMLSelectElement).value,
      maxTokens: parseInt((document.getElementById("max-tokens") as HTMLInputElement).value) || 150,
      temperature: parseFloat((document.getElementById("temperature") as HTMLInputElement).value) || 0.7,
      systemPrompt: (document.getElementById("system-prompt") as HTMLTextAreaElement).value,
    },
  };
}

async function startAutoComplete(): Promise<void> {
  try {
    const config = getConfig();

    // 检查必要的配置
    if (!config.apiConfig.apiKey) {
      showMessage("请输入API密钥", "error");
      return;
    }

    if (!config.apiConfig.model) {
      showMessage("请选择模型", "error");
      return;
    }

    // 禁用启动按钮
    const startButton = document.getElementById("start-auto");
    if (startButton) {
      startButton.setAttribute("disabled", "true");
      (startButton.querySelector(".ms-Button-label") as HTMLElement).textContent = "正在测试...";
    }

    // 创建临时服务实例进行测试
    const testService = new LLMService(config.apiConfig);
    try {
      // 发送测试请求
      const testResult = await testService.complete("这是一个测试请求。只需要返回OK");
      if (!testResult) {
        throw new Error("测试请求返回为空");
      }
    } catch (error) {
      console.error("API test failed:", error);
      showMessage("API连接测试失败，请检查配置", "error");

      // 恢复启动按钮
      if (startButton) {
        startButton.removeAttribute("disabled");
        (startButton.querySelector(".ms-Button-label") as HTMLElement).textContent = "启动";
      }
      return;
    }

    // 保存配置
    ConfigManager.saveConfig(config);

    // 初始化自动完成引擎
    autoCompleteEngine = new AutoCompleteEngine(config);
    setAutoCompleteEngine(autoCompleteEngine);
    await autoCompleteEngine.initialize();

    const stopButton = document.getElementById("stop-auto");
    const configForm = document.getElementById("config-form");
    const runningControls = document.getElementById("running-controls");

    if (startButton && stopButton && configForm && runningControls) {
      startButton.style.display = "none";
      stopButton.style.display = "block";
      runningControls.style.display = "block";

      // 在自动模式下隐藏触发补全按钮
      const triggerButton = document.getElementById("trigger-completion");
      if (triggerButton) {
        triggerButton.style.display = config.triggerMode === "auto" ? "none" : "block";
      }

      // 折叠配置表单
      const configContent = document.getElementById("config-content");
      const toggleButton = document.getElementById("toggle-config");
      if (configContent && toggleButton) {
        configContent.classList.add("collapsed");
        (toggleButton.querySelector(".ms-Button-label") as HTMLElement).textContent = "展开";
      }
    }

    showMessage("自动补全已启动", "success");
  } catch (error) {
    console.error("Failed to start auto-complete:", error);
    showMessage("自动补全启动失败", "error");

    // 恢复启动按钮
    const startButton = document.getElementById("start-auto");
    if (startButton) {
      startButton.removeAttribute("disabled");
      (startButton.querySelector(".ms-Button-label") as HTMLElement).textContent = "启动";
    }
  }
}

function stopAutoComplete(): void {
  try {
    if (autoCompleteEngine) {
      autoCompleteEngine.dispose();
      setAutoCompleteEngine(null);
      autoCompleteEngine = null;
    }

    // 恢复按钮状态
    const startButton = document.getElementById("start-auto");
    const stopButton = document.getElementById("stop-auto");
    const configForm = document.getElementById("config-form");
    const runningControls = document.getElementById("running-controls");

    if (startButton && stopButton && configForm && runningControls) {
      // 恢复启动按钮的文本
      (startButton.querySelector(".ms-Button-label") as HTMLElement).textContent = "启动";
      // 显示/隐藏相关元素
      startButton.style.display = "block";
      stopButton.style.display = "none";
      configForm.style.display = "block";
      runningControls.style.display = "none";
    }

    showMessage("自动补全已停止", "success");
  } catch (error) {
    console.error("Failed to stop auto-complete:", error);
    showMessage("停止自动补全失败", "error");
  }
}

async function triggerCompletion(): Promise<void> {
  try {
    if (!autoCompleteEngine) {
      showMessage("请先启动自动补全功能", "error");
      return;
    }

    const triggerButton = document.getElementById("trigger-completion");
    if (triggerButton) {
      triggerButton.setAttribute("disabled", "true");
    }

    await autoCompleteEngine.triggerCompletion();
    showMessage("已触发补全", "success");
  } catch (error) {
    console.error("Failed to trigger completion:", error);
    showMessage("触发补全失败", "error");
  } finally {
    const triggerButton = document.getElementById("trigger-completion");
    if (triggerButton) {
      triggerButton.removeAttribute("disabled");
    }
  }
}

function toggleConfig(): void {
  const configContent = document.getElementById("config-content");
  const toggleButton = document.getElementById("toggle-config");
  if (configContent && toggleButton) {
    const isCollapsed = configContent.classList.toggle("collapsed");
    (toggleButton.querySelector(".ms-Button-label") as HTMLElement).textContent = isCollapsed ? "展开" : "折叠";
  }
}

function saveConfig(): void {
  try {
    if (autoCompleteEngine) {
      const config = getConfig();
      autoCompleteEngine.updateConfig(config);

      // 保存配置
      ConfigManager.saveConfig(config);

      // 更新触发按钮显示状态
      const triggerButton = document.getElementById("trigger-completion");
      if (triggerButton) {
        triggerButton.style.display = config.triggerMode === "auto" ? "none" : "block";
      }

      showMessage("配置已更新", "success");
    }
  } catch (error) {
    console.error("Failed to save config:", error);
    showMessage("保存配置失败", "error");
  }
}

function showMessage(message: string, type: "success" | "error"): void {
  const messageElement = document.getElementById("message");
  if (messageElement) {
    messageElement.textContent = message;
    messageElement.className = `message ${type}`;
    messageElement.style.display = "block";
    setTimeout(() => {
      messageElement.style.display = "none";
    }, 3000);
  }
}
