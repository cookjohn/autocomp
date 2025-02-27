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
      document.getElementById("save-config")?.addEventListener("click", () => {
        ConfigManager.saveFormConfig();
      });

      // 绑定provider变更事件
      document.getElementById("provider")?.addEventListener("change", () => {
        // 获取当前选中的提供商
        const providerElement = document.getElementById("provider") as HTMLSelectElement;
        const currentProvider = (providerElement?.value as Provider) || "openai";

        // 获取已保存的配置
        const savedConfig = ConfigManager.getSavedConfig() || ConfigManager.getDefaultConfig();

        // 更新 API 密钥输入框
        const apiKeyElement = document.getElementById("api-key") as HTMLInputElement;
        if (apiKeyElement) {
          apiKeyElement.value = savedConfig.apiConfig.providerConfigs[currentProvider].apiKey || "";
        }

        // 标记配置已更改
        ConfigManager.markConfigChanged();

        // 更新模型列表
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
        // 移除自动保存，可以添加UI提示用户需要点击保存按钮
        const saveButton = document.getElementById("save-config");
        if (saveButton) {
          saveButton.classList.add("highlight"); // 添加高亮样式提示用户保存
        }
        ConfigManager.markConfigChanged();
      });

      // 绑定配置折叠事件
      document.getElementById("toggle-config")?.addEventListener("click", toggleConfig);
    })();
  }
});

// 配置管理
const ConfigManager = {
  // 配置是否已更改但未保存
  configChanged: false,

  // 标记配置已更改
  markConfigChanged(): void {
    this.configChanged = true;
    const saveButton = document.getElementById("save-config");
    if (saveButton) {
      saveButton.classList.add("highlight");
    }
  },
  // 获取当前表单配置
  getFormConfig(): AutoCompleteConfig {
    // 获取当前选中的提供商
    const providerElement = document.getElementById("provider") as HTMLSelectElement;
    const currentProvider = (providerElement?.value as Provider) || "openai";

    // 获取已保存的配置（如果有）
    const savedConfig = this.getSavedConfig() || this.getDefaultConfig();

    // 创建新的 providerConfigs，默认使用已保存的值
    const providerConfigs = { ...savedConfig.apiConfig.providerConfigs };

    // 安全地获取表单元素值的辅助函数
    const getElementValue = (id: string, defaultValue: string = ""): string => {
      const element = document.getElementById(id) as HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement;
      return element?.value || defaultValue;
    };

    const getElementValueAsNumber = (id: string, defaultValue: number): number => {
      const value = getElementValue(id);
      const parsedValue = parseInt(value);
      return isNaN(parsedValue) ? defaultValue : parsedValue;
    };

    const getElementValueAsFloat = (id: string, defaultValue: number): number => {
      const value = getElementValue(id);
      const parsedValue = parseFloat(value);
      return isNaN(parsedValue) ? defaultValue : parsedValue;
    };

    // 只更新当前选中提供商的配置
    providerConfigs[currentProvider] = {
      apiKey: getElementValue("api-key"),
      model: getElementValue("model"),
    };

    return {
      triggerMode: getElementValue("trigger-mode", "manual") as "auto" | "manual",
      contextRange: getElementValue("context-range", "paragraph") as ContextRange,
      maxContextLength: getElementValueAsNumber("max-context", 2000),
      customParagraphs: getElementValueAsNumber("custom-paragraphs", 3),
      debounceMs: getElementValueAsNumber("debounce", 1000),
      triggerDelayMs: getElementValueAsNumber("trigger-delay", 2000),
      suggestionPosition: getElementValue("suggestion-position", "sidebar") as "sidebar" | "inline",
      apiConfig: {
        provider: currentProvider,
        maxTokens: getElementValueAsNumber("max-tokens", 150),
        temperature: getElementValueAsFloat("temperature", 0.7),
        systemPrompt: getElementValue("system-prompt", savedConfig.apiConfig.systemPrompt || ""),
        providerConfigs: providerConfigs,
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
        maxTokens: 150,
        temperature: 0.7,
        systemPrompt: "",
        providerConfigs: {
          openai: { apiKey: "", model: "" },
          anthropic: { apiKey: "", model: "" },
          openroute: { apiKey: "", model: "" },
          gemini: { apiKey: "", model: "" },
          custom: { apiKey: "", model: "" },
        },
      },
    };
  },

  // 获取当前provider的配置
  getCurrentProviderConfig(config: AutoCompleteConfig) {
    const provider = config.apiConfig.provider;
    return config.apiConfig.providerConfigs[provider];
  },

  // 将配置应用到表单
  applyConfigToForm(config: AutoCompleteConfig): void {
    // 获取当前选中的提供商
    const provider = config.apiConfig.provider;
    const providerConfig = config.apiConfig.providerConfigs[provider];

    // 设置表单值
    const formConfig = {
      provider: provider,
      "trigger-mode": config.triggerMode,
      "context-range": config.contextRange,
      "custom-paragraphs": config.customParagraphs?.toString(),
      "max-context": config.maxContextLength?.toString(),
      debounce: config.debounceMs?.toString(),
      "trigger-delay": config.triggerDelayMs?.toString(),
      "suggestion-position": config.suggestionPosition,
      "api-key": providerConfig.apiKey,
      "max-tokens": config.apiConfig.maxTokens?.toString(),
      temperature: config.apiConfig.temperature?.toString(),
      "system-prompt": config.apiConfig.systemPrompt,
    };

    // 将配置应用到表单
    Object.entries(formConfig).forEach(([id, value]) => {
      const element = document.getElementById(id);
      if (element) {
        (element as HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement).value = value || "";
      }
    });
  },

  // 保存表单配置并更新引擎
  saveFormConfig(): void {
    try {
      const config = this.getFormConfig();

      // 保存配置
      this.saveConfig(config);

      // 如果引擎已初始化，更新引擎配置
      if (autoCompleteEngine) {
        autoCompleteEngine.updateConfig(config);

        // 更新UI状态
        const triggerButton = document.getElementById("trigger-completion");
        if (triggerButton) {
          triggerButton.style.display = config.triggerMode === "auto" ? "none" : "block";
        }
      }

      // 重置配置变更标记
      this.configChanged = false;
      const saveButton = document.getElementById("save-config");
      if (saveButton) {
        saveButton.classList.remove("highlight");
      }

      showMessage("配置已更新", "success");
    } catch (error) {
      console.error("Failed to save config:", error);
      showMessage("保存配置失败", "error");
    }
  },

  // 初始化配置
  async initializeConfig(): Promise<AutoCompleteConfig> {
    const config = this.getSavedConfig() || this.getDefaultConfig();
    this.applyConfigToForm(config);
    return config;
  },
};

async function updateModelList(): Promise<void> {
  try {
    const config = ConfigManager.getFormConfig();
    const modelElement = document.getElementById("model") as HTMLSelectElement;
    modelElement.innerHTML = "<option value=''>加载中...</option>";

    const provider = config.apiConfig.provider;
    const apiKey = (document.getElementById("api-key") as HTMLInputElement).value;

    if (!apiKey) {
      modelElement.innerHTML = "<option value=''>请先填写API密钥</option>";
      return;
    }

    // 使用当前表单的值创建配置
    const currentConfig = {
      ...config,
      apiConfig: {
        ...config.apiConfig,
        providerConfigs: {
          ...config.apiConfig.providerConfigs,
          [provider]: {
            ...config.apiConfig.providerConfigs[provider],
            apiKey: apiKey,
          },
        },
      },
    };

    llmService = new LLMService(currentConfig.apiConfig);
    const models = await llmService.getAvailableModels();
    if (!models.length) {
      modelElement.innerHTML = "<option value=''>未找到可用模型</option>";
      return;
    }

    const sortedModels = models.sort((a, b) => a.name.localeCompare(b.name));
    const savedConfig = ConfigManager.getSavedConfig();
    const savedProviderConfig = savedConfig?.apiConfig.providerConfigs[provider];
    const savedModel = savedProviderConfig?.model;

    // 生成模型列表选项
    const options = sortedModels.map(
      (model) => `<option value="${model.id}">${model.name} (${model.context_length}tokens)</option>`
    );

    // 如果有上次使用的模型且在列表中，添加选中状态
    if (savedModel && sortedModels.some((model) => model.id === savedModel)) {
      modelElement.innerHTML = options.join("");
      modelElement.value = savedModel;

      // 不再自动保存配置，而是标记配置已更改
      ConfigManager.markConfigChanged();
    } else if (savedModel) {
      // 如果上次使用的模型不在列表中，显示提示
      modelElement.innerHTML =
        `<option value="">上次使用的模型 ${savedModel} 不可用，请重新选择</option>` + options.join("");
      showMessage("模型列表已更新，请重新选择模型", "error");
    } else {
      // 首次使用或没有保存的模型
      modelElement.innerHTML = `<option value="">请选择模型</option>` + options.join("");
      // 如果有可用模型，自动选择第一个
      if (sortedModels.length > 0) {
        setTimeout(() => {
          modelElement.value = sortedModels[0].id;
          ConfigManager.markConfigChanged();
        }, 0);
      }
    }
  } catch (error) {
    console.error("Failed to fetch models:", error);
    const modelElement = document.getElementById("model") as HTMLSelectElement;
    modelElement.innerHTML = "<option value=''>获取模型列表失败</option>";
  }
}

// 已移除 loadSavedConfig 函数，使用 ConfigManager.initializeConfig 替代

async function initializeForm(): Promise<void> {
  try {
    // 初始化配置
    await ConfigManager.initializeConfig();

    // 更新模型列表
    await updateModelList();
    // 为所有表单元素添加变更监听
    addChangeListenersToFormElements();
  } catch (error) {
    console.error("Failed to initialize form:", error);
  }
}

// 为所有表单元素添加变更监听
function addChangeListenersToFormElements(): void {
  const formElements = document.querySelectorAll("#config-form input, #config-form select, #config-form textarea");
  formElements.forEach((element) => {
    element.addEventListener("change", () => {
      ConfigManager.markConfigChanged();
    });
    // 对于文本输入，也监听input事件
    if ((element instanceof HTMLInputElement && element.type === "text") || element instanceof HTMLTextAreaElement) {
      element.addEventListener("input", () => {
        ConfigManager.markConfigChanged();
      });
    }
  });
}

// 已移除 getConfig 函数，使用 ConfigManager.getFormConfig 替代

async function startAutoComplete(): Promise<void> {
  try {
    // 获取表单配置但不自动保存
    const config = ConfigManager.getFormConfig();

    // 检查必要的配置
    const provider = config.apiConfig.provider;
    const providerConfig = config.apiConfig.providerConfigs[provider];

    if (!providerConfig.apiKey) {
      showMessage("请输入API密钥", "error");
      return;
    }

    if (!providerConfig.model) {
      // 尝试从模型列表中获取第一个模型
      const modelElement = document.getElementById("model") as HTMLSelectElement;
      if (modelElement && modelElement.options.length > 1) {
        modelElement.selectedIndex = 1; // 选择第一个非空选项
        // 更新配置中的模型
        config.apiConfig.providerConfigs[provider].model = modelElement.value;
        ConfigManager.markConfigChanged();
      } else {
        showMessage("请选择模型", "error");
        return;
      }
    }

    // 如果有未保存的配置更改，自动保存
    if (ConfigManager.configChanged) {
      ConfigManager.saveFormConfig();
      showMessage("检测到配置已更改，已自动保存", "success");
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
      // 确保移除disabled属性，使按钮可以点击
      startButton.removeAttribute("disabled");
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

// 已移除 saveConfig 函数，使用 ConfigManager.saveFormConfig 替代

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
