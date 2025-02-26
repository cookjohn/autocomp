/* global document, Office, console, localStorage, setTimeout */
/* global HTMLElement, HTMLInputElement, HTMLSelectElement, HTMLTextAreaElement */

import { AutoCompleteEngine, AutoCompleteConfig } from "../autocomplete/engine";
import { LLMService } from "../autocomplete/api";
import { setAutoCompleteEngine } from "../commands/commands";

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

      // 绑定模型列表更新事件
      document.getElementById("provider")?.addEventListener("change", updateModelList);
      document.getElementById("api-key")?.addEventListener("change", updateModelList);
      document.getElementById("api-key")?.addEventListener("input", updateModelList);

      // 绑定模型选择事件
      document.getElementById("model")?.addEventListener("change", () => {
        const config = getConfig();
        localStorage.setItem("autoCompleteConfig", JSON.stringify(config));
      });

      // 绑定配置折叠事件
      document.getElementById("toggle-config")?.addEventListener("click", toggleConfig);
    })();
  }
});

async function updateModelList(): Promise<void> {
  try {
    const providerElement = document.getElementById("provider") as HTMLSelectElement;
    const apiKeyElement = document.getElementById("api-key") as HTMLInputElement;
    const modelElement = document.getElementById("model") as HTMLSelectElement;

    const provider = providerElement.value as Provider;
    const apiKey = apiKeyElement.value;

    modelElement.innerHTML = "<option value=''>加载中...</option>";

    if (provider && apiKey) {
      llmService = new LLMService({
        provider,
        apiKey,
        maxTokens: 100,
        temperature: 0.7,
      });

      const models = await llmService.getAvailableModels();
      const sortedModels = models.sort((a, b) => a.name.localeCompare(b.name));

      // 获取保存的配置中的模型
      const savedConfig = localStorage.getItem("autoCompleteConfig");
      const config = savedConfig ? (JSON.parse(savedConfig) as AutoCompleteConfig) : null;
      const savedModel = config?.apiConfig.model || "";

      // 生成选项并选中保存的模型
      const options = sortedModels.map(
        (model) =>
          `<option value="${model.id}"${model.id === savedModel ? " selected" : ""}>${
            model.name
          } (${model.context_length}tokens)</option>`
      );

      // 如果有保存的模型但不在列表中，添加一个选项
      if (savedModel && !sortedModels.some((model) => model.id === savedModel)) {
        options.unshift(`<option value="${savedModel}" selected>${savedModel}</option>`);
      }

      modelElement.innerHTML = options.join("");
    } else {
      modelElement.innerHTML = "<option value=''>请先填写API密钥</option>";
    }
  } catch (error) {
    console.error("Failed to fetch models:", error);
    const modelElement = document.getElementById("model") as HTMLSelectElement;
    modelElement.innerHTML = "<option value=''>获取模型列表失败</option>";
  }
}

async function initializeForm(): Promise<void> {
  const savedConfigStr = localStorage.getItem("autoCompleteConfig");
  if (!savedConfigStr) {
    return;
  }

  try {
    const config = JSON.parse(savedConfigStr) as AutoCompleteConfig;
    const elements = {
      "api-key": config.apiConfig.apiKey || "",
      provider: config.apiConfig.provider || "openai",
      "max-tokens": config.apiConfig.maxTokens?.toString() || "150",
      temperature: config.apiConfig.temperature?.toString() || "0.7",
      "trigger-mode": config.triggerMode || "manual",
      "context-range": config.contextRange || "paragraph",
      "custom-paragraphs": config.customParagraphs?.toString() || "3",
      "max-context": config.maxContextLength?.toString() || "2000",
      debounce: config.debounceMs?.toString() || "1000",
      "trigger-delay": config.triggerDelayMs?.toString() || "2000",
      "suggestion-position": config.suggestionPosition || "sidebar",
      "system-prompt": config.apiConfig.systemPrompt || "",
    };

    // 设置基本配置
    Object.entries(elements).forEach(([id, value]) => {
      const element = document.getElementById(id);
      if (element) {
        (element as HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement).value = value;
      }
    });

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
    suggestionPosition: (document.getElementById("suggestion-position") as HTMLSelectElement).value as "sidebar" | "inline",
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
    localStorage.setItem("autoCompleteConfig", JSON.stringify(config));

    autoCompleteEngine = new AutoCompleteEngine(config);
    setAutoCompleteEngine(autoCompleteEngine);
    await autoCompleteEngine.initialize();

    const startButton = document.getElementById("start-auto");
    const stopButton = document.getElementById("stop-auto");
    const configForm = document.getElementById("config-form");
    const runningControls = document.getElementById("running-controls");

    if (startButton && stopButton && configForm && runningControls) {
      startButton.style.display = "none";
      stopButton.style.display = "block";
      runningControls.style.display = "block";

      // 折叠配置表单
      const configContent = document.getElementById("config-content");
      const toggleButton = document.getElementById("toggle-config");
      if (configContent && toggleButton) {
        configContent.classList.add("collapsed");
        (toggleButton.querySelector(".ms-Button-label") as HTMLElement).textContent = "展开";
      }
    }

    showMessage("自动完成功能已启动", "success");
  } catch (error) {
    console.error("Failed to start auto-complete:", error);
    showMessage("启动自动完成功能失败", "error");
  }
}

function stopAutoComplete(): void {
  try {
    if (autoCompleteEngine) {
      autoCompleteEngine.dispose();
      setAutoCompleteEngine(null);
      autoCompleteEngine = null;
    }

    const startButton = document.getElementById("start-auto");
    const stopButton = document.getElementById("stop-auto");
    const configForm = document.getElementById("config-form");
    const runningControls = document.getElementById("running-controls");

    if (startButton && stopButton && configForm && runningControls) {
      startButton.style.display = "block";
      stopButton.style.display = "none";
      configForm.style.display = "block";
      runningControls.style.display = "none";
    }

    showMessage("自动完成功能已停止", "success");
  } catch (error) {
    console.error("Failed to stop auto-complete:", error);
    showMessage("停止自动完成功能失败", "error");
  }
}

async function triggerCompletion(): Promise<void> {
  try {
    if (autoCompleteEngine) {
      await autoCompleteEngine.triggerCompletion();
    }
  } catch (error) {
    console.error("Failed to trigger completion:", error);
    showMessage("触发补全失败", "error");
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
      localStorage.setItem("autoCompleteConfig", JSON.stringify(config));
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
