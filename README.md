# Word文本自动补全插件 | Word Autocomplete Add-in

这是一个基于AI的Word文本自动补全插件，可以根据当前文档中正在输入的内容智能预测并提供后续文本建议。

This is an AI-powered Word autocomplete add-in that intelligently predicts and suggests subsequent text based on the current document content.

## 主要功能 | Key Features

1. 多种AI服务支持 | Multiple AI Service Support
   - OpenAI
   - Anthropic
   - OpenRoute
   - Google Gemini
   - 自定义API | Custom API

2. 智能建议显示 | Smart Suggestion Display
   - 侧边栏悬浮窗模式：在配置表单上方显示，大小自适应内容 | Sidebar floating window mode: Displays above the configuration form, size adapts to content
   - 内联模式：在光标位置直接显示灰色斜体建议 | Inline mode: Displays gray italic suggestions directly at cursor position
   - 可随时切换显示模式 | Display mode can be switched anytime

3. 灵活的触发方式 | Flexible Trigger Methods
   - 自动模式：输入停止后自动触发 | Auto mode: Triggers automatically after input stops
   - 手动模式：点击按钮触发 | Manual mode: Triggers by button click
   - Alt+D快捷键：快速接受建议 | Alt+D shortcut: Quickly accept suggestions

4. 丰富的配置选项 | Rich Configuration Options
   - 上下文范围：当前段落/整个文档/自定义段落数 | Context range: Current paragraph/Entire document/Custom paragraph count
   - 触发延迟：可自定义自动触发的等待时间 | Trigger delay: Customizable waiting time for auto-trigger
   - 防抖设置：避免频繁触发 | Debounce settings: Avoid frequent triggering
   - 模型参数：温度、最大生成长度等 | Model parameters: Temperature, maximum generation length, etc.
   - 系统提示词：自定义AI助手的行为 | System prompts: Customize AI assistant behavior

## 使用方法 | Usage Instructions

### 1. 基本设置 | Basic Settings

1. 选择AI服务提供商（OpenAI/Anthropic/OpenRoute/Gemini/自定义）| Select AI service provider (OpenAI/Anthropic/OpenRoute/Gemini/Custom)
2. 输入API密钥 | Enter API key
3. 从自动获取的模型列表中选择要使用的模型 | Select the model to use from automatically fetched model list
4. 点击"保存配置"按钮保存设置 | Click "Save Configuration" button to save settings

### 2. 高级配置 | Advanced Configuration

- **触发模式 | Trigger Mode**：
  - 自动：停止输入后自动提供建议 | Auto: Provides suggestions automatically after input stops
  - 手动：需要点击"触发补全"按钮 | Manual: Requires clicking "Trigger Completion" button

- **上下文范围 | Context Range**：
  - 当前段落：仅使用当前段落作为上下文 | Current paragraph: Uses only current paragraph as context
  - 整个文档：使用整个文档作为上下文 | Entire document: Uses the whole document as context
  - 自定义段落数：指定使用的段落数量 | Custom paragraph count: Specify number of paragraphs to use

- **显示方式 | Display Mode**：
  - 侧边栏悬浮窗：在配置表单上方显示建议 | Sidebar floating window: Displays suggestions above configuration form
  - 内联模式：在光标位置直接显示建议 | Inline mode: Displays suggestions directly at cursor position

- **其他参数 | Other Parameters**：
  - 触发延迟：设置自动模式下的等待时间（默认2秒）| Trigger delay: Set waiting time in auto mode (default 2 seconds)
  - 防抖延迟：避免频繁触发的时间间隔 | Debounce delay: Time interval to avoid frequent triggering
  - 最大生成长度：限制生成文本的长度 | Maximum generation length: Limit the length of generated text
  - 温度：控制生成文本的创造性（0-2）| Temperature: Control text generation creativity (0-2)
  - 系统提示词：自定义AI助手的行为指南 | System prompts: Customize AI assistant behavior guidelines

### 3. 开始使用 | Getting Started

1. 完成配置后点击"启动"按钮 | Click "Start" button after completing configuration
2. 将光标放在需要补全的位置 | Place cursor at the position needing completion
3. 根据选择的模式 | According to selected mode:
   - 自动模式：等待建议自动出现 | Auto mode: Wait for suggestions to appear automatically
   - 手动模式：点击"触发补全"按钮 | Manual mode: Click "Trigger Completion" button

### 4. 接受建议 | Accepting Suggestions

有三种方式接受建议 | Three ways to accept suggestions:
- 按Alt+D快捷键（推荐，适用于所有显示模式）| Press Alt+D shortcut (recommended, works in all display modes)
- 点击建议框（仅侧边栏模式）| Click suggestion box (sidebar mode only)
- 点击"触发补全"按钮 | Click "Trigger Completion" button

### 5. 停止使用 | Stop Using

点击"停止"按钮可以随时停止自动完成功能。| Click "Stop" button to stop autocomplete function anytime.

## 开发者信息 | Developer Information

### 构建和运行 | Build and Run

```bash
# 安装依赖 | Install dependencies
npm install

# 启动开发服务器 | Start development server
npm run dev

# 构建生产版本 | Build production version
npm run build
```

### 注意事项 | Notes

1. 确保API密钥配置正确 | Ensure API key is configured correctly
2. 建议先在小范围文本上测试 | Recommend testing on small text scope first
3. 根据网络状况调整触发延迟 | Adjust trigger delay according to network conditions
4. 可以通过系统提示词优化输出质量 | Can optimize output quality through system prompts
5. 如遇问题，查看浏览器控制台输出 | If problems occur, check browser console output

## 许可证 | License

MIT License
