# Translate Pro - 项目立项文档

> 版本: v1.2.1 | 日期: 2026-07-06 | 作者: SYNTEC

---

## 一、项目概述

### 1.1 项目背景
工业CNC软件（数控系统）的多语言本地化是一项复杂且重复性高的工作。现有通用翻译工具无法直接处理`.res`（GZIP压缩）格式的资源文件，且缺乏针对CNC行业的专业术语管理。此外，许多工业环境存在网络限制，需要支持离线翻译工作流。同时，CNC软件的品牌定制（Logo、快捷方式）也是工程师常见的需求。

### 1.2 项目定位
**Translate Pro** 是一款专为CNC软件本地化设计的Windows桌面翻译工具，支持XML/RES资源文件的批量翻译，覆盖127种语言变体，提供智能缩写、翻译记忆、DiskC工作流等特色功能。特别支持**Excel离线翻译模式**，允许用户导出翻译表进行人工校对，适用于无网络环境或需要专业翻译人员参与的场景。同时提供**Logo定制工具**，支持图片转换和快捷方式创建，满足CNC软件品牌个性化需求。

### 1.3 核心价值
- **行业专注**: 支持CNC厂商（禾川、高创、松下、汇川、台达、雷赛、信捷等）的资源格式
- **效率提升**: Provider 自适应并发、翻译记忆复用、智能缩写
- **离线协作**: 支持导出Excel进行人工翻译，适用于无网络环境或专业翻译人员
- **定制化工具**: Logo转换、快捷方式创建等CNC客制化辅助功能

### 1.4 Excel离线翻译工作流

```
┌─────────────────────────────────────────────────────────────┐
│                    Excel离线翻译流程                         │
├─────────────────────────────────────────────────────────────┤
│  有网络环境                    无网络环境                    │
│  ┌─────────────┐              ┌─────────────┐              │
│  │ 添加源文件  │              │  人工翻译   │              │
│  │ 选择语言    │   导出Excel  │  校对润色   │  导入Excel   │
│  │ 导出Excel ──┼──────────────┼─────────────┼──────────→   │
│  │             │              │             │   更新翻译表  │
│  └─────────────┘              └─────────────┘              │
│                                                             │
│  应用场景:                                                  │
│  • 工厂车间无外网                                           │
│  • 专业翻译公司协作                                         │
│  • 多人并行翻译                                             │
│  • 质量审核流程                                             │
└─────────────────────────────────────────────────────────────┘
```

### 1.5 核心功能清单

| 功能模块 | 功能项 | 说明 |
|----------|--------|------|
| 文件管理 | 添加文件/文件夹 | 支持XML和.res格式 |
| | DiskC工作流 | 自动发现CNC软件目录结构 |
| | 打包.res | 将XML打包为GZIP压缩格式 |
| 翻译引擎 | 多 Provider 翻译 | DeepL、百度、有道、小牛、腾讯云、火山引擎、阿里云及 Google GTX 测试模式 |
| | 翻译记忆 | JSON本地存储，避免重复翻译 |
| Excel协作 | **导出Excel** | 生成.xlsx文件，支持离线人工翻译 |
| | **导入Excel** | 读取人工翻译结果，增量更新翻译表 |
| 智能处理 | 语言缩写 | 按语言特性自动截断过长文本 |
| 输出生成 | XML输出 | 按语言生成独立XML文件 |
| | .res打包 | 可选GZIP压缩输出 |
| Logo定制 | **图片转换** | 将任意图片转换为ICO(128x128)和GIF(75x94) |
| | **自定义命名** | 支持自定义公司名称(如"LOGO"、"禾川") |
| | **快捷方式创建** | 自动创建桌面快捷方式并刷新图标缓存 |
| | **图标缓存清理** | 自动清理Windows图标缓存确保更新生效 |
| 其他 | 日志导出 | 保存翻译过程日志为txt文件 |
| | CLI模式 | 支持命令行无界面批量处理 |

### 1.6 项目目标用户

| 用户类型 | 使用场景 | 核心需求 |
|----------|----------|----------|
| CNC软件工程师 | 本地化翻译、品牌定制 | 高效翻译、一键部署 |
| 翻译人员 | 人工翻译、质量审核 | Excel离线协作 |
| 项目经理 | 批量部署、进度跟踪 | 批量操作、日志导出 |
| 系统集成商 | OEM定制、客户交付 | Logo定制、快捷方式创建 |
| 工厂维护人员 | 现场部署、快速安装 | 离线工具、简单操作 |

---

## 二、竞品分析

### 2.1 开源翻译工具对比

| 工具 | 类型 | 技术栈 | GitHub Stars | 核心特点 | 局限性 |
|------|------|--------|--------------|----------|--------|
| **Weblate** | Web平台 | Python/Django | 6k | 持续本地化、Git集成、团队协作 | 需服务器部署，不支持RES格式 |
| **Pootle** | Web平台 | Python/Django | 1.5k | 在线协作、翻译统计、权限管理 | 仅支持PO/XLIFF，无桌面版，仅Python2 |
| **OmegaT** | 桌面CAT | Java | 520 | 翻译记忆(TMX)、模糊匹配、术语库 | 学习曲线陡峭，XML支持有限 |

**关键差异**: 上述工具均需要网络环境或专业格式支持，而Translate Pro的Excel离线翻译模式允许在完全断网环境下进行人工翻译工作。同时，Logo定制工具为CNC软件提供了品牌个性化的便捷方案，这是其他翻译工具所不具备的。

### 2.2 商业工具参考

| 工具 | 类型 | 价格模式 | 适用场景 |
|------|------|----------|----------|
| **SDL Trados** | 桌面CAT | 订阅制 | 专业翻译公司 |
| **memoQ** | 桌面/Web | 订阅制 | 企业级本地化 |
| **Crowdin** | Web平台 | 免费+付费 | 开源项目本地化 |
| **Lokalise** | Web平台 | 订阅制 | 软件/App本地化 |

### 2.3 Translate Pro 差异化优势

| 特性 | Translate Pro | 竞品 |
|------|---------------|------|
| .res GZIP格式支持 | ✅ 原生支持 | ❌ 不支持 |
| DiskC工作流 | ✅ 自动发现目录 | ❌ 需手动配置 |
| 智能语言缩写 | ✅ 127种语言规则 | ⚠️ 有限支持 |
| 离线翻译记忆 | ✅ JSON本地存储 | ⚠️ 需服务器 |
| Excel离线翻译 | ✅ 导出/导入 | ⚠️ 需专业工具 |
| Logo定制工具 | ✅ 图标转换+快捷方式 | ❌ 无 |
| 单文件部署 | ✅ EXE免安装 | ❌ 需安装依赖 |
| CLI批量模式 | ✅ 无界面命令行 | ⚠️ 需脚本 |

---

## 三、UI风格规范

### 3.1 配色体系

| 角色 | 色值 | 应用场景 |
|------|------|----------|
| 侧边栏背景 | `#2D2B4E` | 左侧导航区域 |
| 主题强调色 | `#6C63FF` | 按钮、进度条、选中状态 |
| 主内容区背景 | `#F5F6FA` | 工作区底色 |
| 卡片背景 | `#FFFFFF` (white) | 内容卡片、统计面板 |
| 主文字色 | `#2D2B4E` | 标题、正文 |
| 副标题文字 | `#6B6B8D` | Subtitle.TLabel |
| 次要文字色 | `#8E8EA0` | 统计标签 |
| 导航未激活 | `#A5A0C0` | 未选中导航项 |
| 导航激活色 | `#FFFFFF` | 选中导航项 |
| Tab未激活背景 | `#E8E8F0` | 未选中标签页 |
| Tab未激活文字 | `#6E6E73` | 未选中标签页文字 |
| Tab悬停背景 | `#D0D0E0` | Tab.TButton active状态 |
| Tab激活背景 | `#FFFFFF` | 选中标签页 |
| Tab激活文字 | `#6C63FF` | 选中标签页文字 |
| Tab激活悬停 | `#F5F5F7` | TabActive.TButton active状态 |
| 进度条底色 | `#E8E8F0` | TProgressbar troughcolor |
| 列表框背景 | `#FAFBFE` | file_listbox |
| 列表框选中 | `#6C63FF` | selectbackground |
| 列表框边框 | `#E8E8F0` | highlightbackground |

### 3.2 字体规范

| 元素 | 样式类 | 字体 | 字号 | 字重 |
|------|--------|------|------|------|
| 侧边栏标题 | Header.TLabel | Segoe UI | 11pt | Bold |
| 导航项 | Nav.TLabel | Segoe UI | 10pt | Regular |
| 导航项(激活) | NavActive.TLabel | Segoe UI | 10pt | Bold |
| 页面标题 | Title.TLabel | Segoe UI | 13pt | Bold |
| 副标题 | Subtitle.TLabel | Segoe UI | 9pt | Regular |
| 统计数字 | StatValue.TLabel | Segoe UI | 24pt | Bold |
| 统计标签 | StatLabel.TLabel | Segoe UI | 9pt | Regular |
| 强调按钮 | Accent.TButton | Segoe UI | 10pt | Bold |
| 普通按钮 | TButton | Segoe UI | 9pt | Regular |
| 标签页按钮 | Tab.TButton | Segoe UI | 9pt | Regular |
| 标签页(激活) | TabActive.TButton | Segoe UI | 9pt | Bold |
| 复选框 | TCheckbutton | Segoe UI | 9pt | Regular |
| 卡片标题 | Card.TLabelframe.Label | Segoe UI | 11pt | Bold |
| 统计区标题 | (内联) | Segoe UI | 14pt | Bold |
| 列表框 | (内联) | Segoe UI | 9pt | Regular |

### 3.3 布局结构

```
┌─────────────┬────────────────────────────┬─────────────┐
│  侧边栏     │        主内容区            │  统计面板   │
│   200px     │         弹性宽度           │   220px     │
│             │                            │             │
│  Logo       │  标题 + 操作按钮           │  文件统计   │
│  导航菜单   │  文件管理卡片              │  语言数量   │
│  (6项)      │    工具栏(含Excel按钮)     │  翻译条目   │
│             │    文件列表                │  进度条     │
│             │  标签页(语言/Logo)         │  提示信息   │
│             │    - 目标语言选择          │             │
│             │    - Logo转换设置          │             │
│             │  日志面板                  │             │
└─────────────┴────────────────────────────┴─────────────┘
```

### 3.4 ttk主题与样式

- **基础主题**: `clam` (line 330)
- **定义样式类** (共19个):
  `TFrame`, `Card.TFrame`, `Sidebar.TFrame`, `Stats.TFrame`,
  `Header.TLabel`, `Nav.TLabel`, `NavActive.TLabel`, `Title.TLabel`,
  `Subtitle.TLabel`, `StatValue.TLabel`, `StatLabel.TLabel`,
  `Accent.TButton`, `TButton`, `Tab.TButton`, `TabActive.TButton`,
  `TCheckbutton`, `TProgressbar`, `Card.TLabelframe`, `Card.TLabelframe.Label`

### 3.5 组件设计语言

- **卡片**: 白色背景、`Card.TLabelframe` 带标题
- **按钮**: 主按钮(`Accent.TButton`)紫色强调，次按钮灰色
- **标签页**: 自定义`Tab.TButton`/`TabActive.TButton`切换，使用`grid_remove()`/`grid()`显隐
- **进度条**: `TProgressbar`厚度8px，紫色(`#6C63FF`)填充
- **列表框**: `#FAFBFE`背景，选中项紫色，边框`#E8E8F0`

---

## 四、交互层设计

### 4.1 导航结构

| 序号 | 导航项 | 图标 | 内部名 | 功能 |
|------|--------|------|--------|------|
| 1 | 文件管理 | 📁 | files | 添加/移除源文件，DiskC工作流 |
| 2 | 语言选择 | 🌐 | langs | 127种语言多选，搜索过滤 |
| 3 | Logo设置 | 🖼️ | logo | 图片转换、快捷方式创建 |
| 4 | 处理进度 | 📊 | progress | 翻译进度，实时统计 |
| 5 | 处理日志 | 📝 | logs | 详细日志，导出功能 |
| 6 | 设置选项 | ⚙️ | settings | API配置，默认语言 |

**导航切换机制**: `<Enter>`/`<Leave>`事件绑定切换`Nav.TLabel`/`NavActive.TLabel`样式，点击调用`show_content(name)`切换内容区。

### 4.2 标签页结构

内容区包含两个标签页，通过`grid_remove()`/`grid()`切换：

| 标签页 | 按钮文字 | 内容 |
|--------|----------|------|
| 目标语言 | 🌍 目标语言 | 语言搜索、全选/取消、API选择、语言列表 |
| Logo设置 | 🖼️ Logo设置 | 图片选择、公司名称、转换/创建按钮 |

### 4.3 工具栏功能 (文件管理页)

| 按钮 | 功能 |
|------|------|
| + 添加文件夹 | `filedialog.askdirectory` |
| + 添加文件 | `filedialog.askopenfilename` (*.xml;*.res) |
| DiskC 工作流 | 自动发现CNC目录结构 |
| 移除选中 | 删除列表选中项 |
| 清空列表 | 清空所有文件 |
| 打包 .res | `BooleanVar`复选框 |
| 翻译表 | 查看翻译记忆库 |
| 导出Excel | `export_to_excel()` |
| 导入Excel | `import_from_excel()` |
| 导出日志 | `export_log()` |

### 4.4 核心用户流程

**在线翻译流程**:
```
启动应用 → 添加文件/文件夹 → 选择目标语言 → 配置翻译引擎
→ 点击"开始翻译" → 实时进度+日志 → 翻译完成 → "确认生成"
→ 输出XML/RES文件
```

**Excel离线翻译流程**:
```
有网络: 添加源文件 → 选择语言 → 导出Excel
无网络: 人工翻译/校对Excel各语言列
有网络: 导入Excel → 自动更新翻译表 → 生成XML/RES
```

**Logo定制流程**:
```
进入Logo标签页 → 选择图片 → 输入公司名称(默认"LOGO")
→ 点击操作按钮 → 自动转换ICO/GIF → 创建快捷方式
```

### 4.5 状态反馈机制

| 状态 | 反馈方式 |
|------|----------|
| 翻译中 | 进度条百分比 + ScrolledText日志滚动 |
| API调用 | 日志显示当前文件名/语言 |
| 完成 | `messagebox.showinfo`弹窗通知 |
| 错误 | 日志错误信息 + `messagebox.showerror`弹窗 |
| Logo转换 | `indeterminate`进度条 + 状态文字 |
| 空闲 | 统计面板显示待处理数量 |

### 4.6 交互动画

- **悬停效果**: 导航项`<Enter>`/`<Leave>`切换颜色（无过渡动画）
- **标签页切换**: `grid_remove()`/`grid()`即时切换（无过渡动画）
- **进度更新**: `root.after(0, callback)`线程安全刷新
- **搜索高亮**: 匹配项背景色`#EDE7FF`即时反馈

---

## 五、工程架构说明

### 5.1 整体架构

```
翻译/
├── translation_gui.py              # 主应用(单体架构，~2800行)
│   └── TranslationApp              # 主类
│       ├── UI层                     # tkinter/ttk组件、样式定义
│       ├── 翻译逻辑                 # API调用、缓存、智能缩写
│       ├── XML解析                   # ElementTree + 正则回退
│       ├── Logo工具                  # 图片转换、快捷方式、图标缓存
│       └── 工具函数                  # Excel、日志、文件操作
├── res_packer_gui.py               # 独立RES打包工具(197行，独立运行)
├── translation_config.json         # 用户配置
├── translation_table.json          # 翻译记忆库(~235K行)
├── TranslationTool_v1.2.0.spec     # PyInstaller构建配置
├── TranslationTool_v1.1.0.spec     # 旧版构建配置
├── version_info.txt                # Windows版本信息
├── build/                          # PyInstaller构建中间产物
└── dist/
    └── TranslationTool_v1.2.0.exe  # 最终可执行文件
```

### 5.2 技术栈

| 层级 | 技术 | 版本/备注 |
|------|------|-----------|
| 语言 | Python | 3.12 |
| GUI | tkinter + ttk | 标准库，clam主题 |
| HTTP | requests | 每个翻译工作线程独立 Session，3 次退避重试；代理 TLS 握手异常时自动尝试直连 |
| Excel | openpyxl | Workbook/load_workbook |
| 图像 | Pillow (PIL) | ICO/GIF转换 |
| XML | xml.etree.ElementTree | 主解析器 |
| 压缩 | gzip / zlib | .res文件处理 |
| 并发 | threading + concurrent.futures | ThreadPoolExecutor |
| 打包 | PyInstaller | 单文件EXE，UPX压缩 |

### 5.3 并发模型

```
主线程 (Tk事件循环)
    │
    ├── ThreadPoolExecutor(max_workers=15)
    │       └── 并行API调用(翻译任务)
    │
    └── root.after(0, callback)
            └── 线程安全UI更新(进度、日志)
```

### 5.4 CLI模式参数

程序支持通过`argparse`进行无界面批量处理：

| 参数 | 短参数 | 类型 | 默认值 | 说明 |
|------|--------|------|--------|------|
| `--input` | `-i` | str | None | 输入文件或文件夹路径 |
| `--mode` | - | choices | `batch` | batch=文件夹, single=单文件 |
| `--output` | `-o` | str | 脚本目录 | 输出目录 |
| `--langs` | - | str | None | 逗号分隔语言缩写(如CHS,ENG) |
| `--pack` | - | flag | False | 处理后打包为.res |
| `--diskc` | - | str | None | DiskC根目录路径 |
| `--log` | - | str | None | 日志输出文件路径 |
| `--nogui` | - | flag | False | 无界面模式(仅命令行) |

### 5.5 构建配置 (PyInstaller)

```bash
# TranslationTool_v1.2.0.spec 配置
入口脚本: translation_gui.py
数据文件: translation_table.json -> .
          快速设定logo和快捷方式/create_shortcut.vbs
隐式导入: requests, urllib3
优化等级: -O2 (optimize=2)
UPX压缩: 启用
控制台: 关闭 (console=False)
版本信息: version_info.txt
输出: TranslationTool_v1.2.0.exe (单文件)
```

**已知问题**: `version_info.txt`中的`FileDescription`和`ProductName`仍为`AnalysisPlatform`（从其他项目复制），未更新为`TranslationTool`。

### 5.6 代码规范

- **编码**: UTF-8
- **注释**: 中文注释
- **命名**: snake_case(函数/变量)，PascalCase(类)
- **日志**: logging模块，支持控制台和文件输出
- **错误处理**: try-except + 日志记录 + messagebox提示

---

## 六、后端技术栈

### 6.1 核心依赖

| 依赖 | 用途 | 备注 |
|------|------|------|
| `requests` | HTTP客户端 | 每个翻译工作线程独立 Session，3 次退避重试 |
| `openpyxl` | Excel读写 | 翻译表导入导出 |
| `Pillow` (PIL) | 图像处理 | Logo转换为ICO/GIF |
| `xml.etree` | XML解析 | 主解析器 |
| `threading` | 多线程 | 并行翻译 |
| `concurrent.futures` | 线程池 | Provider 自适应并发（Google GTX 为 2，其余为 4） |
| `json` | 配置/数据 | 持久化存储 |
| `gzip` / `zlib` | 压缩 | .res文件处理 |
| `hashlib` / `hmac` / `uuid` | 签名 | 百度、有道、腾讯云、火山引擎与阿里云 API 认证 |
| `argparse` | 命令行 | CLI参数解析 |
| `ctypes` | Windows API | 管理员权限检测、图标缓存清理 |
| `subprocess` | 子进程 | 调用VBScript |
| `shutil` | 文件操作 | 文件复制/删除 |
| `logging` | 日志 | 结构化日志输出 |

### 6.2 翻译API集成

工具使用统一的 Provider 接口，配置页按照官方免费额度/付费服务优先顺序显示服务。所有密钥只保存到本机 `translation_config.json`，绝不会写入处理日志。

| Provider | 类型 | 认证方式 | 默认服务地址 |
|---|---|---|---|
| DeepL API Free | 官方免费额度 | Auth Key | `https://api-free.deepl.com/v2` |
| 百度翻译 | 官方免费额度/付费 | App ID + Secret Key | 百度通用文本翻译 API |
| 有道智云 | 官方试用/付费 | App Key + App Secret（v3 签名） | 有道文本翻译 API |
| 小牛翻译 | 官方免费额度/付费 | API Key | 小牛文本翻译 API |
| 腾讯云 TMT | 官方免费额度/付费 | SecretId + SecretKey（TC3-HMAC-SHA256），全局 4 次/秒限流 | `https://tmt.tencentcloudapi.com` |
| 火山引擎机器翻译 | 官方免费额度/付费 | AccessKey + Secret AccessKey（官方 Signature V4；固定 `cn-beijing/translate`） | `https://translate.volcengineapi.com` |
| 阿里云机器翻译 | 官方免费额度/付费 | AccessKey ID + Secret（ACS 签名） | 阿里云通用版 API |
| DeepL API Pro | 官方付费 | Auth Key | `https://api.deepl.com/v2` |
| Google GTX | 非官方测试 | 无 | Google 公共 GTX 端点 |

#### 批量文本翻译

对于提供**有序数组请求和响应**的服务，工具会按**目标语言**聚合待译文本、自动分片，并按原顺序写回翻译表：

| Provider | 请求字段 | 响应字段 | 客户端批次边界 |
|---|---|---|---|
| 腾讯云 TMT | `SourceTextList` | `TargetTextList` | 最多 200 条；每条最多 2,000 字符 |
| DeepL Free / Pro | `text[]` | `translations[]` | 最多 50 条；使用 120 KiB 安全阈值，低于官方 128 KiB 请求体上限 |
| 火山引擎 | `TextList` | `TranslationList` | 最多 16 条；总计最多 5,000 字符 |
| 阿里云 | `SourceText` 索引对象 | `TranslatedList` | 最多 50 条；单条最多 1,000 字符；总计最多 8,000 字符 |

- 每批只消耗一次 API 请求额度；批量失败时整批会在退避后重试。
- 超过服务批量限制的单条文本保持为空并记录原因，不会被截断。
- 腾讯云 TMT 的所有批次共享 **4 次/秒** 请求限流，低于当前账户提示的 5 次/秒限制。
- 百度、有道、小牛和 Google GTX 没有在本工具中采用可可靠逐条回填的官方数组响应，因此继续使用单条调用，避免多段译文错位。
- 火山引擎签名严格使用官方规则：`region=cn-beijing`、`service=translate`，签名请求体字节、`Host`、`Content-Type`、`X-Date` 和 `X-Content-Sha256` 与实际发送内容保持一致。

#### 配置与测试流程

1. 在主界面 API 下拉框选择服务；官方免费额度服务位于列表前方。
2. 点击“配置”，填写该服务的密钥、地域或服务地址。
3. 点击“测试连接”，确认服务、网络和密钥均可用后再点击“保存”。
4. 使用 CLI 时可通过 `--api <provider>` 选择服务，例如 `--api deepl_free`、`--api baidu` 或 `--api tencent`。
5. 在侧栏“设置选项”中选择启动时默认窗口化或默认窗口最大化；点击“保存并应用”后立即生效，并写入本机配置。

旧版仅包含 `api_type`、`baidu_app_id` 和 `baidu_secret_key` 的配置文件会自动迁移，仍可继续使用百度翻译。旧配置中若仍选择或包含已移除的服务，程序启动时会自动清理；如果它曾是当前服务，则会切换为 Google GTX（仅测试），请随后在 API 配置中选择所需的正式服务。

#### 失败处理

- 失败请求会记录服务名、目标语言和实际错误原因，但不会记录密钥。
- 如果系统 `HTTP_PROXY`/`HTTPS_PROXY` 代理在 TLS 握手阶段异常断开，程序会保留代理配置并自动使用不读取系统代理的直连 Session 重试；直连也失败时会同时保留两次错误原因，便于排查网络策略。
- 失败单元格保持为空，不会再将中文原文缓存或误写为译文；下次运行可切换 Provider 后重试。
- GUI 在生成 XML 前会提示未成功的翻译数量。CLI 检测到失败会以状态码 `3` 退出且不会生成 XML。
- Google GTX 仅用于小批量连通性测试；它不是官方生产 API，且并发限制为 2。其他 Provider 默认并发为 4，以减少限流和 TLS 连接中断。
- 腾讯云 TMT 除 4 个工作线程外，另有跨线程的全局 **4 次/秒** 请求限流，以低于账户提示的 5 次/秒硬上限；`RequestLimitExceeded` 会经过退避后重试。

### 6.3 数据持久化

| 文件 | 格式 | 大小 | 用途 |
|------|------|------|------|
| `translation_config.json` | JSON | ~100行 | 用户 Provider 配置 |
| `translation_table.json` | JSON | ~235,394行 | 翻译记忆库 |
| `*.xlsx` (导出) | Excel | 变化 | 离线人工翻译 |
| `translation_log.txt` | TXT | 变化 | 导出的日志 |

---

## 七、数据库架构（JSON文件持久化）

### 7.1 存储方案

采用JSON文件作为轻量级持久化方案，无需数据库依赖：
- **优势**: 零配置、便携性、人类可读
- **限制**: 不支持多个进程同时编辑；大数据量时加载性能下降。单进程保存使用临时文件原子替换，避免读到半写入文件。
- **适用**: 单用户桌面工具场景

### 7.2 文件结构

```
应用目录/
├── translation_config.json    # Provider 配置文件（大小随服务配置增长）
├── translation_table.json     # 翻译表(逐渐增长)
└── exports/                   # Excel导出目录(可选)
    └── translation_*.xlsx     # 导出的翻译文件
```

### 7.3 数据增长策略

| 记录数 | 预估JSON大小 | 加载时间 | 建议 |
|--------|-------------|----------|------|
| <10K | <5MB | <1s | 正常使用 |
| 10K-50K | 5-20MB | 1-3s | 可接受 |
| 50K-100K | 20-40MB | 3-5s | 考虑优化 |
| >100K | >40MB | >5s | 建议分库 |

**Excel导出文件大小**: 每条约100字节(原文+译文)，10K条记录约1MB。

---

## 八、数据字典

### 8.1 translation_config.json

```json
{
  "default_langs": ["ARA","CHT","ENG","ESP","FIN","FRA","GER","ITA","JPN","KOR","NLD","PLK","PTG","RUS","TRK","VIT"],
  "api_type": "google",
  "provider_configs": {
    "deepl_free": {
      "auth_key": "",
      "endpoint": "https://api-free.deepl.com/v2"
    },
    "baidu": {
      "app_id": "",
      "secret_key": ""
    }
  },
  "baidu_app_id": "",
  "baidu_secret_key": ""
}
```

| 字段 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `default_langs` | array[string] | 见上方 | 默认选中的语言代码列表(16种) |
| `api_type` | string | "google" | 当前服务：`deepl_free`、`baidu`、`youdao`、`niutrans`、`tencent`、`volcengine`、`aliyun`、`deepl_pro` 或 `google` |
| `provider_configs` | object | 见上方 | 各 Provider 的密钥、端点、地域及场景配置 |
| `baidu_app_id` | string | "" | 旧版百度 App ID 兼容字段，会同步到 `provider_configs.baidu.app_id` |
| `baidu_secret_key` | string | "" | 旧版百度密钥兼容字段，会同步到 `provider_configs.baidu.secret_key` |

**注意**: 代码中DEFAULT_LANGS常量定义11种语言(`['CAT','CHT','ESP','GER','ITA','KOR','PLK','PTG','RUS','TRK','VIT']`)，但运行时`load_config()`会用JSON文件中的16种覆盖。

### 8.2 translation_table.json

**结构**: 扁平字典，键为中文原文，值为语言翻译字典。

```json
{
  "非绝对式编码器": {
    "CAT": "",
    "CHT": "非絕對式編碼器",
    "CHS": "",
    "USA": "Incremental Encoder",
    "JPN": "インクメンタルエンコーダ",
    ...
  }
}
```

| 层级 | 类型 | 说明 |
|------|------|------|
| 顶层键 | string | 中文原文文本 |
| 二级键 | string | 3字母语言代码(127种) |
| 二级值 | string | 翻译结果，空字符串表示未翻译 |

### 8.3 语言代码映射 (LANG_MAP)

**总计127种语言变体**，格式: `{'3字母代码': {'code': 'BCP47代码', 'name': '语言名称'}}`

| 语系 | 数量 | 示例 |
|------|------|------|
| 阿拉伯语变体 | 16 | ARA(ar-SA), ARI(ar-IQ), ARE(ar-EG)... |
| 中文变体 | 5 | CHT(zh-TW), CHS(zh-CN), ZHH(zh-HK), ZHI(zh-SG), ZHM(zh-MO) |
| 英语变体 | 7 | USA(en-US), ENG(en-GB), ENA(en-AU)... |
| 德语变体 | 5 | GER(de-DE), DES(de-CH), DEA(de-AT)... |
| 西班牙语变体 | 19 | ESP(es-ES), ESM(es-MX), ESG(es-GT)... |
| 法语变体 | 6 | FRA(fr-FR), FRB(fr-BE), FRC(fr-CA)... |
| 意大利语变体 | 2 | ITA(it-IT), ITS(it-CH) |
| 其他语言 | 67 | JPN, KOR, RUS, POR, TUR, VIE, THA等 |

### 8.4 智能缩写规则

按语言设置最大翻译长度，超出则自动截断：

| 语言/语系 | 最大长度 | 缩写策略 |
|-----------|----------|----------|
| JPN (日语) | 22字符 | 移除助词 |
| KOR (韩语) | 22字符 | 音节计数压缩 |
| RUS (俄语) | 25字符 | 默认截断 |
| ESP及所有ES变体 (西班牙语) | 28字符 | 默认截断 |
| ITA/ITS (意大利语) | 30字符 | 默认截断 |
| PTG/PTB (葡萄牙语) | 30字符 | 默认截断 |
| FRA及所有FR变体 (法语) | 32字符 | 默认截断 |
| GER及所有DE变体 (德语) | 35字符 | 默认截断 |
| ARA及所有AR变体 (阿拉伯语) | 40字符 | 默认截断 |
| 默认 | 40字符 | 英文/拉丁语系使用停用词移除 |
| CHT/CHS/ZHH/ZHI/ZHM (中文) | **不缩写** | 完全跳过 |

### 8.5 Excel导入导出格式

**导出格式** (.xlsx):
```
| 原文 | USA - English: United States | JPN - Japanese: Japan | KOR - Korean: Korea | ... |
|------|------------------------------|----------------------|---------------------|-----|
| 启动 | Start                        | 開始                  | 시작                 | ... |
| 停止 | Stop                         | 停止                  | 중지                 | ... |
```

**导入规则**:
- 第一列必须为"原文"
- 列标题格式: `语言代码 - 语言名称`（用` - `分隔）
- 空单元格保持原值不变
- 仅更新有变化的翻译条目（增量更新）
- 导入后自动保存到translation_table.json

### 8.6 Logo输出文件

| 文件 | 格式 | 尺寸 | 路径 |
|------|------|------|------|
| `{公司名称}.ico` | ICO | 128x128 | `OpenCNC/Bin/Logo/` |
| `LoadingImage.gif` | GIF | 75x94 | `OpenCNC/Bin/Logo/` |

**支持的输入图片格式**: `*.png *.jpg *.jpeg *.bmp *.gif *.tiff *.ico`

---

## 九、技术路线图

### 9.1 短期优化 (v1.3)

- [ ] 模块化重构（拆分ui/core/data层）
- [ ] 添加单元测试覆盖
- [ ] 完善错误处理和重试机制
- [ ] Excel导入增加格式校验和错误提示
- [ ] Logo工具支持批量图片转换
- [ ] 修复`_res_converted`临时目录在翻译中断时未清理的问题（建议在app退出时注册清理钩子）

### 9.2 中期功能 (v2.0)

- [ ] 翻译记忆TMX格式支持
- [ ] 模糊匹配功能
- [ ] CNC术语库管理
- [ ] 翻译前后对比预览
- [ ] Excel导出增加翻译进度标记
- [ ] Logo工具支持更多图片格式、自定义尺寸输出
- [ ] 支持XLIFF标准格式
- [ ] BRAND_NAMES常量实际集成到翻译流程中

### 9.3 长期规划 (v3.0)

- [ ] 插件系统（自定义翻译引擎）
- [ ] 多平台支持（macOS/Linux）
- [ ] 云同步翻译记忆
- [ ] AI辅助翻译集成
- [ ] Logo模板库（预设CNC厂商模板）
- [ ] 批量快捷方式创建（多台电脑部署）

---

## 附录

### A. 依赖清单

```
requests>=2.31.0
openpyxl>=3.1.2
Pillow>=10.2.0
```

构建工具:
```
pyinstaller>=6.3.0
```

### B. CNC品牌厂商列表

代码中定义(`BRAND_NAMES`): 禾川、高创、松下、汇川、迪维迅、台达、雷赛、信捷、H系列

**注意**: 该常量在代码中定义但**未被实际引用**，品牌名称目前不会被跳过翻译。

### C. .res文件格式说明

- **格式**: GZIP压缩的Windows资源文件
- **结构**: 标准GZIP头 + FLG=0x00
- **解压**: `gzip.decompress()`直接解压
- **打包**: `zlib.compress(content, 9)` + 自定义GZIP头构造

### D. DiskC工作流说明

自动发现CNC软件目录结构：
1. 加载 `{root}/OpenCNC/Bin/Language/CHS.res` (解压)
2. 扫描 `{root}/OpenCnc Shared/OCRes/CHS/String/` 下所有XML
3. 加载 `{root}/OpenCNC/Bin/Plugin/Config/CHS.xml` (可选)
4. 输出: RES→`{root}/OpenCNC/Bin/Language/{LANG}`, XML→`{root}/OpenCnc Shared/OCRes/{LANG}/String/`

### E. 独立工具: res_packer_gui.py

独立运行的XML打包工具，与translation_gui.py完全独立（无import关系）。
- 功能: 选择XML文件 → `zlib.compress(content, 9)` → 输出`filename.xml.res`
- 界面: 800x600窗口，文件列表+进度条+日志

### F. 快捷方式创建

通过 `_generate_vbs_script()` 方法动态生成VBScript脚本，调用 `cscript //Nologo` 执行创建桌面快捷方式。脚本在运行时自动生成到 `快速设定logo和快捷方式/create_shortcut.vbs`，无需预先打包。

### G. 版本历史

| 版本 | 日期 | 主要变更 |
|------|------|----------|
| v1.0 | 2025-01 | 初始版本 |
| v1.1.0 | 2025-06 | 添加Logo转换和快捷方式创建功能 |
| v1.2.0 | 2025-07 | DiskC工作流、Excel导入导出 |
| v1.2.1 | 2026-07-06 | Bug修复: 添加`sys`导入; 修复`create_shortcut_only`中`script_dir`未定义; 修复`version_info.txt`项目名称错误; 修复`LANG_NATIVE_NAME`翻译错误(ESH/ARY); 删除重复代码; VBS脚本改为运行时动态生成 |
| v1.2.2 | 2026-09-22 | 维护：翻译任务仅处理当前源文件文本，避免误翻译历史翻译记忆；清空文件列表时完整重置 DiskC 工作流状态，避免普通文件误走 DiskC 输出与清理路径 |

### H. 已知问题

| 问题 | 位置 | 说明 |
|------|------|------|
| `_res_converted`临时目录未清理 | translation_gui.py:859,2143 | `prepare_file()`创建临时目录存放RES解压文件，`_cleanup_temp_files()`仅在翻译全部完成后调用。若翻译中断(异常/用户关闭)，该目录不会被删除，残留于应用目录 |
| `BRAND_NAMES`未使用 | translation_gui.py:206 | 常量定义了9个品牌名但未集成到翻译/缩写流程中 |

---

**文档版本**: v1.2.2
**最后更新**: 2026-09-22
**维护者**: SYNTEC开发团队

**文档结束**
