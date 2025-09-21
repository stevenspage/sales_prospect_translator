# Sales Prospector and Translator

一个自动化的客户信息搜索和翻译工具，能够搜索指定关键词的相关网站，提取内容并进行智能翻译，最终生成多种格式的报告。

## 功能特性


- 🚀 **零配置启动**: googlesearch模式无需申请API，开箱即用
- 🌐 **指定搜索结果数量**: 可以指定获取前1000个搜索结果（以Google实际搜索数量为准）
- 🌐 **搜索特定国家**: 可以指定搜索美国、英国、法国、巴西、阿根廷等国家
- 🌐 **搜索特定语言**: 可以指定搜索中文、英语、法语、西班牙语、法语等多种语言
- 🌐 **多搜索关键词**: 可以指定多个搜索关键词一并搜索
- 📄 **内容提取**: 支持提取正文内容
- 🤖 **智能翻译**: 集成智谱翻译API，支持批量并发翻译
- 📊 **多格式输出**: 支持Word、HTML、Excel等多种格式
- 📧 **邮件发送**: 支持通过QQ邮箱发送报告

## 使用场景示例

### 场景1：寻找阿根廷运动相机经销商
```python
SEARCH_KEYWORD = "distribuidor cámaras de acción Argentina"
# 搜索阿根廷的运动相机进口商
```

### 场景2：寻找墨西哥医疗设备分销商
```python
SEARCH_KEYWORD = "distribuidor mexicano de equipos médicos"
# 搜索墨西哥的医疗设备分销商
```

### 场景3：寻找巴西工业机械进口商
```python
SEARCH_KEYWORD = "importador brasileño de maquinaria industrial"
# 搜索巴西的工业机械进口商
```

---

## 快速开始

> 🚀 **5分钟快速上手** - 从下载到运行，只需5个简单步骤

---

### ✨ 第1️⃣步：下载项目文件

#### 🎯 从GitHub页面下载（推荐）

1. 🟢 点击上方绿色的 "Code" 按钮（需要在电脑上打开）
2. 📦 选择 "Download ZIP" 下载完整项目
3. 📁 解压后找到 `sales_prospector_translator.py` 文件

> 💡 **提示**: 确保下载了 `sales_prospector_translator.py` 文件，这是程序的核心文件

---

### ✨ 第2️⃣步：下载Python

<details>
<summary>🐍 点击查看 Python 安装步骤</summary>

如果您的系统还没有安装Python，请先下载并安装Python：

1. 🌐 访问 [Python官网](https://www.python.org/downloads/)
2. ⬇️ 下载最新版本的Python（推荐Python 3.8或更高版本）
3. 🔧 运行安装程序，**务必勾选"Add Python to PATH"**
4. ✅ 其他设置请勾选默认即可，安装结束后无需再管Python

> ⚠️ **重要**: 必须勾选"Add Python to PATH"，否则无法在命令行中使用Python

</details>

---

### ✨ 第3️⃣步：安装依赖（必须，否则程序无法运行）

#### 🚀 一键安装（推荐）

1. 🖱️ 在解压的项目目录中找到 `install.bat` 文件
2. 🖱️ **双击 `install.bat` 文件**，系统会自动安装所有必需的库

> 💡 **提示**: 如果双击后没有反应，请右键点击 `install.bat` 选择"以管理员身份运行"

#### 🔧 手动安装（备选方案）

如果自动安装失败，可以手动运行以下命令：

```bash
pip install pandas openpyxl requests beautifulsoup4 chardet newspaper3k fpdf python-docx trafilatura langdetect google-search-results
```
---

### ✨ 第4️⃣步：配置搜索方式（无需申请API）

#### 📋 配置步骤

1. 📊 打开并填写 `search_config.xlsx`（在本项目中下载）
2. 🚀 在"配置模板"工作表：填写搜索接口，**推荐选择googlesearch模式**（无需申请API，开箱即用）
3. 📝 在"配置模板"工作表：填写搜索关键词等
4. ✅ 其他配置保持默认即可


#### 🔥 googlesearch模式（推荐 - 无需API）

> 🎉 **最佳选择**：googlesearch模式无需申请任何API，配置简单，立即可用！

---

### ✨ 第5️⃣步：双击运行py脚本

#### 🎯 方法一：双击运行（推荐）

1. 📁 找到 `sales_prospector_translator.py` 文件
2. 🖱️ 双击该文件即可运行程序
3. ⚙️ 程序会自动读取 `search_config.xlsx` 中的配置（如果没有会自动创建）

> 🚀 **最简单的方式**：双击即可运行，无需命令行操作

#### 🔧 方法二：在VS CODE 或者 Cursor 中运行（适合开发者）


```bash
在代码编辑器中运行
```

#### 🎬 程序运行流程

1. 🔍 **搜索阶段** - 搜索指定关键词的相关网站
2. 📄 **提取阶段** - 提取每个网站的标题、摘要和正文内容
3. 🤖 **翻译阶段** - 使用AI翻译所有内容（如果启用）
4. 📊 **生成阶段** - 生成Word、HTML、Excel格式的报告
5. 📧 **发送阶段** - 发送邮件报告（如果启用）

---

### ✨ 第6️⃣步：查看运行结果

程序运行完成后，会在当前目录生成以下文件：

- `{搜索关键词}_{时间戳}.docx` - **Word格式报告**，包含超链接，适合阅读和打印
- `{搜索关键词}_{时间戳}.html` - **HTML格式报告**，适合在浏览器中查看
- `{搜索关键词}_{时间戳}.xlsx` - **Excel格式报告**，适合数据分析和处理

**文件命名示例：**
- 搜索关键词：`"importador argentino de máquinas de café"`
- 生成文件：`importador_argentino_de_máquinas_de_café_20250117_1430.xlsx`

---


### ✨ 第7️⃣步：配置翻译API（可选）


如果您需要翻译功能，需要配置智谱翻译API：

1. 访问 [智谱AI开放平台](https://open.bigmodel.cn/) 注册账号
2. 获取API密钥
3. 在 Excel 的"配置模板"工作表中填写 `translation_api_key` 并将"是否启用翻译"设为 True。

> ⚠️ **注意**: 如果不配置API密钥，程序将跳过翻译步骤，只进行搜索和内容提取。


---

### ✨ 第8️⃣步：配置邮件发送（可选）

如果您需要邮件发送功能，需要配置QQ邮箱：

1. 获取授权码（不是邮箱密码）
   - 详细教程：[QQ邮箱获取授权码教程](https://service.mail.qq.com/detail/0/75)
   - 其他主流邮箱如Gmail，Outlook均支持通过授权码发送邮件（具体请询问AI）
3. 在 Excel 的“配置模板”工作表里填写 `sender_email`、`sender_auth_code`、`recipient_email`，并将“是否启用邮件”设为 True。

**注意：** 如果不配置邮箱，程序将跳过邮件发送步骤。

---

## 🔧 其他搜索方案（需要申请API）

如果您需要更多高级功能或遇到googlesearch模式的问题，可以考虑以下API方案：

### 方案 1：🔥 SERP API

<details>
<summary>📋 点击查看 SERP API 申请步骤</summary>

1. 🌐 访问 [SerpApi官网](https://serpapi.com/) 注册账号
2. 🔑 登录后进入 [Dashboard](https://serpapi.com/manage-api-key)
3. 📋 复制 API Key
4. ⚙️ 在 search_config.xlsx 中"搜索接口"选择 `serp_api`
5. 💾 在 search_config.xlsx 中"API密钥"填入刚才获取的 API KEY

**特点：**
- 免费额度：2500次/月
- 搜索结果无限制
- 申请流程简单

</details>

### 方案 2：🔧 Google Custom Search API

<details>
<summary>📋 点击查看 Google Custom Search API 申请步骤</summary>

1. 📖 **详细教程**: [Google Custom Search API 配置指南](https://yb.tencent.com/s/Rbc5eDI2GCTH)
2. 🔧 按照教程获取 `API_KEY` 和 `GOOGLE_SEARCH_ENGINE_ID`
3. 📝 在 Excel 配置文件中填入相应的值

**特点：**
- 免费额度：100次/天
- 官方接口，稳定性高
- 每关键词最多100个结果

</details>

### 搜索方案对比表

| 特性 | 🚀 googlesearch（推荐） | 🔥 SERP API | 🔧 Google Custom Search API |
|------|-------------------------|-------------|------------------------------|
| **申请流程** | ✅ 无需申请 | ⭐ 非常简单 | ⭐⭐⭐ 稍复杂 |
| **免费额度** | ✅ 无限制 | 2500次/月 | 100次/天 |
| **搜索结果限制** | ✅ 无限制 | 无限制 | 每关键词最多100个 |
| **配置复杂度** | ✅ 零配置 | 只需API Key | 需要API Key + 搜索引擎ID |
| **缺点** | ⚠️ 频率太快可能被Google封锁 | ❌ 无限制 | ❌ 无限制 |
| **推荐程度** | 🌟🌟🌟🌟🌟 | 🌟🌟🌟🌟 | 🌟🌟🌟 |

---

## 许可证

本项目采用MIT许可证。

---

## 贡献

欢迎提交Issue和Pull Request来改进这个项目！

### 如何贡献
1. Fork 本项目
2. 创建您的特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交您的修改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 打开一个 Pull Request

---

## 联系方式

如果您有任何问题或建议，请通过以下方式联系：

- 提交 Issue
- 发送邮件
- 创建 Pull Request

---

**祝您使用愉快！** 🎉
