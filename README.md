# Sales Prospect and Translation

一个自动化的销售前景搜索和翻译工具，能够搜索指定关键词的相关网站，提取内容并进行智能翻译，最终生成多种格式的报告。

## 快速开始

### 第一步：下载项目文件

#### 方法一：从GitHub页面下载（推荐）

1. 访问 [GitHub仓库](https://github.com/stevenspage/sales_prospect_translator)
2. 点击绿色的 "Code" 按钮
3. 选择 "Download ZIP" 下载完整项目
4. 解压后找到 `sales_prospect_translator.py` 文件

#### 方法二：使用Git克隆（适合开发者）

```bash
git clone https://github.com/stevenspage/sales_prospect_translator.git
cd sales_prospect_translator
```

**重要提示**：确保下载了 `sales_prospect_translator.py` 文件，这是程序的核心文件。

## 功能特性

- 🔍 **智能搜索**: 使用Google搜索API获取相关网站（默认300个结果）
- 🌐 **多语言支持**: 自动检测搜索语言，支持多种语言搜索
- 📄 **内容提取**: 使用多种库组合提取网页正文内容
- 🤖 **智能翻译**: 集成智谱翻译API，支持批量并发翻译
- 📊 **多格式输出**: 支持Word、HTML、Excel等多种格式
- 📧 **邮件发送**: 支持通过QQ邮箱发送报告

## 快速开始

### 第二步：安装依赖

```bash
pip install googlesearch-python newspaper3k fpdf python-docx trafilatura langdetect pandas openpyxl
```

或者使用requirements.txt：

```bash
pip install -r requirements.txt
```

### 第三步：配置搜索关键词

打开 `sales_prospect_translator.py` 文件，找到第5行，修改搜索关键词：

```python
SEARCH_KEYWORD = "your_search_keyword_here"  # 修改为您的搜索关键词
```

**示例：**
- `"importador argentino de máquinas de café"` - 阿根廷咖啡机进口商
- `"distribuidor mexicano de equipos médicos"` - 墨西哥医疗设备分销商
- `"importador brasileño de maquinaria industrial"` - 巴西工业机械进口商

### 第四步：配置翻译API（可选）

如果您需要翻译功能，需要配置智谱翻译API：

1. 访问 [智谱AI开放平台](https://open.bigmodel.cn/) 注册账号
2. 获取API密钥
3. 在 `sales_prospect_translator.py` 文件中找到第34行，填入API密钥：

```python
TRANSLATION_CONFIG = {
    'api_key': 'your_api_key_here',  # 替换为您的API密钥
    'enable_translation': True,      # 启用翻译功能
    # ... 其他配置
}
```

**注意：** 如果不配置API密钥，程序将跳过翻译步骤，只进行搜索和内容提取。

### 第五步：配置邮件发送（可选）

如果您需要邮件发送功能，需要配置QQ邮箱：

1. 在QQ邮箱设置中开启SMTP服务
2. 获取授权码（不是QQ密码）
3. 在 `sales_prospect_translator.py` 文件中找到第24-27行，填入邮箱信息：

```python
EMAIL_CONFIG = {
    'sender_email': 'your_email@qq.com',        # 您的QQ邮箱
    'sender_password': 'your_auth_code',        # QQ邮箱授权码
    'recipient_email': 'recipient@example.com', # 收件人邮箱
    'enable_email': True                        # 启用邮件发送
}
```

**注意：** 如果不配置邮箱，程序将跳过邮件发送步骤。

### 第六步：调整搜索参数（可选）

您可以根据需要调整搜索参数：

```python
# 搜索结果数量（默认300个）
SEARCH_CONFIG = {
    'num_results': 300,  # 可以调整为10-500之间的数字
    'sleep_interval_min': 1,  # 最小请求间隔（秒）
    'sleep_interval_max': 2   # 最大请求间隔（秒）
}

# 正文截取长度（默认500字符）
TEXT_TRUNCATE_CONFIG = {
    'max_chars': 500,  # 可以调整为100-2000之间的数字
    'enable_truncate': True
}
```

### 第七步：运行程序

配置完成后，运行程序：

```bash
python sales_prospect_translator.py
```

程序将自动执行以下步骤：
1. 🔍 搜索指定关键词的相关网站
2. 📄 提取每个网站的标题、摘要和正文内容
3. 🤖 使用AI翻译所有内容（如果启用）
4. 📊 生成Word、HTML、Excel格式的报告
5. 📧 发送邮件报告（如果启用）

## 输出文件

程序运行完成后，会在当前目录生成以下文件：

- `{搜索关键词}.docx` - **Word格式报告**，包含超链接，适合阅读和打印
- `{搜索关键词}.html` - **HTML格式报告**，适合在浏览器中查看
- `{搜索关键词}.xlsx` - **Excel格式报告**，适合数据分析和处理

**文件命名示例：**
- 搜索关键词：`"importador argentino de máquinas de café"`
- 生成文件：`importador_argentino_de_máquinas_de_café.docx`

## 详细配置说明

### 搜索配置
```python
SEARCH_KEYWORD = "your_search_keyword"  # 搜索关键词
SEARCH_CONFIG = {
    'num_results': 300,           # 搜索结果数量（10-500）
    'sleep_interval_min': 1,      # 最小请求间隔（秒）
    'sleep_interval_max': 2       # 最大请求间隔（秒）
}
```

### 翻译配置
```python
TRANSLATION_CONFIG = {
    'api_key': 'your_api_key',           # 智谱翻译API密钥
    'enable_translation': True,          # 是否启用翻译功能
    'translate_content': True,           # 是否翻译正文内容
    'translate_summary': True,           # 是否翻译摘要
    'max_workers': 10,                   # 并发翻译线程数（建议3-10）
    'request_delay': 0.5                 # 请求间隔（秒）
}
```

### 邮件配置
```python
EMAIL_CONFIG = {
    'sender_email': 'your_email@qq.com',     # 发送方QQ邮箱
    'sender_password': 'your_auth_code',     # QQ邮箱授权码
    'recipient_email': 'recipient@example.com', # 收件人邮箱
    'enable_email': False                    # 是否启用邮件发送
}
```

### 内容截取配置
```python
TEXT_TRUNCATE_CONFIG = {
    'max_chars': 500,        # 最大字符数（100-2000）
    'enable_truncate': True  # 是否启用正文截取
}
```

## 使用场景示例

### 场景1：寻找咖啡机进口商
```python
SEARCH_KEYWORD = "importador argentino de máquinas de café"
# 搜索阿根廷的咖啡机进口商
```

### 场景2：寻找医疗设备分销商
```python
SEARCH_KEYWORD = "distribuidor mexicano de equipos médicos"
# 搜索墨西哥的医疗设备分销商
```

### 场景3：寻找工业机械进口商
```python
SEARCH_KEYWORD = "importador brasileño de maquinaria industrial"
# 搜索巴西的工业机械进口商
```

## 注意事项

### ⚠️ 重要提醒
1. **API密钥安全**: 请勿将包含真实API密钥的代码上传到公共仓库
2. **QQ邮箱授权码**: 需要在QQ邮箱设置中开启SMTP服务并获取授权码（不是QQ密码）
3. **API限制**: 智谱翻译API可能有请求频率限制，建议适当调整并发数
4. **网络连接**: 程序需要稳定的网络连接进行搜索和翻译

### 💡 使用建议
1. **首次使用**: 建议先设置较少的搜索结果数量（如10-50个）进行测试
2. **翻译功能**: 如果不需要翻译，可以设置 `enable_translation: False` 提高运行速度
3. **邮件功能**: 如果不需要邮件发送，可以设置 `enable_email: False`
4. **搜索关键词**: 使用具体、准确的关键词可以获得更好的搜索结果

## 故障排除

### 常见问题

#### 1. 搜索失败
**问题**: 程序无法获取搜索结果
**解决方案**:
- 检查网络连接是否正常
- 确认搜索关键词格式正确
- 尝试减少搜索结果数量
- 检查是否被Google限制访问

#### 2. 翻译失败
**问题**: 翻译功能无法正常工作
**解决方案**:
- 检查API密钥是否正确
- 确认API密钥是否还有余额
- 检查网络连接
- 尝试减少并发翻译线程数

#### 3. 邮件发送失败
**问题**: 无法发送邮件
**解决方案**:
- 检查QQ邮箱授权码是否正确（不是QQ密码）
- 确认已开启SMTP服务
- 检查收件人邮箱地址格式
- 确认网络连接正常

#### 4. 内容提取失败
**问题**: 无法提取网页内容
**解决方案**:
- 某些网站可能有反爬虫机制，这是正常现象
- 程序会自动尝试多种提取方法
- 可以忽略部分失败的结果

### 依赖安装问题

如果遇到依赖安装问题，可以尝试：

```bash
# 升级pip
pip install --upgrade pip

# 安装依赖
pip install -r requirements.txt

# 如果仍有问题，尝试逐个安装
pip install googlesearch-python
pip install newspaper3k
pip install fpdf2
pip install python-docx
pip install trafilatura
pip install langdetect
pip install pandas
pip install openpyxl
pip install requests
pip install beautifulsoup4
```

### 性能优化建议

1. **减少搜索结果数量**: 如果不需要太多结果，可以减少 `num_results` 值
2. **关闭翻译功能**: 如果不需要翻译，设置 `enable_translation: False`
3. **调整并发数**: 根据网络情况调整 `max_workers` 值
4. **增加请求间隔**: 如果遇到限流，可以增加 `request_delay` 值

## 更新日志

### v1.0.0
- 初始版本发布
- 支持Google搜索和内容提取
- 集成智谱翻译API
- 支持多种输出格式（Word、HTML、Excel）
- 支持邮件发送功能

## 许可证

本项目采用MIT许可证。

## 贡献

欢迎提交Issue和Pull Request来改进这个项目！

### 如何贡献
1. Fork 本项目
2. 创建您的特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交您的修改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 打开一个 Pull Request

## 联系方式

如果您有任何问题或建议，请通过以下方式联系：

- 提交 Issue
- 发送邮件
- 创建 Pull Request

---

**祝您使用愉快！** 🎉
