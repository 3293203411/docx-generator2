# 制式文档生成API

一个基于 FastAPI 的 Word 模板占位符替换服务，可用于替代扣子(Coze)工作流中的"制式文件生成助手"插件。

## 功能特性

- ✅ 支持 `{{占位符}}` 格式的模板替换
- ✅ 自动处理 Word 内部 run 拆分问题
- ✅ 支持段落、表格中的占位符替换
- ✅ 保留原始文本格式
- ✅ 兼容扣子工作流 HTTP 请求节点
- ✅ 自动清理过期文件

## 快速开始

### 1. 本地运行

```bash
# 进入项目目录
cd 制式文档生成API

# 创建虚拟环境（推荐）
python -m venv venv
source venv/bin/activate  # Linux/Mac
# 或 venv\Scripts\activate  # Windows

# 安装依赖
pip install -r requirements.txt

# 运行服务
python main.py
```

服务启动后访问 `http://localhost:8000`

### 2. 测试 API

访问 API 文档页面：`http://localhost:8000/docs`

或使用 curl 测试：

```bash
curl -X POST "http://localhost:8000/generate" \
  -H "Content-Type: application/json" \
  -d '{
    "template_file_url": "https://example.com/template.docx",
    "text_keys": ["公司名称", "项目名称"],
    "text_values": ["示例公司", "示例项目"],
    "filename": "output.docx"
  }'
```

## 接口说明

### 生成文档接口

**POST** `/generate`

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| template_file_url | string | ✅ | 模板文件下载URL |
| text_keys | array/string | ✅ | 占位符名称数组 |
| text_values | array/string | ✅ | 替换内容数组 |
| filename | string | 否 | 输出文件名 |

**返回格式：**

```json
{
  "success": true,
  "message": "文档生成成功",
  "full_download_url": "http://xxx:8000/download/output.docx",
  "filename": "output.docx",
  "replaced_count": 2
}
```

### 下载文件接口

**GET** `/download/{filename}`

直接下载生成的文件。

### 健康检查

**GET** `/health`

返回服务健康状态。

### 清理过期文件

**DELETE** `/cleanup?max_age_hours=24`

清理超过指定时间的生成文件。

## 扣子工作流配置

### 步骤1：添加HTTP请求节点

在扣子工作流中，添加一个 **HTTP 请求** 节点。

### 步骤2：配置请求

| 配置项 | 值 |
|--------|-----|
| **请求方法** | POST |
| **请求URL** | `https://your-app.onrender.com/generate` (部署后的URL) |
| **请求体类型** | JSON |
| **请求体内容** | 参考下方JSON模板 |

### 步骤3：请求体JSON模板

```json
{
  "template_file_url": "{{模板文件URL}}",
  "text_keys": ["占位符1", "占位符2"],
  "text_values": ["替换值1", "替换值2"],
  "filename": "输出文件名.docx"
}
```

### 步骤4：提取响应数据

在后续节点中，使用以下表达式获取下载链接：

```
{{HTTP节点名称.full_download_url}}
```

## 部署指南

### 方案一：Render.com（推荐，免费）

#### 1. 准备代码

将项目代码推送到 GitHub 仓库。

#### 2. 创建 Render 账号

访问 [render.com](https://render.com) 并注册账号。

#### 3. 创建 Web Service

1. 点击 **New +** → **Web Service**
2. 连接你的 GitHub 仓库
3. 配置以下选项：

| 配置项 | 值 |
|--------|-----|
| **Name** | docx-generator（自定义） |
| **Region** | Singapore（离中国大陆近） |
| **Branch** | main |
| **Root Directory** | 制式文档生成API |
| **Runtime** | Python |
| **Build Command** | `pip install -r requirements.txt` |
| **Start Command** | `uvicorn main:app --host 0.0.0.0 --port $PORT` |
| **Plan** | Free |

#### 4. 设置环境变量

点击 **Environment** 添加：

| 变量名 | 值 | 说明 |
|--------|-----|------|
| `BASE_URL` | `https://your-app.onrender.com` | 替换为你的实际URL |
| `TEMP_DIR` | `/tmp/docx_generated` | 临时文件目录 |

#### 5. 部署

点击 **Create Web Service** 开始部署。

> ⚠️ **注意**：Render 免费套餐每月有750小时额度，服务会在30分钟无流量后休眠。

### 方案二：Railway（免费额度更多）

#### 1. 准备代码

将项目代码推送到 GitHub 仓库。

#### 2. 创建 Railway 项目

1. 访问 [railway.app](https://railway.app) 并注册
2. 点击 **New Project** → **Deploy from GitHub repo**
3. 选择你的仓库

#### 3. 配置部署

1. Railway 会自动检测为 Python 项目
2. 设置启动命令为：`uvicorn main:app --host 0.0.0.0 --port $PORT`
3. 添加环境变量：
   - `BASE_URL`: 你的 Railway 域名（如 `https://docx-generator.up.railway.app`）

#### 4. 获取域名

部署成功后，Railway 会提供随机域名，如：`docx-generator.up.railway.app`

### 方案三：Sealos（国内推荐）

[Sealos](https://www.sealos.io) 是面向企业的云原生容器平台，个人使用免费。

#### 1. 创建应用

1. 进入 Sealos 桌面
2. 点击 **应用管理** → **新建应用**

#### 2. 配置应用

| 配置项 | 值 |
|--------|-----|
| **应用名称** | docx-generator |
| **镜像** | `python:3.11-slim` |
| **CPU** | 0.5核 |
| **内存** | 512MB |
| **启动命令** | `bash -c "pip install fastapi uvicorn python-docx requests && uvicorn main:app --host 0.0.0.0 --port 8000"` |

#### 3. 暴露端口

添加端口映射：`8000:8000`

#### 4. 配置环境变量

- `BASE_URL`: Sealos 提供的公网地址

### 方案四：Docker 部署

#### 1. 创建 Dockerfile

```dockerfile
FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY main.py .

EXPOSE 8000

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
```

#### 2. 构建和运行

```bash
# 构建镜像
docker build -t docx-generator .

# 运行容器
docker run -d -p 8000:8000 \
  -e BASE_URL=http://localhost:8000 \
  --name docx-generator \
  docx-generator
```

## 模板制作规范

### 占位符格式

使用双花括号包裹占位符名称：

```
{{公司名称}}
{{项目名称}}
{{联系人}}
```

占位符内支持空格：

```
{{ 公司名称 }}
{{ 项目名称 }}
```

### 模板示例

```
合同编号：{{合同编号}}

甲方：{{甲方名称}}
乙方：{{乙方名称}}

项目名称：{{项目名称}}
项目金额：{{项目金额}}

签订日期：{{签订日期}}
```

### 注意事项

1. **区分大小写**：`{{公司名称}}` 和 `{{公司名称}}` 是不同的占位符
2. **避免嵌套**：不要在占位符内包含其他占位符
3. **特殊字符**：如果替换值包含 `{{}}`，可能会影响替换结果
4. **Word格式**：确保模板保存为 `.docx` 格式（非 `.doc`）

## 常见问题

### Q: 部署后下载链接返回404？

检查 `BASE_URL` 环境变量是否设置正确，应为完整的 HTTPS 地址。

### Q: 替换失败，文档内容未变化？

1. 确认模板中占位符格式正确（双花括号）
2. 检查 `text_keys` 和 `text_values` 数量是否一致
3. 确保占位符名称与模板中完全匹配

### Q: 文件下载失败？

Render 免费版服务会休眠，首次请求可能需要等待服务唤醒（10-30秒）。

### Q: 支持表格中的替换吗？

支持。表格单元格内的占位符也会被正确替换。

## 项目结构

```
制式文档生成API/
├── main.py              # FastAPI 主程序
├── requirements.txt     # 依赖列表
├── README.md           # 说明文档
└── .gitignore          # Git忽略文件（可选）
```

## 许可证

MIT License
