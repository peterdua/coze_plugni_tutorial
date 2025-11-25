# 本地 Excel→JSON Coze 插件完整实战教程

> 目标：在本地跑一个 FastAPI 服务，将 Excel 转成 JSON / QA 列表，并通过 **OpenAPI 文档** 导入到本地 Coze Studio，成为可以在智能体里调用的插件。

---

## 目录

1. [整体架构与数据流](#整体架构与数据流)  
2. [环境准备](#环境准备)  
3. [实现本地 Excel→JSON 服务](#实现本地-exceljson-服务)  
   - [项目结构](#项目结构)  
   - [`main.py` 代码说明](#mainpy-代码说明)  
4. [本地运行与自测 API](#本地运行与自测-api)  
5. [编写 OpenAPI 文档](#编写-openapi-文档)  
   - [完整 `openapi_excel_json.yaml`](#完整-openapi_excel_jsonyaml)  
   - [Coze 对 OpenAPI 的特殊限制](#coze-对-openapi-的特殊限制)  
6. [在 Coze Studio 中创建本地插件](#在-coze-studio-中创建本地插件)  
   - [选择正确的 Server URL](#选择正确的-server-url)  
   - [导入 OpenAPI 文档](#导入-openapi-文档)  
7. [在 Coze 中调试插件](#在-coze-中调试插件)  
   - [健康检查工具 `/health`](#健康检查工具-health)  
   - [原始转换 `/convert`](#原始转换-convert)  
   - [QA 扁平化 `/convert_qa`](#qa-扁平化-convert_qa)  
8. [在智能体里如何使用返回数据](#在智能体里如何使用返回数据)  
9. [常见坑与排查清单](#常见坑与排查清单)  
10. [总结](#总结)

---

## 整体架构与数据流

- **本地 FastAPI 服务**（宿主机跑）  
  - 端口：`8001`  
  - 功能：
    - 接受 Excel 文件（`file_url` 或 `file_base64`）
    - 解析为 JSON（全部表 / 指定表）
    - 将每个学员的「题目 + 答案 + 得分」扁平化成 QA 列表

- **Coze Studio（Docker 内）**
  - 插件类型：OpenAPI 插件  
  - 插件通过我们提供的 `openapi_excel_json.yaml` 理解：
    - 有哪些 API（/convert, /convert_qa, /health）
    - 输入参数是什么
    - 输出字段是什么类型（为适配 Coze，统一用字符串封装 JSON）

- **调用路径示意**

1. Coze 插件调用 `GET http://host.docker.internal:8001/convert_qa?...`  
2. FastAPI 读取 Excel 并生成 QA 列表  
3. 返回 JSON：

   ```json
   {
     "items": "[{...}, {...}, ...]"  // 字符串内部是 JSON 数组
   }
   ```

4. 智能体在工具结果中解析 `items` 字符串，再做后续分析 / 生成反馈。

---

## 环境准备

- 操作系统：macOS
- 已运行的 Coze Studio Docker 环境（`docker ps` 能看到 `coze-web`, `coze-server` 等容器）
- Python 3.10+（建议）
- 虚拟环境 & 依赖（在 `excel_json_service` 目录下）：
  - `fastapi`
  - `uvicorn`
  - `httpx`
  - `openpyxl`
  - `pydantic`

> 本教程基于目录：  
> `/Users/duanyangfan/Downloads/reports/cozes/plugiins/excel_json_service`

---

## 实现本地 Excel→JSON 服务

### 项目结构

推荐结构如下（你已经基本是这样）：

```text
excel_json_service/
  main.py
  openapi_excel_json.yaml
  requirements.txt
  excel_base64.txt          # 调试用，可选
  测试试卷.xlsx              # 示例 Excel
  venv/                     # 虚拟环境
```

### `main.py` 代码说明

核心点：

- 使用 FastAPI 暴露三个接口：
  - `GET /health`：健康检查
  - `GET /convert`：返回 Excel→JSON 的原始结构（字符串包装）
  - `GET /convert_qa`：把每个学员的题目+答案+得分扁平化为数组（字符串包装）
- 为了兼容 Coze 的类型校验，**返回模型里的复杂字段（`sheets`、`items`）都定义为 `str`**，内部再用 JSON 字符串承载真正结构。

主要数据模型：

```python
class ConvertRequest(BaseModel):
    file_url: Optional[str] = None
    file_base64: Optional[str] = None
    sheet_name: Optional[str] = None
    header_row: bool = True

class ConvertResponse(BaseModel):
    sheets: str  # JSON 字符串

class ConvertQAResponse(BaseModel):
    items: str   # JSON 字符串
```

核心工具函数：

- `_download_file(url)`：用 `httpx` 下载 Excel 二进制
- `_load_excel(content)`：用 `openpyxl.load_workbook` 加载
- `_sheet_to_rows(ws, header_row)`：把一个 sheet 转成 `List[Dict[col_name, value]]`
- `_build_result(...)`：根据 `file_url` / `file_base64` 解析出 `Dict[sheet_name, List[row_dict]]`
- `_build_qa_items(result)`：把上面的结果「按人+题目+答案+得分」扁平化成数组

三个 HTTP 端点：

- `/health`：简单返回 `{"status": "ok"}`  
- `/convert`：
  - 入参：`file_url` / `file_base64` / `sheet_name` / `header_row`（Query）
  - 内部调用 `_build_result`
  - 返回：

    ```json
    {
      "sheets": "{\"Sheet1\": [...], \"Sheet2\": [...]}"`
    }
    ```

- `/convert_qa`：
  - 入参同上
  - 内部调用 `_build_result` + `_build_qa_items`
  - 返回：

    ```json
    {
      "items": "[{sheet, name, field, question, answer, score}, ...]"
    }
    ```

---

## 本地运行与自测 API

### 启动服务

```bash
cd /Users/duanyangfan/Downloads/reports/cozes/plugiins/excel_json_service

# 激活虚拟环境（根据你的 venv 路径）
source venv/bin/activate

uvicorn main:app --host 0.0.0.0 --port 8001 --reload
```

浏览器访问：

- `http://127.0.0.1:8001/health` → `{"status":"ok"}`

### 用 base64 测本地 Excel

1. 生成 `excel_base64.txt`：

   ```bash
   base64 "测试试卷.xlsx" | tr -d '\n' > excel_base64.txt
   ```

2. 在 Python 里测试 `/convert_qa`：

   ```bash
   python3 - << 'PY'
   import base64, json
   from pathlib import Path
   import httpx

   path = Path("测试试卷.xlsx")
   b64 = base64.b64encode(path.read_bytes()).decode("utf-8")

   resp = httpx.get(
       "http://127.0.0.1:8001/convert_qa",
       params={"file_base64": b64, "header_row": True},
       timeout=60.0,
   )
   print("HTTP status:", resp.status_code)
   print("Response JSON:", json.dumps(resp.json(), ensure_ascii=False, indent=2))
   PY
   ```

你应该能看到类似的结构：

```json
{
  "items": "[{\"sheet\":\"Sheet1\",\"name\":\"张雨晗\",...}]"
}
```

---

## 编写 OpenAPI 文档

### 完整 `openapi_excel_json.yaml`

> 路径：`excel_json_service/openapi_excel_json.yaml`  
> 已适配 Coze 的各种限制，可直接导入。

```yaml
openapi: 3.0.1
info:
  title: Excel to JSON Service
  version: "1.0.0"
  description: |
    Convert Excel files (via URL or base64) to JSON. Designed to be used as a tool/plugin in Coze.

servers:
  - url: http://host.docker.internal:8001
    description: Local Excel to JSON service

paths:
  /health:
    get:
      summary: Health check
      responses:
        "200":
          description: Service is healthy
          content:
            application/json:
              schema:
                type: object
                properties:
                  status:
                    type: string
                    example: ok

  /convert:
    get:
      summary: Convert Excel to JSON
      description: |
        Convert an Excel file to JSON via query parameters.

        You can either provide a public `file_url` or a `file_base64` string.

        - If `sheet_name` is provided, only that sheet will be converted.
        - If `sheet_name` is omitted, all sheets will be converted.
        - If `header_row` is true, the first row is treated as header.
      parameters:
        - in: query
          name: file_url
          description: Public URL of the Excel file.
          schema:
            type: string
        - in: query
          name: file_base64
          description: Base64 encoded Excel content.
          schema:
            type: string
        - in: query
          name: sheet_name
          description: Optional specific sheet name. If omitted, all sheets are converted.
          schema:
            type: string
        - in: query
          name: header_row
          description: Whether to treat the first row as header.
          schema:
            type: boolean
      responses:
        "200":
          description: Conversion result
          content:
            application/json:
              schema:
                type: object
                properties:
                  sheets:
                    type: string

  /convert_qa:
    get:
      summary: Convert Excel to QA items
      description: |
        Flatten Excel sheets into a list of question-answer items by student.

        Each item includes sheet name, student name, field key, question text, answer text, and score.
      parameters:
        - in: query
          name: file_url
          description: Public URL of the Excel file.
          schema:
            type: string
        - in: query
          name: file_base64
          description: Base64 encoded Excel content.
          schema:
            type: string
        - in: query
          name: sheet_name
          description: Optional specific sheet name. If omitted, all sheets are converted.
          schema:
            type: string
        - in: query
          name: header_row
          description: Whether to treat the first row as header.
          schema:
            type: boolean
      responses:
        "200":
          description: Flattened QA items
          content:
            application/json:
              schema:
                type: object
                properties:
                  items:
                    type: string
```

---

### Coze 对 OpenAPI 的特殊限制

根据我们踩坑总结，**要兼容 Coze 插件导入，需要遵守：**

- **Server**
  - `servers` 中必须是 **1 个** URL
  - URL 必须带 `http://` 或 `https://`，且包含 host（例如 `http://host.docker.internal:8001`）

- **Responses**
  - 每个 operation 的 `responses`：
    - 最好只写 `"200"` 一种状态（可以有 `"default"` 但不必要）
    - 不能写 `"400"`, `"500"` 等其它状态，否则报：  
      `response only supports '200' status`
  - `200` 的 `content`：
    - 只能有一种 MIME：`application/json`
    - schema 顶层必须是 `type: object`

- **RequestBody**
  - 如果使用 `requestBody`，它的 Schema 必须是 `type: object`  
    否则会有：`request body only supports 'object' type`
  - 本教程为了简单，全部改成 Query 参数，不再用 `requestBody`

- **格式 `format`**
  - `format: uri` 等 OpenAPI 标准格式，在 Coze 里会被当作“内部 AssistType”，  
    如果不在它支持列表中会报：  
    `the format 'uri' of field 'file_url' is invalid`
  - 所以我们把 `file_url` 统一改成 **纯 `type: string`，不写 `format`**

- **响应字段类型**
  - Coze 会根据响应 schema 递归裁剪字段，只保留声明过的，并且会对类型做严格校验。  
  - 为了避免被裁剪 / 类型不匹配，我们把复杂字段 `sheets`、`items` 统一声明为 `type: string`，内部再用 JSON 承载真实结构。  

---

## 在 Coze Studio 中创建本地插件

### 选择正确的 Server URL

关键点：**Coze 插件代码跑在 Docker 容器内**，而你的 Excel 服务跑在宿主机上。

- 容器内部如果访问 `http://localhost:8001`，指的是容器本身，不是你的 Mac。
- 在 macOS + Docker Desktop 环境下，**`host.docker.internal` 是“容器眼中的宿主机地址”**。

所以在 OpenAPI 里我们配置：

```yaml
servers:
  - url: http://host.docker.internal:8001
```

---

### 导入 OpenAPI 文档

在 Coze Web UI 中：

1. 打开插件 / 开发者中心（本地 Coze Studio）。
2. 新建一个插件（第三方 API / HTTP 插件）。
3. 在插件管理页找到「导入 OpenAPI / API 文档」入口。
4. 选择本地文件：  
   `/Users/duanyangfan/Downloads/reports/cozes/plugiins/excel_json_service/openapi_excel_json.yaml`
5. 导入成功后，你应该能看到 3 个 API：
   - `/health`
   - `/convert`
   - `/convert_qa`

如果导入过程中遇到 `invalid plugin openapi3 document ...` 之类错误，大概率是因为：

- 你导入的是旧版 YAML（含有 `400` 响应、`format: uri` 或复杂 schema）；  
- 确认当前 YAML 内容与本教程里的一致即可。

---

## 在 Coze 中调试插件

### 健康检查工具 `/health`

在插件的 API 列表里选择对应工具（通常名为 `Health check`），点击调试：

- 无需参数
- 应该返回：

  ```json
  { "status": "ok" }
  ```

若失败，检查：

- 本地 `uvicorn` 是否在跑
- `servers.url` 是否是 `http://host.docker.internal:8001`

---

### 原始转换 `/convert`

此工具对应 `GET /convert`，入参：

- `file_url`（string, Query）
- `file_base64`（string, Query）
- `sheet_name`（string, Query, 可选）
- `header_row`（boolean, Query, 可选）

调用方式建议：

- 二选一填：`file_url` 或 `file_base64`（至少一个）
- 调试 `file_base64` 流程示例：
  - `file_url`: 空
  - `file_base64`: 粘贴 `excel_base64.txt` 的整行内容
  - `sheet_name`: 空
  - `header_row`: true

返回类似：

```json
{
  "sheets": "{\"Sheet1\":[{...},{...}]}"
}
```

---

### QA 扁平化 `/convert_qa`

此工具对应 `GET /convert_qa`，入参同 `/convert`。

用同一个 base64 调试：

- `file_url`: 空
- `file_base64`: `excel_base64.txt` 的整行
- `sheet_name`: 空
- `header_row`: true

返回类似：

```json
{
  "items": "[{\"sheet\":\"Sheet1\",\"name\":\"张雨晗\",\"field\":\"gap_fill1\",...}]"
}
```

---

## 在智能体里如何使用返回数据

在 Coze 智能体中，你可以这样指导模型使用工具结果（提示词思路）：

- **使用 `/convert`：**

  - 插件返回：

    ```json
    { "sheets": "<JSON 字符串>" }
    ```

  - 让模型步骤：
    - 先把 `sheets` 字符串解析成 JSON 对象；
    - 然后按 `Sheet1`、`Sheet2` 遍历各行各列进行分析。

- **使用 `/convert_qa`：**

  - 插件返回：

    ```json
    { "items": "<JSON 数组字符串>" }
    ```

  - 让模型步骤：
    - 把 `items` 字段当作 JSON 数组解析；
    - 每个元素都有 `sheet`, `name`, `field`, `question`, `answer`, `score`；
    - 可以按 `name` 分组，为每个学员生成点评，或者按 `field` 分析某类题型。

示例提示（伪英文）：

> 调用工具后，你会得到一个字段 `items`，它是一个 JSON 字符串。  
> 请先把 `items` 解析为数组，每个元素包含：sheet, name, field, question, answer, score。  
> 然后按姓名聚合这些元素，为每个学生生成一段详细的反馈。

---

## 常见坑与排查清单

导入 / 调试过程中，如果遇到以下报错，可以对照处理：

- **`response only supports '200' status`**
  - 确保每个路径的 `responses` 里只有 `"200"`（可有 `"default"`），不要有 `"400"` / `"500"`。

- **`the format 'uri' of field 'file_url' is invalid`**
  - 把所有参数里的 `format: uri` 去掉，只留 `type: string`。

- **`media type 'application/json' not found in response`**
  - 确认 `responses."200".content` 下只包含 `application/json`，名称拼写准确。

- **`request body only supports 'object' type` / `request body schema is required`**
  - 避免使用 `requestBody`；统一用 `parameters` + Query。

- **插件能创建但点进去 500 / 无法查看**
  - 通常是上述 OpenAPI 校验失败（`get_plugin_apis` 报错），检查日志中的 `error: ...` 对照修正 YAML。

- **调试时 `{"sheets": {}}` 或 `{"items": "{}"}`**
  - 如果本地测试（不用 Coze）可正常解析：
    - 检查是不是 Coze 在根据 schema 裁剪 / 校验时丢掉了内容；
    - 使用本教程中的「把复杂字段声明为 `string`，内部再用 JSON 承载」方案。

---

## 总结

通过本教程，你已经完成了：

- 在本地实现一个健壮的 Excel→JSON / QA 服务；
- 深入理解 Coze 对 OpenAPI 插件的各种限制与校验；
- 成功在本地 Coze Studio 中导入插件，并通过 `/convert` 与 `/convert_qa` 工具，把 Excel 试卷结构化为：
  - 按表 / 行 / 列的原始 JSON；
  - 按「姓名 + 题目 + 答案 + 得分」扁平化的 QA 列表。

后续你可以在这个基础上：

- 再增加新的接口，比如：
  - 按姓名聚合成「每个学员一份完整报告」的数据结构；
  - 统计每题正确率、平均得分等；
- 或者接入你已有的讲评生成逻辑，让 Coze 智能体一键生成试卷讲评 / 学员学习报告。

---

## 面向小白：一步步在本地 Coze 创建插件（教学版）

  这一小节用**聊天教学**的口吻，再把整个链路捋一遍，适合第一次接触「HTTP 接口 + OpenAPI + Coze 插件」的小伙伴。

### 0. 先说几个通俗的概念

  - **本地服务（FastAPI）**：
    可以把它想象成一个「会干活的小工人」，只要你给他指令（HTTP 请求），他就会帮你把 Excel 文件读出来，变成 JSON。

  - **OpenAPI 文档（`openapi_excel_json.yaml`）**：
    这是一本「说明书」，上面写着：
    - 小工人住在哪里（服务器地址 / 端口）
    - 他会做哪几件事（路径 `/health`、`/convert`、`/convert_qa`）
    - 每件事需要哪些材料（输入参数）
    - 干完活会给你什么样的结果（响应结构）

  - **Coze 插件**：
    可以理解为：Coze 里的智能体不会自己访问 HTTP 接口，它需要一份「说明书」，告诉它：
    - 这个接口该怎么调用
    - 输入叫什么名、输出叫什么名
    这份说明书就是我们导入的 OpenAPI 文档，导入之后就变成了一个「插件」。

  - **为什么要用 `host.docker.internal`？**
    - Coze 是跑在 Docker 里的，相当于住在一个「隔离的房间」。
    - 你的 FastAPI 服务是跑在宿主机（Mac）上的。
    - `host.docker.internal` 可以理解为「房间里的人看向房间外面的一扇窗户」，通过这扇窗户就能访问你本机的服务。

  理解到这里，就可以开始动手了。

### 1. 在本地启动 Excel 服务（小工人上岗）

  1. 打开终端，进到服务目录：

     ```bash
     cd /Users/duanyangfan/Downloads/reports/cozes/plugiins/excel_json_service
     ```

  2.（可选）创建并激活虚拟环境，安装依赖（如果你还没做）：

     ```bash
     # 创建虚拟环境
     python3 -m venv venv

     # 激活
     source venv/bin/activate

     # 安装依赖
     pip install -r requirements.txt
     ```

  3. 启动 FastAPI 服务：

     ```bash
     uvicorn main:app --host 0.0.0.0 --port 8001 --reload
     ```

     看到类似输出：

     ```text
     Uvicorn running on http://0.0.0.0:8001
     Application startup complete.
     ```

  4. 在浏览器里访问：

     ```text
     http://127.0.0.1:8001/health
     ```

     如果看到：

     ```json
     {"status": "ok"}
     ```

     说明小工人已经在本机正常工作了。

### 2. 准备好「说明书」——OpenAPI 文档

  在这个项目里，我们已经帮你写好了一份 OpenAPI 文档：

  - 文件路径：`openapi_excel_json.yaml`
  - 里面写清楚了：
    - 服务器地址：`http://host.docker.internal:8001`
    - 三个接口：
      - `/health`：健康检查
      - `/convert`：把 Excel 转成原始 JSON 结构
      - `/convert_qa`：把「人 + 题目 + 答案 + 得分」扁平化成列表
    - 每个接口的输入参数（都在 query 里）和输出字段（`sheets` 或 `items`，类型都是字符串，里面再放 JSON）。

  > 如果你只是跟着教程走，不需要改这份 YAML；只有当你的服务端口或者路径变了，才需要同步修改 `servers.url` 或 `paths` 里的内容。

### 3. 在 Coze 本地界面中创建插件

  下面假设你的 Coze Studio Web 已经跑在 `http://localhost:8888`。

  1. 打开浏览器，访问：`http://localhost:8888`，进入 Coze 界面。
  2. 在左侧找到「插件」（或者英文 `Plugins`）菜单，进入插件管理页面。
  3. 点击「新建插件」：
     - 给插件起个名字，比如：`Excel to JSON Service`
     - 描述可以写：`本地 Excel 转 JSON/QA 的工具`
     - 鉴权方式选择「无」或 `None`（我们这个服务没有额外鉴权）。
  4. 找到「从文档导入」或「导入 OpenAPI / API 文档」按钮：
     - 选择本地文件：`openapi_excel_json.yaml`
     - 提交导入。

  如果一切顺利，Coze 会告诉你导入成功，然后在插件详情页中，你能看到 3 个工具：

  - 一个是健康检查 `/health`
  - 一个是原始转换 `/convert`
  - 一个是 QA 扁平化 `/convert_qa`

### 4. 在 Coze 里测试插件（以当前 Excel 插件为例）

#### 4.1 准备一个 Excel 文件

  你已经有一个 `测试试卷.xlsx`，里面有：

  - 第一行：题目列（`gap_fill1`、`biref_aws1`、`open_ques1` 等）
  - 后面每一行：每个学员的答案和得分（`姓名`、每题答案、`客观题得分`）。

#### 4.2 用 base64 方式传 Excel（简单稳定）

  1. 在项目目录下，用命令把 Excel 转成一行的 base64 文本：

     ```bash
     cd /Users/duanyangfan/Downloads/reports/cozes/plugiins/excel_json_service
     base64 "测试试卷.xlsx" | tr -d '\n' > excel_base64.txt
     ```

  2. 打开 `excel_base64.txt`，`全选 + 复制` 里面那一整行。

#### 4.3 在 Coze 里调试 `/convert_qa`

  1. 回到 Coze 的插件详情页，找到 `Convert Excel to QA items` 这个工具（即 `/convert_qa`）。
  2. 点击「调试」或「测试」按钮。
  3. 在参数输入区：
     - `file_url`：留空
     - `file_base64`：粘贴刚才复制的整段 base64 字符串
     - `sheet_name`：留空（表示所有 sheet 都处理）
     - `header_row`：设为 `true`（因为第一行是题目）
  4. 点击发送 / 调试。

  如果你的服务在终端有日志，会看到一条 GET `/convert_qa?...` 的请求；在 Coze 调试界面，会看到类似结果：

  ```json
  {
    "items": "[{\"sheet\": \"Sheet1\", \"name\": \"张雨晗\", \"field\": \"gap_fill1\", ...}]"
  }
  ```

  这就说明：

  - 插件已经能正常调用你本地的 FastAPI 服务；
  - Excel 里的题目+答案+分数已经被成功扁平化成一个列表。

### 5. 让智能体真正用起来

  上面只是「在插件里调试」。要让某个 Bot 能用这个插件，一般需要再做两步：

  1. 在 Bot 的配置里，找到「工具 / 插件」一栏：
     - 勾选你刚刚创建的 `Excel to JSON Service` 插件。
  2. 在系统 Prompt（系统指令）里给模型一点「使用说明」，比如：

     > 当你需要分析 Excel 试卷时，请调用 `Convert Excel to QA items` 工具。  
     > 工具返回的 `items` 字段是一个 JSON 字符串，请先将其解析成数组，然后按 `name` 分组，为每个学生总结他们的答题情况和得分。

  这样，当用户在对话里说「帮我分析这次考试试卷的情况」时，模型就知道：

  - 先调用插件，拿到 `items`；
  - 再解析 `items`，按人汇总，输出反馈。

  > 小结：你可以把这个流程理解为：
  > - FastAPI：干体力活，负责「读 Excel + 转结构」。
  > - OpenAPI：写说明书，解释「怎么找这个人干活」。
  > - Coze 插件：把说明书交给智能体，让它学会「什么时候、怎么叫这个人来帮忙」。

  只要这个链路打通，你之后换成别的 Excel、别的接口，流程都是类似的。
