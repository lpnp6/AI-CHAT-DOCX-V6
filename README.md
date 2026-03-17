# AI-CHAT-DOCX-V6

一个最小可运行的 DOCX 自动填写原型。

它会直接读取 `.docx` 压缩包里的 `word/document.xml`，把 XML、用户填写要求、失败反馈和已锁定 XPath 一起发给 LLM，让模型返回结构化写入指令，再由本地执行器把结果写回新的 DOCX 文件。

## 当前能力

- 读取输入 DOCX，并保留压缩包中的其他文件
- 只针对 `word/document.xml` 做文本修改
- 让 LLM 返回结构化 JSON 指令
- 执行 `set_text` 类型的 XPath 写入
- 对部分常见“节点存在但 `<w:t>` 缺失”的场景做本地修复
- 支持多轮重试
- 使用 `locked_xpaths` 防止后续轮次覆盖已成功写入的内容
- 输出逐轮日志，便于排查失败原因

## 工作流

1. 从输入 `.docx` 中读取 `word/document.xml`
2. 调用 LLM，传入：
   - 用户 prompt
   - 当前 `document.xml`
   - 上一轮失败指令
   - 已成功写入并锁定的 XPath
3. 要求模型只返回如下 JSON：

```json
{
  "instructions": [
    {
      "type": "set_text",
      "xpath": "./w:body/w:p[1]/w:r[1]/w:t[1]",
      "text": "replacement text"
    }
  ]
}
```

4. 本地执行器按 XPath 写入文本
5. 失败项回传给 LLM 继续修复，直到成功或达到最大轮数
6. 写出新的 DOCX 和日志文件

## 目录结构

- `docx_mvp/`: 核心实现
- `input/`: 输入 DOCX
- `output/`: 输出 DOCX 和日志
- `tests/`: 单元测试
- `docx.md`: 调研说明文档

输出路径规则：

- 输出文档：`output/<name>/<name>`
- 输出日志：`output/<name>/<name>.log`

如果不传 `--output-name`，默认使用输入文件名。

## 环境要求

- Python `3.11+`
- `OPENAI_API_KEY`

可选环境变量：

- `OPENAI_MODEL`
- `OPENAI_BASE_URL`

默认模型是 `gpt-5.2`，默认接口地址是 `https://api.openai.com/v1`。

## 安装

```bash
python -m venv .venv
source .venv/bin/activate
pip install -e .
```

## 环境变量示例

```bash
export OPENAI_API_KEY=your_api_key
export OPENAI_MODEL=gpt-5.2
export OPENAI_BASE_URL=https://api.openai.com/v1
```

## 用法

把源文件放到 `input/` 下，然后执行：

```bash
docx-edit-mvp demo.docx "把公司名称替换成小米科技有限责任公司"
```

指定输出名和最大轮数：

```bash
docx-edit-mvp demo.docx "填写雷军简历" -o result.docx --max-rounds 5
```

命令行参数现在都是必填，不再支持交互式输入。

## 日志内容

每次运行都会生成日志，包含：

- 输入文件路径
- 输出文件路径
- 原始 prompt
- 每一轮的模型原始输出
- 每一轮的解析后指令
- 每一轮的失败列表
- 最终剩余失败项（如果有）

## 重试与覆盖保护

多轮修复的核心约束如下：

- 首轮成功写入的 XPath 会加入 `locked_xpaths`
- 后续轮次会把这些锁定位置传回模型
- 执行层也会拒绝覆盖这些锁定位置

这样可以避免模型在修复失败项时，把前一轮已经写对的字段重新覆盖掉。

## 当前限制

- 当前只支持一种指令类型：`set_text`
- 修改范围只在 `word/document.xml`
- 不会自动扩展模板结构
- 如果模板本身没有足够槽位，模型只能在现有节点中尝试修复
- 目前没有样式级修改能力，也不处理图片、页眉页脚、`styles.xml`、`theme` 等其他部件

## 测试

运行单元测试：

```bash
python -m unittest discover -s tests -v
```

## 核心文件

- [`docx_mvp/__main__.py`](/Users/dip/Developer/study/AI-CHAT-DOCX-V6/docx_mvp/__main__.py): CLI 入口、输入输出路径、日志初始化
- [`docx_mvp/workflow.py`](/Users/dip/Developer/study/AI-CHAT-DOCX-V6/docx_mvp/workflow.py): 多轮编排、失败回传、XPath 锁定
- [`docx_mvp/llm.py`](/Users/dip/Developer/study/AI-CHAT-DOCX-V6/docx_mvp/llm.py): LLM 请求和指令解析
- [`docx_mvp/executor.py`](/Users/dip/Developer/study/AI-CHAT-DOCX-V6/docx_mvp/executor.py): XPath 执行、本地修复、失败收集
- [`docx_mvp/package.py`](/Users/dip/Developer/study/AI-CHAT-DOCX-V6/docx_mvp/package.py): DOCX 压缩包读写
