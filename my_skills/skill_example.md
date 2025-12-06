# 我的第一个 Skill

描述: 这是一个示例技能，说明如何写 skills 文件。

输入: 用户的自然语言询问。
输出: 一个结构化的 JSON 响应，包含 `name`, `description`, `run` 等字段。

示例:
```json
{
  "name": "greet",
  "description": "对用户打招呼",
  "run": {
    "input": "名字",
    "output": "Hello, <名字>!"
  }
}
```
