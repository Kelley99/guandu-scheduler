---
name: attribute-collector-lessons
description: |
  属性收集工具开发中的经验教训。当涉及OCR数据流、前后端字段同步、Vue响应式Proxy等问题时参考。
---

# 属性收集工具 - 经验教训

## 核心教训：前后端字段同步必须严格一致

### 问题根因

属性字段从6个扩展到12个（新增了用户不需要的字段），导致：
1. 后端OCR提取12个字段，多余字段返回0或异常值
2. 前端formData合并时被多余字段覆盖（Vue展开运算符 `...result.data` 会把所有字段合并进formData）
3. 前端UI只显示6个输入框，但formData实际包含12个字段，导致逻辑混乱
4. OCR识别成功但表单无法填充

### 规则

1. **任何字段变更必须全链路同步修改**：后端attr_labels → SQL INSERT/UPDATE → API返回 → 前端formData → HTML输入框 → computed计算 → resetForm
2. **不要"过度扩展"**：用户要什么字段就只处理什么字段，不要自作主张添加"可能有用"的字段
3. **OCR后端提取可以多，但返回给前端必须只返回需要的字段**：如果后端需要提取额外字段用于内部计算，在返回前过滤掉多余字段

## OCR调试规范

1. **必须先看原始OCR文本再写提取逻辑**：盲目写正则只会越改越错
2. **调试代码必须有生命周期**：添加调试按钮时，在commit message或TODO中标记"待清理"，修复后立即删除
3. **不要假设OCR文本格式**：用户提供的截图格式可能与预期不同，每张截图都要单独验证

## Vue注意事项（已记录在MEMORY.md，此处不再重复）

- Vue3 reactive Proxy对象传给fetch body时需用 `JSON.parse(JSON.stringify(obj))` 剥离
- Vue production模式静默崩溃：模板引用的变量必须在return中暴露
- Safari无法使用浏览器控制台，只能靠alert调试

## 数据库Schema注意事项

1. **ALTER TABLE添加的列追加在末尾**：不会插入到中间位置
2. **SELECT * 的列顺序 = schema顺序**：用 `PRAGMA table_info(table_name)` 获取动态列名映射，不要硬编码索引
3. **columns数组必须与DB schema一致**：`dict(zip(columns, row))` 如果顺序错位会导致整表数据映射错误
