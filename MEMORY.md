- **2026-04-18**：记忆系统启用

## 技术规范偏好

1. Vue ref/reactive对象直接发给后端会挂起，需先用JSON.parse(JSON.stringify(obj))剥离Proxy再发送
2. jsonify()返回chunked响应（无Content-Length），Tauri等HTTP客户端无法处理；修复：用make_response(json.dumps(data, ensure_ascii=False))手动设置Content-Length
3. axios从CDN加载在某些网络环境会被阻塞，导致请求挂起；优先用原生fetch替代axios
- Flask+Vue分隔符冲突最佳实践：Vue保持默认{{ }}分隔符，Jinja2模板中用{% raw %}...{% endraw %}包裹Vue模板区块；不要修改Vue分隔符（Vue 3配置方式不直观且易失效）
- Safari调试JavaScript网络请求只能靠alert()弹窗或页面内嵌文字，console.log无效；Flask+Vue项目调试时优先检查网络请求是否返回200、文件路径是否正确、Vue初始化是否依赖window.Vue对象
- Python文本替换必须精确匹配目标文本，最好先grep确认原文再写替换逻辑
- edit工具替换时oldText边界可能吞掉相邻行（尤其是函数定义行紧接try/except时），替换后必须python3 -c "import app"验证语法

## 用户身份与偏好

- 用户叫 Kelley
- 已开通的API/服务必须主动记录到备忘，不要等用户提醒
