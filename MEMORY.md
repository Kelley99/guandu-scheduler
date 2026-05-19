- **2026-04-18**：记忆系统启用

## 技术规范偏好

- Flask+Vue项目数据序列化规范：Vue ref/reactive对象直接发给后端会挂起，需先用JSON.parse(JSON.stringify(obj))剥离Proxy再发送
- Flask+Vue项目HTTP响应规范（两条经验教训）：
1. Vue ref/reactive对象直接发给后端会挂起，需先用JSON.parse(JSON.stringify(obj))剥离Proxy再发送
2. jsonify()返回chunked响应（无Content-Length），Tauri等HTTP客户端无法处理；修复：用make_response(json.dumps(data, ensure_ascii=False))手动设置Content-Length
3. axios从CDN加载在某些网络环境会被阻塞，导致请求挂起；优先用原生fetch替代axios
- Safari（ Kelley 的主浏览器）无法打开控制台，调试 JavaScript 网络请求只能靠 alert() 弹窗或页面内嵌文字，console.log 无效；Flask+Vue 项目调试时优先检查网络请求是否返回200、文件路径是否正确、Vue初始化是否依赖window.Vue对象
- Flask+Vue分隔符冲突最佳实践：Vue保持默认{{ }}分隔符，Jinja2模板中用{% raw %}...{% endraw %}包裹Vue模板区块；不要修改Vue分隔符（Vue 3配置方式不直观且易失效）
- macOS端口占用排查：`lsof -i :端口号` 查看监听PID，`pkill -f 进程名` 或 `kill PID` 终止进程
- Safari调试JavaScript网络请求只能靠alert()弹窗或页面内嵌文字，console.log无效；Flask+Vue项目调试时优先检查网络请求是否返回200、文件路径是否正确、Vue初始化是否依赖window.Vue对象
- macOS端口占用排查：lsof -i :端口号查看监听PID，pkill -f进程名或kill PID终止进程
- Vue 3 production模式对未定义变量静默崩溃（不报错，整个app挂掉），只有errorHandler能捕获
- Python文本替换必须精确匹配目标文本，最好先grep确认原文再写替换逻辑
- Vue单文件中所有ref()变量名必须唯一，重复声明不会报错但会导致整个应用崩溃
- Vue重命名变量后需在return语句中同步暴露，否则Vue production模式静默崩溃
- Cloudflared临时隧道URL不稳定，断开后需重新启动获取新URL
- Vue 3 Composition API：ref/reactive声明后必须同步加入setup()的return，否则模板访问为undefined，return遗漏是静默崩溃的常见原因
- edit工具替换时oldText边界可能吞掉相邻行（尤其是函数定义行紧接try/except时），替换后必须python3 -c "import app"验证语法
- 5002 端口是 334 服专属，禁止操作，必须保持运行

## 用户身份与偏好

- 用户叫 Kelley
