# TOOLS.md - Local Notes

Skills define _how_ tools work. This file is for _your_ specifics — the stuff that's unique to your setup.

## What Goes Here

Things like:

- Camera names and locations
- SSH hosts and aliases
- Preferred voices for TTS
- Speaker/room names
- Device nicknames
- Anything environment-specific

## Examples

```markdown
### Cameras

- living-room → Main area, 180° wide angle
- front-door → Entrance, motion-triggered

### SSH

- home-server → 192.168.1.100, user: admin

### TTS

- Preferred voice: "Nova" (warm, slightly British)
- Default speaker: Kitchen HomePod
```

## Why Separate?

Skills are shared. Your setup is yours. Keeping them apart means you can update skills without losing your notes, and share skills without leaking your infrastructure.

---

Add whatever helps you do your job. This is your cheat sheet.

---

## 🔑 API 服务备忘

### 腾讯云 OCR（已开通，正在使用）

- **服务**: 通用文字识别 (GeneralBasicOCR)
- **SecretId**: `AKID_REMOVED_FROM_HISTORY`
- **SecretKey**: `SECRET_KEY_REMOVED_FROM_HISTORY`
- **区域**: `ap-guangzhou`
- **SDK**: `tencentcloud-sdk-python`
- **使用项目**: attribute_collector (5002/450), gc_manager (5001)
- **额度**: 免费额度已用完，后付费模式
- **接口**: `GeneralBasicOCR` — 通用印刷体识别

### 百度 OCR（已开通）

- **API Key**: `OS2wp5hlvvJwJIYg5ayRA8kt`
- **Secret Key**: `VkbZhazXFLM3hswEtikSIiKGUOEpG1Ts`
- **接口**: 通用文字识别（标准版），免费1000次/月
- **用途**: 备选 OCR，腾讯云OCR失败时兜底
- **Token获取**: `https://aip.baidubce.com/oauth/2.0/token?grant_type=client_credentials&client_id={API_KEY}&client_secret={SECRET_KEY}`
- **调用地址**: `https://aip.baidubce.com/rest/2.0/ocr/v1/general_basic?access_token={token}`

---

## 🔐 项目密码

| 项目 | 用途 | 密码 |
|------|------|------|
| gc_manager | 考勤密码验证 | `lingxiao2026` |
| gc_manager | 版本回滚密码 | `334dengni` |
| attribute_collector (334) | 编辑/删除 | `zhuofansoso` |
| attribute_collector (450) | 编辑/删除 | `ZHUO123FANS` |

---

## 🌐 服务端口

| 端口 | 服务 | 路径 | 备注 |
|------|------|------|------|
| 5001 | gc_manager | `/Users/kelley/.qclaw/workspace-agent-29b6e205/` | 官渡+考勤 |
| 5002 | attribute_collector (334) | `/Users/kelley/Desktop/kelley_work/attribute_collector/` | 禁止操作 |
| 450 | attribute_collector (450) | `/Users/kelley/Desktop/kelley_work/attribute_collector_450/` | 隧道不稳定 |
