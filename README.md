# 成都看雪山

基于数值预报与实况数据，计算成都西部四座雪山（西岭、四姑娘山、贡嘎、九顶山）的观山评分与天空形态模拟的 Flask 应用。

## 功能

- 观山评分：云量/能见度/湿度/气溶胶/洗尘/霾层/季节/日出窗口/天气系统多因子综合评分
- 天空形态模拟：路径云层剖面（11 云属识别 + 积雨云/雨层云跨层云塔），预报与历史回算均支持
- 实况观测：METAR 机场实况、风云四号卫星云图（西南裁剪 + 云型识别走廊叠加）、全国/四川雷达拼图
- 历史回顾：任意日期范围回算（ERA5 再分析），趋势图区分预报与回算，Excel 导出
- 地图选点联动：观测点拖拽 → 云图/趋势/评分联动刷新

## 本地运行

```bash
pip install -r requirements.txt
python chengdu_snow_mountain.py
# 打开 http://localhost:5000
```

## 部署

应用入口读取 `PORT` 环境变量（`os.getenv("PORT", "5000")`），自动适配各 PaaS 平台注入的端口，无需改代码。

### Zeabur（推荐 · 免费免绑卡 · 中文界面）

1. 打开 [zeabur.com](https://zeabur.com) 注册（支持 GitHub 登录 / 手机号，**无需信用卡**）
2. 新建项目 → **部署新服务 → GitHub** → 导入本仓库
3. 平台自动识别 Python 并构建（或选择 Dockerfile）
4. 部署完成后在服务设置里绑定域名，得到固定地址 `https://xxx.zeabur.app`

> 免费计划支持 1 个服务 + 1 个自定义域名，适合个人长期使用；国内访问速度快。

### Koyeb（备选 · 免费免绑卡）

1. 打开 [koyeb.com](https://koyeb.com) 注册（GitHub 登录，免费 Starter 计划**不需要信用卡**）
2. Create Web Service → 连接本 GitHub 仓库
3. 自动构建部署，得到固定地址 `https://xxx.koyeb.app`

> 免费实例 1 小时无流量会休眠，下次访问自动唤醒（首屏稍慢属正常）。

### Render（需要绑定银行卡验证）

仓库根目录已提供 `render.yaml`（Blueprint 一键部署），部署后会得到固定地址 `https://chengdu-snow-mountain.onrender.com`。**注意**：Render 新账号创建服务（含免费实例）要求先绑定银行卡做身份验证，无卡用户请改用上方 Zeabur / Koyeb。

### Railway / Fly.io / PythonAnywhere

入口适配各平台 `PORT` 环境变量，按平台文档创建 Web 服务即可；这些平台免费层多要求绑卡，请以官方最新政策为准。

## 数据源

- Open-Meteo 数值预报 / Archive（ERA5 再分析）/ Air Quality（CAMS 气溶胶）
- 中央气象台风云四号 B 星真彩色云图、全国/四川雷达拼图
- METAR 机场实况（AviationWeather）

> 注：所有数据源均为免费公开接口，无 API Key 要求。
