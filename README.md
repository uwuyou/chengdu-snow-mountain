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

### Hugging Face Spaces（免费 · 无需绑卡）

1. 注册 [huggingface.co](https://huggingface.co)（邮箱即可，**不要求信用卡**）
2. 右上角头像 → **New Space**，配置：
   - Space name：`chengdu-snow-mountain`
   - License：随意（如 MIT）
   - **SDK 选择 Docker**（关键，不要选 Gradio/Streamlit）
   - Hardware：CPU basic（免费 2 vCPU / 16GB）
   - Visibility：Public
3. 创建后按页面提示连接仓库：**git push 到该 Space**（或网页直接上传文件）
4. 等待 2-5 分钟构建，访问 `https://你的用户名-chengdu-snow-mountain.hf.space`

应用通过 `PORT` 环境变量自动适配 Space 的 7860 端口，无需改代码。免费额度闲置 48 小时后休眠，下次访问自动唤醒。

### Render（免费额度，可能需要手机验证）

1. 将本仓库导入 Render（New + → Web Service → 连接 GitHub 仓库）
2. 配置：
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `python chengdu_snow_mountain.py`
   - 环境变量 `PORT` 会自动注入

### Railway / Fly.io / PythonAnywhere

应用入口读取 `PORT` 环境变量（`os.getenv("PORT", "5000")`），适配各 PaaS 平台，按平台文档创建 Web 服务即可。

## 数据源

- Open-Meteo 数值预报 / Archive（ERA5 再分析）/ Air Quality（CAMS 气溶胶）
- 中央气象台风云四号 B 星真彩色云图、全国/四川雷达拼图
- METAR 机场实况（AviationWeather）

> 注：所有数据源均为免费公开接口，无 API Key 要求。
