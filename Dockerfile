FROM python:3.10-slim

WORKDIR /app

# 系统依赖（requests 的 CA 证书等）
RUN apt-get update && apt-get install -y --no-install-recommends ca-certificates && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Hugging Face Spaces 固定监听 7860（应用读取 PORT 环境变量，自动适配）
EXPOSE 7860

CMD ["python", "chengdu_snow_mountain.py"]
