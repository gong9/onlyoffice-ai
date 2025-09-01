#!/bin/bash
set -e

# 构建镜像
docker buildx build --platform linux/amd64 --no-cache -t onlyoffice-server:v0.1 .

# 生成带时间戳的文件名
FILE_NAME="onlyoffice-server-$(date +"%Y%m%d-%H%M%S").tar"

# 保存镜像到本地
docker save onlyoffice-server:v0.1 > "$FILE_NAME"

# 上传到远程服务器
scp "$FILE_NAME" root@211.90.219.252:/home/baohui/onlyoffice-ai-main/onlyoffice-server/

# 在远程服务器加载镜像
ssh root@211.90.219.252 "docker load < /home/baohui/onlyoffice-ai-main/onlyoffice-server/$FILE_NAME"

# 启动容器
ssh root@211.90.219.252 "docker-compose -f /home/baohui/onlyoffice-ai-main/onlyoffice-server/docker-compose.yml up -d"
