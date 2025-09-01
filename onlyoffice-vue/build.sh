#!/bin/bash
set -e
 



# 上传到远程服务器
scp -r "dist/" root@211.90.219.252:/home/baohui/onlyoffice-ai-main/nginx

# 在远程服务器加载镜像
ssh root@211.90.219.252 "rsync -a --delete /home/baohui/onlyoffice-ai-main/nginx/dist/ /home/baohui/onlyoffice-ai-main/nginx/frontend/"



# 启动容器
ssh root@211.90.219.252 "docker restart onlyoffice-frontend"