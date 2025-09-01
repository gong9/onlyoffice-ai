set -e
docker buildx build --platform linux/amd64 -t onlyoffice-server:v0.1 . 
docker save onlyoffice-server:v0.1 > onlyoffice-server-v0.1.tar
scp onlyoffice-server-v0.1.tar root@211.90.219.252:/home/baohui/onlyoffice-ai-main/onlyoffice-server
ssh root@211.90.219.252 "docker load < /home/baohui/onlyoffice-ai-main/onlyoffice-server/onlyoffice-server-v0.1.tar"
ssh root@211.90.219.252 "docker-compose -f /home/baohui/onlyoffice-ai-main/onlyoffice-server/docker-compose.yml up -d"