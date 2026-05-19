# 部署指南 — spe-avatar.com

## 服务器配置建议

| 项目 | 建议 |
|------|------|
| CPU / 内存 | **2 核 1GB 可以跑**，但偏紧；已默认 `GUNICORN_WORKERS=1`、容器内存约 700MB |
| 域名 | **spe-avatar.com**（已写入 `nginx/conf.d/app.conf`） |
| 同时在线 | 1GB 建议同时 **≤3～5** 人做 Avatar 访谈；人多请升到 2GB |

## 一、从本机导出密钥到 .env

在项目目录打开 PowerShell：

```powershell
cd "路径\interrogation_app"
.\scripts\export-env.ps1
```

脚本会读取你 Windows **用户/系统环境变量** 中的：

- `DEEPSEEK_API_KEY`
- `LIVEAVATAR_API_KEY`
- `ELEVENLABS_API_KEY`
- `ADMIN_PASSWORD`（可选）
- `SECRET_KEY`（若无则自动生成）

生成根目录 `.env`，**不要提交到 Git**。

## 二、上传代码到服务器

```bash
# 服务器上
mkdir -p /opt/interrogation-app
cd /opt/interrogation-app
```

上传整个项目（可用 git clone / scp），并上传本机生成的 `.env`：

```powershell
scp .env root@你的服务器IP:/opt/interrogation-app/.env
```

## 三、DNS

在域名控制台添加：

| 类型 | 主机 | 值 |
|------|------|-----|
| A | @ | 服务器公网 IP |
| A | www | 服务器公网 IP |

## 四、首次部署

```bash
cd /opt/interrogation-app
chmod 600 .env
chmod +x deploy.sh certbot-setup.sh deploy-update.sh
bash deploy.sh
```

## 五、HTTPS 证书

```bash
bash certbot-setup.sh spe-avatar.com 你的邮箱@example.com
docker compose exec nginx nginx -s reload
```

访问：https://spe-avatar.com

## 六、更新版本

```bash
cd /opt/interrogation-app
bash deploy-update.sh
```

## 七、常用命令

```bash
docker compose ps
docker compose logs -f app
docker compose restart app
```

## 功能说明

- 每次 Avatar 访谈 **最长 10 分钟**，到时自动断开并提示提交判断（训练）或结束会话（练习）。
- 日志目录：`./logs/app.log`
- 实验数据：`experiment_data.xlsx`（请定期备份）
