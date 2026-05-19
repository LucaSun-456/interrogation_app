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

## 三、DNS 与 HTTPS（二选一）

### 方案 A：Cloudflare 反代 + 免费 HTTPS（推荐，无需 Certbot）

适合国内 VPS，**不用在服务器申请证书**。

1. 注册 [Cloudflare](https://dash.cloudflare.com)，添加站点 `spe-avatar.com`（免费计划即可）。
2. 按提示把域名 **Nameserver** 改到 Cloudflare（在 GoDaddy 域名管理里把 NS 换成 Cloudflare 提供的两个地址）。
3. 等 NS 生效后，在 Cloudflare **DNS** 里添加：

| 类型 | 名称 | 内容 | 代理 |
|------|------|------|------|
| A | `@` | `47.103.48.94`（你的 VPS IP） | **已代理（橙色云）** |
| A | `www` | 同上（可选） | **已代理** |

4. **SSL/TLS** → 加密模式选 **灵活（Flexible）**  
   - 用户 ↔ Cloudflare：HTTPS  
   - Cloudflare ↔ 你的服务器：HTTP（80），与当前 Nginx 配置一致  

5. **SSL/TLS** → 边缘证书 → 开启 **始终使用 HTTPS（Always Use HTTPS）**。

6. 服务器上 **只需 HTTP**，不必跑 Certbot：
   ```bash
   # 若曾复制过 SSL 配置且没有证书，可删掉避免 443 报错
   rm -f nginx/conf.d/app-ssl.conf
   docker compose exec nginx nginx -t && docker compose exec nginx nginx -s reload
   ```
7. 阿里云安全组：放行 **80**（必须）；443 可不开放（Flexible 模式下 Cloudflare 连源站用 80）。

访问：**https://spe-avatar.com**

> Avatar / WebRTC 走浏览器直连 LiveAvatar API，一般不受 Cloudflare 影响。若访谈异常，可在 Cloudflare 该子域暂时关闭代理（灰云）排查。

### 方案 B：服务器 Certbot（不用 Cloudflare 时）

在域名控制台（GoDaddy 等）添加：

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

## 五、HTTPS 证书（仅方案 B：不用 Cloudflare 时）

若已用 **Cloudflare Flexible**，跳过本节。

首次部署 **只启用 HTTP**（`app.conf`），避免尚未申请证书时 Nginx 因缺少 `fullchain.pem` 无法启动。证书就绪后：

```bash
bash certbot-setup.sh spe-avatar.com 你的邮箱@example.com
# 脚本会复制 app-ssl.conf.example → app-ssl.conf 并 reload
```

访问：https://spe-avatar.com

若 `app` 容器一直 Restarting，先查日志：`docker compose logs app --tail 30`。

| 日志关键词 | 处理 |
|-----------|------|
| `Permission denied: '/app/logs/app.log'` | 在服务器执行：`bash scripts/fix-server-permissions.sh` |
| `Is a directory: '/app/experiment_data.xlsx'` | `docker compose down && rm -rf experiment_data.xlsx && mkdir -p data`，再 `git pull` 并重建 |
| `location directive is not allowed here` | `git pull` 更新 `nginx/nginx.conf`，或删掉 `http {}` 里错误的 `location` 块 |
| `Missing required environment variables` | 补全 `.env` 后 `docker compose up -d --force-recreate app` |

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
- 实验数据：**`data/experiment_data.xlsx`**（含参与者、问卷、培训、严肃游戏等所有工作表；请定期备份 `data/` 目录）
- 勿将 `experiment_data.xlsx` 单独挂载为 Docker 卷——若宿主机上该路径不存在，Docker 会把它建成**目录**导致启动失败
- 培训材料：`materials/combined_materials.docx`（由启动时或 `python scripts/build_combined_materials.py` 从 Word/PDF 合并生成）
