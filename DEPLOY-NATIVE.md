# 无 Docker 部署（Ubuntu + 用户 spe_avatar）

路径：**`/home/spe_avatar/interrogation-app`**

应用对外 **`http://47.238.75.193:3003`**（Gunicorn `0.0.0.0:3003`）。域名 **`spe-avatar.com`** 仍经 Nginx **80 → 3003**。

| 项目 | 值 |
|------|-----|
| 服务器 IP | **47.238.75.193** |
| 域名 | **spe-avatar.com** |
| 应用端口（本机） | **3003** |

**DNS**：在域名商或 Cloudflare 添加 A 记录 `@` 和 `www` → `47.238.75.193`。  
**云安全组**（阿里云等）：放行 **3003**（IP 直连）、**80**（域名）、**443**（若用 HTTPS）、**22**（SSH）。

### 与同机其他站点共存（例如 rogare.site）

| 站点 | 域名 | 本机端口 | Nginx 配置文件 |
|------|------|----------|----------------|
| 已有 | rogare.site | **3001** | 保持不动（如 `sites-available/rogare.site`） |
| 本应用 | spe-avatar.com | **3003** | 仅写入 `sites-available/spe-avatar.com` |

两台应用可同时监听 80：Nginx 按 `server_name` 分流，互不冲突。部署脚本**不会**修改 rogare 的配置或占用 3001。

不需要 Docker、pnpm、Node。

---

## 零、一键部署（推荐）

在 Ubuntu 服务器上以 **root** 执行：

```bash
# 1) 先在本机上传 .env（PowerShell）
# scp .env spe_avatar@47.238.75.193:/home/spe_avatar/interrogation-app/.env

# 2) 克隆后一键部署
git clone https://github.com/LucaSun-456/interrogation_app.git /home/spe_avatar/interrogation-app
chmod +x /home/spe_avatar/interrogation-app/scripts/deploy-ubuntu-spe-avatar.sh
sudo bash /home/spe_avatar/interrogation-app/scripts/deploy-ubuntu-spe-avatar.sh
```

脚本会自动：安装依赖、创建用户、venv、systemd（3003）、nginx（spe-avatar.com）、UFW、健康检查。

---

## 一、创建用户（root 执行一次）

```bash
# 上传项目后，或在 clone 后执行：
cd /home/spe_avatar/interrogation-app
sudo bash scripts/ubuntu-setup-user.sh
```

或手动：

```bash
sudo useradd -m -s /bin/bash spe_avatar
sudo passwd spe_avatar          # 设置登录密码（可选）
sudo apt update
sudo apt install -y python3 python3-venv python3-pip nginx git curl
```

---

## 二、以 spe_avatar 部署代码

```bash
sudo -u spe_avatar -i
cd ~
git clone https://github.com/LucaSun-456/interrogation_app.git interrogation-app
cd ~/interrogation-app
```

从本机上传 `.env` 到 `/home/spe_avatar/interrogation-app/.env`（`scp` 示例）：

```powershell
scp .env spe_avatar@47.238.75.193:/home/spe_avatar/interrogation-app/.env
```

```bash
chmod 600 .env
chmod +x scripts/deploy-native.sh
bash scripts/deploy-native.sh
```

---

## 三、systemd 常驻（root）

```bash
sudo cp /home/spe_avatar/interrogation-app/scripts/interrogation-app.service.example \
        /etc/systemd/system/interrogation-app.service
sudo systemctl daemon-reload
sudo systemctl enable --now interrogation-app
sudo systemctl status interrogation-app
curl http://127.0.0.1:3003/api/health
```

---

## 四、Nginx 反代（root）

```bash
sudo cp /home/spe_avatar/interrogation-app/scripts/nginx-spe-avatar.conf.example \
        /etc/nginx/sites-available/spe-avatar.com
sudo ln -sf /etc/nginx/sites-available/spe-avatar.com /etc/nginx/sites-enabled/
sudo rm -f /etc/nginx/sites-enabled/default
sudo nginx -t && sudo systemctl reload nginx
```

Cloudflare：A 记录 `@`、`www` → **47.238.75.193**，**Flexible SSL**，源站只开 **80**。

---

## 五、从旧服务器迁移

```bash
# 旧机打包
tar czf backup.tar.gz data .env

# 新机（spe_avatar 用户）
cd ~/interrogation-app
tar xzf backup.tar.gz
bash scripts/deploy-native.sh
```

---

## 六、日常更新

```bash
sudo -u spe_avatar -i
cd ~/interrogation-app && git pull && bash scripts/deploy-native.sh
```
