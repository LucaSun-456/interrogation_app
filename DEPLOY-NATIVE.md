# 无 Docker 部署（Ubuntu + 用户 spe_avatar）

路径：**`/home/spe_avatar/interrogation-app`**

不需要 Docker、pnpm、Node。

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
scp .env spe_avatar@你的服务器IP:/home/spe_avatar/interrogation-app/.env
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
curl http://127.0.0.1:8000/api/health
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

Cloudflare：A 记录 → 新服务器 IP，**Flexible SSL**，源站只开 **80**。

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
