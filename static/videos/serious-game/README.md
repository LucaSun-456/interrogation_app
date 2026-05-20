# 模拟行动游戏 — 本地视频

将 MP4 放在**本目录**，文件名必须与 `app.py` 里 `SeriousStep(video="...")` 完全一致（区分大小写）。

## 需要的文件（共 40 个）

命名规则：`{案件}_{有罪|无罪}_{步骤}.mp4`

### 纵火 Arson · 有罪
- `Arson_Guilty_1.mp4` … `Arson_Guilty_7.mp4`（分支步：`Arson_Guilty_2-1.mp4` 等）

### 纵火 · 无罪
- `Arson_Innocent_1.mp4` … `Arson_Innocent_7.mp4`

### 盗窃 Theft · 有罪
- `Theft_Guilty_1.mp4` … `Theft_Guilty_7.mp4`

### 盗窃 · 无罪
- `Theft_Innocent_1.mp4` … `Theft_Innocent_7.mp4`

从旧文件名批量改名：`.\scripts\rename_serious_game_videos.ps1`

## 从 YouTube 批量下载（需在可访问 YouTube 的网络下执行一次）

1. 安装 [yt-dlp](https://github.com/yt-dlp/yt-dlp)：`pip install yt-dlp`
2. 在项目根目录运行：

```bash
python scripts/download_serious_game_videos.py
```

3. 确认本目录下已有全部 MP4，再启动 Flask 做试玩。

## 自测

浏览器打开（把文件名换成你已有的一个）：

`http://127.0.0.1:5000/static/videos/serious-game/Arson_Guilty_1.mp4`

能播放后，再走实验流程里的「模拟行动游戏」。

## 本地转码（Windows，上传前做）

1. 安装 [ffmpeg](https://ffmpeg.org/download.html) 并加入 PATH  
2. 在项目根目录 PowerShell 执行：

```powershell
.\scripts\transcode_serious_game_local.ps1
```

原文件会备份到 `static/videos/serious-game-backup/`。转完后用 WinSCP 上传 `serious-game` 目录下全部 `.mp4` 到服务器同路径。

## 服务器播放卡顿

1. **Nginx 直出视频**（不要只走 `:3003`）：在 `scripts/nginx-spe-avatar.conf.example` 里已含 `location /static/videos/`。服务器上更新站点配置后 `sudo nginx -t && sudo systemctl reload nginx`。被试请用域名 `https://spe-avatar.com` 访问，不要长期用 IP:3003 看片。
2. **压缩码率**（最有效）：在服务器上对 MP4 批量转码，例如 720p：

```bash
cd ~/interrogation-app/static/videos/serious-game
mkdir -p ../serious-game-backup && cp *.mp4 ../serious-game-backup/
for f in *.mp4; do
  ffmpeg -y -i "$f" -vf "scale=-2:720" -c:v libx264 -preset medium -crf 23 -c:a aac -b:a 128k -movflags +faststart "tmp_$f" && mv "tmp_$f" "$f"
done
```

`-movflags +faststart` 把元数据移到文件头，边下边播更顺。需先 `sudo apt install ffmpeg`。
