# 模拟行动游戏 — 本地视频

将 MP4 放在**本目录**，文件名必须与 `app.py` 里 `SeriousStep(video="...")` 完全一致（区分大小写）。

## 需要的文件（共 40 个）

### 纵火 · 有罪
- `Guilty1.mp4` … `Guilty7.mp4`（含 `Guilty2-1.mp4`、`Guilty2-2.mp4` 等带连字符的）

### 纵火 · 无罪
- `Innocent1.mp4` … `Innocent7.mp4`

### 盗窃 · 有罪
- `Theft_Guilty1.mp4` … `Theft_Guilty7.mp4`

### 盗窃 · 无罪
- `Theft_Innocent1.mp4` … `Theft_Innocent7.mp4`

## 从 YouTube 批量下载（需在可访问 YouTube 的网络下执行一次）

1. 安装 [yt-dlp](https://github.com/yt-dlp/yt-dlp)：`pip install yt-dlp`
2. 在项目根目录运行：

```bash
python scripts/download_serious_game_videos.py
```

3. 确认本目录下已有全部 MP4，再启动 Flask 做试玩。

## 自测

浏览器打开（把文件名换成你已有的一个）：

`http://127.0.0.1:5000/static/videos/serious-game/Guilty1.mp4`

能播放后，再走实验流程里的「模拟行动游戏」。
