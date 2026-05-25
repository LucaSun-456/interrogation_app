# 实验材料目录

本目录存放**合并后的培训与知情同意材料**，不再在根目录或此处散落多个 Word/PDF。

## 结构

```
materials/
  README.md                 # 本说明
  combined_materials.md     # 屏上阅读用合集（改文字只改此文件）
  pdf/                      # 参与者下载用 PDF（见 pdf/README.md）
  prompts/
    avatar_feedback.md      # Avatar 培训反馈 Prompt
  _legacy_sources/          # 可选：历史 Word 备份（程序不读取）
```

**下载 PDF**：将文件放入 `materials/pdf/`（共 8 个：2 个培训 + 每种案件 3 个），文件名须与 `pdf/README.md` 一致。  
**屏上阅读**：继续编辑 `combined_materials.md` 各 `<!-- section:... -->` 段落。

## 章节标记

`combined_materials.md` 内用 HTML 注释分段，供程序读取：

- `consent_suspect` / `consent_interviewer` — 知情同意书
- `theory_sue` — B 组 SUE 理论培训
- `avatar_specific` / `avatar_general` / `control` — 各培训组说明
- `interviewer_bg` / `avatar_persona_settings` / `avatar_demo` / `background_info` — 其他参考材料

原始 Word/PDF 在合并进合集后已从仓库删除，请勿再放回 `materials/` 根目录。

## 首次生成

将原始 `.docx` / `.pdf` 放在项目根目录或本目录后，启动应用会自动合并；也可手动执行：

```bash
python scripts/build_combined_materials.py
```

服务器上若没有源文件，可先使用仓库内的模板（仅供联调，**正式实验前务必替换**）：

```bash
cp materials/combined_materials.md.example materials/combined_materials.md
sudo systemctl restart interrogation-app
```

## LiveAvatar 配置（D/C 组虚拟审讯）

Avatar 的视觉 ID 与语音 ID **不在** `.env` 里，而在数据目录的 JSON 文件中：

| 文件 | 说明 |
|------|------|
| `data/avatars.json` | 运行时读取（已在 `.gitignore`，含真实 ID，仅留在服务器本地） |
| `avatars.json.example` 或 `data/avatars.json.example` | 结构模板（已纳入 Git），复制后改名并填入 ID |

服务器上执行（任选其一）：

```bash
cp avatars.json.example data/avatars.json
# 或
cp data/avatars.json.example data/avatars.json
```

- **C 组（`avatar_general`）**：使用 `generic` 节点下的 `avatar_id`。
- **D 组（`avatar_specific`）**：按配对嫌疑人的问卷（性别、眼镜、发型）匹配 `specific` 下 8 种组合；若无嫌疑人档案则回退到 `generic`。

还需在 `.env` 中配置 `LIVEAVATAR_API_KEY`（及 `ELEVENLABS_API_KEY`）。若仍报「未配置 Avatar ID」，说明 `avatars.json` 缺失或对应节点的 `avatar_id` 为空。

## 实验数据（统计）

所有参与者、预约、问卷、Avatar 培训记录、严肃游戏选择等**统计数据**统一保存在项目根目录的 **`experiment_data.xlsx`**（多工作表），不再使用 `results_serious_game.xlsx` 或 `training_feedback/` 多文件存档。
