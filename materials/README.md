# 实验材料目录

本目录存放**合并后的培训与知情同意材料**，不再在根目录或此处散落多个 Word/PDF。

## 结构

```
materials/
  README.md                 # 本说明
  combined_materials.md     # 培训/知情同意合集（已纳入 Git，改后 push）
  combined_materials.docx     # 自动生成，未纳入 Git（可选 scp 上传）
  prompts/
    avatar_feedback.md      # Avatar 培训反馈 Prompt（由 Feedback prompts.docx 导出）
  _legacy_sources/          # 可选：从外部 IRB 目录复制来的原始文件备份
```

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

## 实验数据（统计）

所有参与者、预约、问卷、Avatar 培训记录、严肃游戏选择等**统计数据**统一保存在项目根目录的 **`experiment_data.xlsx`**（多工作表），不再使用 `results_serious_game.xlsx` 或 `training_feedback/` 多文件存档。
