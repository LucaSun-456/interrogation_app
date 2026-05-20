#!/usr/bin/env python3
"""
One-time download: YouTube -> static/videos/serious-game/<filename>.mp4
Requires network access to YouTube and: pip install yt-dlp
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "static" / "videos" / "serious-game"

# filename -> YouTube video id
VIDEOS: dict[str, str] = {
    "Arson_Guilty_1.mp4": "tKf2BCNEh-M",
    "Arson_Guilty_2-1.mp4": "k0GL1uXqUkk",
    "Arson_Guilty_2-2.mp4": "KALNxqLiJZM",
    "Arson_Guilty_3.mp4": "sQLStPwicr4",
    "Arson_Guilty_4-1.mp4": "zQ88Dzu-D0I",
    "Arson_Guilty_4-2.mp4": "jW4yCLAV2II",
    "Arson_Guilty_5.mp4": "WpBchjwBMec",
    "Arson_Guilty_6-1.mp4": "rHQv0gLh1Ls",
    "Arson_Guilty_6-2.mp4": "3LU89Josrjs",
    "Arson_Guilty_7.mp4": "D1mJTQvoIaY",
    "Arson_Innocent_1.mp4": "KqcOFshJ1UE",
    "Arson_Innocent_2-1.mp4": "k-kScca4P4U",
    "Arson_Innocent_2-2.mp4": "bd5LR81eZfw",
    "Arson_Innocent_3.mp4": "I6GL4QA4qn4",
    "Arson_Innocent_4-1.mp4": "Nh1naKB74ho",
    "Arson_Innocent_4-2.mp4": "LLJq5LG9qmk",
    "Arson_Innocent_5.mp4": "XuePULTX0BU",
    "Arson_Innocent_6-1.mp4": "BmACObyXqmQ",
    "Arson_Innocent_6-2.mp4": "NZaCYGkQ8KI",
    "Arson_Innocent_7.mp4": "25Fj6u28Bqc",
    "Theft_Guilty_1.mp4": "UfGQOGLl9Lc",
    "Theft_Guilty_2-1.mp4": "OqoPLqe9o4Y",
    "Theft_Guilty_2-2.mp4": "0Oo3b7JJku8",
    "Theft_Guilty_3.mp4": "aIh9QzrNcFI",
    "Theft_Guilty_4-1.mp4": "XUbCdq_dCu8",
    "Theft_Guilty_4-2.mp4": "GxAJxGwAgKQ",
    "Theft_Guilty_5.mp4": "G1YveoiofPY",
    "Theft_Guilty_6-1.mp4": "FHnFy5r9IJQ",
    "Theft_Guilty_6-2.mp4": "ZqhndWKBpZk",
    "Theft_Guilty_7.mp4": "S5zHcfVcpd4",
    "Theft_Innocent_1.mp4": "UfGQOGLl9Lc",
    "Theft_Innocent_2-1.mp4": "OqoPLqe9o4Y",
    "Theft_Innocent_2-2.mp4": "0Oo3b7JJku8",
    "Theft_Innocent_3.mp4": "aIh9QzrNcFI",
    "Theft_Innocent_4-1.mp4": "99Gk-iwdAqo",
    "Theft_Innocent_4-2.mp4": "Vqyt3g0HUcc",
    "Theft_Innocent_5.mp4": "Au7twsHYwf8",
    "Theft_Innocent_6-1.mp4": "ZUV3Jk9B1ww",
    "Theft_Innocent_6-2.mp4": "hsKqRAe8Lls",
    "Theft_Innocent_7.mp4": "d1i1SdStq0Y",
}


def main() -> int:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    missing = [f for f in VIDEOS if not (OUT_DIR / f).is_file()]
    if not missing:
        print(f"All {len(VIDEOS)} files already present in {OUT_DIR}")
        return 0

    try:
        subprocess.run(["yt-dlp", "--version"], capture_output=True, check=True)
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("Install yt-dlp first: pip install yt-dlp", file=sys.stderr)
        return 1

    ok, fail = 0, 0
    for name in sorted(missing):
        vid = VIDEOS[name]
        dest = OUT_DIR / name
        url = f"https://www.youtube.com/watch?v={vid}"
        print(f"\n>>> {name} ({vid})")
        cmd = [
            "yt-dlp",
            "-f", "bv*[ext=mp4]+ba[ext=m4a]/b[ext=mp4]/b",
            "--merge-output-format", "mp4",
            "-o", str(dest),
            "--no-overwrites",
            url,
        ]
        r = subprocess.run(cmd)
        if r.returncode == 0 and dest.is_file():
            ok += 1
        else:
            fail += 1
            print(f"FAILED: {name}", file=sys.stderr)

    print(f"\nDone: {ok} downloaded, {fail} failed, {len(VIDEOS) - len(missing)} skipped (already exist)")
    return 0 if fail == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
