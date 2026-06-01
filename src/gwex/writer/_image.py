"""画像の自動縮小（privileged 補助）。

シミュレータのスクショは大きいため、長辺上限で縮小してから添付する。
特に Google Sheets のセル内画像は base64 でサイズ上限があるため必須。
"""

from __future__ import annotations

import os
import tempfile
from typing import Optional

from PIL import Image


def downscale(image_path: str, max_dim: Optional[int]) -> str:
    """長辺が max_dim を超える場合のみ縮小し、新しい一時ファイルのパスを返す。

    max_dim が None / 0、または既に十分小さい場合は元パスをそのまま返す。
    """
    if not max_dim:
        return image_path
    with Image.open(image_path) as im:
        w, h = im.size
        if max(w, h) <= max_dim:
            return image_path
        scale = max_dim / max(w, h)
        # 縮小は LANCZOS（高品質リサンプリング）。既定フィルタはエイリアスで画質が落ちる。
        resized = im.resize(
            (max(1, round(w * scale)), max(1, round(h * scale))), Image.LANCZOS
        )
        suffix = os.path.splitext(image_path)[1] or ".png"
        out = tempfile.mktemp(suffix=suffix)
        if resized.mode in ("RGBA", "P") and suffix.lower() in (".jpg", ".jpeg"):
            resized = resized.convert("RGB")
        # PNG は無損失・最適圧縮で保存（品質劣化なし）
        save_kwargs = {"optimize": True} if suffix.lower() == ".png" else {}
        resized.save(out, **save_kwargs)
        return out
