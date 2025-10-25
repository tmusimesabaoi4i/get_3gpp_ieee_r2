# emoscript.py
from __future__ import annotations
import shutil

class _Emo:
    __slots__ = ("_map",)
    def __init__(self):
        self._map = {
            # ==== base ====
            "ok": "✅",
            "ng": "❌",
            "warn": "⚠️",
            "info": "ℹ️",
            "star": "⭐",
            "spark": "✨",
            "wait": "⏳",
            "done": "🟢",
            "stop": "⛔",

            # ==== data-check ====
            "check": "🔎",
            "found": "📌",
            "missing": "🕳️",
            "added": "➕",
            "removed": "➖",
            "updated": "♻️",
            "duplicate": "🔁",
            "invalid": "🚫",
            "valid": "🟩",
            "mismatch": "🧩",
            "diff": "📄",
            "export": "📤",
            "import": "📥",
            "db": "🗄️",
            "excel": "📊",
            "file": "📄",
            "folder": "📁",

            # ==== download / progress ====
            "dl": "⬇️",          # ダウンロード
            "dl_box": "📥",      # 受け取り
            "net": "🌐",         # ネットワーク
            "save": "💾",        # 保存
            "zip": "🗜️",        # 展開/圧縮
            "retry": "🔁",       # リトライ
            "success": "✅",     # 成功
            "fail": "❌",        # 失敗
            "start": "🚀",       # 開始
            "finish": "🏁",      # 完了

            # ==== dots / spinner (単発表示用) ====
            "dots": "…",         # 三点リーダ
            "dots1": ".",        # . .. ... 用
            "dots2": "..",
            "dots3": "...",
            "bdot": "•",         # 黒丸ドット
            "bdots": "•••",      # 黒丸3つ（CMDでも無難）
            "spin1": "|",        # ASCII スピナー（順に出せば回って見える）
            "spin2": "/",
            "spin3": "-",
            "spin4": "\\",

            # ==== fixed separators (ASCII only) ====
            "sep": "=" * 60,
            "dash": "-" * 60,
            "dotline": "." * 60,
            "thick": "==--" * 15,  # 見た目強め
        }

    def __getattr__(self, name: str) -> str:
        # 動的区切り線（端末幅いっぱい）
        if name == "sep_full":
            return self.line("=")
        if name == "dash_full":
            return self.line("-")
        if name == "dot_full":
            return self.line(".")
        # 未定義は Slack 風プレースホルダ
        return self._map.get(name, f":{name}:")

    def add(self, name: str, value: str) -> None:
        """絵文字/記号を追加・上書き"""
        if not name:
            raise ValueError("name must not be empty")
        self._map[name] = value

    def remove(self, name: str) -> None:
        """絵文字/記号を削除（存在しなくてもOK）"""
        self._map.pop(name, None)

    def __getitem__(self, name: str) -> str:
        # emo["ok"] のようにもアクセス可
        return getattr(self, name)

    def line(self, char: str = "=", margin: int = 0) -> str:
        """端末幅いっぱいの区切り線"""
        width = shutil.get_terminal_size(fallback=(80, 20)).columns
        width = max(1, width - max(0, margin))
        if not char:
            char = "="
        return char[0] * width


# グローバルに1個だけ使う想定
emo = _Emo()


# ============== デモ ==============
if __name__ == "__main__":
    print(f"{emo.start} 開始 {emo.sep_full}")
    print(f"{emo.check} 入力データ検証中{emo.dots3}")
    print(f"{emo.found} レコード検出 / {emo.duplicate} 重複あり / {emo.invalid} 不正行あり")
    print(f"{emo.updated} 3件更新 {emo.added} 5件追加 {emo.removed} 2件削除 {emo.sep}")

    print(f"{emo.dl} ダウンロード開始{emo.dots1}")
    print(f"{emo.net} 接続中 {emo.spin1} {emo.spin2} {emo.spin3} {emo.spin4}")
    print(f"{emo.save} 保存中{emo.dots3}")
    print(f"{emo.zip} 展開中{emo.bdots}")
    print(f"{emo.success} 完了 {emo.finish}")
    print(emo.dash_full)
