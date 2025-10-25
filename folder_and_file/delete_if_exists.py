import sys, pathlib
sys.path.append(str(pathlib.Path(__file__).resolve().parents[1]))
from emoji.emoscript import emo

from pathlib import Path

def delete_if_exists(
        folder_abs_path: str,
        filename: str
    ) -> bool:
    
    p = Path(folder_abs_path) / filename
    if not p.is_absolute():
        raise ValueError(f"{emo.warn} folder_abs_path は絶対パスで指定してください。")
    if not p.exists():
        raise FileNotFoundError(f"{emo.warn} Excel ファイルが見つかりません: {p}")

    if not p.exists():
        return False

    try:
        if p.is_file() or p.is_symlink():
            p.unlink()
            return True
        if p.is_dir():
            raise IsADirectoryError(f"{emo.invalid} 指定のパスはディレクトリです: {p}")
        raise OSError(f"{emo.invalid} 通常のファイルではありません: {p}")
    except PermissionError as e:
        raise PermissionError(f"{emo.invalid} 削除権限がありません: {p}") from e
    except OSError as e:
        raise OSError(f"{emo.invalid} 削除に失敗しました: {p} ({e})") from e