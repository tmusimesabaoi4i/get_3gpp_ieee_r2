import sys, pathlib
sys.path.append(str(pathlib.Path(__file__).resolve().parents[1]))
from emoji.emoscript import emo

from pathlib import Path
from typing import Tuple
from folder_and_file.folder_exists_in_folder import folder_exists_in_folder

def create_subfolder_when_absent(
    folder_abs_path: str,
    subfolder_name: str,
    ) -> Tuple[str, bool]:
    
    parent = Path(folder_abs_path)
    if not parent.is_absolute():
        raise ValueError(f"{emo.invalid} folder_abs_path は絶対パスで指定してください。: {folder_abs_path}")
    if not parent.exists():
        raise FileNotFoundError(f"{emo.ng} 親フォルダが存在しません: {parent}")
    if not parent.is_dir():
        raise NotADirectoryError(f"{emo.invalid} 親フォルダがディレクトリではありません: {parent}")

    name_only = Path(subfolder_name).name
    if not name_only:
        raise ValueError(f"{emo.invalid} subfolder_name が空です。")

    target = parent / name_only

    exists = bool(folder_exists_in_folder(str(parent), name_only))
    if not exists:
        try:
            target.mkdir(parents=False, exist_ok=False)
            print(f"{emo.added} {emo.folder} 作成しました: {target}")
            return (str(target.resolve()), True)
        except FileExistsError:
            if target.is_dir():
                return (str(target.resolve()), False)
            raise FileExistsError(f"{emo.ng} 同名のファイルが存在します: {target}")

    if target.exists():
        if not target.is_dir():
            raise FileExistsError(f"{emo.ng} 同名のファイルが存在します: {target}")
        return (str(target.resolve()), False)

    target.mkdir(parents=False, exist_ok=False)
    print(f"{emo.added} {emo.folder} 作成しました: {target}  {emo.warn}（存在判定との不整合を修復）")
    return (str(target.resolve()), True)

# ============== 実行部 ==============
if __name__ == "__main__":
    folder = r"C:\Users\yohei\Downloads"
    folder1  = "project_2"
    folder2  = "project"

    create_subfolder_when_absent(folder, folder1)
    create_subfolder_when_absent(folder, folder2)