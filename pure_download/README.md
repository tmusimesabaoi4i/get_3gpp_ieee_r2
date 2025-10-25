# pure_download

```
tree /f

└─pure_download
    │  download_file.py
    │  download_html.py
    │  download_util.py
    │  msxml2_util.py
    │  README.md
```

# how to use

```python
import sys, pathlib
sys.path.append(str(pathlib.Path(__file__).resolve().parents[1]))

from pure_download.download_html import download_html_safely_msxml2
from pure_download.download_file import download_file_safely_msxml2
```