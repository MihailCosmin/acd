# acd/paddle_runtime.py
from __future__ import annotations
import os
import subprocess
import sys
import importlib.util
from typing import Optional

def ensure_paddle(auto_install: bool = False) -> bool:
    """
    Returns True if paddle+paddleocr are importable at or above minimum version.
    If auto_install=True, will attempt to install/upgrade once.
    """
    # Final check
    try:
        import paddle  # noqa
        import paddleocr  # noqa
        return True
    except Exception:
        try:
            subprocess.check_call([
                sys.executable, "-m", "pip", "install",
                "paddlepaddle>=3.1.0", 
                "-i", "https://www.paddlepaddle.org.cn/packages/stable/cpu/", "--timeout=1000"
            ])
            subprocess.check_call([
                sys.executable, "-m", "pip", "install",
                "paddlepaddle-gpu>=3.1.0", 
                "-i", "https://www.paddlepaddle.org.cn/packages/stable/cu118/", "--timeout=1000"
            ])
            subprocess.check_call([
                sys.executable, "-m", "pip", "install",
                "paddleocr>=3.1.0"
            ])
        except Exception:
            return False
        import paddle  # noqa
        import paddleocr  # noqa
        return True


def require_paddle() -> None:
    """
    Ensure PaddleOCR is available.
    If it's not installed yet, attempt automatic install once.
    """
    ok = ensure_paddle(auto_install=True)
    if not ok:
        raise RuntimeError(
            "Failed to install PaddleOCR automatically. "
            "Please install manually:\n"
            "  pip install -i https://www.paddlepaddle.org.cn/packages/stable/cpu/ 'paddlepaddle>=3.1.0'\n"
            "  pip install paddleocr>=3.1.0\n"
            "  pip install -i https://www.paddlepaddle.org.cn/packages/stable/cu118/ 'paddlepaddle-gpu>=3.1.0'"
        )
