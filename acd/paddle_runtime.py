# acd/paddle_runtime.py
from __future__ import annotations
import os
import subprocess
import sys
import importlib.util
from typing import Optional

_MIN_PADDLE = "3.1.0"
_MIN_PADDLEOCR = "3.1.0"

# Single-process guards
_paddle_ready = False
_tried_install = False

def _pkg_missing(name: str, min_ver: Optional[str] = None) -> bool:
    if importlib.util.find_spec(name) is None:
        return True
    if not min_ver:
        return False
    try:
        from importlib.metadata import version
        from packaging.version import Version
        return Version(version(name)) < Version(min_ver)
    except Exception:
        # If we can’t compare, assume present
        return False

def ensure_paddle(auto_install: bool = False) -> bool:
    """
    Returns True if paddle+paddleocr are importable at or above minimum version.
    If auto_install=True, will attempt to install/upgrade once.
    """
    global _paddle_ready, _tried_install
    if _paddle_ready:
        return True

    needs_paddle = (
        _pkg_missing("paddle", _MIN_PADDLE) or
        _pkg_missing("paddleocr", _MIN_PADDLE)
    )
    if not needs_paddle:
        _paddle_ready = True
        return True

    if not auto_install:
        return False

    if _tried_install:
        return False  # don’t loop
    _tried_install = True

    # Decide CPU vs GPU by env; default CPU (safer).
    want_gpu = os.getenv("ACD_PADDLE_GPU", "0") in ("1", "true", "TRUE")

    cmds = []
    if want_gpu:
        cmds.append([
            sys.executable, "-m", "pip", "install",
            f"paddlepaddle-gpu>={_MIN_PADDLE}",
            "-i", "https://www.paddlepaddle.org.cn/packages/stable/cu118/",
            "--timeout=1000",
        ])
    else:
        cmds.append([
            sys.executable, "-m", "pip", "install",
            f"paddlepaddle>={_MIN_PADDLE}",
            "-i", "https://www.paddlepaddle.org.cn/packages/stable/cpu/",
            "--timeout=1000",
        ])
    cmds.append([sys.executable, "-m", "pip", "install", f"paddleocr>={_MIN_PADDLE}"])

    for cmd in cmds:
        try:
            subprocess.check_call(cmd)
        except subprocess.CalledProcessError:
            return False

    # Final check
    try:
        import paddle  # noqa
        import paddleocr  # noqa
        _paddle_ready = True
        return True
    except Exception:
        return False


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
