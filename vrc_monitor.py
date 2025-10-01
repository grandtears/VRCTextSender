import psutil

PROCESS_NAMES = {"vrchat.exe", "vrchat"}

def is_vrchat_running() -> bool:
    """VRChatの起動状態をチェック"""
    for proc in psutil.process_iter(["name"]):
        try:
            name = (proc.info.get("name") or "").lower()
            if name in PROCESS_NAMES:
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    return False
