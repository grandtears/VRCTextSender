from pythonosc.udp_client import SimpleUDPClient

class VRChatOSCClient:
    def __init__(self, ip: str, port: int):
        self.set_target(ip, port)

    def set_target(self, ip: str, port: int):
        """接続先を設定"""
        self.ip = ip
        self.port = port
        self.client = SimpleUDPClient(ip, port)

    def send_chat(self, text: str, enter: bool = True):
        """チャットボックスにテキストを送信"""
        # VRChatのOSC仕様変更に対応するため複数パターンを試行
        try:
            self.client.send_message("/chatbox/input", [text, enter, False])
        except Exception:
            try:
                self.client.send_message("/chatbox/input", [text, enter])
            except Exception:
                self.client.send_message("/chatbox/input", text)
