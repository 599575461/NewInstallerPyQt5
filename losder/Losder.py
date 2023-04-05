import json
import base64


class Losder:
    def __init__(self, WhetherItIsEncrypted: bool = False):
        self.WhetherItIsEncrypted = WhetherItIsEncrypted

    @staticmethod
    def read_json_file(file: str) -> dict:
        try:
            with open(file, "r", encoding="utf-8") as w:
                return json.load(w)
        except Exception as e:
            print(e)

    def read_qss_file(self, file: str) -> str:
        """读取QSS文件是否加密

        Args:
            file (str): 文件路径

        Returns:
            str: 返回读取值
        """
        try:
            if self.WhetherItIsEncrypted:
                with open(file, "rb") as w:
                    return base64.b64decode(w.read()).decode("utf-8")
            else:
                with open(file, "r", encoding="utf-8") as w:
                    return w.read()
        except Exception as e:
            print(e)

    def write_qss_file(self, file: str, text) -> None:
        """写入QSS文件

        Args:
            file (str): 文件路径
            text: 写入文本
        """
        try:
            with open(file, "wb") as w:
                if self.WhetherItIsEncrypted:
                    w.write(base64.b64encode(text.encode(encoding="utf-8")))
                else:
                    w.write(text)
        except Exception as e:
            print(e)

    @staticmethod
    def write_json_file(file: str, obj: dict) -> None:
        try:
            with open(file, "w+", encoding="utf-8") as w:
                json.dump(obj, w)
        except Exception as e:
            print(e)

    def change_json_text(self, file: str, var: str, key: str) -> None:
        try:
            temp_json = self.read_json_file(file)
            temp_json[key] = var
            self.write_json_file(file, temp_json)
        except Exception as e:
            print(e)
