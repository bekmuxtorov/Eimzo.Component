import websocket

import asyncio
from websockets.asyncio.client import connect

import win32com.server.util
import pythoncom


class WebSocket:
    _public_methods_ = ["get_author",
                        "get_all_certificates", "send", "get_certificates", "create_pkcs", "load_key"]
    _reg_progid_ = "Eimzo.Component"
    _reg_clsid_ = pythoncom.CreateGuid()

    URI = "wss://127.0.0.1:64443/service/cryptapi"
    ORIGIN = "https://e-imzo.soliq.uz/"

    def get_author(self):
        return f"Asadbek Muxtorov | https://github.com/bekmuxtorov"

    async def async_send(self, message: str, wait_iterations: int = None) -> str:
        async with connect(uri=self.URI, origin=self.ORIGIN) as websocket:
            await websocket.send(message)
            return await websocket.recv()

    async def async_get_all_certificates(self) -> str:
        """
        return: Kompyuterdagi barcha sertifikatlarni qaytaradi.
        """
        async with connect(uri=self.URI, origin=self.ORIGIN) as websocket:
            await websocket.send("""{"plugin": "pfx", "name": "list_all_certificates"}""")
            return await websocket.recv()

    async def async_get_certificates(self, disk: str) -> str:
        """
        Kompyuterdagi {disk}(C, D)diskdagi sertifikatlarni qaytaradi.

        Params:
            disk: Belgilangan disk

        Return: 
            {
            "certificates": [
                {
                    "disk": "C:\\",
                    "path": "DSKEYS",
                    "name": "DS5230703",
                    "alias": "cn=muxtorov asadbek,name=asadbek,surname=muxtorov,l=a tumani,st=b viloyati,c=uz,1.2.860.3.16.1.2=jshshr,serialnumber=2222,validfrom=2024.11.12 22:25:27,validto=2026.11.12 22:25:26"
                }
            ],
            "success": true
            }
        """
        async with connect(uri=self.URI, origin=self.ORIGIN) as websocket:
            message = """{"plugin": "pfx", "name": "list_certificates", "arguments":""" + \
                f"""[\"{disk}\"]""" + "}"
            await websocket.send(message)
            return await websocket.recv()

    async def async_load_key(self, disk: str, path: str, name: str, alias: str) -> str:
        """
        Sertifikatni qabul qilish

        Params:
            disk: Belgilangan disk
            path: Sertifikatni kompyuterda joylashgan manzili, 
            name: Sertifikat faylini nomi, 
            alias: Sertifikat ma'lumotlari

        Return: 
            {
                "keyId": "ebd4978d5",
                "type": "PFX_KEY_STORE",
                "success": true
            }
        """
        async with connect(uri=self.URI, origin=self.ORIGIN) as websocket:
            message = """{\"plugin\"""" + ":" + """\"pfx\", \"name\"""" + ":" + \
                """\"load_key\",\"arguments\"""" + ":" + \
                f"""[\"{disk}\",\"{path}\",\"{name}\",\"{alias}\"]""" + "}"
            await websocket.send(message)
            return await websocket.recv()

    async def async_create_pkcs(self, textbase64: str, keyId: str) -> str:
        """
        Hjjatni imzolash. (Sertifikat parolini kiritish va uni tekshirish uchun ishlatilsa bo'ladi)

        Params:
            textbase64: base64 dagi matn
            keyId: Sertifikat ID si

        Return: 
            {
                "pkcs7_64": "MIAGCSqGSIb3DQEHA
                "signer_serial_number": "444444",
                "signature_hex": "09e01d95
                "success": true
            }
        """
        async with connect(uri=self.URI, origin=self.ORIGIN) as websocket:
            message = """{\"plugin\"""" + ":" + """\"pkcs7\", \"name\"""" + ":" + \
                """\"create_pkcs7\",\"arguments\"""" + ":" + \
                f"""[\"{textbase64}\",\"{keyId}\",\"no\"]""" + "}"
            await websocket.send(message)
            return await websocket.recv()

    def get_all_certificates(self):
        return asyncio.run(self.async_get_all_certificates())

    def get_certificates(self, disk):
        return asyncio.run(self.async_get_certificates(disk))

    def send(self, message: str, wait_iterations: int = None) -> str:
        return asyncio.run(self.async_send(message, wait_iterations))

    def load_key(self, disk, path, name, alias):
        return asyncio.run(self.async_load_key(disk, path, name, alias))

    def create_pkcs(self, textbase64, keyId):
        return asyncio.run(self.async_create_pkcs(textbase64, keyId))


if __name__ == "__main__":
    from win32com.server.register import UseCommandLine
    print("Registering...")
    UseCommandLine(WebSocket)
