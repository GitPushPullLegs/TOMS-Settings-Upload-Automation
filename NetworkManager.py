import socket



class NetworkManager:
    _dnsServer = '1.1.1.1'

    def is_connected(self, host: str = _dnsServer) -> bool:
        try:
            s = socket.create_connection((host, 80), 2)
            s.close()
            return True
        except:
            from EmailManager import EmailManager
            EmailManager().send_email(subject="TOMS Settings Upload Error", body="Unable to connect to network.")
            return False
