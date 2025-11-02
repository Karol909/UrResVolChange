import threading
import struct
import socket
import logging
import time

logging.basicConfig(level=logging.INFO)

class URReceive:
    def __init__(self, host='192.38.66.227', port=30003):
        """
        Create a connection to robot.
        IP -> Computers
        PORT -> Unused port for the data to flow
        """
        self.host = host
        self.port = port

        self.s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.s.bind((self.host, self.port))
        self.s.listen(1)
    
    def listen(self) -> bool:
        logging.info(f"Listening for robot on {self.host}:{self.port}")
        self.conn, self.addr = self.s.accept()
        logging.info(f"Robot connected from {self.addr}")
        return True

    def get_data_UR(self) -> None:
        
        data = self.conn.recv(64)
        if not data:
            return None
        return data.decode("utf-8")
    
    def disconnect(self):
        self.s.close()