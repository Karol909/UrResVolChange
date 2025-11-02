import threading
import struct
import socket
import logging
import time
import string

logging.basicConfig(level=logging.INFO)

class URDashboard:
    def __init__(self, host='192.168.0.64', port=29999) -> None:
        """
        Create a connection to robot.
        IP -> Computers
        PORT -> Unused port for the data to flow
        """
        self.s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.host = host
        self.port = port

        try:
            self.s.connect((self.host,self.port))
            logging.info(f"[Dashboard] Successfully connected to {self.host}:{self.port}")
        except socket.error as e:
            logging.error((f"[Dashboard] Failed to connect to {self.host}:{self.port} - {e}"))

    def load(self, program:str) -> None:
        """
        Loads a program with a name specified when calling this function.
        
        Parameters
        ----------
        program : str
            Name of the program with .urp
        """
        COMMAND = f"load {program}\n"
        self.program = program
        
        try:
            self.s.sendall(COMMAND.encode('utf-8'))
            logging.info(f"[Dashboard] Successfully loaded program {program}")
        except socket.error as e:
            logging.error((f"[Dashboard] Failed to load program {program} - {e}"))
    

    def play(self) -> None:
        """
        Plays a program with a name specified when calling this function.
        
        Parameters
        ----------
        program : str
            Name of the program with .urp
        """
        COMMAND = "play\n"
        
        try:
            self.s.sendall(COMMAND.encode('utf-8'))
            logging.info(f"[Dashboard] Successfully played program {self.program}")
        except socket.error as e:
            logging.error((f"[Dashboard] Failed to play program {self.program} - {e}"))
            
    def stop(self) -> None:
        """
        Stops a program with a name specified when calling this function.
        
        Parameters
        ----------
        program : str
            Name of the program with .urp
        """
        COMMAND = "stop\n"
        
        try:
            self.s.sendall(COMMAND.encode('utf-8'))
            logging.info(f"[Dashboard] Successfully stopped program")
        except socket.error as e:
            logging.error((f"[Dashboard] Failed to stop program - {e}"))
            
    def disconnect(self):
        self.s.close()