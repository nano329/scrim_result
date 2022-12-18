import socket
from _thread import *

client_sockets = []

HOST = "127.0.0.1"
PORT = 9999

print(">> Go")
server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
server_socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
server_socket.bind((HOST, PORT))
server_socket.listen()