from __future__ import print_function

import socket
import sys
import threading


def socket_to_stdout(connection):
    output_stream = sys.stdout.buffer
    try:
        while True:
            data = connection.recv(4096)
            if not data:
                break
            output_stream.write(data)
            output_stream.flush()
    except Exception as error:
        sys.stderr.write("Clippy MCP bridge receive error: {0}\n".format(error))
        sys.stderr.flush()


def main():
    if len(sys.argv) != 3:
        sys.stderr.write("usage: mcp_stdio_forward.py HOST PORT\n")
        return 2

    host = sys.argv[1]
    try:
        port = int(sys.argv[2])
    except ValueError:
        sys.stderr.write("PORT must be an integer\n")
        return 2

    connection = socket.create_connection((host, port), 10)
    receiver = threading.Thread(target=socket_to_stdout, args=(connection,))
    receiver.daemon = True
    receiver.start()

    input_stream = sys.stdin.buffer
    try:
        while True:
            data = input_stream.readline()
            if not data:
                break
            connection.sendall(data)
        try:
            connection.shutdown(socket.SHUT_WR)
        except socket.error:
            pass
        receiver.join()
    finally:
        connection.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
