from __future__ import print_function

import socket
import sys
import threading


BRIDGE_CLOSE_TIMEOUT_SECONDS = 2.0


def socket_to_stdout(connection, output_stream, closing):
    try:
        while True:
            data = connection.recv(4096)
            if not data:
                break
            output_stream.write(data)
            output_stream.flush()
    except Exception as error:
        if not closing.is_set():
            sys.stderr.write("Clippy MCP bridge receive error: {0}\n".format(error))
            sys.stderr.flush()


def forward(
    connection,
    input_stream=None,
    output_stream=None,
    close_timeout=BRIDGE_CLOSE_TIMEOUT_SECONDS,
):
    active_input = input_stream if input_stream is not None else sys.stdin.buffer
    active_output = (
        output_stream if output_stream is not None else sys.stdout.buffer
    )
    closing = threading.Event()
    receiver = threading.Thread(
        target=socket_to_stdout,
        args=(connection, active_output, closing),
    )
    receiver.daemon = True
    receiver.start()

    try:
        while True:
            data = active_input.readline()
            if not data:
                break
            connection.sendall(data)
    finally:
        closing.set()
        try:
            connection.shutdown(socket.SHUT_WR)
        except socket.error:
            pass

        receiver.join(close_timeout)
        if receiver.is_alive():
            try:
                connection.shutdown(socket.SHUT_RDWR)
            except socket.error:
                pass

        connection.close()
        if receiver.is_alive():
            receiver.join(close_timeout)


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
    forward(connection)
    return 0


if __name__ == "__main__":
    sys.exit(main())
