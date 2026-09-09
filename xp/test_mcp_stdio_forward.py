from __future__ import print_function

import io
import socket
import threading
import unittest

from xp.mcp_stdio_forward import forward


class FakeConnection(object):
    def __init__(self, responses=None, block_receive=False):
        self.responses = list(responses or [])
        self.block_receive = block_receive
        self.receive_released = threading.Event()
        self.sent = []
        self.shutdown_calls = []
        self.closed = False

    def recv(self, _size):
        if self.block_receive:
            self.receive_released.wait(1.0)
            return b""
        if self.responses:
            return self.responses.pop(0)
        return b""

    def sendall(self, data):
        self.sent.append(data)

    def shutdown(self, how):
        self.shutdown_calls.append(how)
        if how == socket.SHUT_RDWR:
            self.receive_released.set()

    def close(self):
        self.closed = True
        self.receive_released.set()


class BridgeCleanupTests(unittest.TestCase):
    def test_eof_half_closes_write_and_forwards_remaining_output(self):
        connection = FakeConnection([b'{"result":"ok"}\n', b""])
        output = io.BytesIO()

        forward(connection, io.BytesIO(b'{"method":"ping"}\n'), output, 0.05)

        self.assertEqual([b'{"method":"ping"}\n'], connection.sent)
        self.assertEqual(b'{"result":"ok"}\n', output.getvalue())
        self.assertEqual([socket.SHUT_WR], connection.shutdown_calls)
        self.assertTrue(connection.closed)

    def test_eof_forces_full_close_when_server_does_not_finish(self):
        connection = FakeConnection(block_receive=True)

        forward(connection, io.BytesIO(b""), io.BytesIO(), 0.01)

        self.assertEqual(
            [socket.SHUT_WR, socket.SHUT_RDWR],
            connection.shutdown_calls,
        )
        self.assertTrue(connection.closed)


if __name__ == "__main__":
    unittest.main()
