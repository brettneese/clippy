from __future__ import print_function

import socket
import sys
import threading

try:
    from .mcp_diagnostics import BRIDGE_LOG_PATH
    from .mcp_diagnostics import DiagnosticLog
    from .mcp_diagnostics import emit as _emit_diagnostic
    from .mcp_diagnostics import protocol_metadata
except (ImportError, ValueError, SystemError):
    from mcp_diagnostics import BRIDGE_LOG_PATH
    from mcp_diagnostics import DiagnosticLog
    from mcp_diagnostics import emit as _emit_diagnostic
    from mcp_diagnostics import protocol_metadata


BRIDGE_CLOSE_TIMEOUT_SECONDS = 2.0


def socket_to_stdout(connection, output_stream, closing, diagnostics=None):
    try:
        while True:
            _emit_diagnostic(diagnostics, "bridge_tcp_receive_wait")
            data = connection.recv(4096)
            if not data:
                _emit_diagnostic(diagnostics, "bridge_tcp_eof")
                break
            _emit_diagnostic(
                diagnostics,
                "bridge_tcp_received",
                bytes=len(data),
            )
            _emit_diagnostic(
                diagnostics,
                "bridge_stdout_write_start",
                bytes=len(data),
            )
            output_stream.write(data)
            output_stream.flush()
            _emit_diagnostic(
                diagnostics,
                "bridge_stdout_write_complete",
                bytes=len(data),
            )
    except Exception as error:
        if not closing.is_set():
            _emit_diagnostic(
                diagnostics,
                "bridge_receive_error",
                error_type=type(error).__name__,
            )
            sys.stderr.write("Clippy MCP bridge receive error: {0}\n".format(error))
            sys.stderr.flush()


def forward(
    connection,
    input_stream=None,
    output_stream=None,
    close_timeout=BRIDGE_CLOSE_TIMEOUT_SECONDS,
    diagnostics=None,
):
    active_input = input_stream if input_stream is not None else sys.stdin.buffer
    active_output = (
        output_stream if output_stream is not None else sys.stdout.buffer
    )
    closing = threading.Event()
    receiver = threading.Thread(
        target=socket_to_stdout,
        args=(connection, active_output, closing, diagnostics),
    )
    receiver.daemon = True
    receiver.start()

    try:
        while True:
            _emit_diagnostic(diagnostics, "bridge_stdin_read_wait")
            data = active_input.readline()
            if not data:
                _emit_diagnostic(diagnostics, "bridge_stdin_eof")
                break
            metadata = protocol_metadata(data)
            metadata["bytes"] = len(data)
            _emit_diagnostic(diagnostics, "bridge_stdin_received", **metadata)
            _emit_diagnostic(
                diagnostics,
                "bridge_tcp_send_start",
                bytes=len(data),
            )
            connection.sendall(data)
            _emit_diagnostic(
                diagnostics,
                "bridge_tcp_send_complete",
                bytes=len(data),
            )
    except Exception as error:
        _emit_diagnostic(
            diagnostics,
            "bridge_forward_error",
            error_type=type(error).__name__,
        )
        raise
    finally:
        closing.set()
        _emit_diagnostic(diagnostics, "bridge_shutdown_write_start")
        try:
            connection.shutdown(socket.SHUT_WR)
        except socket.error:
            pass
        _emit_diagnostic(diagnostics, "bridge_shutdown_write_complete")

        _emit_diagnostic(diagnostics, "bridge_receiver_join_start")
        receiver.join(close_timeout)
        if receiver.is_alive():
            _emit_diagnostic(diagnostics, "bridge_receiver_force_close")
            try:
                connection.shutdown(socket.SHUT_RDWR)
            except socket.error:
                pass

        _emit_diagnostic(diagnostics, "bridge_socket_close_start")
        connection.close()
        _emit_diagnostic(diagnostics, "bridge_socket_close_complete")
        if receiver.is_alive():
            receiver.join(close_timeout)
        _emit_diagnostic(
            diagnostics,
            "bridge_close_complete",
            receiver_alive=receiver.is_alive(),
        )


def main():
    diagnostics = DiagnosticLog(BRIDGE_LOG_PATH)
    _emit_diagnostic(diagnostics, "bridge_process_start")
    if len(sys.argv) != 3:
        _emit_diagnostic(diagnostics, "bridge_usage_error")
        sys.stderr.write("usage: mcp_stdio_forward.py HOST PORT\n")
        return 2

    host = sys.argv[1]
    try:
        port = int(sys.argv[2])
    except ValueError:
        _emit_diagnostic(diagnostics, "bridge_port_error")
        sys.stderr.write("PORT must be an integer\n")
        return 2

    _emit_diagnostic(
        diagnostics,
        "bridge_connect_start",
        host=host,
        port=port,
    )
    try:
        connection = socket.create_connection((host, port), 10)
    except Exception as error:
        _emit_diagnostic(
            diagnostics,
            "bridge_connect_error",
            error_type=type(error).__name__,
        )
        raise
    _emit_diagnostic(diagnostics, "bridge_connect_complete")
    forward(connection, diagnostics=diagnostics)
    _emit_diagnostic(diagnostics, "bridge_process_complete")
    return 0


if __name__ == "__main__":
    sys.exit(main())
