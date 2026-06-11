from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
import argparse
import datetime
import hashlib
import json
import socket
import sys
import threading
import traceback


class TFMThreadingHTTPServer(ThreadingHTTPServer):
    allow_reuse_address = True
    daemon_threads = True


class UploadHandler(BaseHTTPRequestHandler):
    server_version = "TFMReceiver/4.0"
    protocol_version = "HTTP/1.1"

    def setup(self):
        super().setup()
        try:
            self.connection.settimeout(self.server.socket_timeout)
        except Exception:
            pass

    def handle_expect_100(self):
        self.send_response_only(100)
        self.end_headers()
        return True

    def _utc_timestamp(self):
        return datetime.datetime.now(datetime.timezone.utc).isoformat()

    def _append_log(self, record):
        record.setdefault("timestamp", self._utc_timestamp())
        self.server.log_file.parent.mkdir(parents=True, exist_ok=True)
        with self.server.log_lock:
            with open(self.server.log_file, "a", encoding="utf-8") as f:
                f.write(json.dumps(record, ensure_ascii=False) + "\n")
                f.flush()

    def _send_body(self, code, body, content_type):
        if isinstance(body, str):
            raw_body = body.encode("utf-8")
        else:
            raw_body = body

        self.close_connection = True
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(raw_body)))
        self.send_header("Connection", "close")
        self.end_headers()
        if raw_body:
            self.wfile.write(raw_body)
        self.wfile.flush()

    def _send_text(self, code, text):
        self._send_body(code, text, "text/plain; charset=utf-8")

    def _send_json(self, code, payload):
        body = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
        self._send_body(code, body, "application/json; charset=utf-8")

    def _log_exception(self, phase, exc):
        try:
            self._append_log({
                "event": "receiver_exception",
                "phase": phase,
                "client_ip": self.client_address[0] if self.client_address else "",
                "client_port": self.client_address[1] if self.client_address else "",
                "method": getattr(self, "command", ""),
                "path": getattr(self, "path", ""),
                "error": str(exc),
                "traceback": traceback.format_exc(),
            })
        except Exception:
            pass

    def do_GET(self):
        try:
            if self.path.split("?", 1)[0] != "/health":
                self._send_text(404, "Not Found\n")
                return

            self._append_log({
                "event": "health",
                "client_ip": self.client_address[0],
                "client_port": self.client_address[1],
                "method": self.command,
                "path": self.path,
                "bytes_received": 0,
                "saved_file": "",
                "sha256": "",
                "user_agent": self.headers.get("User-Agent", ""),
                "content_type": self.headers.get("Content-Type", ""),
            })
            self._send_text(200, "OK\n")
        except (BrokenPipeError, ConnectionResetError, socket.timeout) as exc:
            self._log_exception("GET", exc)
        except Exception as exc:
            self._log_exception("GET", exc)
            try:
                self._send_json(500, {"status": "error", "error": str(exc)})
            except Exception:
                pass

    def do_PUT(self):
        self._receive_upload()

    def do_POST(self):
        self._receive_upload()

    def _receive_upload(self):
        saved_file = ""
        try:
            clean_path = self.path.split("?", 1)[0]
            if clean_path != "/upload":
                self._send_text(404, "Not Found\n")
                return

            raw_length = self.headers.get("Content-Length")
            if raw_length is None:
                self._send_json(411, {"status": "error", "error": "Missing Content-Length"})
                return

            try:
                content_length = int(raw_length)
            except ValueError:
                self._send_json(400, {"status": "error", "error": "Invalid Content-Length"})
                return

            if content_length < 0:
                self._send_json(400, {"status": "error", "error": "Negative Content-Length"})
                return
            if self.server.max_bytes and content_length > self.server.max_bytes:
                self._send_json(413, {
                    "status": "error",
                    "error": "Payload too large",
                    "max_bytes": self.server.max_bytes,
                })
                return

            body = self.rfile.read(content_length)
            if len(body) != content_length:
                self._send_json(400, {
                    "status": "error",
                    "error": "Incomplete request body",
                    "expected": content_length,
                    "received": len(body),
                })
                return

            now = datetime.datetime.now(datetime.timezone.utc)
            stem = now.strftime("upload_%Y%m%d_%H%M%S_%f")
            content_type = self.headers.get("Content-Type", "")
            suffix = ".zip" if "zip" in content_type.lower() else ".bin"

            self.server.out_dir.mkdir(parents=True, exist_ok=True)
            target = self.server.out_dir / f"{stem}{suffix}"
            target.write_bytes(body)
            saved_file = str(target)

            sha256 = hashlib.sha256(body).hexdigest()
            record = {
                "timestamp": now.isoformat(),
                "client_ip": self.client_address[0],
                "client_port": self.client_address[1],
                "method": self.command,
                "path": self.path,
                "bytes_received": len(body),
                "saved_file": saved_file,
                "sha256": sha256,
                "user_agent": self.headers.get("User-Agent", ""),
                "content_type": content_type,
            }
            self._append_log(record)

            with self.server.upload_lock:
                self.server.upload_count += 1
                should_shutdown = (
                    self.server.max_uploads > 0
                    and self.server.upload_count >= self.server.max_uploads
                )

            self._send_json(200, {
                "status": "ok",
                "bytes_received": len(body),
                "saved_file": saved_file,
                "sha256": sha256,
            })

            if should_shutdown:
                threading.Thread(target=self.server.shutdown, daemon=True).start()

        except (BrokenPipeError, ConnectionResetError, socket.timeout) as exc:
            self._log_exception("UPLOAD", exc)
        except Exception as exc:
            self._log_exception("UPLOAD", exc)
            try:
                self._send_json(500, {
                    "status": "error",
                    "error": str(exc),
                    "saved_file": saved_file,
                })
            except Exception:
                pass

    def log_message(self, fmt, *args):
        line = "[{0}] {1} - {2}\n".format(
            datetime.datetime.now(datetime.timezone.utc).isoformat(),
            self.address_string(),
            fmt % args,
        )
        sys.stdout.write(line)
        sys.stdout.flush()


def main():
    parser = argparse.ArgumentParser(description="TFM TEC-009 controlled HTTP receiver")
    parser.add_argument("--host", default="0.0.0.0")
    parser.add_argument("--port", type=int, default=8000)
    parser.add_argument("--out-dir", default="received")
    parser.add_argument("--log-file", default="received/receiver_log.jsonl")
    parser.add_argument("--max-seconds", type=int, default=600)
    parser.add_argument("--max-uploads", type=int, default=5)
    parser.add_argument("--max-bytes", type=int, default=104857600)
    parser.add_argument("--socket-timeout", type=int, default=120)
    args = parser.parse_args()

    out_dir = Path(args.out_dir)
    log_file = Path(args.log_file)
    out_dir.mkdir(parents=True, exist_ok=True)
    log_file.parent.mkdir(parents=True, exist_ok=True)

    server = TFMThreadingHTTPServer((args.host, args.port), UploadHandler)
    server.out_dir = out_dir
    server.log_file = log_file
    server.max_uploads = args.max_uploads
    server.max_bytes = args.max_bytes
    server.socket_timeout = args.socket_timeout
    server.upload_count = 0
    server.upload_lock = threading.Lock()
    server.log_lock = threading.Lock()

    startup = {
        "event": "receiver_start",
        "timestamp": datetime.datetime.now(datetime.timezone.utc).isoformat(),
        "host": args.host,
        "port": args.port,
        "out_dir": str(out_dir.resolve()),
        "log_file": str(log_file.resolve()),
        "max_seconds": args.max_seconds,
        "max_uploads": args.max_uploads,
        "max_bytes": args.max_bytes,
    }
    with open(log_file, "a", encoding="utf-8") as f:
        f.write(json.dumps(startup, ensure_ascii=False) + "\n")
        f.flush()

    print(f"TFM receiver v4 listening on http://{args.host}:{args.port}/upload")
    print(f"Health URL: http://{args.host}:{args.port}/health")
    print(f"Output directory: {out_dir.resolve()}")
    print(f"Log file: {log_file.resolve()}")
    print(f"Limits: max_seconds={args.max_seconds}, max_uploads={args.max_uploads}, max_bytes={args.max_bytes}")
    sys.stdout.flush()

    timer = threading.Timer(args.max_seconds, server.shutdown)
    timer.daemon = True
    timer.start()

    try:
        server.serve_forever()
    finally:
        timer.cancel()
        server.server_close()
        shutdown = {
            "event": "receiver_stop",
            "timestamp": datetime.datetime.now(datetime.timezone.utc).isoformat(),
            "uploads": server.upload_count,
        }
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(json.dumps(shutdown, ensure_ascii=False) + "\n")
            f.flush()
        print("TFM receiver v4 stopped.")


if __name__ == "__main__":
    main()
