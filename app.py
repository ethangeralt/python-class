#!/usr/bin/env python3
"""
Python Version Mismatch Demonstration App
Designed for Docker classroom training.
Shows why containerization is essential to avoid "works on my machine" issues.
"""

import sys
import os
import platform
from http.server import HTTPServer, BaseHTTPRequestHandler

# ==========================================
# 1. VERSION CHECK & ENVIRONMENT VALIDATION
# ==========================================
TARGET_MAJOR = 3
TARGET_MINOR = 9

current_version = f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"
is_python_39 = (sys.version_info.major == TARGET_MAJOR and sys.version_info.minor == TARGET_MINOR)

print("=" * 70)
print(f"🚀  Starting Application...")
print(f"📌  Detected Python Version : {current_version}")
print(f"📌  Operating System        : {platform.system()} ({platform.release()})")
print("=" * 70)

# Check 1: Strict Python Version Enforce
if not is_python_39:
    print("\n" + "!" * 70)
    print("❌ FATAL ERROR: INCOMPATIBLE PYTHON VERSION!")
    print(f"   Required: Python {TARGET_MAJOR}.{TARGET_MINOR}.x (e.g., Python 3.9)")
    print(f"   Found   : Python {current_version}")
    print("\n   [Why this broke]:")
    print("   This application uses legacy standard libraries and binary bindings")
    print("   compiled strictly for Python 3.9.")
    print("\n   [The Classroom Lesson]:")
    print("   Your host machine has a different Python version installed.")
    print("   Instead of forcing every developer/server to downgrade their system Python,")
    print("   Docker provides a clean, isolated Python 3.9 environment!")
    print("!" * 70 + "\n")
    sys.exit(1)

# Check 2: Real-world breaking change simulation (Removed in Python 3.13)
# In Python 3.13, 'cgi' was permanently removed from Python standard library (PEP 594).
# It only works on Python <= 3.12 (and works perfectly in Python 3.9).
try:
    import cgi  # Realistic dependency that breaks across Python versions
except ImportError as e:
    print(f"❌ FATAL ERROR: Missing legacy module: {e}")
    print("   Reason: Module 'cgi' was removed in Python 3.13+ standard library.")
    sys.exit(1)

print("✅ Environment Check Passed: Running on certified Python 3.9 runtime!")
print("✅ Legacy module compatibility verified (cgi loaded successfully).")


# ==========================================
# 2. APPLICATION WEB SERVER (Zero Dependencies)
# ==========================================
PORT = int(os.getenv("PORT", 8080))

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Docker Classroom Demo - Python Version Parity</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            background: #0f172a;
            color: #f8fafc;
            display: flex;
            justify-content: center;
            align-items: center;
            min-height: 100vh;
            padding: 20px;
        }
        .card {
            background: #1e293b;
            border: 1px solid #334155;
            border-radius: 16px;
            max-width: 650px;
            width: 100%;
            padding: 36px;
            box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.5);
        }
        .badge {
            display: inline-block;
            background: #059669;
            color: #ecfdf5;
            padding: 6px 14px;
            border-radius: 9999px;
            font-size: 13px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            margin-bottom: 16px;
        }
        h1 {
            font-size: 26px;
            font-weight: 700;
            margin-bottom: 12px;
            color: #38bdf8;
        }
        p.subtitle {
            color: #94a3b8;
            font-size: 15px;
            line-height: 1.6;
            margin-bottom: 24px;
        }
        .info-grid {
            background: #0f172a;
            border: 1px solid #1e293b;
            border-radius: 10px;
            padding: 16px;
            margin-bottom: 24px;
        }
        .info-row {
            display: flex;
            justify-content: space-between;
            padding: 8px 0;
            border-bottom: 1px solid #1e293b;
            font-size: 14px;
        }
        .info-row:last-child {
            border-bottom: none;
        }
        .info-label {
            color: #64748b;
            font-weight: 500;
        }
        .info-val {
            color: #e2e8f0;
            font-family: monospace;
            font-weight: 600;
        }
        .takeaway {
            background: rgba(56, 189, 248, 0.1);
            border-left: 4px solid #38bdf8;
            padding: 14px 16px;
            border-radius: 4px;
            font-size: 14px;
            line-height: 1.5;
            color: #bae6fd;
        }
    </style>
</head>
<body>
    <div class="card">
        <div class="badge">Success &bull; Isolated Runtime</div>
        <h1>🐳 App Running in Python 3.9</h1>
        <p class="subtitle">
            This application failed on the host system because of a Python version mismatch,
            but runs reliably inside this Docker container.
        </p>

        <div class="info-grid">
            <div class="info-row">
                <span class="info-label">Active Python Version:</span>
                <span class="info-val">{python_version}</span>
            </div>
            <div class="info-row">
                <span class="info-label">Runtime Environment:</span>
                <span class="info-val">{container_status}</span>
            </div>
            <div class="info-row">
                <span class="info-label">Operating System:</span>
                <span class="info-val">{os_info}</span>
            </div>
            <div class="info-row">
                <span class="info-label">Process ID:</span>
                <span class="info-val">{pid}</span>
            </div>
        </div>

        <div class="takeaway">
            <strong>Key Concept for Students:</strong><br>
            "It works on my machine" is solved by bundling the exact runtime (Python 3.9)
            and all dependencies into an immutable Docker image.
        </div>
    </div>
</body>
</html>
"""

class DemoHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        is_docker = os.path.exists("/.dockerenv")
        container_status = "🐳 Docker Container (Isolated)" if is_docker else "💻 Local Host Machine"
        os_info = f"{platform.system()} ({platform.release()})"

        html_content = HTML_TEMPLATE.format(
            python_version=current_version,
            container_status=container_status,
            os_info=os_info,
            pid=os.getpid()
        )

        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(html_content.encode("utf-8"))))
        self.end_headers()
        self.wfile.write(html_content.encode("utf-8"))

    def log_message(self, format, *args):
        print(f"[{self.log_date_time_string()}] HTTP Request: {args[0]}")

def run_server():
    server_address = ("", PORT)
    httpd = HTTPServer(server_address, DemoHandler)
    print(f"🌐 Server started at http://localhost:{PORT}")
    print(f"💡 Press Ctrl+C to stop the server.")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n🛑 Server stopped.")

if __name__ == "__main__":
    run_server()
