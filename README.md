# 🐳 Docker Classroom Demo: Python Version Mismatch ("Works On My Machine")

This sample project is designed for **Docker instructors** to demonstrate to students why containerization is essential in modern software engineering.

---

## 🎯 The Educational Goal
Show students:
1. What happens when a project has strict runtime dependencies (e.g., **Python 3.9**) and a developer or production server has a different version installed (e.g., **Python 3.10, 3.11, 3.12, or 3.13**).
2. How Docker eliminates the *"it works on my machine"* problem by packaging the runtime, standard libraries, and application together.

---

## 📁 Project Structure

```text
python-version-demo/
├── app.py              # Zero-dependency Python app (enforces Python 3.9 & checks legacy modules)
├── Dockerfile          # Builds an isolated Python 3.9 runtime
├── docker-compose.yml  # Optional: Compose definition for one-command startup
└── README.md           # Classroom walkthrough guide
```

---

## 🧑‍🏫 Classroom Walkthrough Guide (Step-by-Step)

### Step 1: Run it directly on the Host (The Failure State ❌)
Ask students to clone/open this directory and run the application using their local Python interpreter:

```bash
python3 app.py
```

#### What happens?
If their laptop runs Python 3.8, 3.10, 3.11, 3.12, or 3.13, the app **immediately crashes**:

```text
======================================================================
🚀  Starting Application...
📌  Detected Python Version : 3.13.7
📌  Operating System        : Darwin (25.2.0)
======================================================================

!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
❌ FATAL ERROR: INCOMPATIBLE PYTHON VERSION!
   Required: Python 3.9.x (e.g., Python 3.9)
   Found   : Python 3.13.7

   [Why this broke]:
   This application uses legacy standard libraries and binary bindings
   compiled strictly for Python 3.9.

   [The Classroom Lesson]:
   Your host machine has a different Python version installed.
   Instead of forcing every developer/server to downgrade their system Python,
   Docker provides a clean, isolated Python 3.9 environment!
!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
```

#### Instructor Discussion Point:
- Ask students: *"How would you fix this without Docker?"*
  - Common answers: *pyenv*, *conda*, downgrading the operating system's Python.
  - Highlight the problem: Managing multiple Python installs on laptops and production servers is tedious, error-prone, and leads to deployment drift.

---

### Step 2: Containerize the Application with Docker (The Solution 🐳)
Show students the `Dockerfile`:

```dockerfile
FROM python:3.9-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PORT=8080

WORKDIR /app
COPY app.py .
EXPOSE 8080

CMD ["python", "app.py"]
```

Explain:
- `FROM python:3.9-slim`: We are telling Docker to pull an official Linux image that has **Python 3.9** pre-installed.
- No matter what OS or Python version the student's laptop is running (Windows, macOS, Ubuntu, Python 3.13), the container will **always run Python 3.9**.

---

### Step 3: Build the Docker Image

Run:

```bash
docker build -t python-39-demo .
```

---

### Step 4: Run the Docker Container (The Success State ✅)

Run:

```bash
docker run -p 8080:8080 python-39-demo
```

Output:
```text
======================================================================
🚀  Starting Application...
📌  Detected Python Version : 3.9.21
📌  Operating System        : Linux (...)
======================================================================
✅ Environment Check Passed: Running on certified Python 3.9 runtime!
✅ Legacy module compatibility verified (cgi loaded successfully).
🌐 Server started at http://localhost:8080
```

---

### Step 5: Test in the Browser 🌐

Have students open:
[http://localhost:8080](http://localhost:8080)

They will see a clean dashboard showing:
- **Active Python Version**: 3.9.x
- **Runtime Environment**: 🐳 Docker Container (Isolated)
- **Status**: Passed

---

## 💡 Key Classroom Takeaways

1. **Isolation**: Docker isolates the runtime environment from the host system.
2. **Reproducibility**: The exact same image will run identically on macOS, Windows, Linux, and Kubernetes.
3. **No Pollution**: The student did not have to alter their host's Python setup.
