#!/usr/bin/env python3
"""
Generate the 'Docker Fundamentals' PowerPoint presentation.
Guarantees 100% faithful adherence to the exact content provided by the user.
Adds:
1. Title slide: 'Docker Fundamentals' + 'The Journey of Sarah: From Deployment Chaos to Container Clarity'
2. Slide with 'Road-map' title and embedded screenshot
3. Agenda / Topics slide in the beginning matching the exact topic list
4. Exact verbatim content across all technical topic slides
"""

import os
import pptx
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

OUTPUT_PPTX = "/Users/aanand-mac/Downloads/Docker_Fundamentals.pptx"
SCREENSHOT_PATH = "/var/folders/3q/jn41f3x56y922gqq0_02kxl80000gn/T/TemporaryItems/NSIRD_screencaptureui_egVht0/Screenshot 2026-11-18 at 8.23.12 PM.png"

# Color Palette Definitions (Dark Modern Tech Style)
COLOR_BG = RGBColor(15, 23, 42)          # Deep Slate/Navy Background
COLOR_CARD_BG = RGBColor(30, 41, 59)     # Slate 800 Card
COLOR_CARD_BORDER = RGBColor(51, 65, 85) # Slate 700 Border
COLOR_TEXT_MAIN = RGBColor(248, 250, 252)# White Header
COLOR_TEXT_MUTED = RGBColor(148, 163, 184)# Slate 400 Subtitle
COLOR_TEXT_BODY = RGBColor(226, 232, 240)# Light Body Text

# Accent Colors
COLOR_ACCENT_CYAN = RGBColor(56, 189, 248)
COLOR_ACCENT_BLUE = RGBColor(14, 165, 233)
COLOR_ACCENT_GREEN = RGBColor(16, 185, 129)
COLOR_ACCENT_AMBER = RGBColor(245, 158, 11)
COLOR_ACCENT_PURPLE = RGBColor(168, 85, 247)

prs = pptx.Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

def set_slide_background(slide):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = COLOR_BG
    bg.line.fill.background()
    return bg

def add_header(slide, title_text, subtitle_text=None):
    tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(1.1))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

    p_title = tf.paragraphs[0]
    p_title.text = title_text
    p_title.font.name = "Arial"
    p_title.font.size = Pt(24)
    p_title.font.bold = True
    p_title.font.color.rgb = COLOR_TEXT_MAIN

    if subtitle_text:
        p_title.space_after = Pt(3)
        p_sub = tf.add_paragraph()
        p_sub.text = subtitle_text
        p_sub.font.name = "Arial"
        p_sub.font.size = Pt(13)
        p_sub.font.color.rgb = COLOR_ACCENT_CYAN
    else:
        p_title.space_after = Pt(8)

def add_panel(slide, left, top, width, height, title, content_lines, accent_color=COLOR_ACCENT_CYAN):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = COLOR_CARD_BG
    card.line.color.rgb = COLOR_CARD_BORDER
    card.line.width = Pt(1.2)

    tb = slide.shapes.add_textbox(left + Inches(0.25), top + Inches(0.2), width - Inches(0.5), height - Inches(0.4))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

    p_first = tf.paragraphs[0]
    p_first.text = title
    p_first.font.name = "Arial"
    p_first.font.size = Pt(16)
    p_first.font.bold = True
    p_first.font.color.rgb = accent_color
    p_first.space_after = Pt(8)

    for line in content_lines:
        p = tf.add_paragraph()
        p.space_after = Pt(5)
        p.font.name = "Arial"
        p.font.size = Pt(11.5)

        line_str = str(line)
        if line_str.startswith("- ") or line_str.startswith("• "):
            p.text = "• " + line_str[2:]
            p.font.color.rgb = COLOR_TEXT_BODY
        elif ":" in line_str and not line_str.startswith("http") and not line_str.startswith("Step"):
            parts = line_str.split(":", 1)
            r1 = p.add_run()
            r1.text = parts[0] + ":"
            r1.font.bold = True
            r1.font.color.rgb = COLOR_TEXT_MAIN
            r2 = p.add_run()
            r2.text = parts[1]
            r2.font.bold = False
            r2.font.color.rgb = COLOR_TEXT_BODY
        else:
            p.text = line_str
            if line_str.startswith("Step") or line_str.startswith("Rule") or line_str.startswith("Phase"):
                p.font.bold = True
                p.font.color.rgb = COLOR_ACCENT_AMBER
            else:
                p.font.color.rgb = COLOR_TEXT_BODY


# ==============================================================================
# SLIDE 1: TITLE SLIDE
# ==============================================================================
s1 = prs.slides.add_slide(blank_layout)
set_slide_background(s1)

tb1 = s1.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(11.333), Inches(4.5))
tf1 = tb1.text_frame
tf1.word_wrap = True

p1 = tf1.paragraphs[0]
p1.text = "Docker Fundamentals"
p1.font.name = "Arial"
p1.font.size = Pt(44)
p1.font.bold = True
p1.font.color.rgb = COLOR_TEXT_MAIN
p1.space_after = Pt(12)

p2 = tf1.add_paragraph()
p2.text = "The Journey of Sarah: From Deployment Chaos to Container Clarity"
p2.font.name = "Arial"
p2.font.size = Pt(20)
p2.font.bold = True
p2.font.color.rgb = COLOR_ACCENT_AMBER
p2.space_after = Pt(18)

p3 = tf1.add_paragraph()
p3.text = "Meet Sarah, a developer who just finished building an awesome web application.\nHer journey will teach us everything about Docker."
p3.font.name = "Arial"
p3.font.size = Pt(15)
p3.font.color.rgb = COLOR_TEXT_BODY


# ==============================================================================
# SLIDE 2: ROAD-MAP SLIDE (With Screenshot)
# ==============================================================================
s2 = prs.slides.add_slide(blank_layout)
set_slide_background(s2)
add_header(s2, "Road-map", "Master curriculum overview from Docker foundations to Kubernetes orchestration and observability.")

if os.path.exists(SCREENSHOT_PATH):
    # Image aspect ratio 317x705 -> width ~ 2.38 in for height 5.3 in
    s2.shapes.add_picture(SCREENSHOT_PATH, Inches(0.8), Inches(1.6), height=Inches(5.4))

add_panel(s2, Inches(3.6), Inches(1.6), Inches(8.9), Inches(5.4),
          "Course Curriculum Journey",
          [
              "Today's Module: Class 1 - Docker Introduction & Container Core Architecture",
              "- Docker Introduction & Deep Kernel Foundations",
              "- Docker Image and Docker Containers",
              "- Building Your Own Docker Image",
              "- Image and Containers Deep Dive",
              "- Docker Volume & Persistent Storage",
              "- Docker Networking Deep Dive",
              "- Hardening the Docker Environment (DevSecOps)",
              "- Introduction to Orchestration & Kubernetes Architecture",
              "- Kubernetes Pods, ReplicaSets, Deployments & Networking",
              "- Scaling in Kubernetes (HPA, CA, VPA) & RBAC Security",
              "- Ingress, Probes, CRD, CNI, ConfigMaps & Secrets",
              "- Deploying Hugging Face AI Models on Kubernetes",
              "- Observability: Prometheus Monitoring, PromQL, Grafana & New Relic"
          ],
          accent_color=COLOR_ACCENT_CYAN)


# ==============================================================================
# SLIDE 3: AGENDA / TOPICS LIST (As explicitly requested by user)
# ==============================================================================
s3 = prs.slides.add_slide(blank_layout)
set_slide_background(s3)
add_header(s3, "Course Topics & Agenda", "All slides will follow this systematic progression.")

add_panel(s3, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.4),
          "Part 1: The Foundations",
          [
              "- Docker Introduction",
              "- What is Docker",
              "- Why Docker Was Created",
              "- What is a Process in Linux",
              "- How Processes Run and Are Isolated",
              "- Containers vs Virtual Machines",
              "- Key Linux Kernel Features:",
              "    • 7.1 Namespaces: Isolation of View",
              "    • 7.2 cgroups: Resource Limits",
              "    • 7.3 Union Filesystems",
              "- How Docker Uses These Kernel Features"
          ],
          accent_color=COLOR_ACCENT_CYAN)

add_panel(s3, Inches(6.9), Inches(1.6), Inches(5.6), Inches(5.4),
          "Part 2: Deep Dive & Architecture",
          [
              "- Lifecycle of a Container Process",
              "- Basic Linux Commands to Understand Containers",
              "- Docker Architecture",
              "- Docker Engine vs CLI vs Daemon",
              "- Container Lifecycle Theory",
              "- Sarah's Aha Moment: Putting It All Together",
              "- Key Takeaways"
          ],
          accent_color=COLOR_ACCENT_AMBER)


# ==============================================================================
# SLIDE 4: 1. DOCKER INTRODUCTION
# ==============================================================================
s4 = prs.slides.add_slide(blank_layout)
set_slide_background(s4)
add_header(s4, "1. Docker Introduction", "Sarah's app works perfectly on her laptop. But when she tries to deploy it to the server, disaster strikes.")

add_panel(s4, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.4),
          "The Deployment Chaos",
          [
              "- 'It works on my machine!' becomes her daily phrase",
              "- Different Python versions cause crashes",
              "- Missing dependencies break everything",
              "- Team members can't run her code"
          ],
          accent_color=COLOR_ACCENT_AMBER)

add_panel(s4, Inches(6.9), Inches(1.6), Inches(5.6), Inches(5.4),
          "The Solution Needed",
          [
              "Sarah needs a solution that packages everything together.",
              "Enter Docker.",
              "- Packages app and all runtime dependencies into one portable unit",
              "- Eliminates environment differences between laptop and server",
              "- Allows team members to run code instantly without local setup headaches"
          ],
          accent_color=COLOR_ACCENT_GREEN)


# ==============================================================================
# SLIDE 5: 2. WHAT IS DOCKER
# ==============================================================================
s5 = prs.slides.add_slide(blank_layout)
set_slide_background(s5)
add_header(s5, "2. What is Docker", "Docker is a platform that packages applications with all their dependencies into standardized units called containers.")

add_panel(s5, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.4),
          "Think of It Like a Shipping Container",
          [
              "- A shipping container holds goods and can be moved anywhere (ship, truck, train)",
              "- A Docker container holds your app and can run anywhere (laptop, server, cloud)",
              "- Standardized dimensions and handling mechanisms mean the transport system doesn't need to know what is inside the box"
          ],
          accent_color=COLOR_ACCENT_CYAN)

add_panel(s5, Inches(6.9), Inches(1.6), Inches(5.6), Inches(5.4),
          "Key Insight: The Furnished Apartment",
          [
              "Docker containers are not just about your app—they include the OS libraries, dependencies, and configuration needed to run it.",
              "(Just like moving into a fully furnished apartment where you just bring your clothes).",
              "- Everything needed for execution is bundled inside",
              "- No need to install specific Python versions or libraries on the host system"
          ],
          accent_color=COLOR_ACCENT_AMBER)


# ==============================================================================
# SLIDE 6: 3. WHY DOCKER WAS CREATED
# ==============================================================================
s6 = prs.slides.add_slide(blank_layout)
set_slide_background(s6)
add_header(s6, "3. Why Docker Was Created", "Solving the 'Dependency Hell' problem and resource waste before containerization.")

add_panel(s6, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.4),
          "The 'Dependency Hell' Problem Before Docker",
          [
              "Before Docker, Sarah faced:",
              "- Environment mismatch: Dev uses Ubuntu, production uses CentOS",
              "- Version conflicts: App needs Python 3.8, but server has 3.6",
              "- Manual setup: Every new developer spends days setting up their environment",
              "- Resource waste: Running multiple apps requires multiple virtual machines"
          ],
          accent_color=COLOR_ACCENT_PURPLE)

add_panel(s6, Inches(6.9), Inches(1.6), Inches(5.6), Inches(5.4),
          "What Developers Wanted & The Solution",
          [
              "What developers wanted:",
              "- Package app + dependencies together",
              "- Run anywhere without changes",
              "- Lightweight (not heavy like VMs)",
              "- Fast startup times",
              "Docker solved this by using Linux kernel features smartly."
          ],
          accent_color=COLOR_ACCENT_GREEN)


# ==============================================================================
# SLIDE 7: 4. WHAT IS A PROCESS IN LINUX
# ==============================================================================
s7 = prs.slides.add_slide(blank_layout)
set_slide_background(s7)
add_header(s7, "4. What is a Process in Linux", "Sarah needs to understand processes first. A process is a running program in memory.")

add_panel(s7, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.4),
          "When Sarah Runs: python app.py",
          [
              "The kernel:",
              "- Loads python executable into RAM",
              "- Allocates memory for the process",
              "- Assigns a Process ID (PID)",
              "- Gives it CPU time to execute"
          ],
          accent_color=COLOR_ACCENT_CYAN)

add_panel(s7, Inches(6.9), Inches(1.6), Inches(5.6), Inches(5.4),
          "Key Characteristics & Linux Behavior",
          [
              "Key characteristics:",
              "- Every process has a PID (unique number)",
              "- Processes have a parent-child relationship",
              "- PID 1 is the first process (init/systemd)",
              "- Processes can see other processes",
              "- Processes share the same kernel",
              "Example: If Sarah runs ps aux, she sees ALL processes on the system—web servers, databases, other users' apps. This is normal Linux behavior."
          ],
          accent_color=COLOR_ACCENT_AMBER)


# ==============================================================================
# SLIDE 8: 5. HOW PROCESSES RUN AND ARE ISOLATED
# ==============================================================================
s8 = prs.slides.add_slide(blank_layout)
set_slide_background(s8)
add_header(s8, "5. How Processes Run and Are Isolated", "Traditional Linux offers no isolation. Sarah needs isolation across processes.")

add_panel(s8, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.4),
          "Traditional Linux: No Isolation",
          [
              "When Sarah runs two Python apps:",
              "- Both see all system processes",
              "- Both use the same network interfaces",
              "- Both access the same filesystem",
              "- One app can interfere with another",
              "This is a problem for:",
              "- Security (apps can spy on each other)",
              "- Resource conflicts (port 8080 already in use!)",
              "- Stability (one app crashes, affects others)"
          ],
          accent_color=COLOR_ACCENT_AMBER)

add_panel(s8, Inches(6.9), Inches(1.6), Inches(5.6), Inches(5.4),
          "What Sarah Needs: Kernel Isolation",
          [
              "What Sarah needs:",
              "- Each app thinks it's alone on the system",
              "- Each app has its own network namespace",
              "- Each app has controlled resource limits",
              "Docker provides this isolation using kernel features."
          ],
          accent_color=COLOR_ACCENT_GREEN)


# ==============================================================================
# SLIDE 9: 6. CONTAINERS VS VIRTUAL MACHINES
# ==============================================================================
s9 = prs.slides.add_slide(blank_layout)
set_slide_background(s9)
add_header(s9, "6. Containers vs Virtual Machines", "Sarah considers using VMs but learns the difference: VMs virtualize hardware, containers virtualize the OS.")

add_panel(s9, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.4),
          "Virtual Machines",
          [
              "- Full OS: Each VM runs complete Linux/Windows",
              "- Heavy: GBs of disk, minutes to boot",
              "- Isolation: Hardware-level via hypervisor",
              "- Overhead: Each VM has its own kernel",
              "Example: Running 10 apps needs 10 VMs = 10 kernels = 40GB+ RAM",
              "(Analogy: Building 10 entirely separate houses)"
          ],
          accent_color=COLOR_ACCENT_PURPLE)

add_panel(s9, Inches(6.9), Inches(1.6), Inches(5.6), Inches(5.4),
          "Containers",
          [
              "- Shared kernel: All containers use host Linux kernel",
              "- Lightweight: MBs of disk, milliseconds to start",
              "- Isolation: Process-level via namespaces",
              "- Efficient: 10 containers = 1 kernel = 2GB RAM",
              "(Analogy: Building 10 apartments inside one shared building)",
              "Key difference: VMs virtualize hardware, containers virtualize the OS."
          ],
          accent_color=COLOR_ACCENT_CYAN)


# ==============================================================================
# SLIDE 10: 7. KEY LINUX KERNEL FEATURES & 7.1 NAMESPACES
# ==============================================================================
s10 = prs.slides.add_slide(blank_layout)
set_slide_background(s10)
add_header(s10, "7. Key Linux Kernel Features & 7.1 Namespaces", "Docker containers are just Linux processes with special isolation. Three kernel features make this possible.")

add_panel(s10, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.4),
          "7.1 Namespaces: Isolation of View",
          [
              "Namespaces control what a process can SEE. (These are the soundproof walls and locked doors of your apartment).",
              "Think of namespaces as individual bubbles—each container lives in its own bubble.",
              "PID Namespace:",
              "- Container sees only its own processes",
              "- Inside container: process has PID 1",
              "- Outside container: same process has different PID (like 5421)",
              "- Prevents containers from killing each other's processes",
              "Example: Sarah's web app thinks it's the only thing running (PID 1), but the host sees it as PID 8374."
          ],
          accent_color=COLOR_ACCENT_CYAN)

add_panel(s10, Inches(6.9), Inches(1.6), Inches(5.6), Inches(5.4),
          "Network, Mount & Other Namespaces",
          [
              "Network Namespace:",
              "- Each container has its own network stack",
              "- Own IP addresses, routes, firewall rules",
              "- Containers can't see each other's network traffic",
              "- Prevents port conflicts (Two containers can both use port 8080 internally)",
              "Mount Namespace:",
              "- Each container sees its own filesystem",
              "- Container can't access host files (unless explicitly mounted)",
              "UTS Namespace: Each container has its own hostname ('web-server').",
              "IPC Namespace: Isolates inter-process communication (shared memory/semaphores).",
              "User Namespace: Maps container's root user to non-root user on host."
          ],
          accent_color=COLOR_ACCENT_GREEN)


# ==============================================================================
# SLIDE 11: 7.2 CGROUPS: RESOURCE LIMITS
# ==============================================================================
s11 = prs.slides.add_slide(blank_layout)
set_slide_background(s11)
add_header(s11, "7.2 cgroups: Resource Limits", "Namespaces control what you SEE, cgroups control what you USE.")

add_panel(s11, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.4),
          "cgroups (Control Groups) Overview",
          [
              "(These are the utility meters limiting how much water or electricity your specific apartment can use).",
              "Cgroups (control groups) limit and measure resource usage:",
              "Resource Limits:",
              "- CPU: Container can use max 50% of one CPU core",
              "- Memory: Container limited to 512MB RAM",
              "- Disk I/O: Container can read/write max 100MB/s",
              "- Network: Bandwidth limits",
              "Example: Sarah limits her web app to 1GB RAM. If it tries to use more, kernel kills it."
          ],
          accent_color=COLOR_ACCENT_AMBER)

add_panel(s11, Inches(6.9), Inches(1.6), Inches(5.6), Inches(5.4),
          "Resource Accounting & Importance",
          [
              "Resource Accounting:",
              "- Track how much CPU/memory each container uses",
              "- Enables monitoring and billing",
              "Why this matters:",
              "- Without cgroups, one container could consume all CPU and starve others.",
              "- Provides isolation and fair-sharing across all container workloads on the system"
          ],
          accent_color=COLOR_ACCENT_CYAN)


# ==============================================================================
# SLIDE 12: 7.3 UNION FILESYSTEMS
# ==============================================================================
s12 = prs.slides.add_slide(blank_layout)
set_slide_background(s12)
add_header(s12, "7.3 Union Filesystems", "Union filesystems allow layering multiple directories into one view.")

add_panel(s12, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.4),
          "Layering Concept & How It Works",
          [
              "(Like moving into a furnished apartment: the base furniture is locked in place, but you can add your own decorations in a separate layer).",
              "- Traditional filesystem: One directory structure",
              "- Union filesystem: Stack multiple read-only layers + one writable layer",
              "How it works:",
              "- Base layer: Ubuntu OS files (read-only)",
              "- Layer 2: Python installed (read-only)",
              "- Layer 3: Your app code (read-only)",
              "- Layer 4: Runtime changes (writable)",
              "Two main types: AUFS (older, stable) & OverlayFS (newer, faster, default in modern Docker)."
          ],
          accent_color=COLOR_ACCENT_PURPLE)

add_panel(s12, Inches(6.9), Inches(1.6), Inches(5.6), Inches(5.4),
          "Benefits of Union Filesystems",
          [
              "Benefits:",
              "- Sharing: 10 containers share the same Ubuntu base layer",
              "- Speed: Start containers instantly (layers already on disk)",
              "- Storage: Save disk space (no duplication)",
              "Example: Sarah has 5 containers using Ubuntu. Instead of 5 copies of Ubuntu (5GB), she has 1 copy shared by all (1GB)."
          ],
          accent_color=COLOR_ACCENT_GREEN)


# ==============================================================================
# SLIDE 13: 8. HOW DOCKER USES THESE KERNEL FEATURES
# ==============================================================================
s13 = prs.slides.add_slide(blank_layout)
set_slide_background(s13)
add_header(s13, "8. How Docker Uses These Kernel Features", "When Sarah runs a container, Docker orchestrates kernel features step-by-step.")

add_panel(s13, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.4),
          "Step 1 & Step 2",
          [
              "Step 1: Create Namespaces",
              "- PID namespace → container sees only its processes",
              "- Network namespace → container gets own IP",
              "- Mount namespace → container gets own filesystem view",
              "- UTS namespace → container gets own hostname",
              "Step 2: Apply cgroups",
              "- CPU limit: max 1 core",
              "- Memory limit: max 512MB",
              "- Disk I/O: limit write speed"
          ],
          accent_color=COLOR_ACCENT_CYAN)

add_panel(s13, Inches(6.9), Inches(1.6), Inches(5.6), Inches(5.4),
          "Step 3, Step 4 & Result",
          [
              "Step 3: Set up Union Filesystem",
              "- Layer 1: Base OS (shared)",
              "- Layer 2: Dependencies (shared)",
              "- Layer 3: App code (shared)",
              "- Layer 4: Container writes (unique)",
              "Step 4: Start Process",
              "- Launch app as PID 1 inside container",
              "- Process thinks it's alone on the system",
              "- Actually isolated by namespaces + cgroups",
              "Result: A lightweight, isolated environment that starts in milliseconds."
          ],
          accent_color=COLOR_ACCENT_GREEN)


# ==============================================================================
# SLIDE 14: 9. LIFECYCLE OF A CONTAINER PROCESS
# ==============================================================================
s14 = prs.slides.add_slide(blank_layout)
set_slide_background(s14)
add_header(s14, "9. Lifecycle of a Container Process", "Sarah learns that containers are just processes with a lifecycle.")

add_panel(s14, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.4),
          "Container States & Process Flow",
          [
              "States:",
              "- Created: Container exists but not started. Namespaces prepared, filesystem ready. No running process yet.",
              "- Running: Main process (PID 1) is executing. Container is 'alive'. Uses CPU, memory, network.",
              "- Paused: Process frozen (via cgroups). Still in memory but not executing. Can resume quickly.",
              "- Stopped: Main process exited or killed. Container still exists (filesystem preserved). Can be restarted.",
              "- Deleted: Container removed completely. Filesystem destroyed (unless volumes used). Cannot be recovered.",
              "Process Flow: docker run → Create namespaces + cgroups → Start PID 1 → Running"
          ],
          accent_color=COLOR_ACCENT_CYAN)

add_panel(s14, Inches(6.9), Inches(1.6), Inches(5.6), Inches(5.4),
          "The PID 1 Rule & What Happens During 'docker run'",
          [
              "Key insight: When PID 1 exits, the container stops. No PID 1 = dead container.",
              "Example: Sarah's web app is PID 1. If it crashes, container stops automatically.",
              "What happens during docker run:",
              "- Pull image layers (if not present)",
              "- Create container filesystem (union of layers)",
              "- Create namespaces (PID, network, mount, etc.)",
              "- Set up cgroups (resource limits)",
              "- Configure networking (assign IP, set up iptables)",
              "- Start process as PID 1",
              "- Return control to user (Container is now running—just a Linux process in a special namespace!)"
          ],
          accent_color=COLOR_ACCENT_AMBER)


# ==============================================================================
# SLIDE 15: 10. BASIC LINUX COMMANDS TO UNDERSTAND CONTAINERS (PART 1)
# ==============================================================================
s15 = prs.slides.add_slide(blank_layout)
set_slide_background(s15)
add_header(s15, "10. Basic Linux Commands to Understand Containers", "Sarah learns commands to peek behind Docker's curtain: Process & Namespace inspection.")

add_panel(s15, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.4),
          "Process Inspection",
          [
              "- ps aux: Lists ALL processes on system. Shows PID, CPU%, memory, command. Use: See that container processes are just normal Linux processes.",
              "- ps -ef: Alternative format showing parent-child relationships. Shows PPID (parent process ID).",
              "- top / htop: Real-time process monitoring. Shows CPU/memory usage per process. htop is prettier and more interactive.",
              "Usage: Sarah runs ps aux | grep python and sees her containerized Python app listed like any other process!"
          ],
          accent_color=COLOR_ACCENT_CYAN)

add_panel(s15, Inches(6.9), Inches(1.6), Inches(5.6), Inches(5.4),
          "Namespace Inspection",
          [
              "- lsns: Lists all namespaces on system. Shows type (pid, net, mount, etc.). Shows which processes belong to which namespace.",
              "Usage: Sarah runs lsns -t net and sees separate network namespaces for each container.",
              "- ip netns list: Lists network namespaces specifically. More detailed than lsns -t net.",
              "- ip netns exec <ns> <command>: Run command inside a specific network namespace.",
              "Example: ip netns exec container1 ip addr shows IPs inside that namespace."
          ],
          accent_color=COLOR_ACCENT_GREEN)


# ==============================================================================
# SLIDE 16: 10. BASIC LINUX COMMANDS TO UNDERSTAND CONTAINERS (PART 2)
# ==============================================================================
s16 = prs.slides.add_slide(blank_layout)
set_slide_background(s16)
add_header(s16, "10. Basic Linux Commands to Understand Containers (Cont.)", "Filesystem & Network commands, and putting it all together.")

add_panel(s16, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.4),
          "Filesystem & Network Commands",
          [
              "Filesystem Commands:",
              "- mount: Shows all mounted filesystems. Reveals overlay/aufs mounts used by Docker (mount | grep overlay).",
              "- chroot <newroot> <command>: Changes root directory for a process. Simulates container filesystem isolation.",
              "- df -h & du -sh <directory>: Shows disk space and directory usage.",
              "Network Commands:",
              "- ip addr show & ip link show: Shows network interfaces, docker0 bridge and container veth pairs.",
              "- ss -tulpn / netstat -tulpn: Shows listening ports and processes (verifies port 8080 mapped to 80)."
          ],
          accent_color=COLOR_ACCENT_PURPLE)

add_panel(s16, Inches(6.9), Inches(1.6), Inches(5.6), Inches(5.4),
          "Putting it Together: Sarah Verifies Docker's Magic",
          [
              "Sarah can now verify Docker's magic:",
              "- Run docker ps → see container ID 'abc123'",
              "- Run docker inspect abc123 → find PID '5421'",
              "- Run ps aux | grep 5421 → see it's just a Linux process",
              "- Run lsns -p 5421 → see it's in isolated namespaces",
              "- Run mount | grep abc123 → see overlay filesystem",
              "- Run ip addr → see veth pair for container networking",
              "Docker is just organizing Linux kernel features cleverly!"
          ],
          accent_color=COLOR_ACCENT_CYAN)


# ==============================================================================
# SLIDE 17: 11. DOCKER ARCHITECTURE
# ==============================================================================
s17 = prs.slides.add_slide(blank_layout)
set_slide_background(s17)
add_header(s17, "11. Docker Architecture", "Docker has a client-server architecture: Tenant (Client) and Building Manager (Daemon).")

add_panel(s17, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.4),
          "Architecture Components",
          [
              "- Docker Client (CLI): What Sarah types: docker run, docker ps, etc. Talks to Docker Daemon via REST API. Can connect to local or remote daemon. Location: /usr/bin/docker.",
              "- Docker Daemon (dockerd): The 'brain' that does actual work. Manages containers, images, networks, volumes. Listens on socket: /var/run/docker.sock. Talks to containerd.",
              "- containerd: Lower-level container runtime. Manages container lifecycle (start, stop, pause). Daemon delegates actual work here.",
              "- runc: Lowest level: actually creates containers. Implements OCI spec. Creates namespaces, cgroups, starts process.",
              "- Docker Registry: Stores Docker images (Docker Hub, private registries). Client pulls images from here."
          ],
          accent_color=COLOR_ACCENT_CYAN)

add_panel(s17, Inches(6.9), Inches(1.6), Inches(5.6), Inches(5.4),
          "Execution Flow: Sarah types docker run nginx",
          [
              "Flow:",
              "Sarah types: docker run nginx",
              "↓",
              "Docker CLI → sends API request",
              "↓",
              "Docker Daemon → receives request",
              "↓",
              "Checks for nginx image locally → not found",
              "↓",
              "Pulls from Docker Hub Registry → Passes to containerd",
              "↓",
              "containerd calls runc → runc creates namespaces + cgroups",
              "↓",
              "Starts nginx process as PID 1 → Container running!"
          ],
          accent_color=COLOR_ACCENT_GREEN)


# ==============================================================================
# SLIDE 18: 12. DOCKER ENGINE VS CLI VS DAEMON
# ==============================================================================
s18 = prs.slides.add_slide(blank_layout)
set_slide_background(s18)
add_header(s18, "12. Docker Engine vs CLI vs Daemon", "Sarah gets confused by terminology. Clarifying the components and communication.")

add_panel(s18, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.4),
          "Definitions & Car Analogy",
          [
              "- Docker Engine: The complete Docker system. Includes: daemon + CLI + containerd + runc. What you install when you 'install Docker'.",
              "- Docker CLI: Just the command-line tool. Binary: /usr/bin/docker. Can work with remote daemons: docker -H remote.server.com ps.",
              "- Docker Daemon: Background service (dockerd). Does the heavy lifting. Runs as root (needs kernel access). Manages: images, containers, networks, volumes.",
              "Analogy:",
              "- Engine = entire car",
              "- CLI = steering wheel (how you control it)",
              "- Daemon = engine (does the work)"
          ],
          accent_color=COLOR_ACCENT_AMBER)

add_panel(s18, Inches(6.9), Inches(1.6), Inches(5.6), Inches(5.4),
          "Communication & Security",
          [
              "Communication:",
              "CLI and Daemon talk via:",
              "- Unix socket: /var/run/docker.sock (local)",
              "- TCP: tcp://IP:2375 (remote, unencrypted)",
              "- TCP+TLS: tcp://IP:2376 (remote, encrypted)",
              "Security note:",
              "Docker socket = root access. Protect it!"
          ],
          accent_color=COLOR_ACCENT_CYAN)


# ==============================================================================
# SLIDE 19: 13. CONTAINER LIFECYCLE THEORY
# ==============================================================================
s19 = prs.slides.add_slide(blank_layout)
set_slide_background(s19)
add_header(s19, "13. Container Lifecycle Theory", "Sarah understands the complete lifecycle phases and the 4 important rules.")

add_panel(s19, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.4),
          "Phases of Container Lifecycle",
          [
              "Phase 1: Image Building (Dockerfile recipe, layers cached and reused, image = read-only template)",
              "Phase 2: Container Creation (docker create/run, copy layers, add writable layer, allocate namespaces, configure network)",
              "Phase 3: Container Execution (Start PID 1, apply cgroups limits, state: Running)",
              "Phase 4: Container Stopping (Graceful: SIGTERM → 10s wait → SIGKILL; Forced: SIGKILL kill -9; state: Stopped, writable layer preserved)",
              "Phase 5: Container Restart (Reuse container, don't recreate namespaces, start PID 1 again, data still there)",
              "Phase 6: Container Removal (Delete container & writable layer, namespaces destroyed, cgroups removed, data lost unless volumes)",
              "States Diagram: Created → Running → Stopped → Deleted"
          ],
          accent_color=COLOR_ACCENT_PURPLE)

add_panel(s19, Inches(6.9), Inches(1.6), Inches(5.6), Inches(5.4),
          "4 Important Rules",
          [
              "Important Rules:",
              "- Rule 1: Container lives as long as PID 1 lives. PID 1 exits → container stops. Keep PID 1 running to keep container alive.",
              "- Rule 2: Data in writable layer is ephemeral. Stopping preserves data. Deleting destroys data. Use volumes for persistence.",
              "- Rule 3: Images are immutable, containers are mutable. Image = template (never changes). Container = running instance (can change). Delete container, image stays.",
              "- Rule 4: One process per container (best practice). Not enforced, but recommended. Makes containers simple and replaceable. Follow Unix philosophy: do one thing well."
          ],
          accent_color=COLOR_ACCENT_GREEN)


# ==============================================================================
# SLIDE 20: 14. SARAH'S AHA MOMENT: PUTTING IT ALL TOGETHER
# ==============================================================================
s20 = prs.slides.add_slide(blank_layout)
set_slide_background(s20)
add_header(s20, "14. Sarah's Aha Moment: Putting It All Together", "Sarah now connects every concept: containers are just isolated Linux processes!")

add_panel(s20, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.4),
          "Sarah Now Understands & The Big Picture",
          [
              "Sarah now understands:",
              "- Docker containers aren't magic → just Linux processes with namespaces + cgroups",
              "- Not VMs → shared kernel, much lighter",
              "- Isolation through kernel features → namespaces (what you see), cgroups (what you use)",
              "- Layered filesystem → images share layers, containers get unique writable layer",
              "- Client-server architecture → CLI talks to daemon, daemon uses containerd/runc",
              "- Process lifecycle → container = PID 1 + its children in isolated namespaces",
              "The big picture:",
              "- Container = isolated Linux process",
              "- Docker = tool that orchestrates namespaces, cgroups, and union filesystems",
              "- Images = templates (read-only layers)",
              "- Containers = running instances (images + writable layer)"
          ],
          accent_color=COLOR_ACCENT_AMBER)

add_panel(s20, Inches(6.9), Inches(1.6), Inches(5.6), Inches(5.4),
          "When Sarah Runs docker run",
          [
              "When Sarah runs:",
              "docker run -d --name demo-nginx -p 8080:80 nginx",
              "- Create namespaces (PID, network, mount)",
              "- Apply cgroups (512MB memory limit)",
              "- Set up port mapping (80→8080)",
              "- Start myapp as PID 1",
              "- Return immediately (-d = detached)"
          ],
          accent_color=COLOR_ACCENT_GREEN)


# Save presentation
prs.save(OUTPUT_PPTX)
print(f"✅ Generated presentation with {len(prs.slides)} slides at: {OUTPUT_PPTX}")
