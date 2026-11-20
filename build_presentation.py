#!/usr/bin/env python3
"""
Script to generate the 'Docker Fundamentals' presentation.
Incorporates:
- Modern widescreen 16:9 layout
- Professional dark tech palette (Tailwind Slate/Sky/Emerald theme)
- Story-driven narrative: Aanand's Journey from Chaos to Clarity
- Embedded Course Roadmap screenshot
- Rich cards, comparisons, command cheat sheets, and architectural flows
"""

import os
import pptx
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

# Output presentation file
OUTPUT_PPTX = "/Users/aanand-mac/Downloads/Docker_Fundamentals.pptx"
SCREENSHOT_PATH = "/var/folders/3q/jn41f3x56y922gqq0_02kxl80000gn/T/TemporaryItems/NSIRD_screencaptureui_egVht0/Screenshot 2026-11-18 at 8.23.12 PM.png"

# Color Palette Definitions
COLOR_BG = RGBColor(15, 23, 42)          # Slate 900 (Deep Navy Background)
COLOR_CARD_BG = RGBColor(30, 41, 59)     # Slate 800 (Panel Card)
COLOR_CARD_BORDER = RGBColor(51, 65, 85) # Slate 700 (Subtle Border)
COLOR_TEXT_MAIN = RGBColor(248, 250, 252)# Slate 50 (Crisp White Heading)
COLOR_TEXT_MUTED = RGBColor(148, 163, 184)# Slate 400 (Secondary Text)
COLOR_TEXT_BODY = RGBColor(226, 232, 240)# Slate 200 (Readable Body Text)

# Accent Colors
COLOR_ACCENT_CYAN = RGBColor(56, 189, 248)   # Sky 400 (Primary Brand Accent)
COLOR_ACCENT_BLUE = RGBColor(14, 165, 233)   # Sky 500
COLOR_ACCENT_GREEN = RGBColor(16, 185, 129)  # Emerald 500 (Success / Isolation)
COLOR_ACCENT_AMBER = RGBColor(245, 158, 11)  # Amber 500 (Warning / Chaos)
COLOR_ACCENT_PURPLE = RGBColor(168, 85, 247) # Purple 500 (Architecture)
COLOR_ACCENT_RED = RGBColor(239, 68, 68)     # Red 500 (Fatal errors / Limits)

prs = pptx.Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

def set_slide_background(slide):
    """Draws the dark background on the slide."""
    bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5)
    )
    bg.fill.solid()
    bg.fill.fore_color.rgb = COLOR_BG
    bg.line.fill.background() # no border
    return bg

def add_header(slide, category_tag, title_text, subtitle_text=None):
    """Adds a standard modern header with tag and subtitle."""
    header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.45), Inches(11.733), Inches(1.3))
    tf = header_box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

    # Tag / Category tracker
    p_tag = tf.paragraphs[0]
    p_tag.text = category_tag.upper()
    p_tag.font.name = "Arial"
    p_tag.font.size = Pt(11)
    p_tag.font.bold = True
    p_tag.font.color.rgb = COLOR_ACCENT_CYAN
    p_tag.space_after = Pt(4)

    # Main Title
    p_title = tf.add_paragraph()
    p_title.text = title_text
    p_title.font.name = "Arial"
    p_title.font.size = Pt(26)
    p_title.font.bold = True
    p_title.font.color.rgb = COLOR_TEXT_MAIN

    # Optional Subtitle
    if subtitle_text:
        p_title.space_after = Pt(3)
        p_sub = tf.add_paragraph()
        p_sub.text = subtitle_text
        p_sub.font.name = "Arial"
        p_sub.font.size = Pt(13)
        p_sub.font.color.rgb = COLOR_TEXT_MUTED
    else:
        p_title.space_after = Pt(10)

def add_card(slide, left, top, width, height, title, items, badge=None, accent_color=COLOR_ACCENT_CYAN):
    """Adds a stylish container card with title, items/bullets, and optional badge."""
    # Background shape
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = COLOR_CARD_BG
    card.line.color.rgb = COLOR_CARD_BORDER
    card.line.width = Pt(1.2)

    # Text box
    tb = slide.shapes.add_textbox(left + Inches(0.25), top + Inches(0.2), width - Inches(0.5), height - Inches(0.4))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

    p_first = tf.paragraphs[0]
    if badge:
        p_first.text = f"[{badge}]  {title}"
    else:
        p_first.text = title
    p_first.font.name = "Arial"
    p_first.font.size = Pt(17)
    p_first.font.bold = True
    p_first.font.color.rgb = accent_color
    p_first.space_after = Pt(10)

    for item in items:
        p = tf.add_paragraph()
        p.space_after = Pt(7)
        p.font.name = "Arial"
        p.font.size = Pt(12.5)

        if isinstance(item, tuple):
            strong_part, rest_part = item
            run1 = p.add_run()
            run1.text = "• " + strong_part + ": "
            run1.font.bold = True
            run1.font.color.rgb = COLOR_TEXT_MAIN

            run2 = p.add_run()
            run2.text = rest_part
            run2.font.bold = False
            run2.font.color.rgb = COLOR_TEXT_BODY
        else:
            run = p.add_run()
            run.text = "• " + item
            run.font.color.rgb = COLOR_TEXT_BODY

    return card


# ==============================================================================
# SLIDE 1: TITLE SLIDE
# ==============================================================================
s1 = prs.slides.add_slide(blank_layout)
set_slide_background(s1)

tb1 = s1.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(11.333), Inches(4.2))
tf1 = tb1.text_frame
tf1.word_wrap = True

p1 = tf1.paragraphs[0]
p1.text = "DEVOPS & CONTAINER MASTERCLASS"
p1.font.name = "Arial"
p1.font.size = Pt(14)
p1.font.bold = True
p1.font.color.rgb = COLOR_ACCENT_CYAN
p1.space_after = Pt(12)

p2 = tf1.add_paragraph()
p2.text = "Docker Fundamentals"
p2.font.name = "Arial"
p2.font.size = Pt(46)
p2.font.bold = True
p2.font.color.rgb = COLOR_TEXT_MAIN
p2.space_after = Pt(14)

p3 = tf1.add_paragraph()
p3.text = "The Journey of Aanand: From Deployment Chaos to Container Clarity"
p3.font.name = "Arial"
p3.font.size = Pt(20)
p3.font.color.rgb = COLOR_ACCENT_AMBER
p3.space_after = Pt(24)

p4 = tf1.add_paragraph()
p4.text = "Deep-dive into Linux Processes, Namespaces, cgroups, Union Filesystems & Docker Architecture"
p4.font.name = "Arial"
p4.font.size = Pt(14)
p4.font.color.rgb = COLOR_TEXT_MUTED


# ==============================================================================
# SLIDE 2: ROADMAP SLIDE
# ==============================================================================
s2 = prs.slides.add_slide(blank_layout)
set_slide_background(s2)
add_header(s2, "Curriculum Overview", "Course Roadmap: From Docker to Kubernetes & Cloud",
           "Complete end-to-end journey from container basics to production-grade orchestration.")

# Embed Screenshot on Left
if os.path.exists(SCREENSHOT_PATH):
    img_left = Inches(0.8)
    img_top = Inches(1.9)
    img_height = Inches(5.1)
    # 317x705 aspect ratio -> width ~ 2.29 inches
    pic = s2.shapes.add_picture(SCREENSHOT_PATH, img_left, img_top, height=img_height)

# Cards on Right
add_card(s2, Inches(3.6), Inches(1.9), Inches(4.3), Inches(5.1),
         "Stage 1: Container Core",
         [
             ("Class 1 (Today)", "Docker Introduction & Kernel Isolation"),
             ("Class 2-3", "Images, Containers & Building Custom Images"),
             ("Class 4-5", "Deep Dive: Storage Volumes & Container Networking"),
             ("Class 6-7", "Hardening Docker Environments & DevSecOps"),
             ("Milestone", "Zero 'works on my machine' issues across all environments")
         ],
         badge="FOUNDATION", accent_color=COLOR_ACCENT_CYAN)

add_card(s2, Inches(8.2), Inches(1.9), Inches(4.3), Inches(5.1),
         "Stage 2: Orchestration & AI",
         [
             ("Kubernetes Core", "Pods, ReplicaSets, Deployments, Services"),
             ("Scaling & Autoscaling", "HPA, VPA, Cluster Autoscaler, Ingress"),
             ("Advanced K8s", "CRDs, CNIs, Probes, Secrets & RBAC"),
             ("AI Deployment", "Deploying Hugging Face ML Models on K8s"),
             ("Observability", "Prometheus, Grafana, PromQL & New Relic")
         ],
         badge="PRODUCTION", accent_color=COLOR_ACCENT_PURPLE)


# ==============================================================================
# SLIDE 3: 1. DOCKER INTRODUCTION
# ==============================================================================
s3 = prs.slides.add_slide(blank_layout)
set_slide_background(s3)
add_header(s3, "Chapter 1", "Docker Introduction: The Deployment Chaos",
           "Meet Aanand, a developer who built an awesome web app, only to hit a wall in production.")

add_card(s3, Inches(0.8), Inches(1.9), Inches(5.6), Inches(5.0),
         "Aanand's Laptop Experience",
         [
             ("Status", "Works perfectly! Fast, responsive, passes tests"),
             ("Developer Environment", "macOS, Python 3.9, locally cached pip libraries"),
             ("The Big Day", "Aanand pushes code to GitHub and prepares to deploy to production"),
             ("The Expectation", "'It took me 10 minutes locally, server deploy will be effortless!'")
         ],
         badge="LOCAL DEV", accent_color=COLOR_ACCENT_GREEN)

add_card(s3, Inches(6.9), Inches(1.9), Inches(5.6), Inches(5.0),
         "Production Reality: Disaster Strikes",
         [
             ("'Works on My Machine'", "Became Aanand's frustrated daily defense"),
             ("Python Version Clashes", "Host had Python 3.6; code required Python 3.9 features"),
             ("Missing System Libs", "Missing C-extensions and system header dependencies"),
             ("Team Onboarding Broken", "New teammates spent 3 days just setting up local runtimes"),
             ("The Urgent Need", "A reliable unit that bundles app + runtime + OS libraries together")
         ],
         badge="PROD DISASTER", accent_color=COLOR_ACCENT_RED)


# ==============================================================================
# SLIDE 4: 2. WHAT IS DOCKER
# ==============================================================================
s4 = prs.slides.add_slide(blank_layout)
set_slide_background(s4)
add_header(s4, "Chapter 2", "What is Docker? Standardized Software Units",
           "Packaging applications with all dependencies into self-contained, portable containers.")

add_card(s4, Inches(0.8), Inches(1.9), Inches(5.6), Inches(5.0),
         "The Shipping Container Analogy",
         [
             ("Historical Chaos", "Before 1956, cargo shipping was chaotic; barrels, boxes, and sacks had to be manually re-handled at every port"),
             ("The Intermodal Container", "A standardized steel box carried identically by cargo ship, train, and truck without unpacking"),
             ("Docker's Solution", "A standardized software container running identically on a laptop, on-prem server, or cloud"),
             ("Predictability", "No matter what is inside, the outside interface is identical")
         ],
         badge="ANALOGY 1", accent_color=COLOR_ACCENT_BLUE)

add_card(s4, Inches(6.9), Inches(1.9), Inches(5.6), Inches(5.0),
         "The Furnished Apartment Analogy",
         [
             ("Unfurnished (Traditional)", "You bring your own bed, plumbing fixtures, electricity wiring, and kitchen appliances. Half of them don't fit!"),
             ("Fully Furnished (Docker)", "You walk in with just your clothes. Bed, electricity, plumbing, and utensils are pre-configured and guaranteed to work"),
             ("Key Architectural Insight", "Containers package not just app code, but system libraries, environment variables, binaries, and configurations"),
             ("Portability", "Pick up the entire furnished room and drop it anywhere")
         ],
         badge="ANALOGY 2", accent_color=COLOR_ACCENT_AMBER)


# ==============================================================================
# SLIDE 5: 3. WHY DOCKER WAS CREATED
# ==============================================================================
s5 = prs.slides.add_slide(blank_layout)
set_slide_background(s5)
add_header(s5, "Chapter 3", "Why Docker Was Created: Overcoming Dependency Hell",
           "Solving the matrix of doom between applications, runtimes, and operating systems.")

add_card(s5, Inches(0.8), Inches(1.9), Inches(3.6), Inches(5.0),
         "The 'Dependency Hell'",
         [
             ("OS Inconsistencies", "Ubuntu locally vs CentOS on prod"),
             ("Runtime Drift", "Python 3.6 vs 3.8 vs 3.9"),
             ("Library Clashes", "App A needs libssl 1.0, App B needs libssl 1.1"),
             ("Manual Runbooks", "20-page setup manuals that broke every month")
         ],
         badge="THE PROBLEM", accent_color=COLOR_ACCENT_RED)

add_card(s5, Inches(4.8), Inches(1.9), Inches(3.6), Inches(5.0),
         "What Developers Wanted",
         [
             ("Hermetic Packaging", "App + dependencies in one single artifact"),
             ("Guaranteed Parity", "Exact same behavior in Dev, Staging & Prod"),
             ("Lightweight Footprint", "Fast execution without hypervisor lag"),
             ("Sub-Second Starts", "Instant boot compared to slow VMs")
         ],
         badge="THE DREAM", accent_color=COLOR_ACCENT_AMBER)

add_card(s5, Inches(8.8), Inches(1.9), Inches(3.7), Inches(5.0),
         "Docker's Breakthrough",
         [
             ("Smart Kernel Usage", "Didn't invent virtualization; leveraged built-in Linux primitives"),
             ("Namespaces", "Carved out private process views"),
             ("cgroups", "Enforced strict resource governance"),
             ("UnionFS", "Shared common files across images")
         ],
         badge="THE SOLUTION", accent_color=COLOR_ACCENT_GREEN)


# ==============================================================================
# SLIDE 6: 4. WHAT IS A PROCESS IN LINUX
# ==============================================================================
s6 = prs.slides.add_slide(blank_layout)
set_slide_background(s6)
add_header(s6, "Chapter 4", "What is a Process in Linux?",
           "Before understanding containers, Aanand must understand what a Linux process actually is.")

add_card(s6, Inches(0.8), Inches(1.9), Inches(5.6), Inches(5.0),
         "How the Linux Kernel Runs a Program",
         [
             ("A Process Defined", "A program loaded into RAM actively executing instructions"),
             ("Binary Execution", "When Aanand executes 'python app.py':"),
             ("1. RAM Allocation", "Kernel loads python binary & bytecode into memory"),
             ("2. PID Assigned", "Kernel issues a unique Process ID (PID)"),
             ("3. CPU Slicing", "Kernel scheduler assigns CPU time slices"),
             ("4. File Descriptors", "Standard input (0), output (1), and error (2) opened")
         ],
         badge="EXECUTION", accent_color=COLOR_ACCENT_CYAN)

add_card(s6, Inches(6.9), Inches(1.9), Inches(5.6), Inches(5.0),
         "Key Process Characteristics",
         [
             ("Tree Hierarchy", "Every process has a parent (PPID). Root of all processes is PID 1 (init/systemd)"),
             ("Shared Kernel", "All processes talk to the exact same Linux kernel via system calls (syscalls)"),
             ("Global Visibility", "Running 'ps aux' exposes ALL processes running across the entire operating system"),
             ("Aanand's Discovery", "By default, regular Linux processes have ZERO isolation from each other")
         ],
         badge="BEHAVIOR", accent_color=COLOR_ACCENT_PURPLE)


# ==============================================================================
# SLIDE 7: 5. HOW PROCESSES RUN AND ARE ISOLATED
# ==============================================================================
s7 = prs.slides.add_slide(blank_layout)
set_slide_background(s7)
add_header(s7, "Chapter 5", "How Processes Run and Are Isolated",
           "Traditional Linux vs the Containerized Isolation Aanand needs.")

add_card(s7, Inches(0.8), Inches(1.9), Inches(5.6), Inches(5.0),
         "Traditional Linux: Zero Isolation",
         [
             ("Shared Filesystem", "Every process can read `/tmp` and potentially inspect config files"),
             ("Port Collisions", "If App 1 binds to port 8080, App 2 crashes with 'Address already in use'"),
             ("Process Snooping", "Processes can send signals (`kill`) and trace other processes with `ptrace`"),
             ("No Memory Caps", "A memory leak in one rogue script crashes the entire server (OOM panic)"),
             ("Security Risk", "A compromised web app gives an attacker visibility into all host operations")
         ],
         badge="TRADITIONAL", accent_color=COLOR_ACCENT_RED)

add_card(s7, Inches(6.9), Inches(1.9), Inches(5.6), Inches(5.0),
         "Container Isolation: What Aanand Needs",
         [
             ("Private Bubble", "The app believes it is the sole application on the computer"),
             ("Dedicated PID 1", "The app operates as PID 1 inside its own isolated tree"),
             ("Own Network Stack", "Isolated loopback, unique IP, and private port 8080 binding"),
             ("Guaranteed CPU/RAM", "Hard limits prevent one app from starving neighbor apps"),
             ("Isolated Root Filesystem", "Can modify `/etc` or `/var` without touching host files")
         ],
         badge="CONTAINERIZED", accent_color=COLOR_ACCENT_GREEN)


# ==============================================================================
# SLIDE 8: 6. CONTAINERS VS VIRTUAL MACHINES
# ==============================================================================
s8 = prs.slides.add_slide(blank_layout)
set_slide_background(s8)
add_header(s8, "Chapter 6", "Containers vs Virtual Machines",
           "Hardware-level virtualization vs Operating System-level virtualization.")

add_card(s8, Inches(0.8), Inches(1.9), Inches(5.6), Inches(5.0),
         "Virtual Machines (Hypervisor Based)",
         [
             ("Virtualization Level", "Hardware Virtualization via Hypervisor (Type 1 or 2)"),
             ("Operating System", "Each VM runs a complete Guest OS (kernel + binaries + drivers)"),
             ("Size & Overhead", "Heavyweight: 10GB - 40GB disk per VM; high baseline RAM usage"),
             ("Boot Speed", "Minutes (must initialize virtual hardware, bios & full OS boot)"),
             ("Analogy", "Building 10 completely independent standalone houses on separate plots")
         ],
         badge="HEAVYWEIGHT", accent_color=COLOR_ACCENT_AMBER)

add_card(s8, Inches(6.9), Inches(1.9), Inches(5.6), Inches(5.0),
         "Containers (Kernel Sharing)",
         [
             ("Virtualization Level", "OS-Level Virtualization via Linux Kernel Features"),
             ("Operating System", "Shared Host Kernel; only packages userland binaries & libraries"),
             ("Size & Overhead", "Lightweight: Tens of megabytes; near-zero virtualization overhead"),
             ("Boot Speed", "Milliseconds (just launching a standard Linux process)"),
             ("Analogy", "Building 10 modern private apartments inside one shared building structure")
         ],
         badge="LIGHTWEIGHT", accent_color=COLOR_ACCENT_CYAN)


# ==============================================================================
# SLIDE 9: 7. KEY LINUX KERNEL FEATURES
# ==============================================================================
s9 = prs.slides.add_slide(blank_layout)
set_slide_background(s9)
add_header(s9, "Chapter 7", "The Triad of Containerization: Linux Kernel Features",
           "Docker containers are not magic. They are plain Linux processes governed by 3 kernel pillars.")

add_card(s9, Inches(0.8), Inches(1.9), Inches(3.6), Inches(5.0),
         "1. Namespaces",
         [
             ("Core Purpose", "Isolation of View"),
             ("Key Question", "What can a process SEE?"),
             ("Mechanism", "Soundproof walls and one-way mirrors"),
             ("Scope", "PIDs, Network interfaces, Mount points, Hostnames, IPC, Users"),
             ("Result", "The process believes it has the entire computer to itself")
         ],
         badge="VIEW", accent_color=COLOR_ACCENT_CYAN)

add_card(s9, Inches(4.8), Inches(1.9), Inches(3.6), Inches(5.0),
         "2. cgroups",
         [
             ("Core Purpose", "Resource Limits & Metering"),
             ("Key Question", "What can a process USE?"),
             ("Mechanism", "Utility meters and governor limits"),
             ("Scope", "CPU cores/percentages, Memory (RAM), Disk I/O speed, Network bandwidth"),
             ("Result", "No single container can starve or crash the host machine")
         ],
         badge="LIMITS", accent_color=COLOR_ACCENT_AMBER)

add_card(s9, Inches(8.8), Inches(1.9), Inches(3.7), Inches(5.0),
         "3. Union Filesystems",
         [
             ("Core Purpose", "Layered Storage & Deduplication"),
             ("Key Question", "How are files STACKED?"),
             ("Mechanism", "Transparent layering (OverlayFS)"),
             ("Scope", "Read-only image layers stacked under one thin writable layer"),
             ("Result", "Massive disk savings; instant container launches via CoW")
         ],
         badge="STORAGE", accent_color=COLOR_ACCENT_PURPLE)


# ==============================================================================
# SLIDE 10: 7.1 NAMESPACES: ISOLATION OF VIEW
# ==============================================================================
s10 = prs.slides.add_slide(blank_layout)
set_slide_background(s10)
add_header(s10, "Chapter 7.1", "Linux Namespaces: Controlling What Processes Can SEE",
           "Soundproof walls dividing the operating system into independent visibility bubbles.")

add_card(s10, Inches(0.8), Inches(1.9), Inches(5.6), Inches(5.0),
         "Primary Namespaces",
         [
             ("PID Namespace", "Inside: App is PID 1. Outside on Host: Process has normal PID (e.g. 8421). Prevents containers from signaling host processes"),
             ("Network (net) Namespace", "Each container gets private loopback (127.0.0.1), virtual ethernet interface (`eth0`), routing table, and private ports"),
             ("Mount (mnt) Namespace", "Isolates filesystem mount points. Container sees `/` rooted in its own image, unable to traverse host files")
         ],
         badge="ESSENTIAL", accent_color=COLOR_ACCENT_CYAN)

add_card(s10, Inches(6.9), Inches(1.9), Inches(5.6), Inches(5.0),
         "Supporting Namespaces",
         [
             ("UTS Namespace", "Isolates hostname and domain name. Container can be named `api-server` without altering host hostname"),
             ("IPC Namespace", "Isolates Inter-Process Communication (POSIX message queues, SysV shared memory). Stops cross-container memory snooping"),
             ("User Namespace", "Maps UID 0 (root inside container) to an unprivileged UID (e.g., 10001 on host). Critical security barrier against breakouts")
         ],
         badge="ADVANCED", accent_color=COLOR_ACCENT_PURPLE)


# ==============================================================================
# SLIDE 11: 7.2 CGROUPS: RESOURCE LIMITS
# ==============================================================================
s11 = prs.slides.add_slide(blank_layout)
set_slide_background(s11)
add_header(s11, "Chapter 7.2", "Control Groups (cgroups): Controlling What Processes Can USE",
           "Enforcing resource limits, prioritization, and granular accounting across containers.")

add_card(s11, Inches(0.8), Inches(1.9), Inches(5.6), Inches(5.0),
         "Enforceable Resource Limits",
         [
             ("CPU Throttling", "Limit a container to 50% of 1 core (`--cpus=0.5`) or pin it to specific physical cores (`--cpuset-cpus=0,1`)"),
             ("Memory Hard Caps", "Limit container RAM (`-m 512m`). If the process exceeds this limit, the Linux kernel Out-Of-Memory (OOM) killer terminates it"),
             ("Block I/O (Disk)", "Throttle read/write speeds (`--device-write-bps /dev/sda:10mb`) so logging doesn't saturate disk drives"),
             ("Process Limits (pids)", "Prevent fork-bombs by limiting max concurrent threads/processes inside the container")
         ],
         badge="GOVERNANCE", accent_color=COLOR_ACCENT_AMBER)

add_card(s11, Inches(6.9), Inches(1.9), Inches(5.6), Inches(5.0),
         "Resource Accounting & Monitoring",
         [
             ("Accurate Metering", "Kernel continuously logs exact CPU cycles, memory pages, and disk bytes consumed per cgroup"),
             ("Exposed via `/sys/fs/cgroup`", "Metrics are plain text files on Linux, easily scraped by monitoring daemons"),
             ("Foundation for Cloud Billing", "Powers cloud container pricing (AWS Fargate, Google Cloud Run)"),
             ("Classroom Insight", "Without cgroups, one rogue loop in Aanand's app would take down the entire server")
         ],
         badge="ACCOUNTING", accent_color=COLOR_ACCENT_GREEN)


# ==============================================================================
# SLIDE 12: 7.3 UNION FILESYSTEMS
# ==============================================================================
s12 = prs.slides.add_slide(blank_layout)
set_slide_background(s12)
add_header(s12, "Chapter 7.3", "Union Filesystems: Layering & Copy-on-Write (CoW)",
           "How Docker achieves instant startup times and massive disk storage savings.")

add_card(s12, Inches(0.8), Inches(1.9), Inches(5.6), Inches(5.0),
         "How OverlayFS Stacks Layers",
         [
             ("Layer 4: Container (Writable)", "Ephemeral top layer where app writes logs, temp files, and runtime state"),
             ("Layer 3: App Code (Read-Only)", "Aanand's Python web application code and scripts"),
             ("Layer 2: Python Runtime (Read-Only)", "Python 3.9 interpreter, stdlib, pip packages"),
             ("Layer 1: Base OS (Read-Only)", "Debian/Alpine slim root filesystem"),
             ("Copy-on-Write (CoW)", "When a file is modified, it is copied up to the writable layer; base layers remain pristine")
         ],
         badge="OVERLAYFS", accent_color=COLOR_ACCENT_PURPLE)

add_card(s12, Inches(6.9), Inches(1.9), Inches(5.6), Inches(5.0),
         "Huge Storage & Speed Advantages",
         [
             ("Deduplication", "If 10 containers use Ubuntu 22.04 (1GB base), total disk used is 1GB + tiny writable layers, NOT 10GB!"),
             ("Instant Container Boot", "Docker does not copy gigabytes of OS files; it just mounts existing read-only layers in milliseconds"),
             ("Immutability", "Shared layers are cryptographically checksummed and tamper-proof"),
             ("Evolution", "AUFS (older, complex) has been superseded by OverlayFS (native Linux kernel standard)")
         ],
         badge="EFFICIENCY", accent_color=COLOR_ACCENT_CYAN)


# ==============================================================================
# SLIDE 13: 8. HOW DOCKER USES KERNEL FEATURES
# ==============================================================================
s13 = prs.slides.add_slide(blank_layout)
set_slide_background(s13)
add_header(s13, "Chapter 8", "How Docker Orchestrates Kernel Features",
           "The step-by-step mechanical execution sequence when Aanand runs a container.")

add_card(s13, Inches(0.8), Inches(1.9), Inches(2.7), Inches(5.0),
         "Step 1: Namespaces",
         [
             ("PID Namespace", "Create isolated process tree"),
             ("Net Namespace", "Create virtual eth & IP"),
             ("Mnt Namespace", "Isolate view of root filesystem"),
             ("UTS Namespace", "Set container hostname")
         ],
         badge="STEP 1", accent_color=COLOR_ACCENT_CYAN)

add_card(s13, Inches(3.75), Inches(1.9), Inches(2.7), Inches(5.0),
         "Step 2: cgroups",
         [
             ("Directory Created", "Entry in `/sys/fs/cgroup/`"),
             ("CPU Quota", "Apply core allocations"),
             ("Memory Max", "Set RAM limit (e.g. 512MB)"),
             ("I/O Weights", "Set disk transfer throttles")
         ],
         badge="STEP 2", accent_color=COLOR_ACCENT_AMBER)

add_card(s13, Inches(6.7), Inches(1.9), Inches(2.7), Inches(5.0),
         "Step 3: UnionFS",
         [
             ("Overlay Mount", "Stack image layers"),
             ("LowerDir", "Point to read-only layers"),
             ("UpperDir", "Create container writable dir"),
             ("Merged View", "Mount as root (`/`) for process")
         ],
         badge="STEP 3", accent_color=COLOR_ACCENT_PURPLE)

add_card(s13, Inches(9.65), Inches(1.9), Inches(2.85), Inches(5.0),
         "Step 4: Launch PID 1",
         [
             ("Pivot Root", "Switch process root to OverlayFS"),
             ("Exec App", "Start python app as PID 1"),
             ("Isolated Reality", "App runs thinking it owns the entire machine!"),
             ("Done in 50ms", "Ready to serve traffic")
         ],
         badge="STEP 4", accent_color=COLOR_ACCENT_GREEN)


# ==============================================================================
# SLIDE 14: 9. LIFECYCLE OF A CONTAINER PROCESS
# ==============================================================================
s14 = prs.slides.add_slide(blank_layout)
set_slide_background(s14)
add_header(s14, "Chapter 9", "Lifecycle of a Container Process",
           "A container is a process with states. When PID 1 exits, the container dies.")

add_card(s14, Inches(0.8), Inches(1.9), Inches(5.6), Inches(5.0),
         "The 5 Container States",
         [
             ("Created", "Filesystem and namespaces prepared, but main process not yet executed (`docker create`)"),
             ("Running", "Main process (PID 1) is active and executing CPU cycles (`docker start` / `docker run`)"),
             ("Paused", "Process frozen using cgroups freezer subsystem; remains in RAM without CPU usage (`docker pause`)"),
             ("Stopped", "Process exited or killed with SIGTERM/SIGKILL; writable filesystem preserved (`docker stop`)"),
             ("Deleted", "Container and its writable layer permanently removed from disk (`docker rm`)")
         ],
         badge="STATES", accent_color=COLOR_ACCENT_CYAN)

add_card(s14, Inches(6.9), Inches(1.9), Inches(5.6), Inches(5.0),
         "The Golden Rule of PID 1",
         [
             ("The Lifecycle Anchor", "The container's entire existence is tied to PID 1. If PID 1 completes or crashes, container stops immediately"),
             ("Graceful Stop Flow", "1. Docker sends `SIGTERM` to PID 1\n2. Waits 10-second grace period for cleanup\n3. Sends `SIGKILL` if still running"),
             ("Zombie Reaping", "PID 1 is responsible for reaping orphan child processes inside its namespace"),
             ("Aanand's Takeaway", "To keep a web container alive, its primary process must run in the foreground")
         ],
         badge="GOLDEN RULE", accent_color=COLOR_ACCENT_RED)


# ==============================================================================
# SLIDE 15: 10. BASIC LINUX COMMANDS TO UNDERSTAND CONTAINERS
# ==============================================================================
s15 = prs.slides.add_slide(blank_layout)
set_slide_background(s15)
add_header(s15, "Chapter 10", "Peeking Behind the Curtain with Linux Commands",
           "Verifying that Docker containers are just ordinary Linux processes.")

add_card(s15, Inches(0.8), Inches(1.9), Inches(5.6), Inches(5.0),
         "Inspection Commands",
         [
             ("ps aux / ps -ef", "Lists all host processes. Running `ps aux | grep python` reveals the containerized process running on the host kernel!"),
             ("lsns", "Lists all system namespaces, their IDs, types, and the host process assigned to each"),
             ("ip netns exec <ns> ip a", "Inspects IP addresses and routing tables inside an isolated network namespace"),
             ("mount | grep overlay", "Reveals Docker's active OverlayFS mount layers, UpperDir, and LowerDir on the host filesystem")
         ],
         badge="COMMANDS", accent_color=COLOR_ACCENT_CYAN)

add_card(s15, Inches(6.9), Inches(1.9), Inches(5.6), Inches(5.0),
         "Aanand's 5-Step Proof of Reality",
         [
             ("1. Run Container", "`docker run -d --name demo nginx`"),
             ("2. Find Host PID", "`docker inspect demo --format '{{.State.Pid}}'` -> e.g. `5421`"),
             ("3. Verify on Host", "`ps -fp 5421` -> See standard nginx process running!"),
             ("4. Check Namespaces", "`lsns -p 5421` -> See distinct PID, NET, and MNT namespaces"),
             ("5. Conclusion", "There is no 'container' entity in the kernel—just an isolated process!")
         ],
         badge="LAB EXPERIMENT", accent_color=COLOR_ACCENT_GREEN)


# ==============================================================================
# SLIDE 16: 11. DOCKER ARCHITECTURE
# ==============================================================================
s16 = prs.slides.add_slide(blank_layout)
set_slide_background(s16)
add_header(s16, "Chapter 11", "Docker Architecture: Client-Server Model",
           "Deconstructing the Docker Client, Daemon, containerd, and runc.")

add_card(s16, Inches(0.8), Inches(1.9), Inches(5.6), Inches(5.0),
         "Core System Components",
         [
             ("Docker CLI (`docker`)", "Command-line interface used by developers. Translates flags into REST API payloads. Can control local or remote daemons"),
             ("Docker Daemon (`dockerd`)", "Persistent background service. Manages high-level images, volumes, networks, and coordinates with containerd"),
             ("containerd", "Industry-standard container supervisor. Manages container lifecycle, pulls image layers, and unpacks storage"),
             ("runc", "Lightweight OCI CLI wrapper. Actually calls Linux kernel syscalls to configure namespaces and cgroups, then execs the process")
         ],
         badge="COMPONENTS", accent_color=COLOR_ACCENT_PURPLE)

add_card(s16, Inches(6.9), Inches(1.9), Inches(5.6), Inches(5.0),
         "End-to-End Execution Flow (`docker run nginx`)",
         [
             ("1. CLI Command", "Aanand runs `docker run nginx` -> CLI converts to HTTP POST"),
             ("2. Socket Dispatch", "Sent across `/var/run/docker.sock` to `dockerd`"),
             ("3. Image Pull", "`dockerd` contacts Docker Hub, pulls missing image layers"),
             ("4. Handover to containerd", "`dockerd` passes container spec to `containerd`"),
             ("5. runc Invocation", "`containerd` spawns `runc` to configure namespaces/cgroups"),
             ("6. Container Alive", "Nginx process starts as PID 1; runc exits; container runs!")
         ],
         badge="EXECUTION FLOW", accent_color=COLOR_ACCENT_CYAN)


# ==============================================================================
# SLIDE 17: 12. DOCKER ENGINE VS CLI VS DAEMON
# ==============================================================================
s17 = prs.slides.add_slide(blank_layout)
set_slide_background(s17)
add_header(s17, "Chapter 12", "Docker Engine vs CLI vs Daemon",
           "Clarifying the terminology that confuses beginners and interviewees.")

add_card(s17, Inches(0.8), Inches(1.9), Inches(3.6), Inches(5.0),
         "Docker Engine",
         [
             ("Definition", "The entire software product suite"),
             ("What's Included", "CLI + dockerd + containerd + runc + plugins"),
             ("Package", "What gets installed when you install Docker on Ubuntu or macOS"),
             ("Analogy", "The entire car including wheels, engine, and controls")
         ],
         badge="THE WHOLE CAR", accent_color=COLOR_ACCENT_BLUE)

add_card(s17, Inches(4.8), Inches(1.9), Inches(3.6), Inches(5.0),
         "Docker CLI",
         [
             ("Definition", "The command-line client binary (`/usr/bin/docker`)"),
             ("Role", "The user interface; sends API calls to daemon"),
             ("Flexibility", "Can target remote servers: `docker -H prod.server:2376 ps`"),
             ("Analogy", "The steering wheel and dashboard gauges")
         ],
         badge="THE STEERING WHEEL", accent_color=COLOR_ACCENT_CYAN)

add_card(s17, Inches(8.8), Inches(1.9), Inches(3.7), Inches(5.0),
         "Docker Daemon",
         [
             ("Definition", "The `dockerd` background service running as root"),
             ("Role", "Heavy lifter managing storage, networking, and containers"),
             ("Socket", "Listens on `/var/run/docker.sock`"),
             ("Security Alert", "Access to docker.sock equals ROOT ACCESS on the host!")
         ],
         badge="THE ENGINE", accent_color=COLOR_ACCENT_AMBER)


# ==============================================================================
# SLIDE 18: 13. CONTAINER LIFECYCLE THEORY
# ==============================================================================
s18 = prs.slides.add_slide(blank_layout)
set_slide_background(s18)
add_header(s18, "Chapter 13", "Container Lifecycle Theory: 4 Golden Rules",
           "Essential mental models for building resilient, production-ready container architectures.")

add_card(s18, Inches(0.8), Inches(1.9), Inches(5.6), Inches(5.0),
         "Rules 1 & 2: Process & State",
         [
             ("Rule 1: PID 1 Controls Life", "A container only lives as long as its PID 1 process. If PID 1 completes its script or crashes, the container stops immediately. Keep processes in the foreground!"),
             ("Rule 2: Ephemeral Writable Layer", "Data written to the container layer survives `docker stop` and `docker restart`, but is PERMANENTLY DESTROYED upon `docker rm`. Always use Docker Volumes for databases and state")
         ],
         badge="RULES 1 & 2", accent_color=COLOR_ACCENT_CYAN)

add_card(s18, Inches(6.9), Inches(1.9), Inches(5.6), Inches(5.0),
         "Rules 3 & 4: Immutability & Scope",
         [
             ("Rule 3: Images Are Immutable", "Images are read-only templates. You never modify an existing image in place; you build a new tag. Consistency is guaranteed across environments"),
             ("Rule 4: One Primary Concern Per Container", "Follow the Unix philosophy: one service per container (e.g. separate Web App and Database containers). Makes apps horizontally scalable, decoupled, and easy to debug")
         ],
         badge="RULES 3 & 4", accent_color=COLOR_ACCENT_GREEN)


# ==============================================================================
# SLIDE 19: 14. AANAND'S AHA MOMENT
# ==============================================================================
s19 = prs.slides.add_slide(blank_layout)
set_slide_background(s19)
add_header(s19, "Chapter 14", "Aanand's 'Aha!' Moment: Putting It All Together",
           "Connecting every theoretical concept into a unified mental model.")

add_card(s19, Inches(0.8), Inches(1.9), Inches(5.6), Inches(5.0),
         "What Aanand Realized",
         [
             ("No Magic Boxes", "Containers are NOT miniature virtual machines. They are standard Linux processes running directly on the host kernel"),
             ("The Isolation Formula", "Container = Process + Namespaces (View) + cgroups (Limits) + OverlayFS (Storage)"),
             ("The End of 'It Works on My Machine'", "By packaging exact Python 3.9 runtime + libraries into read-only image layers, local and production run identical code"),
             ("Lightweight Scale", "Can boot 50 containers in 3 seconds using 1GB RAM, where 50 VMs would require a datacenter")
         ],
         badge="CLARITY", accent_color=COLOR_ACCENT_AMBER)

add_card(s19, Inches(6.9), Inches(1.9), Inches(5.6), Inches(5.0),
         "The Production Triumph",
         [
             ("One Command", "Aanand runs: `docker run -d --name web-app -p 8080:8080 python-app:3.9`"),
             ("Zero Conflicts", "Container binds to host port 8080 without touching server libraries"),
             ("Complete Predictability", "App starts in 80 milliseconds and runs flawlessly on the production server"),
             ("Team Onboarding", "New engineers run `docker compose up` and are productive on Day 1!")
         ],
         badge="SUCCESS", accent_color=COLOR_ACCENT_GREEN)


# ==============================================================================
# SLIDE 20: 15. KEY TAKEAWAYS & WHAT'S NEXT
# ==============================================================================
s20 = prs.slides.add_slide(blank_layout)
set_slide_background(s20)
add_header(s20, "Summary & Next Class", "Key Takeaways & Upcoming Topics",
           "Reviewing today's foundational concepts and looking ahead to Class 2.")

add_card(s20, Inches(0.8), Inches(1.9), Inches(5.6), Inches(5.0),
         "Today's Core Takeaways",
         [
             ("Docker Solves Parity", "Eliminates environment drift between Dev, Staging & Production"),
             ("Kernel Triad", "Namespaces (what you see) + cgroups (what you use) + UnionFS (how files stack)"),
             ("Process Lifecycle", "A container lives and dies with its PID 1 process"),
             ("Client-Server Model", "CLI talks REST API to dockerd; dockerd delegates to containerd & runc")
         ],
         badge="RECAP", accent_color=COLOR_ACCENT_CYAN)

add_card(s20, Inches(6.9), Inches(1.9), Inches(5.6), Inches(5.0),
         "Preview: Class 2 & 3",
         [
             ("Class 2: Images & Containers Deep Dive", "Anatomy of a Docker Image, Layer Caching, Multi-Stage Builds"),
             ("Class 3: Building Custom Images", "Writing production Dockerfiles, reducing image sizes from 1GB to 50MB"),
             ("Upcoming Labs", "Hands-on container networking, persistent volume mounts, and security scanning"),
             ("Roadmap Checkpoint", "Preparing the foundation for Kubernetes Orchestration!")
         ],
         badge="NEXT UP", accent_color=COLOR_ACCENT_PURPLE)


# Save presentation
prs.save(OUTPUT_PPTX)
print(f"✅ Presentation successfully generated with {len(prs.slides)} slides at: {OUTPUT_PPTX}")
