#!/usr/bin/env python3
"""
Generate the 'Docker Images & Containers' PowerPoint Presentation.
Full 16:9 Widescreen Dark-Theme Deck incorporating:
- Complete Table of Contents
- The Bakery Analogy 🍰 (Recipe, Cake, Layers, Registry)
- The Story of Maya: From Bloated Images to Optimized Containers
- Modules 1 through 10 with verbatim technical depth and full code examples
- Key Takeaways & Master Cheat Sheet
"""

import os
import shutil
import pptx
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

OUTPUT_LOCAL = "Docker_Images_and_Containers.pptx"
OUTPUT_DOWNLOADS = "/Users/aanand-mac/Downloads/Docker_Images_and_Containers.pptx"

# Modern Slate / Tailwind Dark Tech Palette
COLOR_BG = RGBColor(15, 23, 42)          # Slate 900
COLOR_CARD_BG = RGBColor(30, 41, 59)     # Slate 800
COLOR_CARD_BORDER = RGBColor(51, 65, 85) # Slate 700
COLOR_CODE_BG = RGBColor(10, 15, 29)     # Slate 950 (Deep code box)
COLOR_CODE_BORDER = RGBColor(40, 53, 72) # Code border
COLOR_TEXT_MAIN = RGBColor(248, 250, 252)# White Header
COLOR_TEXT_MUTED = RGBColor(148, 163, 184)# Slate 400 Subtitle
COLOR_TEXT_BODY = RGBColor(226, 232, 240)# Light Body Text

# Accent Colors
COLOR_ACCENT_CYAN = RGBColor(56, 189, 248)   # Sky 400
COLOR_ACCENT_BLUE = RGBColor(59, 130, 246)   # Blue 500
COLOR_ACCENT_GREEN = RGBColor(16, 185, 129)  # Emerald 500
COLOR_ACCENT_AMBER = RGBColor(245, 158, 11)  # Amber 500
COLOR_ACCENT_PURPLE = RGBColor(168, 85, 247) # Purple 500
COLOR_ACCENT_RED = RGBColor(239, 68, 68)     # Red 500
COLOR_ACCENT_PINK = RGBColor(236, 72, 153)   # Pink 500

prs = pptx.Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

def set_slide_background(slide):
    """Fills slide with dark slate background."""
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = COLOR_BG
    bg.line.fill.background()
    return bg

def add_header(slide, category_tag, title_text, subtitle_text=None):
    """Draws consistent, high-contrast header."""
    header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(1.25))
    tf = header_box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

    p_tag = tf.paragraphs[0]
    p_tag.text = category_tag.upper()
    p_tag.font.name = "Arial"
    p_tag.font.size = Pt(11)
    p_tag.font.bold = True
    p_tag.font.color.rgb = COLOR_ACCENT_CYAN
    p_tag.space_after = Pt(3)

    p_title = tf.add_paragraph()
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
        p_sub.font.size = Pt(12)
        p_sub.font.color.rgb = COLOR_TEXT_MUTED
    else:
        p_title.space_after = Pt(8)

def add_card(slide, left, top, width, height, title, items, badge=None, accent_color=COLOR_ACCENT_CYAN,
             font_size=11, bg_color=COLOR_CARD_BG, border_color=COLOR_CARD_BORDER):
    """Creates a rounded container card with title, badge, and content items."""
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    card.line.color.rgb = border_color
    card.line.width = Pt(1.2)

    tb = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.16), width - Inches(0.4), height - Inches(0.32))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

    p_first = tf.paragraphs[0]
    if badge:
        r_b = p_first.add_run()
        r_b.text = f"[{badge}] "
        r_b.font.name = "Arial"
        r_b.font.size = Pt(12.5)
        r_b.font.bold = True
        r_b.font.color.rgb = accent_color

        r_t = p_first.add_run()
        r_t.text = title
        r_t.font.name = "Arial"
        r_t.font.size = Pt(14.5)
        r_t.font.bold = True
        r_t.font.color.rgb = COLOR_TEXT_MAIN
    else:
        r_t = p_first.add_run()
        r_t.text = title
        r_t.font.name = "Arial"
        r_t.font.size = Pt(15)
        r_t.font.bold = True
        r_t.font.color.rgb = accent_color
    p_first.space_after = Pt(6)

    for item in items:
        p = tf.add_paragraph()
        p.space_after = Pt(3.5)
        p.font.name = "Arial"
        p.font.size = Pt(font_size)

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
            line_str = str(item)
            if line_str.startswith("• ") or line_str.startswith("- "):
                run = p.add_run()
                run.text = "• " + line_str.lstrip("•- ")
                run.font.color.rgb = COLOR_TEXT_BODY
            elif ":" in line_str and not line_str.startswith("http") and not line_str.startswith("    "):
                parts = line_str.split(":", 1)
                run1 = p.add_run()
                run1.text = "• " + parts[0].strip() + ": "
                run1.font.bold = True
                run1.font.color.rgb = COLOR_TEXT_MAIN

                run2 = p.add_run()
                run2.text = parts[1].strip()
                run2.font.bold = False
                run2.font.color.rgb = COLOR_TEXT_BODY
            else:
                run = p.add_run()
                run.text = line_str
                run.font.color.rgb = COLOR_TEXT_BODY

    return card

def add_code_box(slide, left, top, width, height, title, code_lines, badge="CODE", accent_color=COLOR_ACCENT_GREEN, font_size=10):
    """Renders a dark terminal / code block."""
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = COLOR_CODE_BG
    card.line.color.rgb = COLOR_CODE_BORDER
    card.line.width = Pt(1.2)

    tb = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.15), width - Inches(0.4), height - Inches(0.3))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

    p_first = tf.paragraphs[0]
    r_b = p_first.add_run()
    r_b.text = f"[{badge}] "
    r_b.font.name = "Arial"
    r_b.font.size = Pt(11.5)
    r_b.font.bold = True
    r_b.font.color.rgb = accent_color

    r_t = p_first.add_run()
    r_t.text = title
    r_t.font.name = "Arial"
    r_t.font.size = Pt(12.5)
    r_t.font.bold = True
    r_t.font.color.rgb = COLOR_TEXT_MAIN
    p_first.space_after = Pt(5)

    for line in code_lines:
        p = tf.add_paragraph()
        p.space_after = Pt(2)
        p.font.name = "Courier New"
        p.font.size = Pt(font_size)

        line_str = str(line)
        run = p.add_run()
        run.text = line_str
        if line_str.startswith("#"):
            run.font.color.rgb = COLOR_TEXT_MUTED
            run.font.italic = True
        elif line_str.startswith("FROM") or line_str.startswith("RUN") or line_str.startswith("COPY") or line_str.startswith("CMD") or line_str.startswith("WORKDIR") or line_str.startswith("EXPOSE"):
            run.font.color.rgb = COLOR_ACCENT_CYAN
            run.font.bold = True
        elif line_str.startswith("docker ") or line_str.startswith("$ "):
            run.font.color.rgb = COLOR_ACCENT_AMBER
            run.font.bold = True
        elif "←" in line_str or "→" in line_str:
            run.font.color.rgb = COLOR_ACCENT_GREEN
        else:
            run.font.color.rgb = COLOR_TEXT_BODY

    return card

def add_metric_banner(slide, left, top, width, height, stat_number, stat_label, subtext, accent_color=COLOR_ACCENT_CYAN):
    """Draws a metric badge for highlighting significant optimizations."""
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = COLOR_CARD_BG
    card.line.color.rgb = accent_color
    card.line.width = Pt(1.5)

    tb = slide.shapes.add_textbox(left + Inches(0.15), top + Inches(0.1), width - Inches(0.3), height - Inches(0.2))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

    p_num = tf.paragraphs[0]
    p_num.text = stat_number
    p_num.font.name = "Arial"
    p_num.font.size = Pt(20)
    p_num.font.bold = True
    p_num.font.color.rgb = accent_color
    p_num.space_after = Pt(2)

    p_lbl = tf.add_paragraph()
    p_lbl.text = stat_label
    p_lbl.font.name = "Arial"
    p_lbl.font.size = Pt(11.5)
    p_lbl.font.bold = True
    p_lbl.font.color.rgb = COLOR_TEXT_MAIN
    p_lbl.space_after = Pt(2)

    p_sub = tf.add_paragraph()
    p_sub.text = subtext
    p_sub.font.name = "Arial"
    p_sub.font.size = Pt(9.5)
    p_sub.font.color.rgb = COLOR_TEXT_MUTED


# ==============================================================================
# SLIDE 1: TITLE SLIDE
# ==============================================================================
s1 = prs.slides.add_slide(blank_layout)
set_slide_background(s1)

tb1 = s1.shapes.add_textbox(Inches(1.0), Inches(1.5), Inches(11.333), Inches(4.8))
tf1 = tb1.text_frame
tf1.word_wrap = True

p1 = tf1.paragraphs[0]
p1.text = "DEVOPS & CONTAINER MASTERCLASS"
p1.font.name = "Arial"
p1.font.size = Pt(14)
p1.font.bold = True
p1.font.color.rgb = COLOR_ACCENT_CYAN
p1.space_after = Pt(14)

p2 = tf1.add_paragraph()
p2.text = "Docker Images & Containers"
p2.font.name = "Arial"
p2.font.size = Pt(46)
p2.font.bold = True
p2.font.color.rgb = COLOR_TEXT_MAIN
p2.space_after = Pt(14)

p3 = tf1.add_paragraph()
p3.text = "The Story of Maya: From Bloated Images to Optimized Containers"
p3.font.name = "Arial"
p3.font.size = Pt(22)
p3.font.bold = True
p3.font.color.rgb = COLOR_ACCENT_AMBER
p3.space_after = Pt(18)

p4 = tf1.add_paragraph()
p4.text = "Featuring the continuous Bakery Analogy 🍰 (Recipes, Cakes, Layers & Registries)\nMastering UnionFS, Manifests, Container Lifecycle, Optimization, Multi-Stage Builds & Versioning"
p4.font.name = "Arial"
p4.font.size = Pt(14)
p4.font.color.rgb = COLOR_TEXT_BODY
p4.space_after = Pt(20)

p5 = tf1.add_paragraph()
p5.text = "16:9 Widescreen • Comprehensive Architectural & Operational Guide"
p5.font.name = "Arial"
p5.font.size = Pt(12)
p5.font.color.rgb = COLOR_TEXT_MUTED


# ==============================================================================
# SLIDE 2: TABLE OF CONTENTS / EXECUTIVE AGENDA
# ==============================================================================
s2 = prs.slides.add_slide(blank_layout)
set_slide_background(s2)
add_header(s2, "Course Agenda", "Docker Images & Containers: Table of Contents",
           "Complete 10-module curriculum tracking Maya's journey from confusion to production mastery.")

add_card(s2, Inches(0.8), Inches(1.8), Inches(5.7), Inches(4.3),
         "Modules 1 — 5: Core Primitives & Lifecycle",
         [
             ("Module 1", "What is a Docker Image? Anatomy & Characteristics"),
             ("Module 2", "Image Layers & Union Filesystem (Overlay2, CoW)"),
             ("Module 3", "Image Manifests & Metadata (Digests vs Tags)"),
             ("Module 4", "How Images Are Built & Stored (Caching & Storage Driver)"),
             ("Module 5", "Docker Container Lifecycle (States, Transitions & Volumes)")
         ],
         badge="FOUNDATION", accent_color=COLOR_ACCENT_CYAN, font_size=11.5)

add_card(s2, Inches(6.8), Inches(1.8), Inches(5.7), Inches(4.3),
         "Modules 6 — 10: Operations, Optimization & CI/CD",
         [
             ("Module 6", "Important Container Commands (run, start, stop, exec, ps)"),
             ("Module 7", "Important Image Commands (images, pull, push, tag, rmi)"),
             ("Module 8", "Image Optimization Strategies (Multi-Stage, Alpine, .dockerignore)"),
             ("Module 9", "Docker Registries Overview (Docker Hub, ECR, GCR, ACR)"),
             ("Module 10", "Handling Image Tags & Versioning (SemVer, SHA, Rollbacks)"),
             ("Summary", "Key Takeaways & Master Architecture Reference")
         ],
         badge="PRODUCTION", accent_color=COLOR_ACCENT_GREEN, font_size=11.5)

# Maya narrative banner across bottom
add_card(s2, Inches(0.8), Inches(6.25), Inches(11.7), Inches(0.9),
         "The Story of Maya",
         [
             ("Maya's Challenge", "DevOps engineer tasked with containerizing enterprise apps; starts with 2GB bloated images and broken :latest tags, transforming into an optimization expert.")
         ],
         badge="STORYLINE", accent_color=COLOR_ACCENT_AMBER, font_size=10.5)


# ==============================================================================
# SLIDE 3: THE LEARNING ANALOGY: THE BAKERY 🍰
# ==============================================================================
s3 = prs.slides.add_slide(blank_layout)
set_slide_background(s3)
add_header(s3, "Mental Model", "The Learning Analogy: The Bakery 🍰",
           "To help your learners grasp these concepts effortlessly, we use a continuous baking analogy throughout.")

col_w = Inches(2.78)
gap = Inches(0.2)
top_pos = Inches(1.8)
h_pos = Inches(4.4)

add_card(s3, Inches(0.8), top_pos, col_w, h_pos,
         "Docker Image",
         [
             ("Bakery Item", "The Master Recipe Card"),
             ("Nature", "Static, read-only file on disk"),
             ("Mutability", "Never changes once written"),
             ("Role", "Complete formula containing all ingredients, steps, and measurements"),
             ("Power", "Used to bake unlimited identical cakes without modifying the card")
         ],
         badge="RECIPE", accent_color=COLOR_ACCENT_CYAN, font_size=10.5)

add_card(s3, Inches(0.8) + (col_w + gap), top_pos, col_w, h_pos,
         "Docker Container",
         [
             ("Bakery Item", "The Actual Baked Cake"),
             ("Nature", "Running, active process in memory"),
             ("Mutability", "Can be eaten, decorated, modified, or thrown away"),
             ("Role", "Consumes real kitchen resources (oven heat, counter space)"),
             ("Lifespan", "Ephemeral instance created directly from the recipe")
         ],
         badge="CAKE", accent_color=COLOR_ACCENT_GREEN, font_size=10.5)

add_card(s3, Inches(0.8) + (col_w + gap) * 2, top_pos, col_w, h_pos,
         "Image Layers",
         [
             ("Bakery Item", "Stacked Ingredients"),
             ("Nature", "Step-by-step filesystem layers"),
             ("Sequence", "Flour at base, then sugar, then cocoa, then frosting"),
             ("Immutability", "Baked ingredient layers stay frozen in place"),
             ("Top Layer", "Only the frosting layer on top can be changed")
         ],
         badge="INGREDIENTS", accent_color=COLOR_ACCENT_AMBER, font_size=10.5)

add_card(s3, Inches(0.8) + (col_w + gap) * 3, top_pos, col_w, h_pos,
         "Docker Registry",
         [
             ("Bakery Item", "The Master Recipe Book"),
             ("Nature", "Central library / distribution shelf"),
             ("Role", "Where bakers store and share recipes with the global community"),
             ("Distribution", "Download proven recipes (Docker Hub) or keep secret family recipes private (ECR)"),
             ("Access", "Protected with bakery member credentials")
         ],
         badge="LIBRARY", accent_color=COLOR_ACCENT_PURPLE, font_size=10.5)

# Bottom Takeaway Callout
add_card(s3, Inches(0.8), Inches(6.35), Inches(11.7), Inches(0.8),
         "Maya's Aha! Moment",
         [
             ("Key Insight", "One single Recipe Card (Image) can bake 1,000 delicious Cakes (Containers) simultaneously without altering a single word on the card!")
         ],
         badge="AHA!", accent_color=COLOR_ACCENT_PINK, font_size=11)


# ==============================================================================
# SLIDE 4: MODULE 1 — WHAT IS A DOCKER IMAGE
# ==============================================================================
s4 = prs.slides.add_slide(blank_layout)
set_slide_background(s4)
add_header(s4, "Module 1", "What is a Docker Image? Anatomy & Characteristics",
           "Maya's first question: 'What exactly is an image, and what lives inside it?'")

add_card(s4, Inches(0.8), Inches(1.8), Inches(5.7), Inches(5.2),
         "Core Anatomy: What Lives Inside?",
         [
             ("Definition", "A read-only template containing everything needed to run an application reliably:"),
             ("Operating System Files", "Base Linux rootfs (/bin, /lib, /usr, system binaries)"),
             ("Application Code", "Source scripts, compiled binaries, templates, assets"),
             ("Dependencies & Libs", "Python virtualenvs, node_modules, shared C libraries (.so)"),
             ("Environment Variables", "Default runtime paths, ports, configuration keys (PATH, PORT)"),
             ("Default Commands", "ENTRYPOINT and CMD specifying default execution behavior"),
             ("Bakery Connection", "The complete recipe card listing flour, sugar, temperature, and bake timer!")
         ],
         badge="INSIDE IMAGE", accent_color=COLOR_ACCENT_CYAN, font_size=11)

add_card(s4, Inches(6.8), Inches(1.8), Inches(5.7), Inches(3.3),
         "The 4 Key Characteristics",
         [
             ("1. Immutable", "Once built, an image NEVER changes. No in-place editing or silent drift"),
             ("2. Layered", "Constructed from multiple stacked filesystem deltas that snap together"),
             ("3. Shareable", "Easily packaged, pushed, and pulled across registries worldwide"),
             ("4. Versioned", "Uniquely tagged with human-readable versions (e.g., nginx:1.21, python:3.9)")
         ],
         badge="PROPERTIES", accent_color=COLOR_ACCENT_GREEN, font_size=11.5)

add_card(s4, Inches(6.8), Inches(5.25), Inches(5.7), Inches(1.75),
         "Maya's Realization",
         [
             ("The End of Chaos", "An image is self-contained. It packages code + dependencies + OS files into a single immutable artifact, eliminating 'works on my machine' forever.")
         ],
         badge="MAYA'S STORY", accent_color=COLOR_ACCENT_AMBER, font_size=11)


# ==============================================================================
# SLIDE 5: MODULE 1 — IMAGE VS CONTAINER DEEP DIVE
# ==============================================================================
s5 = prs.slides.add_slide(blank_layout)
set_slide_background(s5)
add_header(s5, "Module 1", "Image vs. Container: Blueprint vs. Active Process",
           "Understanding the fundamental distinction between static templates and running runtimes.")

add_card(s5, Inches(0.8), Inches(1.8), Inches(5.7), Inches(4.3),
         "Docker Image (Static Blueprint)",
         [
             ("What It Is", "Static file stored on disk; inactive blueprint"),
             ("State", "Completely frozen, passive, read-only"),
             ("Resource Usage", "Consumes ZERO CPU and ZERO RAM (disk only)"),
             ("Lifecycle", "Built once, stored, tagged, shared, and reused"),
             ("Bakery Analogy", "The Master Recipe Card in the kitchen binder"),
             ("Distinction", "You cannot execute an image directly; it is merely instructions and frozen files")
         ],
         badge="STATIC BLUEPRINT", accent_color=COLOR_ACCENT_CYAN, font_size=11.5)

add_card(s5, Inches(6.8), Inches(1.8), Inches(5.7), Inches(4.3),
         "Docker Container (Active Instance)",
         [
             ("What It Is", "Running instance of an image; active OS process (PID 1)"),
             ("State", "Dynamic runtime with a thin Read-Write container layer"),
             ("Resource Usage", "Actively uses host CPU, RAM, Network & Disk I/O"),
             ("Lifecycle", "Created → Running → Paused → Stopped → Removed"),
             ("Bakery Analogy", "The actual Baked Cake ready to eat or decorate"),
             ("Distinction", "Containers have isolated namespaces, network interfaces, and ephemeral writable storage")
         ],
         badge="RUNNING INSTANCE", accent_color=COLOR_ACCENT_GREEN, font_size=11.5)

add_card(s5, Inches(0.8), Inches(6.25), Inches(11.7), Inches(0.9),
         "The 1-to-Many Multiplier",
         [
             ("Core Rule", "Maya can create infinite isolated containers from one single image, just like baking multiple cakes from one master recipe card! Each container gets its own private writable layer.")
         ],
         badge="POWER OF DOCKER", accent_color=COLOR_ACCENT_PURPLE, font_size=11)


# ==============================================================================
# SLIDE 6: MODULE 2 — IMAGE LAYERS & HOW THEY WORK
# ==============================================================================
s6 = prs.slides.add_slide(blank_layout)
set_slide_background(s6)
add_header(s6, "Module 2", "Image Layers: Building the Stack",
           "Maya discovers images aren't single files—they're stacks of stacked filesystem deltas.")

# Left Card: Dockerfile Instruction Mapping
add_code_box(s6, Inches(0.8), Inches(1.8), Inches(5.7), Inches(5.2),
             "Dockerfile Instruction → Layer Generation",
             [
                 "# Each instruction creates a distinct immutable layer:",
                 "FROM ubuntu:20.04          → Layer 1 (Base OS rootfs)",
                 "RUN apt-get update         → Layer 2 (Package lists)",
                 "RUN apt-get install python → Layer 3 (Python runtime)",
                 "COPY app.py /app/          → Layer 4 (Application code)",
                 "CMD python /app/app.py     → Layer 5 (Metadata, no files)",
                 "",
                 "# Result:",
                 "# Image with 5 layers stacked like pancakes!",
                 "# Lower layers are 100% read-only and immutable.",
                 "# Each layer stores only the DELTA (changes from previous)."
             ],
             badge="DOCKERFILE", accent_color=COLOR_ACCENT_CYAN, font_size=10)

# Right Card: Visual Layer Stack
add_card(s6, Inches(6.8), Inches(1.8), Inches(5.7), Inches(5.2),
         "The Visual Layer Stack (Like Pancakes)",
         [
             ("Top Layer", "[Layer 5: Metadata] Execution CMD & entrypoint (0 MB, metadata only)"),
             ("Layer 4", "[Layer 4: App Code] COPY app.py /app/ (50 KB application logic)"),
             ("Layer 3", "[Layer 3: Binaries] /usr/bin/python3, standard libraries (120 MB)"),
             ("Layer 2", "[Layer 2: Indices] /var/lib/apt/lists package catalogs (35 MB)"),
             ("Bottom Base", "[Layer 1: Base OS] /bin/bash, /lib, core system rootfs (77 MB)"),
             ("Cryptographic Integrity", "Each layer has its own unique SHA256 content hash"),
             ("Stacking Order", "Lower layers are locked; higher layers add, modify, or mask lower files")
         ],
         badge="LAYER STACK", accent_color=COLOR_ACCENT_AMBER, font_size=11)


# ==============================================================================
# SLIDE 7: MODULE 2 — UNION FILESYSTEM & LAYER SHARING
# ==============================================================================
s7 = prs.slides.add_slide(blank_layout)
set_slide_background(s7)
add_header(s7, "Module 2", "Union Filesystem Magic & Layer Sharing",
           "How Docker merges layers into one filesystem and cuts disk usage in half.")

add_card(s7, Inches(0.8), Inches(1.8), Inches(5.7), Inches(3.2),
         "Union Filesystem Magic (Overlay2)",
         [
             ("Unified View", "All stacked layers are transparently merged into a single filesystem view"),
             ("Lower Layers", "Underlying image layers remain strictly read-only"),
             ("Top Layer", "Thin writable layer added on top for container execution"),
             ("Seamless Linux Tree", "The running container sees all files in one cohesive structure:"),
             ("• Layer 1", "/bin/bash, /usr/lib/"),
             ("• Layer 2", "/usr/bin/python"),
             ("• Layer 3", "/app/app.py (Container sees everything in standard / hierarchy)")
         ],
         badge="OVERLAY2", accent_color=COLOR_ACCENT_CYAN, font_size=10.5)

add_card(s7, Inches(6.8), Inches(1.8), Inches(5.7), Inches(3.2),
         "Layer Sharing Across Multiple Images",
         [
             ("Maya's Three Images", "Maya hosts three different services on the exact same server:"),
             ("• Image A", "Ubuntu + Python + App 1"),
             ("• Image B", "Ubuntu + Python + App 2"),
             ("• Image C", "Ubuntu + Node.js + App 3"),
             ("On-Disk Storage", "Ubuntu layer: stored ONCE (shared by A, B, and C)"),
             ("Python layer", "stored ONCE (shared by A and B)"),
             ("App layers", "Unique to each specific image")
         ],
         badge="DEDUPLICATION", accent_color=COLOR_ACCENT_GREEN, font_size=10.5)

# Bottom Storage Savings Banner
add_card(s7, Inches(0.8), Inches(5.15), Inches(11.7), Inches(1.95),
         "The Storage Savings Payoff: 3GB → 1.5GB",
         [
             ("Traditional Virtual Machines", "3 separate VMs require 3 full OS installs + separate runtimes = ~3.0 GB to 6.0 GB of duplicate disk space!"),
             ("Docker Layer Deduplication", "Docker stores common base layers exactly once on host disk. Total footprint drops to ~1.5 GB!"),
             ("Network Bandwidth Bonus", "When pulling Image B, Docker skips downloading Ubuntu and Python because they already exist on disk! Only the unique App 2 layer is transferred.")
         ],
         badge="50% SAVINGS", accent_color=COLOR_ACCENT_AMBER, font_size=11)


# ==============================================================================
# SLIDE 8: MODULE 2 — COPY-ON-WRITE (CoW) MECHANICS
# ==============================================================================
s8 = prs.slides.add_slide(blank_layout)
set_slide_background(s8)
add_header(s8, "Module 2", "Copy-on-Write (CoW) Mechanics",
           "Why containers start in milliseconds and safely share image layers without data corruption.")

col3_w = Inches(3.75)
gap3 = Inches(0.22)

add_card(s8, Inches(0.8), Inches(1.8), col3_w, Inches(3.6),
         "Step 1: Read-Only Access",
         [
             ("Initial State", "File exists inside the shared read-only image layer (e.g. /etc/app.conf)"),
             ("Fast Reads", "Container reads directly from the underlying immutable layer"),
             ("Zero Overhead", "No copying takes place; multiple containers read the same physical block on disk")
         ],
         badge="STEP 1", accent_color=COLOR_ACCENT_CYAN, font_size=11)

add_card(s8, Inches(0.8) + (col3_w + gap3), Inches(1.8), col3_w, Inches(3.6),
         "Step 2: Modification Trigger",
         [
             ("Write Request", "Container process attempts to edit or delete /etc/app.conf"),
             ("CoW Intercept", "Storage driver intercepts write operation"),
             ("Instant Copy", "File is copied up from read-only layer into container's writable layer"),
             ("Original Intact", "Underlying image layer file remains completely untouched!")
         ],
         badge="STEP 2", accent_color=COLOR_ACCENT_AMBER, font_size=11)

add_card(s8, Inches(0.8) + (col3_w + gap3) * 2, Inches(1.8), col3_w, Inches(3.6),
         "Step 3: Isolated Mutation",
         [
             ("Changes Applied", "Edits happen strictly to the copied version in writable layer"),
             ("Isolation", "Container sees its modified version; other containers see original"),
             ("Pristine Image", "If container is removed, base image layers remain 100% clean")
         ],
         badge="STEP 3", accent_color=COLOR_ACCENT_GREEN, font_size=11)

# Why CoW Matters
add_card(s8, Inches(0.8), Inches(5.6), Inches(11.7), Inches(1.5),
         "Why Copy-on-Write is Revolutionary for Containers",
         [
             ("⚡ Sub-Second Boot", "Containers start instantly because Docker doesn't copy gigabytes of OS files upon startup"),
             ("🛡️ Safe Concurrency", "Hundreds of containers can safely share the same base image without interfering or corrupting data"),
             ("💾 Minimal Disk Footprint", "Disk space is only consumed when files are actively modified or new files are created")
         ],
         badge="ARCHITECTURAL IMPACT", accent_color=COLOR_ACCENT_PURPLE, font_size=11)


# ==============================================================================
# SLIDE 9: MODULE 3 — IMAGE MANIFESTS AND METADATA
# ==============================================================================
s9 = prs.slides.add_slide(blank_layout)
set_slide_background(s9)
add_header(s9, "Module 3", "Image Manifests and Metadata",
           "Maya learns images contain more than just files—they are defined by JSON specifications.")

add_card(s9, Inches(0.8), Inches(1.8), Inches(5.7), Inches(3.4),
         "The Image Manifest (JSON Descriptor)",
         [
             ("What It Is", "A standardized JSON document describing the complete image package"),
             ("Layer Digests", "Ordered array of cryptographic SHA256 hashes verifying layer integrity"),
             ("Config Reference", "Pointer to the image configuration blob"),
             ("Target Architecture", "CPU architecture (amd64, arm64, s390x)"),
             ("Target OS", "Target operating system (linux, windows)"),
             ("Size Information", "Compressed and uncompressed byte totals"),
             ("Purpose", "Registries use manifests to verify layer integrity, support multi-architecture images, and enable parallel layer downloads")
         ],
         badge="MANIFEST.JSON", accent_color=COLOR_ACCENT_CYAN, font_size=10.5)

add_code_box(s9, Inches(6.8), Inches(1.8), Inches(5.7), Inches(3.4),
             "Image Config (Metadata Anatomy)",
             [
                 "{",
                 '  "ExposedPorts": { "8080/tcp": {} },',
                 '  "Env": [ "NODE_ENV=production", "PORT=8080" ],',
                 '  "WorkingDir": "/app",',
                 '  "Entrypoint": [ "/usr/bin/node" ],',
                 '  "Cmd": [ "server.js" ],',
                 '  "Labels": { "maintainer": "maya@company.com" },',
                 '  "architecture": "arm64",',
                 '  "os": "linux"',
                 "}"
             ],
             badge="CONFIG.JSON", accent_color=COLOR_ACCENT_GREEN, font_size=10)

add_card(s9, Inches(0.8), Inches(5.35), Inches(11.7), Inches(1.75),
         "Digests vs. Tags: The Crucial Difference",
         [
             ("Digest (SHA256 Hash)", "Immutable 256-bit cryptographic hash (e.g. sha256:abc12345...). NEVER changes. Mathematically unique identifier of the exact content"),
             ("Tag (Human Label)", "Human-readable label (e.g. nginx:latest, app:v1.0). Mutable pointer—can be moved to point to an entirely different image tomorrow!"),
             ("Production Best Practice", "Always reference images by Digest or strict immutable SemVer tags in production CI/CD pipelines to guarantee 100% reproducibility.")
         ],
         badge="PROD CRITICAL", accent_color=COLOR_ACCENT_AMBER, font_size=11)


# ==============================================================================
# SLIDE 10: MODULE 4 — HOW IMAGES ARE BUILT AND STORED
# ==============================================================================
s10 = prs.slides.add_slide(blank_layout)
set_slide_background(s10)
add_header(s10, "Module 4", "How Images Are Built and Stored",
           "Maya creates her first Dockerfile and watches the Docker engine build & store layers.")

add_card(s10, Inches(0.8), Inches(1.8), Inches(5.7), Inches(5.2),
         "Step-by-Step Build Process",
         [
             ("Step 1: Parse Dockerfile", "Docker reads instructions top to bottom, validating syntax"),
             ("Step 2: Pull Base Image", "Fetches FROM image from local cache or pulls from registry"),
             ("Step 3: Execute Instructions", "Each RUN/COPY spins up a temporary container, applies changes, and commits a new layer"),
             ("Step 4: Layer Caching", "Reuses unchanged layers from previous builds if hash matches"),
             ("Step 5: Generate Manifest", "Creates image configuration blob and manifest JSON"),
             ("Step 6: Tag Image", "Assigns repository name and tag alias to the resulting image ID"),
             ("⚠️ Build Context Warning", "Docker sends the ENTIRE current directory to the daemon. Keep it small!")
         ],
         badge="ENGINE PIPELINE", accent_color=COLOR_ACCENT_CYAN, font_size=11)

add_card(s10, Inches(6.8), Inches(1.8), Inches(5.7), Inches(5.2),
         "Where Images Are Stored on Local Disk (Linux)",
         [
             ("Main Storage Root", "/var/lib/docker/ (Host daemon directory)"),
             ("Image Layers", "/var/lib/docker/overlay2/ (Contains diffs, merged trees, and work directories)"),
             ("Metadata & Configs", "/var/lib/docker/image/ (Stores manifest JSON, layer metadata, and repositories.json)"),
             ("Storage Driver Options", "overlay2 (Default & industry standard), aufs, btrfs, devicemapper"),
             ("Overlay2 Structure", "lowerdir (read-only image layers) + upperdir (container writable layer) + merged (unified mount)"),
             ("🚨 Strict Security Rule", "NEVER manually edit, move, or delete files inside /var/lib/docker/! Always manage assets via Docker CLI commands.")
         ],
         badge="DISK STORAGE", accent_color=COLOR_ACCENT_AMBER, font_size=11)


# ==============================================================================
# SLIDE 11: MODULE 4 — BUILD LAYER CACHING & ORDERING
# ==============================================================================
s11 = prs.slides.add_slide(blank_layout)
set_slide_background(s11)
add_header(s11, "Module 4", "Build Layer Caching & Instruction Ordering",
           "Maya learns caching mechanics to turn 10-minute builds into 10-second builds.")

add_card(s11, Inches(0.8), Inches(1.8), Inches(5.7), Inches(5.2),
         "The Golden Rules of Layer Caching",
         [
             ("Cache Rule", "If instruction text AND referenced files are unchanged → Docker reuses cached layer instantly"),
             ("The Invalidation Cascade", "If ANY instruction or file changes, that layer AND ALL subsequent layers must rebuild from scratch!"),
             ("Strategic Ordering", "Put slow, rarely changing instructions at the top; put fast, frequently changing files at the bottom"),
             ("Maya's Discovery", "Putting 'COPY . .' before dependency install re-installed all libraries on every single commit!")
         ],
         badge="CACHE RULES", accent_color=COLOR_ACCENT_CYAN, font_size=11.5)

add_code_box(s11, Inches(6.8), Inches(1.8), Inches(5.7), Inches(5.2),
             "Maya's Caching Optimization in Action",
             [
                 "# Inefficient way (Rebuilds everything on every code edit):",
                 "# COPY . .                  <- breaks cache immediately!",
                 "# RUN pip install -r req... <- downloads for 5 minutes!",
                 "",
                 "# Optimized way (Maya's Caching Solution):",
                 "FROM python:3.9           <- CACHED (Never changes)",
                 "COPY requirements.txt .   <- CACHED (File unchanged)",
                 "RUN pip install -r req... <- CACHED (Reused from cache!)",
                 "COPY app.py .             <- REBUILDS (app.py changed)",
                 "CMD python app.py         <- REBUILDS (fast metadata)",
                 "",
                 "# Outcome: Dependency install skipped 99% of the time!",
                 "# Build time dropped from 5 minutes to 4 seconds."
             ],
             badge="OPTIMIZED DOCKERFILE", accent_color=COLOR_ACCENT_GREEN, font_size=10)


# ==============================================================================
# SLIDE 12: MODULE 5 — DOCKER CONTAINER LIFECYCLE
# ==============================================================================
s12 = prs.slides.add_slide(blank_layout)
set_slide_background(s12)
add_header(s12, "Module 5", "Docker Container Lifecycle: The 5 States",
           "Maya understands containers are ephemeral processes transitioning through distinct lifecycle states.")

box_w = Inches(2.18)
gap_s = Inches(0.18)
box_top = Inches(1.8)
box_h = Inches(3.6)

add_card(s12, Inches(0.8), box_top, box_w, box_h,
         "Created",
         [
             ("Layers", "Image layers referenced"),
             ("Writable", "RW layer created"),
             ("Namespaces", "Prepared"),
             ("State", "Not running yet")
         ],
         badge="STATE 1", accent_color=COLOR_ACCENT_CYAN, font_size=10.5)

add_card(s12, Inches(0.8) + (box_w + gap_s), box_top, box_w, box_h,
         "Running",
         [
             ("Process", "PID 1 executing"),
             ("Resources", "Consumes CPU & RAM"),
             ("Network", "Accepts connections"),
             ("Logs", "Generates stdout/stderr")
         ],
         badge="STATE 2", accent_color=COLOR_ACCENT_GREEN, font_size=10.5)

add_card(s12, Inches(0.8) + (box_w + gap_s) * 2, box_top, box_w, box_h,
         "Paused",
         [
             ("Process", "Frozen via SIGSTOP"),
             ("Memory", "Retained in RAM"),
             ("CPU", "0% CPU consumed"),
             ("Resume", "Instant with SIGCONT")
         ],
         badge="STATE 3", accent_color=COLOR_ACCENT_AMBER, font_size=10.5)

add_card(s12, Inches(0.8) + (box_w + gap_s) * 3, box_top, box_w, box_h,
         "Stopped",
         [
             ("Process", "PID 1 terminated"),
             ("Signal", "Exited via SIGTERM/KILL"),
             ("Data", "Writable layer intact"),
             ("Restart", "Can be restarted")
         ],
         badge="STATE 4", accent_color=COLOR_ACCENT_PURPLE, font_size=10.5)

add_card(s12, Inches(0.8) + (box_w + gap_s) * 4, box_top, box_w, box_h,
         "Removed",
         [
             ("Container", "Completely deleted"),
             ("Writable", "Layer destroyed"),
             ("Data", "Permanently lost"),
             ("Recovery", "Cannot be recovered")
         ],
         badge="STATE 5", accent_color=COLOR_ACCENT_RED, font_size=10.5)

# State Transitions Box
add_code_box(s12, Inches(0.8), Inches(5.6), Inches(11.7), Inches(1.5),
             "State Machine Transitions (Docker CLI Mapping)",
             [
                 "Image  ──[ docker run ]──>  Created  ──[ start ]──>  Running",
                 "                                                       │   ▲",
                 "                              [ pause ]  ┌─────────────┘   │ [ unpause ]",
                 "                                         ▼                 │",
                 "                                      Paused  ─────────────┘",
                 "                                         │",
                 "                              [ stop ]   ▼                 ┌── [ restart ]",
                 "                                      Stopped ─────────────┤",
                 "                                         │                 └── [ docker start ]",
                 "                             [ docker rm ]",
                 "                                         ▼",
                 "                                      Removed (Destroyed permanently)"
             ],
             badge="TRANSITIONS", accent_color=COLOR_ACCENT_CYAN, font_size=9.5)


# ==============================================================================
# SLIDE 13: MODULE 5 — LIFECYCLE BAKERY ANALOGY & DATA PERSISTENCE
# ==============================================================================
s13 = prs.slides.add_slide(blank_layout)
set_slide_background(s13)
add_header(s13, "Module 5", "Lifecycle Bakery Analogy & Data Persistence",
           "Connecting lifecycle states to baking stages and mastering volume persistence.")

add_card(s13, Inches(0.8), Inches(1.8), Inches(5.7), Inches(4.3),
         "The Bakery Analogy for Lifecycle States 🍰",
         [
             ("Created State", "Ingredients and batter are poured into the pan (ready to bake, not in oven yet)"),
             ("Running State", "The cake is actively baking inside the hot oven (using energy and heat)"),
             ("Paused State", "The oven door is opened, pausing the timer while the batter stays hot"),
             ("Stopped State", "The baked cake is taken out and resting on the counter (intact but cooled)"),
             ("Removed State", "The cake has been eaten or thrown into the trash (completely gone!)")
         ],
         badge="BAKERY ANALOGY", accent_color=COLOR_ACCENT_CYAN, font_size=11.5)

add_card(s13, Inches(6.8), Inches(1.8), Inches(5.7), Inches(4.3),
         "Data Persistence: Writable Layer vs. Volumes",
         [
             ("Writable Layer Data", "Tied directly to the container lifespan. Persists through stop/start, but permanently lost when container is removed (docker rm). Never shared!"),
             ("Docker Volumes", "Persistent storage hosted outside container on host disk (/var/lib/docker/volumes)"),
             ("Survives Deletion", "Volumes completely survive container removal"),
             ("Shareable", "Can be safely mounted by multiple containers concurrently"),
             ("High Performance", "Bypasses UnionFS storage driver for native host disk speed")
         ],
         badge="PERSISTENCE", accent_color=COLOR_ACCENT_GREEN, font_size=11.5)

add_card(s13, Inches(0.8), Inches(6.25), Inches(11.7), Inches(0.9),
         "Maya's Golden Storage Rule",
         [
             ("Rule of Thumb", "Never store database records, user uploads, or persistent state in the container writable layer—ALWAYS mount external Docker Volumes (-v)!")
         ],
         badge="GOLDEN RULE", accent_color=COLOR_ACCENT_AMBER, font_size=11)


# ==============================================================================
# SLIDE 14: MODULE 6 — IMPORTANT CONTAINER COMMANDS (PART 1)
# ==============================================================================
s14 = prs.slides.add_slide(blank_layout)
set_slide_background(s14)
add_header(s14, "Module 6", "Essential Container Commands: Mastering 'docker run'",
           "Syntax, critical runtime flags, and fine-grained resource governance.")

add_code_box(s14, Inches(0.8), Inches(1.8), Inches(11.7), Inches(1.1),
             "Master Syntax & Complete Command Example",
             [
                 "Syntax: docker run [options] image [command]",
                 "Example: docker run -d --name myapp -p 8080:80 -v /data:/app/data --memory 512m --cpus 1.5 --restart unless-stopped nginx:alpine"
             ],
             badge="SYNTAX", accent_color=COLOR_ACCENT_CYAN, font_size=11)

add_card(s14, Inches(0.8), Inches(3.05), Inches(5.7), Inches(4.1),
         "Essential Runtime Flags (Part 1)",
         [
             ("-d (Detached)", "Runs container in the background; prints container ID"),
             ("-p 8080:80 (Port Mapping)", "Maps host port 8080 to container port 80"),
             ("--name myapp", "Assigns a custom readable name instead of random string"),
             ("-e KEY=value", "Sets environment variables (e.g., -e NODE_ENV=production)"),
             ("-v /host:/container", "Mounts host folder or volume into container for persistent storage")
         ],
         badge="NETWORKING & MOUNTS", accent_color=COLOR_ACCENT_GREEN, font_size=11.5)

add_card(s14, Inches(6.8), Inches(3.05), Inches(5.7), Inches(4.1),
         "Governance & Lifecycle Flags (Part 2)",
         [
             ("--rm", "Automatically deletes container upon exit (great for one-off tasks)"),
             ("-it (Interactive + TTY)", "Keeps STDIN open and allocates a pseudo-TTY (for shells)"),
             ("--memory 512m", "Sets hard memory limit (enforced via Linux cgroups)"),
             ("--cpus 1.5", "Limits execution to a maximum of 1.5 CPU cores"),
             ("--restart unless-stopped", "Automatically restarts container if it crashes or host reboots")
         ],
         badge="LIMITS & GOVERNANCE", accent_color=COLOR_ACCENT_AMBER, font_size=11.5)


# ==============================================================================
# SLIDE 15: MODULE 6 — IMPORTANT CONTAINER COMMANDS (PART 2)
# ==============================================================================
s15 = prs.slides.add_slide(blank_layout)
set_slide_background(s15)
add_header(s15, "Module 6", "Essential Container Commands: Operations & Debugging",
           "Maya masters day-to-day operations, live log streaming, and container troubleshooting.")

cw_grid = Inches(3.75)
ch_grid = Inches(2.45)
cg_grid = Inches(0.22)

add_card(s15, Inches(0.8), Inches(1.8), cw_grid, ch_grid,
         "docker start & stop",
         [
             ("docker start", "Wakes an EXISTING stopped container (unlike run which creates a brand new one)"),
             ("docker stop", "Sends SIGTERM to PID 1 (graceful shutdown), waits 10s, then sends SIGKILL"),
             ("docker kill", "Force-stop: sends SIGKILL immediately")
         ],
         badge="LIFECYCLE", accent_color=COLOR_ACCENT_CYAN, font_size=10.5)

add_card(s15, Inches(0.8) + (cw_grid + cg_grid), Inches(1.8), cw_grid, ch_grid,
         "docker logs",
         [
             ("Purpose", "View container stdout and stderr output"),
             ("-f (Follow)", "Stream logs live in real time"),
             ("--tail 100", "Show only the last 100 log lines"),
             ("-t (Timestamps)", "Prepend timestamps to each log entry")
         ],
         badge="LOGGING", accent_color=COLOR_ACCENT_GREEN, font_size=10.5)

add_card(s15, Inches(0.8) + (cw_grid + cg_grid) * 2, Inches(1.8), cw_grid, ch_grid,
         "docker exec",
         [
             ("Syntax", "docker exec -it myapp bash"),
             ("Function", "Runs a new process inside an active running container"),
             ("Use Case", "Get an interactive shell inside container for live debugging without restarting")
         ],
         badge="DEBUGGING", accent_color=COLOR_ACCENT_AMBER, font_size=10.5)

add_card(s15, Inches(0.8), Inches(4.5), cw_grid, ch_grid,
         "docker ps",
         [
             ("docker ps", "List all currently running containers"),
             ("-a (All)", "Show all containers (running, stopped, exited)"),
             ("-q (Quiet)", "Display numeric container IDs only (great for shell script piping)")
         ],
         badge="INSPECTION", accent_color=COLOR_ACCENT_PURPLE, font_size=10.5)

add_card(s15, Inches(0.8) + (cw_grid + cg_grid), Inches(4.5), cw_grid, ch_grid,
         "docker inspect",
         [
             ("Syntax", "docker inspect myapp"),
             ("Output", "Returns massive detailed JSON configuration"),
             ("Data Extracted", "IP addresses, port bindings, mounted volumes, state, and environment variables")
         ],
         badge="METADATA", accent_color=COLOR_ACCENT_PINK, font_size=10.5)

add_card(s15, Inches(0.8) + (cw_grid + cg_grid) * 2, Inches(4.5), cw_grid, ch_grid,
         "docker stats",
         [
             ("Syntax", "docker stats"),
             ("Function", "Live streaming resource usage (like top)"),
             ("Metrics Shown", "CPU %, Memory usage & limit, Net I/O, Block I/O, and PIDs for all active containers")
         ],
         badge="METRICS", accent_color=COLOR_ACCENT_BLUE, font_size=10.5)


# ==============================================================================
# SLIDE 16: MODULE 7 — IMPORTANT IMAGE COMMANDS
# ==============================================================================
s16 = prs.slides.add_slide(blank_layout)
set_slide_background(s16)
add_header(s16, "Module 7", "Important Image Commands: Lifecycle Management",
           "Maya learns how to list, pull, push, tag, and safely clean up local images.")

add_card(s16, Inches(0.8), Inches(1.8), Inches(5.7), Inches(2.45),
         "docker images (or docker image ls)",
         [
             ("Purpose", "List all images stored locally on the host"),
             ("Options", "-a (show intermediate images), -q (numeric IDs only)"),
             ("Output Columns", "REPOSITORY, TAG, IMAGE ID, CREATED, SIZE"),
             ("Example", "docker images --filter reference='python*'")
         ],
         badge="LIST", accent_color=COLOR_ACCENT_CYAN, font_size=11)

add_card(s16, Inches(6.8), Inches(1.8), Inches(5.7), Inches(2.45),
         "docker rmi (or docker image rm)",
         [
             ("Purpose", "Remove one or more images from local host storage"),
             ("Example", "docker rmi nginx:latest"),
             ("Safety Guard", "Cannot remove an image if any existing container (even stopped!) is currently using it"),
             ("Cleanup Flag", "docker image prune (removes dangling untagged images)")
         ],
         badge="DELETE", accent_color=COLOR_ACCENT_RED, font_size=11)

add_card(s16, Inches(0.8), Inches(4.5), Inches(3.75), Inches(2.45),
         "docker pull",
         [
             ("Purpose", "Download image from registry"),
             ("Example", "docker pull nginx:1.21"),
             ("Layer Efficiency", "Downloads missing layers only; skips layers already in host cache")
         ],
         badge="DOWNLOAD", accent_color=COLOR_ACCENT_GREEN, font_size=10.5)

add_card(s16, Inches(4.78), Inches(4.5), Inches(3.75), Inches(2.45),
         "docker push",
         [
             ("Purpose", "Upload local image to remote registry"),
             ("Example", "docker push myuser/myapp:v1"),
             ("Prerequisite", "Requires authentication via docker login prior to push")
         ],
         badge="UPLOAD", accent_color=COLOR_ACCENT_AMBER, font_size=10.5)

add_card(s16, Inches(8.75), Inches(4.5), Inches(3.75), Inches(2.45),
         "docker tag",
         [
             ("Purpose", "Create an alias/tag for an image"),
             ("Example", "docker tag myapp:latest myapp:v1.0"),
             ("Zero Disk Waste", "Creates a lightweight pointer; does NOT copy the physical image layers")
         ],
         badge="TAG POINTER", accent_color=COLOR_ACCENT_PURPLE, font_size=10.5)


# ==============================================================================
# SLIDE 17: MODULE 8 — IMAGE OPTIMIZATION: MULTI-STAGE BUILDS
# ==============================================================================
s17 = prs.slides.add_slide(blank_layout)
set_slide_background(s17)
add_header(s17, "Module 8", "Optimization Strategy 1: Multi-Stage Builds",
           "Maya's first image weighed 2GB! The journey to shrinking it down to under 15MB.")

add_card(s17, Inches(0.8), Inches(1.8), Inches(5.7), Inches(5.2),
         "The Problem with Large Images",
         [
             ("Build & Deploy Lag", "Slow to build, slow to transfer across CI/CD runners"),
             ("Network Bandwidth Waste", "Massive gigabyte uploads clog pipelines and cloud bandwidth"),
             ("Slow Autoscaling", "Production servers take minutes to cold-start new containers"),
             ("Security Vulnerabilities", "Compilers, package managers, and debug tools expand the attack surface"),
             ("Expensive Storage", "Bloated images run up costly cloud container registry bills"),
             ("🍰 Bakery Analogy", "Optimization is like realizing you don't need to ship the mixing bowls, the oven, and the measuring cups with the cake. You just box up the final, baked cake!")
         ],
         badge="THE 2GB PROBLEM", accent_color=COLOR_ACCENT_RED, font_size=11)

add_code_box(s17, Inches(6.8), Inches(1.8), Inches(5.7), Inches(3.6),
             "Multi-Stage Build Pattern (Golang Case Study)",
             [
                 "# Stage 1: Build stage (contains compilers & dev tools)",
                 "FROM golang:1.19 AS builder",
                 "WORKDIR /app",
                 "COPY . .",
                 "RUN go build -o app app.go",
                 "",
                 "# Stage 2: Runtime stage (ultra-minimal production image)",
                 "FROM alpine:3.18          <- 5MB minimal base image",
                 "WORKDIR /app",
                 "COPY --from=builder /app/app /app",
                 'CMD ["/app"]'
             ],
             badge="MULTI-STAGE", accent_color=COLOR_ACCENT_GREEN, font_size=10)

add_metric_banner(s17, Inches(6.8), Inches(5.6), Inches(5.7), Inches(1.4),
                  "850 MB  →  15 MB",
                  "Multi-Stage Savings: 57x Smaller!",
                  "The heavy Go compiler and tools stay behind in builder stage; only the static binary runs.",
                  accent_color=COLOR_ACCENT_GREEN)


# ==============================================================================
# SLIDE 18: MODULE 8 — MINIMIZING LAYERS & BASE IMAGE SELECTION
# ==============================================================================
s18 = prs.slides.add_slide(blank_layout)
set_slide_background(s18)
add_header(s18, "Module 8", "Optimization Strategies 2 & 3: Layers & Base Images",
           "Chaining RUN commands and selecting the optimal base operating system.")

add_code_box(s18, Inches(0.8), Inches(1.8), Inches(5.7), Inches(5.2),
             "Strategy 2: Minimizing Layers by Chaining",
             [
                 "# Bad Practice (Creates 4 distinct heavy layers):",
                 "# RUN apt-get update",
                 "# RUN apt-get install -y curl git",
                 "# RUN apt-get clean",
                 "# RUN rm -rf /var/lib/apt/lists/*",
                 "",
                 "# Optimized Practice (1 Single Clean Layer):",
                 "RUN apt-get update && \\",
                 "    apt-get install -y --no-install-recommends curl git && \\",
                 "    apt-get clean && \\",
                 "    rm -rf /var/lib/apt/lists/*",
                 "",
                 "# Golden Rule:",
                 "# Deleting files in a later layer does NOT reduce image size!",
                 "# The earlier layer still contains the deleted bytes.",
                 "# Clean up temporary files in the SAME RUN instruction!"
             ],
             badge="LAYER CHAINING", accent_color=COLOR_ACCENT_AMBER, font_size=10)

add_card(s18, Inches(6.8), Inches(1.8), Inches(5.7), Inches(5.2),
         "Strategy 3: Lightweight Base Image Hierarchy",
         [
             ("Ubuntu (77 MB)", "Full-featured, standard tools. Great for development, troubleshooting, and legacy apps"),
             ("Debian (124 MB)", "Extremely stable, rich package ecosystem. Standard base for many official language images"),
             ("Alpine (5 MB)", "Minimalist security-oriented Linux (uses musl libc + busybox). Industry standard for lean microservices"),
             ("Distroless (2–20 MB)", "Google-curated image with runtime only—NO shell, NO package manager. Maximum production security!"),
             ("Scratch (0 MB)", "Completely empty base. Ideal for statically compiled Go, Rust, or C binaries"),
             ("Maya's Decision", "Switched from Ubuntu (77MB) to Alpine (5MB), saving 72MB immediately!")
         ],
         badge="BASE IMAGES", accent_color=COLOR_ACCENT_CYAN, font_size=11)


# ==============================================================================
# SLIDE 19: MODULE 8 — BUILD CONTEXT, .DOCKERIGNORE & RESULTS
# ==============================================================================
s19 = prs.slides.add_slide(blank_layout)
set_slide_background(s19)
add_header(s19, "Module 8", "Optimization Strategies 4-6 & Maya's Final Victory",
           "Slashing build context from 801MB to 2MB and Maya's total transformation.")

add_card(s19, Inches(0.8), Inches(1.8), Inches(5.7), Inches(3.2),
         "Strategies 4 & 5: Build Context & .dockerignore",
         [
             ("Build Context Problem", "Docker sends the ENTIRE directory to the daemon when you run docker build"),
             ("The Fix (.dockerignore)", "Block unnecessary and sensitive files from entering build context:"),
             ("• node_modules/ & venv/", "Avoid uploading huge local package caches"),
             ("• .git/ directory", "Prevent leaking git history and saving 500MB+"),
             ("• logs, test fixtures, .env", "Boost build speed and protect secrets"),
             ("Result", "Build context dropped from 801 MB → 2 MB!")
         ],
         badge="BUILD CONTEXT", accent_color=COLOR_ACCENT_CYAN, font_size=10.5)

add_card(s19, Inches(6.8), Inches(1.8), Inches(5.7), Inches(3.2),
         "Strategy 6: Additional Optimizations",
         [
             ("Order by Change Frequency", "Put static dependencies first, active source code last"),
             ("Purge Package Caches", "Always add 'pip cache purge' or 'rm -rf /var/cache/apk/*'"),
             ("Omit Recommended Packages", "Use '--no-install-recommends' on Debian/Ubuntu"),
             ("Pin Exact Runtimes", "Use 'python:3.9.16-alpine3.18' instead of loose 'python:alpine'"),
             ("Eliminate Shell Wrappers", "Use exec form CMD ['/app'] instead of shell form")
         ],
         badge="BEST PRACTICES", accent_color=COLOR_ACCENT_GREEN, font_size=10.5)

# Maya's Optimization Journey Card across bottom
add_card(s19, Inches(0.8), Inches(5.18), Inches(11.7), Inches(1.95),
         "Maya's Optimization Journey: 2GB → 15MB",
         [
             ("Original Image (The Bloated Giant)", "Size: 2.0 GB • Build Time: 5 minutes • Hundreds of OS packages & CVEs"),
             ("After Maya's Optimization", "Size: 15 MB (99.2% reduction!) • Build Time: 30 seconds • Zero unnecessary dependencies"),
             ("Production Impact", "Deployments that used to take 10 minutes now complete in seconds. Cloud registry costs cut by 90%!")
         ],
         badge="MAYA'S WIN", accent_color=COLOR_ACCENT_AMBER, font_size=11)


# ==============================================================================
# SLIDE 20: MODULE 9 — DOCKER REGISTRIES OVERVIEW
# ==============================================================================
s20 = prs.slides.add_slide(blank_layout)
set_slide_background(s20)
add_header(s20, "Module 9", "Docker Registries Overview",
           "Maya needs to store, manage, and share images with her team and production clusters.")

add_card(s20, Inches(0.8), Inches(1.8), Inches(5.7), Inches(3.2),
         "What is a Registry? (Hierarchy & Concepts)",
         [
             ("Definition", "A centralized service that stores and distributes Docker images"),
             ("Hierarchy Format", "registry_host / repository : tag"),
             ("Registry", "Server hosting images (e.g. docker.io, Amazon ECR, ghcr.io)"),
             ("Repository", "Collection of related versioned images (e.g. nginx, myapp)"),
             ("Tag", "Specific version identifier (e.g. 1.21, v1.0, latest)"),
             ("🍰 Bakery Analogy", "The central Recipe Library where bakers publish and browse recipes")
         ],
         badge="CONCEPTS", accent_color=COLOR_ACCENT_CYAN, font_size=10.5)

add_card(s20, Inches(6.8), Inches(1.8), Inches(5.7), Inches(3.2),
         "Popular Registries in the Cloud",
         [
             ("Docker Hub (docker.io)", "Public default; free tier, official verified base images"),
             ("Amazon ECR", "Private, AWS-native, IAM-governed, deep ECS/EKS integration"),
             ("Google Container Registry (GCR)", "Private, GCP-native, automated vulnerability scanning"),
             ("GitHub Container Registry (ghcr.io)", "CI/CD integrated directly with GitHub Actions"),
             ("Azure Container Registry (ACR)", "Private, Azure-native, built-in geo-replication"),
             ("Private / Self-Hosted", "Harbor (CNCF), JFrog Artifactory, Sonatype Nexus")
         ],
         badge="PROVIDERS", accent_color=COLOR_ACCENT_GREEN, font_size=10.5)

add_card(s20, Inches(0.8), Inches(5.18), Inches(11.7), Inches(1.95),
         "Registry Authentication Best Practices",
         [
             ("Login Command", "docker login <registry-domain> (authenticates Docker CLI against remote server)"),
             ("Production Security", "NEVER use raw personal passwords in CI/CD pipelines!"),
             ("Token-Based Auth", "Use short-lived Personal Access Tokens (PATs), IAM roles, or robot accounts with least privilege"),
             ("Secure Storage", "Docker stores auth tokens in ~/.docker/config.json or native OS credentials store (Keychain, pass)")
         ],
         badge="SECURITY", accent_color=COLOR_ACCENT_AMBER, font_size=11)


# ==============================================================================
# SLIDE 21: MODULE 10 — HANDLING IMAGE TAGS AND VERSIONING
# ==============================================================================
s21 = prs.slides.add_slide(blank_layout)
set_slide_background(s21)
add_header(s21, "Module 10", "Handling Image Tags and Versioning Strategies",
           "Maya's team kept overwriting :latest and breaking production. Time for disciplined versioning!")

add_card(s21, Inches(0.8), Inches(1.8), Inches(5.7), Inches(3.2),
         "The ':latest' Tag Problem: Maya's Realization",
         [
             ("The Big Myth", "Developers assume :latest means 'the most recent build'"),
             ("The Dangerous Reality", "It is NOT 'most recent'. It is merely whatever was last tagged with that name!"),
             ("Silent Production Breakages", "Overwriting :latest introduces untracked code, breaks rollback capability, and ruins Docker caching"),
             ("Maya's Golden Rule", "NEVER use :latest in production environments!")
         ],
         badge="THE LATEST TRAP", accent_color=COLOR_ACCENT_RED, font_size=10.5)

add_card(s21, Inches(6.8), Inches(1.8), Inches(5.7), Inches(3.2),
         "Semantic Versioning (SemVer) for Images",
         [
             ("SemVer Format", "MAJOR . MINOR . PATCH (e.g., myapp:2.4.7)"),
             ("MAJOR (2)", "Breaking changes, major architectural shifts, API rewrites"),
             ("MINOR (4)", "New features added in a backwards-compatible manner"),
             ("PATCH (7)", "Backwards-compatible bug fixes and security patches only"),
             ("Production Guarantee", "Always deploy immutable versioned tags (e.g. myapp:2.4.7)")
         ],
         badge="SEMVER", accent_color=COLOR_ACCENT_GREEN, font_size=10.5)

add_card(s21, Inches(0.8), Inches(5.18), Inches(11.7), Inches(1.95),
         "Git-Based & Environment Versioning Best Practices",
         [
             ("Git SHA Tagging", "Tag image with short Git commit SHA: myapp:a1b2c3d (100% traceable to exact source code commit)"),
             ("Environment Tagging", "myapp:v2.4.7-dev, myapp:v2.4.7-staging, myapp:v2.4.7-prod"),
             ("Multi-Tagging Best Practice", "Assign multiple tags to the same image ID (e.g., tag both 'v2.4.7' and 'a1b2c3d')"),
             ("Tag Immutability Rule", "Never move or overwrite production version tags once pushed! Date-tag backups for safety.")
         ],
         badge="VERSIONING RULES", accent_color=COLOR_ACCENT_PURPLE, font_size=11)


# ==============================================================================
# SLIDE 22: MODULE 10 — MAYA'S COMPLETE PRODUCTION WORKFLOW
# ==============================================================================
s22 = prs.slides.add_slide(blank_layout)
set_slide_background(s22)
add_header(s22, "Module 10", "Maya's Complete Production CI/CD Workflow",
           "Building, multi-tagging, pushing, deploying, and executing instant zero-downtime rollbacks.")

col_wf = Inches(3.75)
gap_wf = Inches(0.22)

add_code_box(s22, Inches(0.8), Inches(1.8), col_wf, Inches(5.2),
             "1. Build & Multi-Tag",
             [
                 "# Set version and Git commit SHA",
                 'VERSION="v2.4.7"',
                 'GIT_SHA=$(git rev-parse --short HEAD)',
                 "",
                 "# Build image with version tag",
                 "docker build -t myapp:$VERSION .",
                 "",
                 "# Add commit SHA tag pointer",
                 "docker tag myapp:$VERSION \\",
                 "  myapp:$GIT_SHA",
                 "",
                 "# Result: Dual-tagged image",
                 "# Traceable to exact Git commit!"
             ],
             badge="STEP 1", accent_color=COLOR_ACCENT_CYAN, font_size=10)

add_code_box(s22, Inches(0.8) + (col_wf + gap_wf), Inches(1.8), col_wf, Inches(5.2),
             "2. Push to Registry",
             [
                 "# Authenticate with registry token",
                 "echo $REG_TOKEN | docker login \\",
                 "  -u $USER --password-stdin myregistry.com",
                 "",
                 "# Push versioned image",
                 "docker push \\",
                 "  myregistry.com/myapp:$VERSION",
                 "",
                 "# Push Git SHA image pointer",
                 "docker push \\",
                 "  myregistry.com/myapp:$GIT_SHA",
                 "",
                 "# Note: Shared layers uploaded ONCE"
             ],
             badge="STEP 2", accent_color=COLOR_ACCENT_GREEN, font_size=10)

add_code_box(s22, Inches(0.8) + (col_wf + gap_wf) * 2, Inches(1.8), col_wf, Inches(5.2),
             "3. Deploy & Instant Rollback",
             [
                 "# Production Deployment:",
                 "docker run -d \\",
                 "  --name myapp-current \\",
                 "  -p 8080:80 \\",
                 "  myregistry.com/myapp:v2.4.7",
                 "",
                 "# If bug detected -> Instant Rollback!",
                 "# 1. Stop problematic container",
                 "docker stop myapp-current",
                 "",
                 "# 2. Run previous proven image",
                 "docker run -d \\",
                 "  --name myapp-current \\",
                 "  -p 8080:80 \\",
                 "  myregistry.com/myapp:v2.4.6"
             ],
             badge="STEP 3", accent_color=COLOR_ACCENT_AMBER, font_size=10)


# ==============================================================================
# SLIDE 23: KEY TAKEAWAYS & MASTER REFERENCE
# ==============================================================================
s23 = prs.slides.add_slide(blank_layout)
set_slide_background(s23)
add_header(s23, "Summary", "Key Takeaways & Master Reference",
           "The six core pillars of Docker mastery synthesized from Maya's journey.")

w_k = Inches(3.75)
h_k = Inches(2.45)
g_k = Inches(0.22)

add_card(s23, Inches(0.8), Inches(1.8), w_k, h_k,
         "1. Images: Immutable Templates",
         [
             ("Definition", "Read-only blueprints built from stacked layers"),
             ("UnionFS", "Common base layers shared across images on disk, cutting storage by 50%+"),
             ("Analogy", "The master Recipe Card that never changes")
         ],
         badge="IMAGES", accent_color=COLOR_ACCENT_CYAN, font_size=10.5)

add_card(s23, Inches(0.8) + (w_k + g_k), Inches(1.8), w_k, h_k,
         "2. Containers: Active Instances",
         [
             ("Definition", "Ephemeral running processes with a thin writable layer"),
             ("Copy-on-Write", "Instant startup and isolated memory/disk modifications"),
             ("Analogy", "The actual Baked Cake in the oven")
         ],
         badge="CONTAINERS", accent_color=COLOR_ACCENT_GREEN, font_size=10.5)

add_card(s23, Inches(0.8) + (w_k + g_k) * 2, Inches(1.8), w_k, h_k,
         "3. Essential Commands",
         [
             ("Containers", "docker run, start, stop, exec, logs, ps, inspect, stats"),
             ("Images", "docker images, pull, push, tag, rmi"),
             ("Governance", "Always set --memory and --restart policies")
         ],
         badge="COMMANDS", accent_color=COLOR_ACCENT_AMBER, font_size=10.5)

add_card(s23, Inches(0.8), Inches(4.5), w_k, h_k,
         "4. Optimization Playbook",
         [
             ("Multi-Stage", "Separate build tools from runtime (850MB → 15MB)"),
             ("Base OS", "Pick Alpine (5MB) or Distroless for production"),
             ("Hygiene", "Order lines by change frequency and use .dockerignore")
         ],
         badge="OPTIMIZATION", accent_color=COLOR_ACCENT_PURPLE, font_size=10.5)

add_card(s23, Inches(0.8) + (w_k + g_k), Inches(4.5), w_k, h_k,
         "5. Registries & Security",
         [
             ("Distribution", "Use cloud-native registries (ECR, GCR, ACR) or Docker Hub"),
             ("Auth", "Always authenticate via docker login with tokens/IAM"),
             ("Integrity", "Pin layer digests for tamper-proof reproducibility")
         ],
         badge="REGISTRIES", accent_color=COLOR_ACCENT_PINK, font_size=10.5)

add_card(s23, Inches(0.8) + (w_k + g_k) * 2, Inches(4.5), w_k, h_k,
         "6. Versioning Discipline",
         [
             ("The Latest Trap", "NEVER use :latest in production environments!"),
             ("SemVer", "Enforce MAJOR.MINOR.PATCH (e.g., myapp:2.4.7)"),
             ("Traceability", "Tag with short Git commit SHAs for instant rollback")
         ],
         badge="VERSIONING", accent_color=COLOR_ACCENT_BLUE, font_size=10.5)


# ==============================================================================
# SAVE PRESENTATION
# ==============================================================================
print(f"Saving presentation locally to: {OUTPUT_LOCAL}...")
prs.save(OUTPUT_LOCAL)
print("Saved local copy successfully!")

try:
    print(f"Copying to Downloads folder: {OUTPUT_DOWNLOADS}...")
    shutil.copyfile(OUTPUT_LOCAL, OUTPUT_DOWNLOADS)
    print("Copied to Downloads successfully!")
except Exception as e:
    print(f"Warning: Could not copy to Downloads: {e}")

print(f"\nPresentation successfully generated with {len(prs.slides)} slides!")
