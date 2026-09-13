#!/usr/bin/env python3
"""
CampGuard — Campsite Wildlife Detection Prototype
UT Austin Design Thinking Class

Uses YOLOv8 + webcam to detect wildlife and alert via desktop notification.
Designed for Logitech Brio 100 on MacBook Pro.

Usage:
    python3 campguard.py                  # defaults: camera 0, confidence 0.5
    python3 campguard.py --camera 1       # use camera index 1 (if Brio isn't at 0)
    python3 campguard.py --threshold 0.3  # lower confidence = more sensitive
    python3 campguard.py --gui            # launch with Tkinter control panel
"""

import argparse
import csv
import json
import os
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

import requests

import cv2
from ultralytics import YOLO

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------

# COCO class IDs that are animals
ANIMAL_CLASSES = {
    14: "bird", 15: "cat", 16: "dog", 17: "horse",
    18: "sheep", 19: "cow", 20: "elephant", 21: "bear",
    22: "zebra", 23: "giraffe",
}

# Also alert on "person" if you want human intruder mode (off by default)
PERSON_CLASS = {0: "person"}

DEFAULT_CONFIDENCE = 0.5
DEFAULT_COOLDOWN = 30  # seconds between repeat alerts
INFERENCE_INTERVAL = 0.5  # seconds between YOLO runs (~2 FPS)
SNAPSHOT_DIR = Path("campguard_snapshots")
LOG_FILE = Path("campguard_log.csv")

# Pushover credentials — loaded from config.json
CONFIG_FILE = Path(__file__).parent / "config.json"

def load_config():
    """Load Pushover credentials from config.json. Returns (user_key, app_token) or (None, None)."""
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE) as f:
                cfg = json.load(f)
            return cfg.get("pushover_user_key"), cfg.get("pushover_app_token")
        except (json.JSONDecodeError, KeyError) as e:
            print(f"WARNING: Could not read {CONFIG_FILE}: {e}")
    return None, None

PUSHOVER_USER_KEY, PUSHOVER_APP_TOKEN = load_config()

# ---------------------------------------------------------------------------
# NOTIFICATION
# ---------------------------------------------------------------------------

def notify_pushover(title: str, message: str, snapshot_path: str = None):
    """Send push notification to iPhone via Pushover, with optional snapshot."""
    data = {
        "token": PUSHOVER_APP_TOKEN,
        "user": PUSHOVER_USER_KEY,
        "title": title,
        "message": message,
        "priority": 0,  # normal priority; use 1 for high (vibration + bypass quiet hours)
    }
    files = {}
    if snapshot_path and os.path.exists(snapshot_path):
        files["attachment"] = ("snapshot.jpg", open(snapshot_path, "rb"), "image/jpeg")

    try:
        resp = requests.post("https://api.pushover.net/1/messages.json",
                             data=data, files=files, timeout=10)
        if resp.status_code == 200:
            print("  -> Pushover notification sent")
        else:
            print(f"  -> Pushover failed ({resp.status_code}): {resp.text}")
    except requests.RequestException as e:
        print(f"  -> Pushover error: {e}")


def notify_desktop(title: str, message: str, snapshot_path: str = None):
    """Send a desktop notification on macOS or Windows."""
    if sys.platform == "darwin":
        # macOS — use osascript
        script = f'display notification "{message}" with title "{title}"'
        try:
            subprocess.run(["osascript", "-e", script], check=True, capture_output=True)
        except (FileNotFoundError, subprocess.CalledProcessError):
            pass
    elif sys.platform == "win32":
        # Windows — use plyer (pip install plyer)
        try:
            from plyer import notification as plyer_notify
            plyer_notify.notify(
                title=title,
                message=message,
                app_name="CampGuard",
                timeout=10,
            )
        except ImportError:
            print("  -> Desktop notification skipped (install plyer: pip install plyer)")
        except Exception as e:
            print(f"  -> Desktop notification failed: {e}")
    else:
        # Linux or other — try plyer as fallback
        try:
            from plyer import notification as plyer_notify
            plyer_notify.notify(title=title, message=message, app_name="CampGuard", timeout=10)
        except Exception:
            pass


def notify(title: str, message: str, snapshot_path: str = None, use_pushover: bool = True):
    """Send alert via all enabled channels."""
    # Always send macOS desktop notification
    notify_desktop(title, message, snapshot_path)
    # Send to iPhone if Pushover is enabled and configured
    if use_pushover:
        if PUSHOVER_USER_KEY and PUSHOVER_APP_TOKEN:
            notify_pushover(title, message, snapshot_path)
        else:
            print("  -> Pushover skipped (no config.json found — see SETUP-GUIDE.md)")


# ---------------------------------------------------------------------------
# SNAPSHOT
# ---------------------------------------------------------------------------

def save_snapshot(frame, detections, snapshot_dir: Path) -> str:
    """Draw bounding boxes on frame and save as JPEG. Returns file path."""
    snapshot_dir.mkdir(exist_ok=True)
    annotated = frame.copy()

    for det in detections:
        x1, y1, x2, y2 = [int(c) for c in det["bbox"]]
        label = f"{det['class']} {det['confidence']:.0%}"
        # Green box + label
        cv2.rectangle(annotated, (x1, y1), (x2, y2), (0, 255, 0), 2)
        cv2.putText(annotated, label, (x1, y1 - 10),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 255, 0), 2)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = snapshot_dir / f"detection_{timestamp}.jpg"
    cv2.imwrite(str(filename), annotated)
    return str(filename)


# ---------------------------------------------------------------------------
# EVENT LOG
# ---------------------------------------------------------------------------

def init_log(log_file: Path):
    """Create CSV log with headers if it doesn't exist."""
    if not log_file.exists():
        with open(log_file, "w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["timestamp", "class", "confidence", "bbox", "snapshot"])


def log_event(log_file: Path, detection: dict, snapshot_path: str):
    """Append a detection event to the CSV log."""
    with open(log_file, "a", newline="") as f:
        writer = csv.writer(f)
        writer.writerow([
            datetime.now().isoformat(),
            detection["class"],
            f"{detection['confidence']:.3f}",
            detection["bbox"],
            snapshot_path,
        ])


# ---------------------------------------------------------------------------
# DETECTION ENGINE
# ---------------------------------------------------------------------------

class CampGuard:
    def __init__(self, camera_index=0, confidence_threshold=DEFAULT_CONFIDENCE,
                 cooldown=DEFAULT_COOLDOWN, include_person=False, use_pushover=True):
        self.camera_index = camera_index
        self.confidence_threshold = confidence_threshold
        self.cooldown = cooldown
        self.use_pushover = use_pushover
        self.last_alert_time = 0
        self.running = False

        # Build target class set
        self.target_classes = dict(ANIMAL_CLASSES)
        if include_person:
            self.target_classes.update(PERSON_CLASS)

        # Load model
        print("Loading YOLOv8n model...")
        self.model = YOLO("yolov8n.pt")
        print("Model loaded.")

    def process_frame(self, frame):
        """Run YOLO on a frame, return list of animal detections."""
        results = self.model(frame, verbose=False)
        detections = []
        for box in results[0].boxes:
            cls_id = int(box.cls[0])
            conf = float(box.conf[0])
            if cls_id in self.target_classes and conf >= self.confidence_threshold:
                detections.append({
                    "class": self.target_classes[cls_id],
                    "confidence": conf,
                    "bbox": box.xyxy[0].tolist(),
                })
        return detections

    def should_alert(self) -> bool:
        """Check if enough time has passed since last alert (cooldown)."""
        return (time.time() - self.last_alert_time) >= self.cooldown

    def run_headless(self):
        """Main loop — no GUI, just camera + detection + alerts."""
        cap = cv2.VideoCapture(self.camera_index)
        if not cap.isOpened():
            print(f"ERROR: Cannot open camera at index {self.camera_index}")
            print("Try --camera 1 or --camera 2 if the Brio isn't at index 0.")
            sys.exit(1)

        w = cap.get(cv2.CAP_PROP_FRAME_WIDTH)
        h = cap.get(cv2.CAP_PROP_FRAME_HEIGHT)
        print(f"Camera opened: {w:.0f}x{h:.0f}")
        print(f"Confidence threshold: {self.confidence_threshold}")
        print(f"Alert cooldown: {self.cooldown}s")
        print(f"Watching for: {', '.join(sorted(self.target_classes.values()))}")
        print("Press Ctrl+C to stop.\n")

        init_log(LOG_FILE)
        self.running = True

        try:
            while self.running:
                ret, frame = cap.read()
                if not ret:
                    print("WARNING: Failed to read frame, retrying...")
                    time.sleep(1)
                    continue

                detections = self.process_frame(frame)

                if detections and self.should_alert():
                    # Take the highest-confidence detection for the alert
                    best = max(detections, key=lambda d: d["confidence"])
                    snapshot = save_snapshot(frame, detections, SNAPSHOT_DIR)

                    # Log all detections
                    for det in detections:
                        log_event(LOG_FILE, det, snapshot)

                    # Alert
                    msg = f"{best['class']} detected ({best['confidence']:.0%} confidence)"
                    notify("⚠️ CampGuard Alert", msg, snapshot, self.use_pushover)
                    print(f"[{datetime.now().strftime('%H:%M:%S')}] ALERT: {msg}")
                    print(f"  Snapshot: {snapshot}")

                    self.last_alert_time = time.time()
                else:
                    # Heartbeat every 10 seconds so you know it's running
                    pass

                time.sleep(INFERENCE_INTERVAL)

        except KeyboardInterrupt:
            print("\nStopping CampGuard...")
        finally:
            cap.release()
            print("Camera released. Done.")


# ---------------------------------------------------------------------------
# GUI MODE (Tkinter)
# ---------------------------------------------------------------------------

def run_gui(camera_index, confidence_threshold, cooldown, include_person, use_pushover=True):
    """Launch CampGuard with a Tkinter control panel."""
    try:
        import tkinter as tk
        from tkinter import ttk
    except ImportError:
        print("ERROR: Tkinter not available. Run without --gui flag.")
        sys.exit(1)

    from PIL import Image, ImageTk  # pip install Pillow (included with ultralytics)

    guard = CampGuard(camera_index, confidence_threshold, cooldown, include_person, use_pushover)

    cap = cv2.VideoCapture(camera_index)
    if not cap.isOpened():
        print(f"ERROR: Cannot open camera at index {camera_index}")
        sys.exit(1)

    init_log(LOG_FILE)

    # --- Build window ---
    root = tk.Tk()
    root.title("CampGuard — Wildlife Detection")

    # Video frame
    video_label = tk.Label(root)
    video_label.pack(padx=10, pady=10)

    # Controls frame
    controls = ttk.Frame(root)
    controls.pack(fill="x", padx=10, pady=5)

    ttk.Label(controls, text="Sensitivity (confidence threshold):").pack(side="left")

    threshold_var = tk.DoubleVar(value=confidence_threshold)

    def on_threshold_change(val):
        guard.confidence_threshold = float(val)

    threshold_slider = ttk.Scale(controls, from_=0.1, to=0.9,
                                  variable=threshold_var, orient="horizontal",
                                  command=on_threshold_change)
    threshold_slider.pack(side="left", fill="x", expand=True, padx=10)

    threshold_label = ttk.Label(controls, text=f"{confidence_threshold:.1f}")
    threshold_label.pack(side="left")

    # Status bar
    status_var = tk.StringVar(value="Monitoring...")
    status_bar = ttk.Label(root, textvariable=status_var, relief="sunken", anchor="w")
    status_bar.pack(fill="x", padx=10, pady=(0, 10))

    # Detection log
    log_frame = ttk.LabelFrame(root, text="Detection Log")
    log_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    log_text = tk.Text(log_frame, height=6, state="disabled", font=("Menlo", 11))
    log_text.pack(fill="both", expand=True, padx=5, pady=5)

    def append_log(msg):
        log_text.config(state="normal")
        log_text.insert("end", msg + "\n")
        log_text.see("end")
        log_text.config(state="disabled")

    # --- Video loop ---
    def update_frame():
        ret, frame = cap.read()
        if not ret:
            root.after(500, update_frame)
            return

        detections = guard.process_frame(frame)

        # Draw boxes on display frame
        display = frame.copy()
        for det in detections:
            x1, y1, x2, y2 = [int(c) for c in det["bbox"]]
            label = f"{det['class']} {det['confidence']:.0%}"
            cv2.rectangle(display, (x1, y1), (x2, y2), (0, 255, 0), 2)
            cv2.putText(display, label, (x1, y1 - 10),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 255, 0), 2)

        # Alert logic
        if detections and guard.should_alert():
            best = max(detections, key=lambda d: d["confidence"])
            snapshot = save_snapshot(frame, detections, SNAPSHOT_DIR)
            for det in detections:
                log_event(LOG_FILE, det, snapshot)
            msg = f"{best['class']} detected ({best['confidence']:.0%})"
            notify("⚠️ CampGuard Alert", msg, snapshot, guard.use_pushover)
            append_log(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")
            status_var.set(f"Last alert: {msg}")
            guard.last_alert_time = time.time()

        # Update threshold label
        threshold_label.config(text=f"{guard.confidence_threshold:.2f}")

        # Convert frame for Tkinter display
        rgb = cv2.cvtColor(display, cv2.COLOR_BGR2RGB)
        img = Image.fromarray(rgb)
        # Scale to fit window (max 800px wide)
        max_w = 800
        if img.width > max_w:
            ratio = max_w / img.width
            img = img.resize((max_w, int(img.height * ratio)))
        imgtk = ImageTk.PhotoImage(image=img)
        video_label.imgtk = imgtk
        video_label.configure(image=imgtk)

        # Schedule next frame (~2 FPS for detection)
        root.after(int(INFERENCE_INTERVAL * 1000), update_frame)

    update_frame()

    def on_close():
        cap.release()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="CampGuard — Campsite Wildlife Detection")
    parser.add_argument("--camera", type=int, default=0,
                        help="Camera index (default: 0). Try 1 or 2 if Brio isn't at 0.")
    parser.add_argument("--threshold", type=float, default=DEFAULT_CONFIDENCE,
                        help=f"Confidence threshold 0.1-0.9 (default: {DEFAULT_CONFIDENCE})")
    parser.add_argument("--cooldown", type=int, default=DEFAULT_COOLDOWN,
                        help=f"Seconds between alerts (default: {DEFAULT_COOLDOWN})")
    parser.add_argument("--person", action="store_true",
                        help="Also detect people (human intruder mode)")
    parser.add_argument("--gui", action="store_true",
                        help="Launch with Tkinter GUI (live preview + sensitivity slider)")
    parser.add_argument("--no-pushover", action="store_true",
                        help="Disable Pushover phone notifications (desktop only)")
    parser.add_argument("--list-cameras", action="store_true",
                        help="Scan for available cameras and exit")
    args = parser.parse_args()

    use_pushover = not args.no_pushover

    # Camera scan mode
    if args.list_cameras:
        print("Scanning cameras...")
        for i in range(5):
            cap = cv2.VideoCapture(i)
            if cap.isOpened():
                w = cap.get(cv2.CAP_PROP_FRAME_WIDTH)
                h = cap.get(cv2.CAP_PROP_FRAME_HEIGHT)
                print(f"  Camera {i}: {w:.0f}x{h:.0f}")
                cap.release()
            else:
                print(f"  Camera {i}: not found")
        return

    if args.gui:
        run_gui(args.camera, args.threshold, args.cooldown, args.person, use_pushover)
    else:
        guard = CampGuard(args.camera, args.threshold, args.cooldown, args.person, use_pushover)
        guard.run_headless()


if __name__ == "__main__":
    main()
