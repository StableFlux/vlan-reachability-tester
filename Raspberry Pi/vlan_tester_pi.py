#!/usr/bin/env python3
"""
VLAN Reachability Tester
Automatically detects which VLAN you are on and continuously
pings all target devices every 5 seconds, building a reachability profile.
Press SPACE to pause/resume pinging. IP detection continues while paused.
Run with: sudo python3 vlan_tester.py
"""

import subprocess
import socket
import time
import os
import sys
import json
import threading
import tty
import termios
import select
from datetime import datetime
from collections import defaultdict

# ─── Network Configuration ────────────────────────────────────────────────────

VLANS = {
    "CORE":     {"subnet": "10.10.10.", "target": "10.10.10.10",  "label": "Pihole Server"},
    "TRUSTED":  {"subnet": "10.10.20.", "target": "10.10.20.43",  "label": "Beelink PC"},
    "CAMERA":   {"subnet": "10.10.30.", "target": "10.10.30.10",  "label": "Protect NVR"},
    "IOT":      {"subnet": "10.10.40.", "target": "10.10.40.3",   "label": "Home Assistant"},
    "VPN":      {"subnet": "10.10.50.", "target": "10.10.50.3",   "label": "Laptop"},
    "PENTEST":  {"subnet": "10.10.75.", "target": "10.10.75.100", "label": "Pentest PC"},
    "GUEST":    {"subnet": "10.10.200.","target": "10.10.200.159","label": "Rack Pi 1"},
    "DMZ":      {"subnet": "10.10.254.","target": "10.10.254.2",  "label": "TDI Webserver"},
}

PING_INTERVAL   = 5    # seconds between full sweep
PING_TIMEOUT    = 2    # seconds per ping
PING_COUNT      = 1    # pings per target per sweep
RESULTS_FILE    = "vlan_results.json"

# ─── Helpers ──────────────────────────────────────────────────────────────────

def clear():
    os.system("clear")

def get_local_ips():
    """Return all non-loopback IPv4 addresses on this host."""
    ips = []
    try:
        import netifaces
        for iface in netifaces.interfaces():
            addrs = netifaces.ifaddresses(iface).get(netifaces.AF_INET, [])
            for addr in addrs:
                ip = addr.get("addr", "")
                if ip and not ip.startswith("127."):
                    ips.append(ip)
    except ImportError:
        try:
            out = subprocess.check_output(["hostname", "-I"], text=True).strip()
            ips = [ip for ip in out.split() if not ip.startswith("127.")]
        except Exception:
            pass
    return ips

def detect_vlan(ips):
    """Match local IPs against known VLAN subnets."""
    for ip in ips:
        for vlan_name, vlan in VLANS.items():
            if ip.startswith(vlan["subnet"]):
                return vlan_name, ip
    return None, ips[0] if ips else "unknown"

def ping(host, count=1, timeout=2):
    """Returns (reachable: bool, rtt_ms: float or None)."""
    try:
        result = subprocess.run(
            ["ping", "-c", str(count), "-W", str(timeout), host],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        if result.returncode == 0:
            for line in result.stdout.splitlines():
                if "rtt" in line or "round-trip" in line:
                    try:
                        rtt = float(line.split("=")[1].strip().split("/")[1])
                        return True, rtt
                    except Exception:
                        pass
            return True, None
        return False, None
    except Exception:
        return False, None

def save_results(results):
    try:
        with open(RESULTS_FILE, "w") as f:
            json.dump(results, f, indent=2)
    except Exception:
        pass

def load_results():
    try:
        if os.path.exists(RESULTS_FILE):
            with open(RESULTS_FILE) as f:
                return json.load(f)
    except Exception:
        pass
    return {}

# ─── Display ──────────────────────────────────────────────────────────────────

# Colour codes
C_GREEN  = "\033[92m"
C_RED    = "\033[91m"
C_YELLOW = "\033[93m"
C_CYAN   = "\033[96m"
C_RESET  = "\033[0m"
C_BOLD   = "\033[1m"

VLAN_NAMES = list(VLANS.keys())

def render_matrix(results):
    """Print the full 8x8 reachability matrix — colour-only cells, no text inside."""
    BG_GREEN  = "\033[42m"
    BG_RED    = "\033[41m"
    BG_DARK   = "\033[100m"
    BG_RESET  = "\033[0m"
    GREY      = "\033[90m"

    cell_w = 9
    lbl_w  = 9
    indent = lbl_w + 3   # 2 spaces + lbl_w + 1 space before first border

    def make_cell(frm, to):
        key   = f"{frm}->{to}"
        entry = results.get(key)
        if entry is None:
            bg = BG_DARK
        elif entry.get("last"):
            bg = BG_GREEN
        else:
            bg = BG_RED
        return f"{GREY}│{BG_RESET}{bg}{' ' * cell_w}{BG_RESET}"

    # Column headers: centred over (cell_w+1) to account for the leading │ per cell
    print(" " * indent + "DESTINATION  →")
    hdrs = "".join(f"{n:^{cell_w + 1}}" for n in VLAN_NAMES)
    print(f"{GREY}{' ' * indent}{hdrs}{BG_RESET}")

    # Top border
    print(f"{' ' * indent}{GREY}┌" + ("─" * cell_w + "┬") * (len(VLAN_NAMES) - 1) + "─" * cell_w + f"┐{BG_RESET}")

    for i, frm in enumerate(VLAN_NAMES):
        row_label = f"{C_CYAN}{frm:<{lbl_w}}{C_RESET}"
        row       = "".join(make_cell(frm, to) for to in VLAN_NAMES) + f"{GREY}│{BG_RESET}"
        print(f"  {row_label} {row}")

        if i < len(VLAN_NAMES) - 1:
            print(f"{' ' * indent}{GREY}├" + ("─" * cell_w + "┼") * (len(VLAN_NAMES) - 1) + "─" * cell_w + f"┤{BG_RESET}")

    # Bottom border
    print(f"{' ' * indent}{GREY}└" + ("─" * cell_w + "┴") * (len(VLAN_NAMES) - 1) + "─" * cell_w + f"┘{BG_RESET}")

    print()
    print(f"  Key:  {BG_GREEN}{' ' * 5}{BG_RESET} Reachable   "
          f"{BG_RED}{' ' * 5}{BG_RESET} Blocked   "
          f"{BG_DARK}{' ' * 5}{BG_RESET} Not yet tested")

def render_current_sweep(current_vlan, my_ip, results):
    """Print the live ping results for the current VLAN sweep."""
    print(f"  {'DESTINATION':<10} {'TARGET':<18} {'DEVICE':<22} {'STATUS':<9} {'RTT':>7}")
    print("  " + "─" * 68)
    for vlan_name, vlan in VLANS.items():
        key     = f"{current_vlan or 'UNKNOWN'}->{vlan_name}"
        entry   = results.get(key, {})
        state   = entry.get("last")
        rtt     = entry.get("rtt")
        rtt_str = f"{rtt:.1f}ms" if rtt is not None else "  —  "
        marker  = f" {C_CYAN}← you{C_RESET}" if vlan_name == current_vlan else ""

        if state is True:
            sym = f"{C_GREEN}REACH{C_RESET}"
        elif state is False:
            sym = f"{C_RED}BLOCK{C_RESET}"
        else:
            sym = f"{C_YELLOW} ??? {C_RESET}"

        print(f"  {vlan_name:<10} {vlan['target']:<18} {vlan['label']:<22} {sym}   {rtt_str:>7}{marker}")

def render(current_vlan, my_ip, sweep_count, results, countdown, paused=False):
    clear()
    width = 76
    ts    = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    print(f"{C_BOLD}{'═' * width}{C_RESET}")
    print(f"{C_BOLD}  VLAN REACHABILITY TESTER  │  {ts}{C_RESET}")
    print(f"{C_BOLD}{'═' * width}{C_RESET}")

    vlan_display = f"{C_CYAN}{current_vlan}{C_RESET}" if current_vlan else f"{C_YELLOW}UNKNOWN (not on a known subnet){C_RESET}"
    print(f"  Current VLAN : {vlan_display}")
    print(f"  Local IP     : {my_ip}")

    if paused:
        print(f"  Sweep #      : {sweep_count}  │  {C_YELLOW}PAUSED{C_RESET} — press SPACE to resume  │  IP tracking active")
    else:
        print(f"  Sweep #      : {sweep_count}  │  {C_GREEN}ACTIVE{C_RESET} — next sweep in {countdown:2d}s  │  SPACE to pause  │  Ctrl+C to stop")

    print(f"{'─' * width}")

    # ── Live sweep results ──
    print(f"\n  {C_BOLD}CURRENT VLAN SWEEP  (source: {current_vlan or 'UNKNOWN'}){C_RESET}\n")
    render_current_sweep(current_vlan, my_ip, results)

    # ── Matrix ──
    print(f"\n  {C_BOLD}FULL REACHABILITY MATRIX  (rows = source, columns = destination){C_RESET}\n")
    render_matrix(results)

    tested_vlans = set(k.split("->")[0] for k in results)
    total_cells  = len(VLAN_NAMES) * len(VLAN_NAMES)
    tested_cells = len(results)
    reach_cells  = sum(1 for v in results.values() if v.get("last") is True)
    block_cells  = sum(1 for v in results.values() if v.get("last") is False)

    print(f"\n{'─' * width}")
    print(f"  VLANs tested so far : {C_CYAN}{', '.join(sorted(tested_vlans)) or 'none yet'}{C_RESET}")
    print(f"  Matrix coverage     : {tested_cells}/{total_cells} pairs  │  "
          f"{C_GREEN}{reach_cells} reachable{C_RESET}  │  {C_RED}{block_cells} blocked{C_RESET}")
    print(f"  Results saved to    : {RESULTS_FILE}")
    print(f"{'═' * width}")

def export_matrix_txt(results, filename="vlan_matrix.txt"):
    """Write a plain-text (no colour codes) matrix to a file."""
    lines  = []
    lbl_w  = 9

    lines.append("VLAN REACHABILITY MATRIX")
    lines.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("")
    lines.append(" " * lbl_w + "  " + "".join(f"{n[:7]:^8}" for n in VLAN_NAMES))
    lines.append(" " * lbl_w + "  " + "─" * (8 * len(VLAN_NAMES)))

    for frm in VLAN_NAMES:
        row = f"{frm:<{lbl_w}}  "
        for to in VLAN_NAMES:
            key   = f"{frm}->{to}"
            entry = results.get(key)
            if entry is None:
                sym = "   ?   "
            else:
                sym = "   R   " if entry.get("last") else "   X   "
            row += f" {sym}"
        lines.append(row)

    lines.append("")
    lines.append("R = Reachable   X = Blocked   ? = Not yet tested")

    try:
        with open(filename, "w") as f:
            f.write("\n".join(lines))
    except Exception:
        pass

# ─── Keyboard listener (background thread) ────────────────────────────────────

def kb_monitor(paused_flag):
    """Toggle paused_flag on SPACE. Send SIGINT on Ctrl-C or q."""
    fd  = sys.stdin.fileno()
    old = termios.tcgetattr(fd)
    try:
        tty.setcbreak(fd)
        while True:
            r, _, _ = select.select([sys.stdin], [], [], 0.1)
            if r:
                ch = sys.stdin.read(1)
                if ch == ' ':
                    paused_flag[0] = not paused_flag[0]
                elif ch in ('\x03', 'q', 'Q'):
                    os.kill(os.getpid(), 2)
    except Exception:
        pass
    finally:
        termios.tcsetattr(fd, termios.TCSADRAIN, old)

# ─── Main Loop ────────────────────────────────────────────────────────────────

def main():
    if os.geteuid() != 0:
        print("\n[!] This script requires root to send ICMP pings.")
        print("    Run with: sudo python3 vlan_tester.py\n")
        sys.exit(1)

    results     = load_results()
    sweep_count = 0
    paused      = [False]   # list so the kb thread can mutate it

    # Start keyboard listener in background daemon thread
    threading.Thread(target=kb_monitor, args=(paused,), daemon=True).start()

    print("Starting VLAN Reachability Tester... press SPACE to pause, Ctrl+C to stop.")
    time.sleep(1)

    try:
        while True:
            # Always re-detect IP — even while paused
            local_ips           = get_local_ips()
            current_vlan, my_ip = detect_vlan(local_ips)

            if paused[0]:
                render(current_vlan, my_ip, sweep_count, results, 0, paused=True)
                time.sleep(1)
                continue

            # ── Run a full ping sweep ──
            sweep_count += 1
            for vlan_name, vlan in VLANS.items():
                target       = vlan["target"]
                key          = f"{current_vlan or 'UNKNOWN'}->{vlan_name}"
                reached, rtt = ping(target, count=PING_COUNT, timeout=PING_TIMEOUT)

                results[key] = {
                    "last":    reached,
                    "rtt":     rtt,
                    "time":    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "from_ip": my_ip,
                }

                # Render after each ping so results appear live
                render(current_vlan, my_ip, sweep_count, results, 0, paused=False)

            save_results(results)
            export_matrix_txt(results)

            # ── Countdown to next sweep ──
            for remaining in range(PING_INTERVAL, 0, -1):
                # Re-detect IP every second during countdown
                local_ips           = get_local_ips()
                current_vlan, my_ip = detect_vlan(local_ips)

                # If paused mid-countdown, hold until resumed
                while paused[0]:
                    local_ips           = get_local_ips()
                    current_vlan, my_ip = detect_vlan(local_ips)
                    render(current_vlan, my_ip, sweep_count, results, 0, paused=True)
                    time.sleep(1)

                render(current_vlan, my_ip, sweep_count, results, remaining, paused=False)
                time.sleep(1)

    except KeyboardInterrupt:
        save_results(results)
        export_matrix_txt(results)
        print(f"\n\n  Stopped. Files saved:")
        print(f"    {RESULTS_FILE}   (full JSON results)")
        print(f"    vlan_matrix.txt  (plain text matrix)")
        sys.exit(0)

if __name__ == "__main__":
    main()
