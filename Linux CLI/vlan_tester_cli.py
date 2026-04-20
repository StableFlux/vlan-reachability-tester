#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
VLAN Reachability Tester - Linux CLI Edition

Continuously pings user-configured target IPs across multiple VLANs and
shows a live colour-coded reachability matrix in the terminal.

Usage:
  sudo python3 vlan_tester_cli.py          Run with saved config
  sudo python3 vlan_tester_cli.py --setup  (Re-)run the setup wizard

Keys during sweep:
  SPACE   pause / resume
  c       open config menu
  Ctrl+C  quit

Configuration is stored in vlan_config.json next to this script.
"""

import json
import os
import re
import select
import socket
import subprocess
import sys
import termios
import threading
import time
import tty
from datetime import datetime

# ── Paths ─────────────────────────────────────────────────────────────────────

BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE  = os.path.join(BASE_DIR, "vlan_config.json")
RESULTS_FILE = os.path.join(BASE_DIR, "vlan_results.json")

DEFAULT_CONFIG = {
    "vlans":         [],
    "ping_interval": 5,
    "ping_timeout":  2,
    "ping_count":    1,
    "selected_nic":  None,   # ignored on Linux; kept so configs interop with Windows
}

# ── ANSI colours ──────────────────────────────────────────────────────────────

C_GREEN  = "\033[92m"
C_RED    = "\033[91m"
C_YELLOW = "\033[93m"
C_CYAN   = "\033[96m"
C_MUTED  = "\033[90m"
C_RESET  = "\033[0m"
C_BOLD   = "\033[1m"

BG_GREEN = "\033[42m"
BG_RED   = "\033[41m"
BG_DARK  = "\033[100m"
BG_RESET = "\033[0m"

# ── Capability detection ──────────────────────────────────────────────────────

def is_tty():
    try:
        return sys.stdin.isatty()
    except Exception:
        return False


def check_ping_available():
    """Check that `ping` exists and supports our flags. Returns (ok, error_msg)."""
    try:
        result = subprocess.run(
            ["ping", "-c", "1", "-W", "1", "127.0.0.1"],
            capture_output=True, timeout=5,
        )
    except FileNotFoundError:
        return False, "`ping` command not found. Install it with:\n  sudo apt install iputils-ping"
    except Exception as e:
        return False, f"Could not run `ping`: {e}"

    if result.returncode == 0:
        return True, None

    err = (result.stderr or b"").decode(errors="replace").lower()
    if "unknown option" in err or "invalid option" in err or "usage:" in err:
        return False, (
            "Your system's `ping` does not support -c/-W flags "
            "(BusyBox/Alpine?). This script requires iputils-ping."
        )
    if "permission" in err or "operation not permitted" in err:
        return False, (
            "`ping` does not have permission to send ICMP on this system.\n"
            "Try one of:\n"
            "  sudo python3 vlan_tester_cli.py\n"
            "  sudo setcap cap_net_raw+ep $(which ping)"
        )
    # Non-zero exit against 127.0.0.1 is suspicious but not fatal
    return True, None


# ── Local IP detection (cascading fallbacks) ──────────────────────────────────

def _filter_ips(ips):
    return [ip for ip in ips
            if ip and not ip.startswith("127.") and not ip.startswith("169.254.")]


def get_local_ips():
    """Return non-loopback IPv4 addresses using whichever method is available."""
    # 1. netifaces (most accurate, optional package)
    try:
        import netifaces  # type: ignore
        found = []
        for iface in netifaces.interfaces():
            for addr in netifaces.ifaddresses(iface).get(netifaces.AF_INET, []):
                found.append(addr.get("addr", ""))
        ips = _filter_ips(found)
        if ips:
            return ips
    except Exception:
        pass

    # 2. hostname -I (iputils, default on Debian/Ubuntu/RPi OS/Fedora)
    try:
        out = subprocess.check_output(["hostname", "-I"], text=True, timeout=3).strip()
        ips = _filter_ips(out.split())
        if ips:
            return ips
    except Exception:
        pass

    # 3. ip -4 addr show (iproute2, default on all modern distros incl. Arch)
    try:
        out = subprocess.check_output(["ip", "-4", "addr", "show"], text=True, timeout=3)
        ips = _filter_ips(re.findall(r"inet (\d+\.\d+\.\d+\.\d+)", out))
        if ips:
            return ips
    except Exception:
        pass

    # 4. ifconfig (net-tools, legacy)
    try:
        out = subprocess.check_output(["ifconfig"], text=True, timeout=3)
        ips = _filter_ips(re.findall(r"inet (?:addr:)?(\d+\.\d+\.\d+\.\d+)", out))
        if ips:
            return ips
    except Exception:
        pass

    # 5. Python socket resolution (last resort)
    try:
        return _filter_ips(socket.gethostbyname_ex(socket.gethostname())[2])
    except Exception:
        return []


def detect_vlan(ips, vlans):
    for ip in ips:
        for v in vlans:
            if ip.startswith(v["subnet"]):
                return v["name"], ip
    return None, (ips[0] if ips else "unknown")


def ping(host, count=1, timeout=2):
    """Returns (reachable: bool, rtt_ms: float or None)."""
    try:
        result = subprocess.run(
            ["ping", "-c", str(count), "-W", str(timeout), host],
            capture_output=True, text=True,
            timeout=count * (timeout + 1) + 2,
        )
    except Exception:
        return False, None

    if result.returncode != 0:
        return False, None
    for line in result.stdout.splitlines():
        if "rtt" in line or "round-trip" in line:
            try:
                return True, float(line.split("=")[1].strip().split("/")[1])
            except Exception:
                pass
    return True, None


# ── Config I/O ────────────────────────────────────────────────────────────────

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE) as f:
                cfg = json.load(f)
            for k, v in DEFAULT_CONFIG.items():
                cfg.setdefault(k, v)
            return cfg
        except Exception as e:
            print(f"{C_YELLOW}Warning: could not load config: {e}{C_RESET}")
    return dict(DEFAULT_CONFIG)


def save_config(cfg):
    try:
        with open(CONFIG_FILE, "w") as f:
            json.dump(cfg, f, indent=2)
    except Exception as e:
        print(f"{C_RED}Could not save config: {e}{C_RESET}")


def load_results():
    if os.path.exists(RESULTS_FILE):
        try:
            with open(RESULTS_FILE) as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_results(results):
    try:
        with open(RESULTS_FILE, "w") as f:
            json.dump(results, f, indent=2)
    except Exception:
        pass


# ── Interactive prompts ───────────────────────────────────────────────────────

def _input(prompt_str):
    """input() wrapped so a plain Ctrl+C in prompts returns cleanly."""
    try:
        return input(prompt_str)
    except EOFError:
        raise KeyboardInterrupt


def prompt_text(label, default=None, required=False):
    hint = f" [{default}]" if default else ""
    while True:
        val = _input(f"  {label}{hint}: ").strip()
        if not val and default is not None:
            return default
        if val or not required:
            return val
        print(f"  {C_RED}This field is required.{C_RESET}")


def prompt_yes_no(label, default=True):
    suffix = "Y/n" if default else "y/N"
    while True:
        val = _input(f"  {label} ({suffix}): ").strip().lower()
        if not val:
            return default
        if val in ("y", "yes"):
            return True
        if val in ("n", "no"):
            return False


def prompt_vlan(existing=None):
    existing = existing or {}
    print()
    name   = prompt_text("VLAN name (e.g. CORE)", existing.get("name"), required=True).upper()
    subnet = prompt_text("Subnet prefix (e.g. 10.10.10.)", existing.get("subnet"), required=True)
    if not subnet.endswith("."):
        subnet += "."
    target = prompt_text("Target IP to ping (e.g. 10.10.10.1)", existing.get("target"), required=True)
    label  = prompt_text("Device label (optional)", existing.get("label", ""), required=False)
    return {"name": name, "subnet": subnet, "target": target, "label": label}


# ── First-run wizard ──────────────────────────────────────────────────────────

def run_setup_wizard(cfg):
    print()
    print(f"{C_BOLD}{C_CYAN}{'═' * 60}{C_RESET}")
    print(f"{C_BOLD}  VLAN REACHABILITY TESTER  -  SETUP WIZARD{C_RESET}")
    print(f"{C_BOLD}{C_CYAN}{'═' * 60}{C_RESET}")
    print()
    print("  Add each VLAN you want to test. You can add more later via")
    print("  the 'c' config menu during a sweep.")
    print()

    cfg = dict(cfg)
    cfg["vlans"] = []

    while True:
        print(f"{C_CYAN}Adding VLAN #{len(cfg['vlans']) + 1}:{C_RESET}")
        try:
            cfg["vlans"].append(prompt_vlan())
        except KeyboardInterrupt:
            print()
            if not cfg["vlans"]:
                print(f"{C_RED}Setup cancelled - at least one VLAN is required.{C_RESET}")
                sys.exit(1)
            break
        if not prompt_yes_no("Add another VLAN?", default=True):
            break

    print()
    print(f"{C_CYAN}Ping settings (press Enter to keep defaults):{C_RESET}")
    try:
        cfg["ping_interval"] = max(1, int(prompt_text("Sweep interval (seconds)", str(cfg["ping_interval"]))))
        cfg["ping_timeout"]  = max(1, int(prompt_text("Ping timeout (seconds)",   str(cfg["ping_timeout"]))))
        cfg["ping_count"]    = max(1, int(prompt_text("Pings per target",         str(cfg["ping_count"]))))
    except ValueError:
        print(f"{C_YELLOW}Invalid number - using defaults.{C_RESET}")

    save_config(cfg)
    print()
    print(f"{C_GREEN}Configuration saved to {CONFIG_FILE}{C_RESET}")
    print()
    return cfg


# ── In-app config menu ────────────────────────────────────────────────────────

def config_menu(cfg):
    while True:
        print()
        print(f"{C_BOLD}{C_CYAN}{'─' * 60}{C_RESET}")
        print(f"{C_BOLD}  CONFIG MENU{C_RESET}")
        print(f"{C_BOLD}{C_CYAN}{'─' * 60}{C_RESET}")
        print()
        print("  1) List VLANs")
        print("  2) Add VLAN")
        print("  3) Remove VLAN")
        print("  4) Edit VLAN")
        print("  5) Ping settings")
        print("  6) Resume sweep")
        print()
        try:
            choice = _input("  Choose: ").strip()
        except KeyboardInterrupt:
            print()
            return cfg

        if choice == "1":
            print()
            if not cfg["vlans"]:
                print("  (no VLANs configured)")
            else:
                for i, v in enumerate(cfg["vlans"], 1):
                    print(f"  {i}. {C_CYAN}{v['name']:<12}{C_RESET} "
                          f"subnet={v['subnet']:<18} target={v['target']:<16} "
                          f"{C_MUTED}{v.get('label', '')}{C_RESET}")

        elif choice == "2":
            try:
                v = prompt_vlan()
                if any(x["name"] == v["name"] for x in cfg["vlans"]):
                    print(f"{C_RED}A VLAN named '{v['name']}' already exists.{C_RESET}")
                else:
                    cfg["vlans"].append(v)
                    save_config(cfg)
                    print(f"{C_GREEN}Added {v['name']}.{C_RESET}")
            except KeyboardInterrupt:
                print()
                print(f"{C_YELLOW}Cancelled.{C_RESET}")

        elif choice == "3":
            if not cfg["vlans"]:
                print(f"{C_YELLOW}No VLANs to remove.{C_RESET}")
                continue
            print()
            for i, v in enumerate(cfg["vlans"], 1):
                print(f"  {i}. {v['name']}")
            try:
                idx = int(_input("  Remove which number? ").strip()) - 1
                if 0 <= idx < len(cfg["vlans"]):
                    name = cfg["vlans"][idx]["name"]
                    if prompt_yes_no(f"Remove '{name}'?", default=False):
                        del cfg["vlans"][idx]
                        save_config(cfg)
                        print(f"{C_GREEN}Removed {name}.{C_RESET}")
            except (ValueError, KeyboardInterrupt):
                pass

        elif choice == "4":
            if not cfg["vlans"]:
                print(f"{C_YELLOW}No VLANs to edit.{C_RESET}")
                continue
            print()
            for i, v in enumerate(cfg["vlans"], 1):
                print(f"  {i}. {v['name']}")
            try:
                idx = int(_input("  Edit which number? ").strip()) - 1
                if 0 <= idx < len(cfg["vlans"]):
                    v = prompt_vlan(existing=cfg["vlans"][idx])
                    cfg["vlans"][idx] = v
                    save_config(cfg)
                    print(f"{C_GREEN}Updated {v['name']}.{C_RESET}")
            except (ValueError, KeyboardInterrupt):
                pass

        elif choice == "5":
            try:
                cfg["ping_interval"] = max(1, int(prompt_text("Sweep interval (seconds)", str(cfg["ping_interval"]))))
                cfg["ping_timeout"]  = max(1, int(prompt_text("Ping timeout (seconds)",   str(cfg["ping_timeout"]))))
                cfg["ping_count"]    = max(1, int(prompt_text("Pings per target",         str(cfg["ping_count"]))))
                save_config(cfg)
                print(f"{C_GREEN}Ping settings saved.{C_RESET}")
            except (ValueError, KeyboardInterrupt):
                print(f"{C_YELLOW}No changes.{C_RESET}")

        elif choice == "6" or choice == "":
            return cfg


# ── Display ───────────────────────────────────────────────────────────────────

def clear():
    os.system("clear")


def render_current_sweep(vlans, current_vlan, results):
    print(f"  {'DESTINATION':<12} {'TARGET':<18} {'DEVICE':<22} {'STATUS':<9} {'RTT':>7}")
    print("  " + "─" * 70)
    for v in vlans:
        entry   = results.get(f"{current_vlan or 'UNKNOWN'}->{v['name']}", {})
        state   = entry.get("last")
        rtt     = entry.get("rtt")
        rtt_str = f"{rtt:.1f}ms" if rtt is not None else "  —  "
        marker  = f" {C_CYAN}← you{C_RESET}" if v["name"] == current_vlan else ""

        if state is True:
            sym = f"{C_GREEN}REACH{C_RESET}"
        elif state is False:
            sym = f"{C_RED}BLOCK{C_RESET}"
        else:
            sym = f"{C_YELLOW} ??? {C_RESET}"

        print(f"  {v['name']:<12} {v['target']:<18} {v.get('label',''):<22} "
              f"{sym}   {rtt_str:>7}{marker}")


def render_matrix(vlan_names, results):
    cell_w = 9
    lbl_w  = max((len(n) for n in vlan_names), default=4) + 1
    indent = lbl_w + 3
    GREY   = "\033[90m"

    def cell(frm, to):
        entry = results.get(f"{frm}->{to}")
        bg = BG_DARK if entry is None else (BG_GREEN if entry.get("last") else BG_RED)
        return f"{GREY}│{BG_RESET}{bg}{' ' * cell_w}{BG_RESET}"

    print(" " * indent + "DESTINATION  →")
    hdrs = "".join(f"{n[:7]:^{cell_w + 1}}" for n in vlan_names)
    print(f"{GREY}{' ' * indent}{hdrs}{BG_RESET}")
    print(f"{' ' * indent}{GREY}┌"
          + ("─" * cell_w + "┬") * (len(vlan_names) - 1)
          + "─" * cell_w + f"┐{BG_RESET}")

    for i, frm in enumerate(vlan_names):
        row_label = f"{C_CYAN}{frm:<{lbl_w}}{C_RESET}"
        row = "".join(cell(frm, to) for to in vlan_names) + f"{GREY}│{BG_RESET}"
        print(f"  {row_label} {row}")
        if i < len(vlan_names) - 1:
            print(f"{' ' * indent}{GREY}├"
                  + ("─" * cell_w + "┼") * (len(vlan_names) - 1)
                  + "─" * cell_w + f"┤{BG_RESET}")

    print(f"{' ' * indent}{GREY}└"
          + ("─" * cell_w + "┴") * (len(vlan_names) - 1)
          + "─" * cell_w + f"┘{BG_RESET}")
    print()
    print(f"  Key:  {BG_GREEN}{' ' * 5}{BG_RESET} Reachable   "
          f"{BG_RED}{' ' * 5}{BG_RESET} Blocked   "
          f"{BG_DARK}{' ' * 5}{BG_RESET} Not yet tested")


def render(cfg, current_vlan, my_ip, sweep_count, results, countdown,
           paused=False, tty_mode=True):
    if tty_mode:
        clear()

    width = 76
    ts    = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    print(f"{C_BOLD}{'═' * width}{C_RESET}")
    print(f"{C_BOLD}  VLAN REACHABILITY TESTER  │  {ts}{C_RESET}")
    print(f"{C_BOLD}{'═' * width}{C_RESET}")

    vlan_disp = (f"{C_CYAN}{current_vlan}{C_RESET}" if current_vlan
                 else f"{C_YELLOW}UNKNOWN (not on a known subnet){C_RESET}")
    print(f"  Current VLAN : {vlan_disp}")
    print(f"  Local IP     : {my_ip}")

    keys = "SPACE pause  │  c config  │  Ctrl+C quit" if tty_mode else "Ctrl+C to stop"
    if paused:
        print(f"  Sweep #      : {sweep_count}  │  {C_YELLOW}PAUSED{C_RESET}  │  {keys}")
    else:
        print(f"  Sweep #      : {sweep_count}  │  {C_GREEN}ACTIVE{C_RESET}  "
              f"│  next in {countdown:2d}s  │  {keys}")

    print(f"{'─' * width}")

    print(f"\n  {C_BOLD}CURRENT VLAN SWEEP  (source: {current_vlan or 'UNKNOWN'}){C_RESET}\n")
    render_current_sweep(cfg["vlans"], current_vlan, results)

    vlan_names = [v["name"] for v in cfg["vlans"]]
    print(f"\n  {C_BOLD}FULL REACHABILITY MATRIX  (rows = source, columns = destination){C_RESET}\n")
    render_matrix(vlan_names, results)

    total   = len(vlan_names) * len(vlan_names)
    tested  = len(results)
    reach   = sum(1 for v in results.values() if v.get("last") is True)
    blocked = sum(1 for v in results.values() if v.get("last") is False)

    print(f"\n{'─' * width}")
    print(f"  Matrix coverage: {tested}/{total} pairs  │  "
          f"{C_GREEN}{reach} reachable{C_RESET}  │  {C_RED}{blocked} blocked{C_RESET}")
    print(f"{'═' * width}")


# ── Keyboard listener (raw mode) ──────────────────────────────────────────────

class KbState:
    __slots__ = ("paused", "config_requested", "stop")
    def __init__(self):
        self.paused           = False
        self.config_requested = False
        self.stop             = False


def kb_loop(state):
    """Runs in a background thread. Reads one char at a time from stdin."""
    fd  = sys.stdin.fileno()
    old = termios.tcgetattr(fd)
    try:
        tty.setcbreak(fd)
        while not state.stop:
            r, _, _ = select.select([sys.stdin], [], [], 0.1)
            if not r:
                continue
            ch = sys.stdin.read(1)
            if ch == ' ':
                state.paused = not state.paused
            elif ch in ('c', 'C'):
                state.config_requested = True
                state.stop = True       # thread exits; main will restart it
            elif ch in ('\x03', 'q', 'Q'):
                state.stop = True
                os.kill(os.getpid(), 2)  # send SIGINT to main
    finally:
        try:
            termios.tcsetattr(fd, termios.TCSADRAIN, old)
        except Exception:
            pass


def start_kb_thread():
    state  = KbState()
    thread = threading.Thread(target=kb_loop, args=(state,), daemon=True)
    thread.start()
    return state, thread


# ── Main loop ─────────────────────────────────────────────────────────────────

def main():
    if os.name == "nt":
        print("This CLI is for Linux/macOS. Use the GUI version on Windows.")
        sys.exit(1)

    force_setup = "--setup" in sys.argv

    ok, err = check_ping_available()
    if not ok:
        print(f"{C_RED}ERROR:{C_RESET} {err}")
        sys.exit(1)

    tty_mode = is_tty()
    if not tty_mode:
        print(f"{C_YELLOW}Note: stdin is not a TTY - keyboard controls disabled. "
              f"Use Ctrl+C to stop.{C_RESET}")

    cfg = load_config()
    if force_setup or not cfg.get("vlans"):
        cfg = run_setup_wizard(cfg)

    results     = load_results()
    sweep_count = 0

    state, kb_thread = (None, None)
    if tty_mode:
        state, kb_thread = start_kb_thread()

    print(f"\n{C_GREEN}Starting sweep...{C_RESET}")
    time.sleep(0.5)

    try:
        while True:
            # Config menu request (only from TTY)
            if state and state.config_requested:
                # kb thread has already exited and restored terminal
                try:
                    kb_thread.join(timeout=1.0)
                except Exception:
                    pass
                clear()
                cfg = config_menu(cfg)
                # Restart kb listener for the next sweep cycle
                state, kb_thread = start_kb_thread()

            paused = state.paused if state else False

            # Always re-detect IP, even while paused
            local_ips           = get_local_ips()
            current_vlan, my_ip = detect_vlan(local_ips, cfg["vlans"])

            if paused:
                render(cfg, current_vlan, my_ip, sweep_count, results, 0,
                       paused=True, tty_mode=tty_mode)
                time.sleep(1)
                continue

            # Full sweep
            sweep_count += 1
            for v in cfg["vlans"]:
                if state and state.config_requested:
                    break
                key = f"{current_vlan or 'UNKNOWN'}->{v['name']}"
                reached, rtt = ping(v["target"],
                                    count=cfg["ping_count"],
                                    timeout=cfg["ping_timeout"])
                results[key] = {
                    "last":    reached,
                    "rtt":     rtt,
                    "time":    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "from_ip": my_ip,
                }
                render(cfg, current_vlan, my_ip, sweep_count, results, 0,
                       paused=False, tty_mode=tty_mode)

            save_results(results)

            # Countdown to next sweep, re-detect IP each second
            interval = cfg["ping_interval"]
            for remaining in range(interval, 0, -1):
                if state and state.config_requested:
                    break
                local_ips           = get_local_ips()
                current_vlan, my_ip = detect_vlan(local_ips, cfg["vlans"])

                while state and state.paused and not state.config_requested:
                    render(cfg, current_vlan, my_ip, sweep_count, results, 0,
                           paused=True, tty_mode=tty_mode)
                    time.sleep(1)
                    local_ips           = get_local_ips()
                    current_vlan, my_ip = detect_vlan(local_ips, cfg["vlans"])

                render(cfg, current_vlan, my_ip, sweep_count, results, remaining,
                       paused=False, tty_mode=tty_mode)
                time.sleep(1)

    except KeyboardInterrupt:
        if state:
            state.stop = True
        save_results(results)
        print(f"\n\n  Stopped. Results saved to {RESULTS_FILE}")
        sys.exit(0)


if __name__ == "__main__":
    main()
