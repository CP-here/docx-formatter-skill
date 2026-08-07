"""
Helper for running LibreOffice (soffice) in headless mode.

Cross-platform: supports both Linux (including sandboxed VMs where AF_UNIX
sockets may be blocked) and Windows.

Linux sandbox handling: detects the AF_UNIX restriction at runtime and
applies an LD_PRELOAD shim if needed. On Windows the shim is unnecessary.

Usage:
    from soffice import run_soffice

    result = run_soffice(["--headless", "--convert-to", "pdf", "input.docx"])

Call soffice through run_soffice, not through subprocess with get_soffice_env():
the env dict carries the shim but names no user profile, and a non-root sandbox
cannot bootstrap the default one -- soffice aborts with "User installation could
not be completed" and converts nothing. get_soffice_env() stays public for the
callers that build their own argv (they must pass -env:UserInstallation too).
"""

import contextlib
import os
import platform
import shutil
import socket
import subprocess
import sys
import tempfile
from collections.abc import Iterable
from pathlib import Path


# ============================================================
# PLATFORM DETECTION
# ============================================================

_IS_WINDOWS = platform.system() == "Windows"
_IS_LINUX = platform.system() == "Linux"


def _find_soffice() -> str:
    """Locate the soffice executable on any platform."""
    # Check PATH first
    exe = shutil.which("soffice") or shutil.which("soffice.exe")
    if exe:
        return exe

    if _IS_WINDOWS:
        # Common Windows install locations
        win_paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        for p in win_paths:
            if Path(p).exists():
                return p

    # Linux/macOS common locations
    linux_paths = [
        "/usr/bin/soffice",
        "/usr/local/bin/soffice",
        "/opt/libreoffice/program/soffice",
        "/snap/bin/libreoffice",
    ]
    for p in linux_paths:
        if Path(p).exists():
            return p

    # Fall back to bare name (will fail with FileNotFoundError if not installed)
    return "soffice"


# ============================================================
# ENVIRONMENT
# ============================================================

def get_soffice_env() -> dict:
    """Return environment dict with LibreOffice headless settings.

    On Linux sandboxes where AF_UNIX sockets are blocked, an LD_PRELOAD
    shim is compiled and added to the environment. On Windows this step
    is skipped.
    """
    env = os.environ.copy()
    env["SAL_USE_VCLPLUGIN"] = "svp"

    if _IS_LINUX and _needs_shim():
        shim = _ensure_shim()
        env["LD_PRELOAD"] = str(shim)

    return env


# ============================================================
# MAIN ENTRY POINT
# ============================================================

def run_soffice(args: Iterable[str], **kwargs) -> subprocess.CompletedProcess:
    """Run soffice with a temporary user profile and headless env.

    Args:
        args: Arguments to pass to soffice
        **kwargs: Additional kwargs passed to subprocess.run

    Returns:
        subprocess.CompletedProcess result
    """
    args = list(args)
    soffice_exe = _find_soffice()

    with contextlib.ExitStack() as stack:
        if not any(str(a).startswith("-env:UserInstallation") for a in args):
            profile = stack.enter_context(
                tempfile.TemporaryDirectory(prefix="lo_profile_")
            )
            profile_uri = Path(profile).resolve().as_uri()
            args = [f"-env:UserInstallation={profile_uri}"] + args
        return subprocess.run([soffice_exe] + args, env=get_soffice_env(), **kwargs)


# ============================================================
# LINUX SHIM (LD_PRELOAD) — not used on Windows
# ============================================================

if _IS_LINUX:
    _SHIM_SO = Path(tempfile.gettempdir()) / "lo_socket_shim.so"

    def _needs_shim() -> bool:
        """Check if AF_UNIX sockets are blocked (sandboxed environment)."""
        try:
            s = socket.socket(socket.AF_UNIX, socket.SOCK_STREAM)
            s.close()
            return False
        except OSError:
            return True

    def _ensure_shim() -> Path:
        """Compile the LD_PRELOAD shim if not already present."""
        if _SHIM_SO.exists():
            return _SHIM_SO

        src = Path(tempfile.gettempdir()) / "lo_socket_shim.c"
        src.write_text(_SHIM_SOURCE)
        subprocess.run(
            ["gcc", "-shared", "-fPIC", "-o", str(_SHIM_SO), str(src), "-ldl"],
            check=True,
            capture_output=True,
        )
        src.unlink()
        return _SHIM_SO

    _SHIM_SOURCE = r"""
#define _GNU_SOURCE
#include <dlfcn.h>
#include <errno.h>
#include <signal.h>
#include <stdio.h>
#include <stdlib.h>
#include <sys/socket.h>
#include <unistd.h>

static int (*real_socket)(int, int, int);
static int (*real_socketpair)(int, int, int, int[2]);
static int (*real_listen)(int, int);
static int (*real_accept)(int, struct sockaddr *, socklen_t *);
static int (*real_close)(int);
static int (*real_read)(int, void *, size_t);

/* Per-FD bookkeeping (FDs >= 1024 are passed through unshimmed). */
static int is_shimmed[1024];
static int peer_of[1024];
static int wake_r[1024];            /* accept() blocks reading this */
static int wake_w[1024];            /* close()  writes to this      */
static int listener_fd = -1;        /* FD that received listen()    */

__attribute__((constructor))
static void init(void) {
    real_socket     = dlsym(RTLD_NEXT, "socket");
    real_socketpair = dlsym(RTLD_NEXT, "socketpair");
    real_listen     = dlsym(RTLD_NEXT, "listen");
    real_accept     = dlsym(RTLD_NEXT, "accept");
    real_close      = dlsym(RTLD_NEXT, "close");
    real_read       = dlsym(RTLD_NEXT, "read");
    for (int i = 0; i < 1024; i++) {
        peer_of[i] = -1;
        wake_r[i]  = -1;
        wake_w[i]  = -1;
    }
}

/* ---- socket ---------------------------------------------------------- */
int socket(int domain, int type, int protocol) {
    if (domain == AF_UNIX) {
        int fd = real_socket(domain, type, protocol);
        if (fd >= 0) return fd;
        /* socket(AF_UNIX) blocked – fall back to socketpair(). */
        int sv[2];
        if (real_socketpair(domain, type, protocol, sv) == 0) {
            if (sv[0] >= 0 && sv[0] < 1024) {
                is_shimmed[sv[0]] = 1;
                peer_of[sv[0]]    = sv[1];
                int wp[2];
                if (pipe(wp) == 0) {
                    wake_r[sv[0]] = wp[0];
                    wake_w[sv[0]] = wp[1];
                }
            }
            return sv[0];
        }
        errno = EPERM;
        return -1;
    }
    return real_socket(domain, type, protocol);
}

/* ---- listen ---------------------------------------------------------- */
int listen(int sockfd, int backlog) {
    if (sockfd >= 0 && sockfd < 1024 && is_shimmed[sockfd]) {
        listener_fd = sockfd;
        return 0;
    }
    return real_listen(sockfd, backlog);
}

/* ---- accept ---------------------------------------------------------- */
int accept(int sockfd, struct sockaddr *addr, socklen_t *addrlen) {
    if (sockfd >= 0 && sockfd < 1024 && is_shimmed[sockfd]) {
        /* Block until close() writes to the wake pipe. */
        if (wake_r[sockfd] >= 0) {
            char buf;
            real_read(wake_r[sockfd], &buf, 1);
        }
        errno = ECONNABORTED;
        return -1;
    }
    return real_accept(sockfd, addr, addrlen);
}

/* ---- close ----------------------------------------------------------- */
int close(int fd) {
    if (fd >= 0 && fd < 1024 && is_shimmed[fd]) {
        int was_listener = (fd == listener_fd);
        is_shimmed[fd] = 0;

        if (wake_w[fd] >= 0) {              /* unblock accept() */
            char c = 0;
            write(wake_w[fd], &c, 1);
            real_close(wake_w[fd]);
            wake_w[fd] = -1;
        }
        if (wake_r[fd] >= 0) { real_close(wake_r[fd]); wake_r[fd]  = -1; }
        if (peer_of[fd] >= 0) { real_close(peer_of[fd]); peer_of[fd] = -1; }

        if (was_listener)
            _exit(0);                        /* conversion done – exit */
    }
    return real_close(fd);
}
"""

else:
    # Windows / macOS: no shim needed, provide stubs so imports don't fail
    def _needs_shim() -> bool:
        return False

    def _ensure_shim() -> Path:
        raise NotImplementedError("LD_PRELOAD shim is Linux-only")


# ============================================================
# CLI ENTRY
# ============================================================

if __name__ == "__main__":
    result = run_soffice(sys.argv[1:])
    sys.exit(result.returncode)
