# -*- coding: utf-8 -*-
"""
Haftalık Koçluk Çizelgesi - Giriş noktası.
Bu dosyadan veya excel_organizer.py'den çalıştırılabilir.
"""

import os
import sys

# Proje kökü
_ROOT = os.path.dirname(os.path.abspath(__file__))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

# DPI (Windows)
if sys.platform.startswith("win"):
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

from excel_organizer import CoachingApp
import tkinter as tk


def main():
    root = tk.Tk()
    app = CoachingApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
