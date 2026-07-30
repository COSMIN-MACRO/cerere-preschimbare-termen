# -*- coding: utf-8 -*-
"""Creare shortcut pe Desktop pentru Porneste_Aplicatia.bat (pywin32).
Daca automatul esueaza, foloseste procedura manuala din CITESTE_MA.txt."""
import os

def desktop_dir():
    cand = [
        os.path.join(os.path.expanduser("~"), "OneDrive", "Desktop"),
        os.path.join(os.path.expanduser("~"), "Desktop"),
    ]
    for d in cand:
        if os.path.isdir(d):
            return d
    return cand[-1]

def main():
    here = os.path.dirname(os.path.abspath(__file__))
    target = os.path.join(here, "Porneste_Aplicatia.bat")
    icon = os.path.join(here, "icon.ico")
    lnk = os.path.join(desktop_dir(), "Cerere Preschimbare Termen.lnk")
    try:
        from win32com.client import Dispatch
        sh = Dispatch("WScript.Shell")
        sc = sh.CreateShortCut(lnk)
        sc.Targetpath = target
        sc.WorkingDirectory = here
        if os.path.exists(icon):
            sc.IconLocation = icon
        sc.save()
        print("Shortcut creat pe Desktop:", lnk)
    except Exception as e:
        print("Nu s-a putut crea automat shortcut-ul:", e)
        print("Foloseste procedura manuala din CITESTE_MA.txt.")

if __name__ == "__main__":
    main()
    try:
        input("\nApasa Enter pentru a inchide...")
    except Exception:
        pass
