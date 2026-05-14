"""
EXE Build Script
================
BIST_Finansal_Tablolar.py -> BIST_Finansal_Tablolar.exe

Kullanim:
    python build_exe.py

Cikti:
    dist/BIST_Finansal_Tablolar.exe
"""
import os
import subprocess
import sys

# PyInstaller kur
try:
    import PyInstaller  # noqa: F401
except ImportError:
    print("PyInstaller kuruluyor...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])

HIDDEN = [
    "borsapy", "isyatirimhisse", "openpyxl", "openpyxl.styles", "openpyxl.utils",
    "pandas", "requests", "selenium", "selenium.webdriver", "selenium.webdriver.chrome",
    "selenium.webdriver.chrome.options", "selenium.webdriver.chrome.service",
    "selenium.webdriver.common.by", "selenium.webdriver.support.ui",
    "selenium.webdriver.support.expected_conditions",
    "webdriver_manager", "webdriver_manager.chrome",
    "bs4", "lxml", "re", "importlib", "importlib.metadata", "tkinter", "tkinter.ttk",
]

komut = [
    sys.executable, "-m", "PyInstaller",
    "--onefile",
    "--windowed",
    "--name", "BIST_Finansal_Tablolar",
    "--clean",
] + [f"--hidden-import={h}" for h in HIDDEN] + ["BIST_Finansal_Tablolar.py"]

print("Build basliyor (birkac dakika surebilir)...\n")
try:
    subprocess.check_call(komut)
    yol = os.path.abspath("dist/BIST_Finansal_Tablolar.exe")
    print("\n" + "=" * 55)
    print("Build tamamlandi!")
    print(f"   {yol}")
    print("\n   Bu .exe dosyasini istediginiz bilgisayara")
    print("   kopyalayin ve cift tiklayin.")
    print("=" * 55)
except subprocess.CalledProcessError as e:
    print(f"\nBuild hatasi: {e}")
    print("\nManuel komut:")
    print("  pyinstaller --onefile --windowed --name BIST_Finansal_Tablolar BIST_Finansal_Tablolar.py")
