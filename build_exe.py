<<<<<<< HEAD
"""
EXE Build Script
================
BIST_Finansal_Tablolar.py → BIST_Finansal_Tablolar.exe

Kullanım:
    python build_exe.py

Çıktı:
    dist/BIST_Finansal_Tablolar.exe
"""
import subprocess, sys, os

# PyInstaller kur
try:
    import PyInstaller
except ImportError:
    print("📦 PyInstaller kuruluyor...")
    subprocess.check_call([sys.executable,"-m","pip","install","pyinstaller"])

HIDDEN = [
    "borsapy","isyatirimhisse","openpyxl","openpyxl.styles","openpyxl.utils",
    "pandas","requests","selenium","selenium.webdriver","selenium.webdriver.chrome",
    "selenium.webdriver.chrome.options","selenium.webdriver.chrome.service",
    "selenium.webdriver.common.by","selenium.webdriver.support.ui",
    "selenium.webdriver.support.expected_conditions",
    "webdriver_manager","webdriver_manager.chrome",
    "bs4","lxml","re","importlib","importlib.metadata","tkinter","tkinter.ttk",
]

komut = [sys.executable,"-m","PyInstaller",
    "--onefile",
    "--windowed",                          # Konsol penceresi açılmaz (GUI uygulama)
    "--name","BIST_Finansal_Tablolar",
    "--clean",
] + [f"--hidden-import={h}" for h in HIDDEN] + ["BIST_Finansal_Tablolar.py"]

print("🔨 Build başlıyor (birkaç dakika sürebilir)...\n")
try:
    subprocess.check_call(komut)
    yol = os.path.abspath("dist/BIST_Finansal_Tablolar.exe")
    print("\n" + "="*55)
    print("✅ Build tamamlandı!")
    print(f"   📁 {yol}")
    print("\n   Bu .exe dosyasını istediğiniz bilgisayara")
    print("   kopyalayın ve çift tıklayın.")
    print("="*55)
except subprocess.CalledProcessError as e:
    print(f"\n❌ Build hatası: {e}")
    print("\nManüel komut:")
=======
"""
EXE Build Script
================
BIST_Finansal_Tablolar.py → BIST_Finansal_Tablolar.exe

Kullanım:
    python build_exe.py

Çıktı:
    dist/BIST_Finansal_Tablolar.exe
"""
import subprocess, sys, os

# PyInstaller kur
try:
    import PyInstaller
except ImportError:
    print("📦 PyInstaller kuruluyor...")
    subprocess.check_call([sys.executable,"-m","pip","install","pyinstaller"])

HIDDEN = [
    "borsapy","isyatirimhisse","openpyxl","openpyxl.styles","openpyxl.utils",
    "pandas","requests","selenium","selenium.webdriver","selenium.webdriver.chrome",
    "selenium.webdriver.chrome.options","selenium.webdriver.chrome.service",
    "selenium.webdriver.common.by","selenium.webdriver.support.ui",
    "selenium.webdriver.support.expected_conditions",
    "webdriver_manager","webdriver_manager.chrome",
    "bs4","lxml","re","importlib","importlib.metadata","tkinter","tkinter.ttk",
]

komut = [sys.executable,"-m","PyInstaller",
    "--onefile",
    "--windowed",                          # Konsol penceresi açılmaz (GUI uygulama)
    "--name","BIST_Finansal_Tablolar",
    "--clean",
] + [f"--hidden-import={h}" for h in HIDDEN] + ["BIST_Finansal_Tablolar.py"]

print("🔨 Build başlıyor (birkaç dakika sürebilir)...\n")
try:
    subprocess.check_call(komut)
    yol = os.path.abspath("dist/BIST_Finansal_Tablolar.exe")
    print("\n" + "="*55)
    print("✅ Build tamamlandı!")
    print(f"   📁 {yol}")
    print("\n   Bu .exe dosyasını istediğiniz bilgisayara")
    print("   kopyalayın ve çift tıklayın.")
    print("="*55)
except subprocess.CalledProcessError as e:
    print(f"\n❌ Build hatası: {e}")
    print("\nManüel komut:")
>>>>>>> d94de0d (Güncelleme)
    print("  pyinstaller --onefile --windowed --name BIST_Finansal_Tablolar BIST_Finansal_Tablolar.py")