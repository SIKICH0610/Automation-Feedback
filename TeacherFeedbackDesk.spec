# -*- mode: python ; coding: utf-8 -*-
# Build: .venv/bin/python -m PyInstaller TeacherFeedbackDesk.spec --noconfirm
#
# One .app around desktop_app.py. The same binary is also every child process:
# the server relaunches it as `<app> --worker paste_sender|feedback_generator`,
# so those modules and their lazy platform imports must be collected even though
# nothing imports them at module scope.

a = Analysis(
    ["desktop_app.py"],
    pathex=[],
    binaries=[],
    datas=[
        ("frontend", "frontend"),
        # Seed workbook: imported once when a fresh install creates its database.
        ("Geo_TTh_Student_Script_fixed_rows_only.xlsx", "."),
    ],
    hiddenimports=[
        # Lazy imports the runners reach for at job time.
        "wecom_mac",
        "whatsapp_mac",
        "paste_sender",
        "feedback_generator",
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[
        # Windows-only automation; keep the Mac bundle lean.
        "pywinauto",
        "tkinter",
    ],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="Teacher Feedback Desk",
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    name="Teacher Feedback Desk",
)

app = BUNDLE(
    coll,
    name="Teacher Feedback Desk.app",
    icon=None,
    bundle_identifier="com.thinkacademy.teacher-feedback-desk",
    info_plist={
        "CFBundleDisplayName": "Teacher Feedback Desk",
        "CFBundleShortVersionString": "1.0.0",
        "NSHighResolutionCapable": True,
        # Shown when macOS asks about controlling System Events / WeCom for the
        # supervised paste automation.
        "NSAppleEventsUsageDescription": "需要通过系统辅助功能定位企业微信聊天窗口，并把生成的反馈粘贴到输入框（不会自动发送）。",
    },
)
