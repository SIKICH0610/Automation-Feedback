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
        # Blank seed template: imported once when a fresh install creates its
        # database. Deliberately NOT the real roster workbook -- installers get
        # handed to colleagues, and student names/uids must not ride along.
        ("roster_template.xlsx", "."),
        # Quiz-bank seed CSVs: first-run data for the sqlite store.
        ("data/quiz_banks", "data/quiz_banks"),
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
    # Self-signed identity from the login keychain. Signing every build with the
    # SAME certificate is what lets macOS keep the Accessibility grant across
    # updates -- ad-hoc signatures change fingerprint per build and lose it.
    codesign_identity="Think Academy Automation",
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
