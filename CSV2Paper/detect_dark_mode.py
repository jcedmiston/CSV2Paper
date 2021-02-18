import subprocess
import platform

def is_system_dark():
    def detect_dark_mode_windows(): 
        try:
            import winreg
        except ImportError:
            return False
        registry = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
        reg_keypath = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize'
        try:
            reg_key = winreg.OpenKey(registry, reg_keypath)
        except FileNotFoundError:
            return False

        for i in range(1024):
            try:
                value_name, value, _ = winreg.EnumValue(reg_key, i)
                if value_name == 'AppsUseLightTheme':
                    return value == 0
            except OSError:
                break
        return False

    def detect_dark_mode_gnome():
        '''Detects dark mode in GNOME'''
        getArgs = ['gsettings', 'get', 'org.gnome.desktop.interface', 'gtk-theme']

        currentTheme = subprocess.run(
            getArgs, capture_output=True
        ).stdout.decode("utf-8").strip().strip("'")

        darkIndicator = '-dark'
        if currentTheme.endswith(darkIndicator):
            return True
        return False

    def detect_dark_mode_mac():
        """Checks DARK/LIGHT mode of macos."""
        cmd = 'defaults read -g AppleInterfaceStyle'
        p = subprocess.Popen(cmd, stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE, shell=True)
        return bool(p.communicate()[0])

    
    _is_system_dark = False
    if platform.system() == 'Darwin':       # macOS
        _is_system_dark = detect_dark_mode_mac()
    elif platform.system() == 'Windows':    # Windows
        _is_system_dark = detect_dark_mode_windows()
    else:                                   # linux variants
        _is_system_dark = detect_dark_mode_gnome()
    
    return _is_system_dark