import sys
import ctypes
import win32con
import pytoml as toml
from clutch import VK_MOD, VK_CODE
from ctypes import wintypes
from pycaw.pycaw import AudioUtilities  # https://github.com/AndreMiras/pycaw


class app_info(object):  # debug
    name = 'Clutch'
    version = '0.2.0'
    author = 'Stephen Lorenz'
    contact = 'steviejlorenz@gmail.com'

    def __init__(self):
        print('Application: {}\nAuthor: {} ({})\nVersion: {}\n'
              .format(self.name, self.author, self.contact, self.version))


class ConfigurationInterface(object):
    def __init__(self, config_filename='conf.toml'):
        print('* Initializing Configuration Interface')
        self.config_filename = config_filename
        self.config = self.load_configuration()
        self.keybinds = {}
        print('    ', end='')  # HUGE DEBUG - NOT NEEDED

        # Start of Configuration
        # [settings]
        self.settings = self.config['settings']  # readability
        self.whitelist = self.settings['whitelist']
        self.music_app = self.settings['music_app']
        self.unmute_on_quit = self.settings['unmute_on_quit']

        # [keybindings]
        self.keybindings = self.config['keybindings']  # readability
        # VK_CODE is a dictionary that maps a string representation of a key to the key's hexadecimal value (clutch.py)
        # toggle keybind
        if self.keybindings['toggle'] in VK_CODE:
            self.keybindings['toggle'] = VK_CODE[self.keybindings['toggle']]
        else:
            # todo: key capture to rewrite config here
            print('Error: "{}" is not a valid keybind for toggle in {}. Exiting program.'
                  .format(self.keybindings['toggle'], config_filename))
            sys.exit(1)
        # toggle mod
        if self.keybindings['toggle_mod'] in VK_MOD:
            self.keybindings['toggle_mod'] = VK_MOD[self.keybindings['toggle_mod']]
        else:
            # todo: key capture to rewrite config here
            print('Error: "{}" is not a valid keybind for toggle_mod in {}. Exiting program.'
                  .format(self.keybindings['toggle_mod'], config_filename))
            sys.exit(1)

        # quit keybind
        if self.keybindings['quit'] in VK_CODE:
            self.keybindings['quit'] = VK_CODE[self.keybindings['quit']]
        else:
            # todo: key capture to rewrite config here
            print('Error: "{}" is not a valid keybind for quit in {}. Exiting program.'
                  .format(self.keybindings['quit'], config_filename))
            sys.exit(1)

        # quit mod
        if self.keybindings['quit_mod'] in VK_MOD:
            self.keybindings['quit_mod'] = VK_MOD[self.keybindings['quit_mod']]
        else:
            # todo: key capture to rewrite config here
            print('Error: "{}" is not a valid keybind for quit_mod in {}. Exiting program.'
                  .format(self.keybindings['quit_mod'], config_filename))
            sys.exit(1)

        # suspend keybind
        if self.keybindings['suspend'] in VK_CODE:
            self.keybindings['suspend'] = VK_CODE[self.keybindings['suspend']]
        else:
            # todo: key capture to rewrite config here
            print('Error: "{}" is not a valid keybind for suspend in {}. Exiting program.'
                  .format(self.keybindings['suspend'], config_filename))
            sys.exit(1)

        # suspend mod
        if self.keybindings['suspend_mod'] in VK_MOD:
            self.keybindings['suspend_mod'] = VK_MOD[self.keybindings['suspend_mod']]
        else:
            # todo: key capture to rewrite config here
            print('Error: "{}" is not a valid keybind for suspend_mod in {}. Exiting program.'
                  .format(self.keybindings['suspend_mod'], config_filename))
            sys.exit(1)
        # End of Configuration
        print(self.config_filename, 'successfully loaded.')
        print('* Configuration Interface started\n')

    # load configuration from a .toml file into a dictionary
    def load_configuration(self):
        print('    ', end='')  # HUGE DEBUG - NOT NEEDED
        print('Loading', self.config_filename)
        try:
            with open(self.config_filename, 'rb') as config_file:
                config = toml.load(config_file)
        except toml.TomlError:
            # config = None  # debug
            print('    ', end='')  # HUGE DEBUG - NOT NEEDED
            parse_error = sys.exc_info()  # debug
            print(parse_error)  # debug
            print('    ', end='')  # HUGE DEBUG - NOT NEEDED
            print('An error occurred when loading {}. Exiting program'.format(self.config_filename))
            sys.exit(1)
        return config


class AudioInterface(object):
    def __init__(self, whitelist, muted=False):
        print('* Initializing Audio Interface')
        self.whitelist = whitelist
        self.muted = muted
        print('* Audio Interface started\n')

    # unmutes the volume of all processes that are in the whitelist
    def unmute(self):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            volume = session.SimpleAudioVolume
            if session.Process and session.Process.name() not in self.whitelist:
                print(session.Process.name() + ' has been unmuted.')  # debug
                volume.SetMute(0, None)
        print()  # debug
        self.muted = False  # for toggle, not ideal

    # mutes the volume of all processes that are in the whitelist
    def mute(self):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            volume = session.SimpleAudioVolume
            if session.Process and session.Process.name() not in self.whitelist:
                print(session.Process.name() + ' has been muted.')  # debug
                volume.SetMute(1, None)
        print()  # debug
        self.muted = True  # for toggle, not ideal


class HotkeyInterface(object):
    def __init__(self, keybindings, audio_interface, suspended=False):
        print('* Initializing Hotkey Interface')
        self.byref = ctypes.byref  # windows
        self.user32 = ctypes.windll.user32  # windows

        self.suspended = suspended
        self.keybindings = keybindings
        self.audio_interface = audio_interface
        self.hotkeys = {
            1: (keybindings['toggle'], keybindings['toggle_mod']),
            2: (keybindings['quit'], keybindings['quit_mod']),
            3: (keybindings['suspend'], keybindings['suspend_mod'])
        }
        self.hotkey_actions = {
            1: self.handle_toggle,
            2: self.handle_quit,
            3: self.handle_suspend
        }
        self.necessary_hotkeys = (2, 3)
        self.register_hotkeys()
        print('* Hotkey Interface started\n')

    def register_hotkeys(self):
        #  registerHotKey takes:
        #  window handle for WM_HOTKEY messages (None = this thread)
        #  arbitrary id unique within the thread
        #  modifiers (MOD_SHIFT, MOD_ALT, MOD_CONTROL, MOD_WIN)
        #  VK code (either ord ('x') or one of win32con.VK_*)
        register_error = False
        for hotkey_id, (vk, modifiers) in self.hotkeys.items():
            # prints the ID being registered, the key's hex value, and searches VK_CODE for the human-readable name
            print('    ', end='')  # HUGE DEBUG - NOT NEEDED
            print('Registering ID', hotkey_id, 'for key:', vk, '({})'
                  .format(list(VK_CODE.keys())[list(VK_CODE.values()).index(vk)]))  # only for debugging

            if not self.user32.RegisterHotKey(None, hotkey_id, modifiers, vk):
                register_error = True
                print('    ', end='')  # HUGE DEBUG - NOT NEEDED
                print('Error: Unable to register ID:', hotkey_id, 'for key:', vk, '({})'
                      .format(list(VK_CODE.keys())[list(VK_CODE.values()).index(vk)]))
                print('    ', end='')  # HUGE DEBUG - NOT NEEDED
                print('This key may be unavailable for keybinding. Is Clutch already running?')
                print()

        if register_error:
            print('    ', end='')  # HUGE DEBUG - NOT NEEDED
            print('Error: Unable to register all hotkeys. Exiting program.')
            sys.exit(1)
        else:
            print('    ', end='')  # HUGE DEBUG - NOT NEEDED
            print('All hotkeys were successfully registered.')

    def unregister_hotkeys(self):
        for hotkey_id, (vk, modifiers) in self.hotkeys.items():
            print('Unregistering ID', hotkey_id, 'for key:', vk, '({})'
                  .format(list(VK_CODE.keys())[list(VK_CODE.values()).index(vk)]))
            self.user32.UnregisterHotKey(None, hotkey_id)
        print()  # debug

    # the functions below could be better optimized - but they work for now
    def suspend_hotkeys(self):
        for hotkey_id in self.hotkeys.keys():
            if hotkey_id not in self.necessary_hotkeys:  # sketchy debug, proof of concept
                self.user32.UnregisterHotKey(None, hotkey_id)

    def unsuspend_hotkeys(self):
        register_error = False
        for hotkey_id, (vk, modifiers) in self.hotkeys.items():
            if hotkey_id not in self.necessary_hotkeys:
                # prints the ID being registered, the key's hex value, and searches VK_CODE for the human-readable name
                print('Registering ID', hotkey_id, 'for key:', vk, '({})'
                      .format(list(VK_CODE.keys())[list(VK_CODE.values()).index(vk)]))  # debug
                if not self.user32.RegisterHotKey(None, hotkey_id, modifiers, vk):
                    register_error = True
                    print('Error: Unable to register ID:', hotkey_id)
                    print('This key may be unavailable for keybinding. Is Clutch already running?')

        if register_error:
            print('     Error: Unable to re-register hotkey. Exiting program.\n')  # wait till loop finishes for more info
            sys.exit(1)
        else:
            print('All hotkeys were successfully re-registered.\n')

    def handle_toggle(self):
        if self.audio_interface.muted:
            self.audio_interface.unmute()
        elif not self.audio_interface.muted:
            self.audio_interface.mute()
        else:
            print('Error: Unable to toggle audio. Exiting program.')
            self.handle_quit()

    def handle_quit(self):
        if True:
            print('Unmuting all processes before closing the application.')
            if self.audio_interface.muted:
                self.audio_interface.unmute()
            else:
                print('All processes already unmuted.\n')
        self.unregister_hotkeys()
        self.user32.PostQuitMessage(0)

    def handle_suspend(self):
        if self.suspended:
            self.suspended = False
            print('Application has been unsuspended.\n')
            self.unsuspend_hotkeys()
        elif not self.suspended:
            self.suspended = True
            print('Application has been suspended.')
            self.suspend_hotkeys()
        else:
            print('Error: Unable to toggle suspend. Exiting program.')
            self.handle_quit()


class WindowsInterface(object):
    def __init__(self, hotkey_interface):
        print('* Initializing Windows Interface')
        self.byref = ctypes.byref
        self.user32 = ctypes.windll.user32
        self.hotkey_interface = hotkey_interface
        print('* Windows Interface started\n')

    def message_loop(self):
        # Home-grown Windows message loop: does
        # just enough to handle the WM_HOTKEY
        # messages and pass everything else along.
        # todo: replace with qt
        try:
            msg = wintypes.MSG()
            while self.user32.GetMessageA(self.byref(msg), None, 0, 0) != 0:
                if msg.message == win32con.WM_HOTKEY:
                    action_to_take = self.hotkey_interface.hotkey_actions.get(msg.wParam)
                    if action_to_take:
                        action_to_take()

                    self.user32.TranslateMessage(self.byref(msg))
                    self.user32.DispatchMessageA(self.byref(msg))
        finally:
            print('Exiting Clutch.')


class Clutch(object):
    def __init__(self):
        app_info()
        self.config_interface = ConfigurationInterface()
        self.audio_interface = AudioInterface(self.config_interface.whitelist)
        self.hotkey_interface = HotkeyInterface(self.config_interface.keybindings, self.audio_interface)
        self.windows_interface = WindowsInterface(self.hotkey_interface)

    def run(self):
        print('* Running Cluch\n')
        self.windows_interface.message_loop()


def main():
    app = Clutch()
    print('hey')
    app.run()

if __name__ == "__main__":
    main()
