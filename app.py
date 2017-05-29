import sys
import ctypes
import win32con
import pytoml as toml
from msvcrt import getch
from clutch import VK_MOD, VK_CODE
from ctypes import wintypes
from pycaw.pycaw import AudioUtilities  # https://github.com/AndreMiras/pycaw


# debug functions
def indent():
    print('    ', end='')


class app_info(object):  # debug
    name = 'Clutch'
    version = '0.3.1'
    author = 'Stephen Lorenz'
    contact = 'steviejlorenz@gmail.com'

    def __init__(self):
        print('Application: {}\nAuthor: {} ({})\nVersion: {}\n'
              .format(self.name, self.author, self.contact, self.version))


class ConfigurationInterface(object):
    def __init__(self, config_filename='conf.toml'):
        print('* Initializing Configuration Interface')
        self.config_filename = config_filename
        self.config = self.load_configuration()  # HUGE DEBUG - NOT NEEDED

        # Start of Configuration
        # [settings]
        self.settings = self.config['settings']  # readability
        self.whitelist = self.settings['whitelist']

        self.unmute_on_quit = self.settings['unmute_on_quit']

        # [music]
        self.music = self.config['music']

        # [keybindings]
        self.keybindings = self.config['keybindings']  # readability
        keybind_pairs = [
            ('toggle', 'toggle_mod'),
            ('volume_up', 'volume_up_mod'),
            ('volume_down', 'volume_down_mod'),
            ('quit', 'quit_mod'),
            ('suspend', 'suspend_mod')
        ]
        # VK_CODE is a dictionary that maps a string representation of a key to the key's hexadecimal value (clutch.py)
        # VK_MOD is a dictionary that maps a string representation of a mod key to the key's integer value (clutch.py)
        indent()
        for keybind_pair in keybind_pairs:
            self.import_keybind(keybind_pair[0], self.keybindings[keybind_pair[0]], VK_CODE)
            self.import_keybind(keybind_pair[1], self.keybindings[keybind_pair[1]], VK_MOD)
        # End of Configuration
        print(self.config_filename, 'successfully imported.')
        print('* Configuration Interface started\n')

    def import_keybind(self, bind, key, key_dictionary):
        if key in key_dictionary:
            self.keybindings[bind] = key_dictionary[key]
        else:
            # todo: key capture to rewrite config here
            print('Error: "{}" is not a valid keybind for {} in {}. Exiting program.'
                  .format(key, bind, self.config_filename))
            print('\nPress Enter to exit the window...')  # debug
            getch()  # debug
            sys.exit(1)

    # load configuration from a .toml file into a dictionary
    def load_configuration(self):
        indent()  # HUGE DEBUG - NOT NEEDED
        print('Loading', self.config_filename)
        try:
            with open(self.config_filename, 'rb') as config_file:
                config = toml.load(config_file)
        except toml.TomlError:
            # config = None  # debug
            indent()  # HUGE DEBUG - NOT NEEDED
            parse_error = sys.exc_info()  # debug
            print(parse_error)  # debug
            indent()  # HUGE DEBUG - NOT NEEDED
            print('An error occurred when loading {}. Exiting program'.format(self.config_filename))
            print('\nPress Enter to exit the window...')  # debug
            getch()  # debug
            sys.exit(1)
        return config


class AudioController(object):
    def __init__(self, whitelist, music_config, muted=False):
        print('* Initializing Audio Controller')
        self.whitelist = whitelist
        self.music_config = music_config
        self.music_player = music_config['music_player']
        self.volume_increase = music_config['volume_increase']
        self.volume_decrease = music_config['volume_decrease']
        self.muted = muted
        self.volume = self.process_volume()
        print('* Audio Controller started\n')

    # unmutes the volume of all processes that are in the whitelist
    def unmute(self):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() not in self.whitelist:
                indent()
                print(session.Process.name() + ' has been unmuted.')  # debug
                interface.SetMute(0, None)
        print()  # debug
        self.muted = False  # for toggle, not ideal

    # mutes the volume of all processes that are in the whitelist
    def mute(self):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() not in self.whitelist:
                indent()
                print(session.Process.name() + ' has been muted.')  # debug
                interface.SetMute(1, None)
        print()  # debug
        self.muted = True  # for toggle, not ideal

    def process_volume(self):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == self.music_player:
                indent()
                print('Volume:', interface.GetMasterVolume())  # debug
                return interface.GetMasterVolume()

    def set_volume(self, decibels):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == self.music_player:
                self.volume = min(1.0, max(0.0, decibels))  # only set volume in the range 0.0 to 1.0
                indent()
                interface.SetMasterVolume(self.volume, None)
                print('Volume set to', self.volume)  # debug

    def decrease_volume(self):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == self.music_player:
                self.volume = max(0.0, self.volume-self.volume_decrease)  # 0.0 is the min value, reduce by decibels
                indent()
                interface.SetMasterVolume(self.volume, None)
                print('Volume reduced to', int(self.volume*100))  # debug

    def increase_volume(self):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == self.music_player:
                self.volume = min(1.0, self.volume + self.volume_increase)  # 1.0 is the max value, raise by decibels
                interface.SetMasterVolume(self.volume, None)
                indent()
                print('Volume raised to', int(self.volume*100))  # debug


class HotkeyInterface(object):
    def __init__(self, keybindings, audio_controller, suspended=False):
        print('* Initializing Hotkey Interface')
        self.byref = ctypes.byref  # windows
        self.user32 = ctypes.windll.user32  # windows

        self.suspended = suspended
        self.keybindings = keybindings
        self.audio_controller = audio_controller
        self.hotkeys = {
            1: (keybindings['toggle'], keybindings['toggle_mod']),
            2: (keybindings['volume_up'], keybindings['volume_up_mod']),
            3: (keybindings['volume_down'], keybindings['volume_down_mod']),
            4: (keybindings['quit'], keybindings['quit_mod']),
            5: (keybindings['suspend'], keybindings['suspend_mod'])
        }
        self.hotkey_actions = {
            1: self.handle_toggle,
            2: self.handle_volume_up,
            3: self.handle_volume_down,
            4: self.handle_quit,
            5: self.handle_suspend
        }
        self.necessary_hotkeys = (4, 5)
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
            indent()  # HUGE DEBUG - NOT NEEDED
            print('Registering ID', hotkey_id, 'for key:', vk, '({})'
                  .format(list(VK_CODE.keys())[list(VK_CODE.values()).index(vk)]))  # only for debugging

            if not self.user32.RegisterHotKey(None, hotkey_id, modifiers, vk):
                register_error = True
                indent()  # HUGE DEBUG - NOT NEEDED
                print('Error: Unable to register ID:', hotkey_id, 'for key:', vk, '({})'
                      .format(list(VK_CODE.keys())[list(VK_CODE.values()).index(vk)]))
                indent()  # HUGE DEBUG - NOT NEEDED
                print('This key may be unavailable for keybinding. Is Clutch already running?')
                print()

        if register_error:
            indent()  # HUGE DEBUG - NOT NEEDED
            print('Error: Unable to register all hotkeys. Exiting program.')
            print('\nPress Enter to exit the window...')  # debug
            getch()  # debug
            sys.exit(1)
        else:
            indent()  # HUGE DEBUG - NOT NEEDED
            print('All hotkeys were successfully registered.')

    def unregister_hotkeys(self):
        for hotkey_id, (vk, modifiers) in self.hotkeys.items():
            indent()
            print('Unregistering ID', hotkey_id, 'for key:', vk, '({})'
                  .format(list(VK_CODE.keys())[list(VK_CODE.values()).index(vk)]))
            self.user32.UnregisterHotKey(None, hotkey_id)

    # the functions below could be better optimized - but they work for now
    def suspend_hotkeys(self):
        for hotkey_id in self.hotkeys.keys():
            if hotkey_id not in self.necessary_hotkeys:  # sketchy debug, proof of concept
                self.user32.UnregisterHotKey(None, hotkey_id)

    def unsuspend_hotkeys(self):
        register_error = False
        for hotkey_id, (vk, modifiers) in self.hotkeys.items():
            if hotkey_id not in self.necessary_hotkeys:
                indent()
                # prints the ID being registered, the key's hex value, and searches VK_CODE for the human-readable name
                print('Registering ID', hotkey_id, 'for key:', vk, '({})'
                      .format(list(VK_CODE.keys())[list(VK_CODE.values()).index(vk)]))  # debug
                if not self.user32.RegisterHotKey(None, hotkey_id, modifiers, vk):
                    register_error = True
                    print('Error: Unable to register ID:', hotkey_id)
                    print('This key may be unavailable for keybinding. Is Clutch already running?')

        if register_error:
            print('Error: Unable to re-register hotkey. Exiting program.\n')  # wait till loop finishes for more info
            print('\nPress Enter to exit the window...')  # debug
            getch()  # debug
            sys.exit(1)
        else:
            indent()
            print('All hotkeys were successfully re-registered.\n')

    def handle_toggle(self):
        if self.audio_controller.muted:
            self.audio_controller.unmute()
        elif not self.audio_controller.muted:
            self.audio_controller.mute()
        else:
            print('Error: Unable to toggle audio. Exiting program.')
            self.handle_quit()

    def handle_volume_up(self):
        self.audio_controller.increase_volume()

    def handle_volume_down(self):
        self.audio_controller.decrease_volume()

    def handle_quit(self):
        if True:
            indent()
            print('Unmuting all processes before closing the application.')
            if self.audio_controller.muted:
                self.audio_controller.unmute()
            else:
                indent()
                print('All processes already unmuted.\n')
        self.unregister_hotkeys()
        self.user32.PostQuitMessage(0)

    def handle_suspend(self):
        if self.suspended:
            self.suspended = False
            indent()
            print('Application has been unsuspended.\n')
            self.unsuspend_hotkeys()
        elif not self.suspended:
            self.suspended = True
            indent()
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
            print('* Exiting Clutch')


class Clutch(object):
    def __init__(self):
        app_info()
        self.config_interface = ConfigurationInterface()
        self.audio_controller = AudioController(self.config_interface.whitelist, self.config_interface.music)
        self.hotkey_interface = HotkeyInterface(self.config_interface.keybindings, self.audio_controller)
        self.windows_interface = WindowsInterface(self.hotkey_interface)

    def run(self):
        print('* Running Cluch')
        self.windows_interface.message_loop()


def main():
    app = Clutch()
    app.run()
    print('\nPress Enter to exit the window...')  # debug
    getch()  # debug

if __name__ == "__main__":
    main()
