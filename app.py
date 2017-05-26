import sys
import ctypes
import win32con
import pytoml as toml
from clutch import VK_MOD, VK_CODE
from ctypes import wintypes
from pycaw.pycaw import AudioUtilities  # https://github.com/AndreMiras/pycaw


def application_information():  # debug
        name = 'Clutch'
        version = '0.1.0'
        author = 'Stephen Lorenz'
        contact = 'steviejlorenz@gmail.com'

        print('Application: {}\nAuthor: {} ({})\nVersion: {}\n'
              .format(name, author, contact, version))


class Clutch(object):
    def __init__(self, config_filename='conf.toml', muted=False):
        application_information()  # debug

        self.muted = muted
        self.config_filename = config_filename
        self.config = self.load_configuration()

        # Start of Configuration
        # [settings]
        settings = self.config['settings']  # readability
        self.whitelist = settings['whitelist']  # readability
        self.unmute_on_quit = settings['unmute_on_quit']

        # [keybindings]
        keybindings = self.config['keybindings']  # readability
        # VK_CODE is a dictionary that maps a string representation of a key to the key's hexadecimal value (clutch.py)
        # The next four lines replace the string with the hexademical value that is used to communicate with windows
        # toggle keybind
        if keybindings['toggle'] in VK_CODE:
            self.toggle_keybind = VK_CODE[keybindings['toggle']]
        else:
            # todo: key capture to rewrite config here
            print('Error: "{}" is not a valid keybind for toggle in {}. Exiting program.'
                  .format(keybindings['toggle'], config_filename))
            sys.exit(1)

        # toggle mod
        if keybindings['toggle_mod'] in VK_MOD:
            self.toggle_mod = VK_MOD[keybindings['toggle_mod']]
        else:
            # todo: key capture to rewrite config here
            print('Error: "{}" is not a valid keybind for toggle_mod in {}. Exiting program.'
                  .format(keybindings['toggle_mod'], config_filename))
            sys.exit(1)

        # quit keybind
        if keybindings['quit'] in VK_CODE:
            self.quit_keybind = VK_CODE[keybindings['quit']]
        else:
            # todo: key capture to rewrite config here
            print('Error: "{}" is not a valid keybind for quit in {}. Exiting program.'
                  .format(keybindings['quit'], config_filename))
            sys.exit(1)

        # quit mod
        if keybindings['quit_mod'] in VK_MOD:
            self.quit_mod = VK_MOD[keybindings['quit_mod']]
        else:
            # todo: key capture to rewrite config here
            print('Error: "{}" is not a valid keybind for quit_mod in {}. Exiting program.'
                  .format(keybindings['quit_mod'], config_filename))
            sys.exit(1)

        # End of Configuration
        print(self.config_filename, 'successfully loaded.\n')

    # load configuration from a .toml file into a dictionary
    def load_configuration(self):
        print('Loading', self.config_filename)
        try:
            with open(self.config_filename, 'rb') as config_file:
                config = toml.load(config_file)
        except toml.TomlError:
            config = None
            parse_error = sys.exc_info()  # debug
            print(parse_error)  # debug
            print('An error occurred when loading {}. Exiting program'.format(self.config_filename))
            sys.exit(1)
        return config

    # unmutes the volume of all processes that are in the whitelist
    def unmute(self):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            volume = session.SimpleAudioVolume
            if session.Process and session.Process.name() not in self.whitelist:
                print(session.Process.name() + ' has been unmuted.')  # debug
                volume.SetMute(0, None)
        print()  # debug
        self.muted = False

    # mutes the volume of all processes that are in the whitelist
    def mute(self):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            volume = session.SimpleAudioVolume
            if session.Process and session.Process.name() not in self.whitelist:
                print(session.Process.name() + ' has been muted.')  # debug
                volume.SetMute(1, None)
        print()  # debug
        self.muted = True

# instantiate clutch
clutch = Clutch()

byref = ctypes.byref
user32 = ctypes.windll.user32

HOTKEYS = {
    1: (clutch.toggle_keybind, clutch.toggle_mod),
    2: (clutch.quit_keybind, clutch.quit_mod)
}


def handle_toggle():
    if clutch.muted:
        clutch.unmute()
    elif not clutch.muted:
        clutch.mute()
    else:
        print('Error: Unable to toggle. Exiting program.')
        handle_quit()


def handle_quit():
    if clutch.unmute_on_quit:
        print('Unmuting all processes before closing the application.')
        clutch.unmute()
    print('Exiting Clutch.')
    user32.PostQuitMessage(0)


HOTKEY_ACTIONS = {
    1: handle_toggle,
    2: handle_quit
}


def main():
    #  RegisterHotKey takes:
    #  Window handle for WM_HOTKEY messages (None = this thread)
    #  arbitrary id unique within the thread
    #  modifiers (MOD_SHIFT, MOD_ALT, MOD_CONTROL, MOD_WIN)
    #  VK code (either ord ('x') or one of win32con.VK_*)
    for hotkey_id, (vk, modifiers) in HOTKEYS.items():
        # prints the ID being registered, the key's hex value, and searches VK_CODE for the human-readable name
        print('Registering ID', hotkey_id, 'for key:', vk,
              '({})'.format(list(VK_CODE.keys())[list(VK_CODE.values()).index(vk)]))  # unnecessary, only for debugging

        if not user32.RegisterHotKey(None, hotkey_id, modifiers, vk):
            print('Error: Unable to register ID:', hotkey_id)
            print('This key may be unavailable for keybinding. Is Clutch already running?')

    print('All keys were successfully registered.\n')

    #  Home-grown Windows message loop: does
    #  just enough to handle the WM_HOTKEY
    #  messages and pass everything else along.
    try:
        msg = wintypes.MSG()
        while user32.GetMessageA(byref(msg), None, 0, 0) != 0:
            if msg.message == win32con.WM_HOTKEY:
                action_to_take = HOTKEY_ACTIONS.get(msg.wParam)
                if action_to_take:
                    action_to_take()

                user32.TranslateMessage(byref(msg))
            user32.DispatchMessageA(byref(msg))

    finally:
        for hotkey_id in HOTKEYS.keys():
            user32.UnregisterHotKey(None, hotkey_id)

if __name__ == "__main__":
    main()
