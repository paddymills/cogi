
import pyautogui

from pynput import mouse
from time import sleep


class ObjectLocation:

    DEFAULT_DELAY = 0.1
    
    def __init__(self, name, xy=None, lazy_load=True, delay=DEFAULT_DELAY):
        self.name = name
        self.delay = delay
        self._x = None
        self._y  = None

        if type(xy) in (tuple, pyautogui.Point):
            self._x, self._y = xy
        elif not lazy_load:
            self.get_location()

    @property
    def x(self):
        if self._x is None:
            self.get_location()

        return self._x

    @property
    def y(self):
        if self._y is None:
            self.get_location()

        return self._y

    def _set_location(self, x, y, button, state):
        if button == mouse.Button.left and state == False:
            self._x = x
            self._y = y
            
            return False

    def get_location(self):
        print("[{}] click on location".format(self.name))
        with mouse.Listener(on_click=self._set_location) as listener:
            listener.join()

    def click(self):
        pyautogui.click(self.x, self.y)
        sleep(self.delay)
