# src/lazy_loader.py
from importlib import import_module

class LazyLoader:
    def __init__(self, lib_name):
        self.lib_name = lib_name
        self._lib = None

    def __getattr__(self, name):
        if self._lib is None:
            self._lib = import_module(self.lib_name)
        return getattr(self._lib, name)

    def __dir__(self):
        if self._lib is None:
            self._lib = import_module(self.lib_name)
        return dir(self._lib)