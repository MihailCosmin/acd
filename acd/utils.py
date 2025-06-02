""" Functions and objects needed for gui creation """
import sys

# Qt imports
if sys.version_info >= (3, 12):
    from PySide6.QtUiTools import QUiLoader
    from PySide6.QtCore import QMetaObject
    from PySide6.QtWidgets import QMainWindow
    from PySide6.QtCore import Signal
else:
    from PySide2.QtUiTools import QUiLoader
    from PySide2.QtCore import QMetaObject
    from PySide2.QtWidgets import QMainWindow
    from PySide2.QtCore import Signal


class UiLoader(QUiLoader):
    """_summary_

    Args:
        QUiLoader (_type_): _description_
    """

    def __init__(self, baseinstance, custom_widgets=None):
        """_summary_

        Args:
            baseinstance (_type_): _description_
            custom_widgets (_type_, optional): _description_. Defaults to None.
        """
        QUiLoader.__init__(self, baseinstance)
        self.baseinstance = baseinstance
        self.customWidgets = custom_widgets

    def createWidget(self, class_name, parent=None, name=""):
        """_summary_

        Args:
            class_name (_type_): _description_
            parent (_type_, optional): _description_. Defaults to None.
            name (str, optional): _description_. Defaults to ''.

        Raises:
            Exception: _description_

        Returns:
            _type_: _description_
        """
        if parent is None and self.baseinstance:
            return self.baseinstance

        else:
            if class_name in self.availableWidgets():
                widget = QUiLoader.createWidget(self, class_name, parent, name)

            else:
                try:
                    widget = self.customWidgets[class_name](parent)

                except (TypeError, KeyError):
                    raise Exception(
                        "No custom widget "
                        + class_name
                        + " found in customWidgets param of"
                        + "UiLoader __init__."
                    )

            if self.baseinstance:
                setattr(self.baseinstance, name, widget)
            return widget


def load_ui(ui_file, baseinstance=None, custom_widgets=None, wd=None):
    """_summary_

    Args:
        ui_file (_type_): _description_
        baseinstance (_type_, optional): _description_. Defaults to None.
        custom_widgets (_type_, optional): _description_. Defaults to None.
        wd (_type_, optional): _description_. Defaults to None.

    Returns:
        _type_: _description_
    """   
    loader = UiLoader(baseinstance, custom_widgets)
    if wd is not None:
        loader.setWorkingDirectory(wd)
    widget = loader.load(ui_file)
    QMetaObject.connectSlotsByName(widget)
    return widget