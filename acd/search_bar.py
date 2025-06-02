from functools import partial
import sys
if sys.version_info >= (3, 12):
    from PySide6.QtWidgets import QWidget
    from PySide6.QtWidgets import QHBoxLayout
    from PySide6.QtWidgets import QLineEdit
else:
    from PySide2.QtWidgets import QLineEdit
    from PySide2.QtWidgets import QMainWindow
    from PySide2.QtWidgets import QWidget

def include_search_bar(layout):
    """Creates the layout and widgets for the searchbar

    Args:
        layout (_type_): The QT Layout object which should contain the search bar
    """
    # Create search widget
    search_bar_layout = QHBoxLayout()
    search_edit = QLineEdit()
    search_edit.setPlaceholderText("Filter...")

    # Add widget to the layout
    search_bar_layout.addWidget(search_edit)
    layout.addLayout(search_bar_layout)

    # Connect the line edit to a function that filters the buttons when typing smthg into the box
    search_edit.textChanged.connect(
        partial(_filter_widgets, layout=layout, search_edit=search_edit)
    )

def _filter_widgets(
    search_text: str = None,
    layout=None,
    search_edit=None,
    exception_names: list = None,
    exception_widgets: list = None,
):
    """Function gets called whenever a change in the searchbar is made.
    Implements logic for hiding and showing the QWidgets that don't match the input.

    Args:
        search_text (str, optional): _description_. Defaults to None.
        layout (_type_, optional): _description_. Defaults to None.
        search_edit (_type_, optional): _description_. Defaults to None.
    """
    for i in range(layout.count()):
        if layout.itemAt(i).widget() is not None:
            cond1, cond2 = True, True
            if exception_names is not None:
                cond1 = layout.itemAt(i).widget(
                ).objectName() not in exception_names
            if exception_widgets is not None:
                cond2 = not any(isinstance(layout.itemAt(i).widget(), widget)
                                for widget in exception_widgets)
            if isinstance(layout.itemAt(i).widget(), QWidget) and cond1 and cond2:
                if search_text.lower() in layout.itemAt(i).widget().text().lower() or \
                        search_text.lower() in layout.itemAt(i).widget().text().lower().replace("_", " "):
                    layout.itemAt(i).widget().show()
                elif layout.itemAt(i).widget() != search_edit:
                    layout.itemAt(i).widget().hide()
