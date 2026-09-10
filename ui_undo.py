import tkinter as tk
from tkinter import ttk


_MAX_HISTORY = 50


def _winfo_class(widget):
    try:
        return widget.winfo_class()
    except Exception:
        return ""


def _is_text_widget(widget):
    return isinstance(widget, tk.Text) or _winfo_class(widget) == "Text"


def _is_entry_like(widget):
    return _winfo_class(widget) in {"Entry", "TEntry", "TCombobox", "Spinbox", "TSpinbox"}


def _is_overrideredirect(window):
    try:
        value = window.wm_overrideredirect()
    except Exception:
        return False
    return str(value).strip().lower() in {"1", "true"}


def _get_widget_value(widget):
    try:
        if _is_text_widget(widget):
            return widget.get("1.0", "end-1c")
        if _is_entry_like(widget):
            return widget.get()
    except Exception:
        return None
    return None


def _set_widget_value(widget, value):
    try:
        if _is_text_widget(widget):
            state = str(widget.cget("state"))
            if state == "disabled":
                return False
            widget.delete("1.0", "end")
            widget.insert("1.0", value)
            widget.mark_set("insert", "end-1c")
            return True
        if _is_entry_like(widget):
            state = ""
            try:
                state = str(widget.cget("state"))
            except Exception:
                state = ""
            if state == "disabled":
                return False
            if _winfo_class(widget) == "TCombobox":
                widget.set(value)
            else:
                widget.delete(0, "end")
                widget.insert(0, value)
            try:
                widget.icursor("end")
            except Exception:
                pass
            return True
    except Exception:
        return False
    return False


def _remember_focus(event):
    widget = event.widget
    try:
        top = widget.winfo_toplevel()
        top._last_undo_widget = widget
    except Exception:
        pass
    if _is_text_widget(widget):
        try:
            widget.configure(undo=True, autoseparators=True, maxundo=_MAX_HISTORY)
        except Exception:
            pass
    _record_widget_snapshot(widget, force=True)


def _record_widget_snapshot(widget, force=False):
    if not (_is_text_widget(widget) or _is_entry_like(widget)):
        return
    current = _get_widget_value(widget)
    if current is None:
        return
    history = getattr(widget, "_undo_history", [])
    if force or not history or history[-1] != current:
        history.append(current)
        if len(history) > _MAX_HISTORY:
            history = history[-_MAX_HISTORY:]
        widget._undo_history = history


def _record_after_idle(widget):
    try:
        widget.after_idle(lambda w=widget: _record_widget_snapshot(w))
    except Exception:
        _record_widget_snapshot(widget)


def _on_key_release(event):
    _record_after_idle(event.widget)


def _on_combo_selected(event):
    _record_after_idle(event.widget)


def _target_widget(window):
    widget = None
    try:
        widget = window.focus_get()
    except Exception:
        widget = None
    if widget and (_is_text_widget(widget) or _is_entry_like(widget)):
        return widget
    return getattr(window, "_last_undo_widget", None)


def undo_focused_widget(window):
    widget = _target_widget(window)
    if widget is None:
        return "break"

    if _is_text_widget(widget):
        before = _get_widget_value(widget)
        used_tk_undo = False
        try:
            widget.edit_undo()
            used_tk_undo = True
        except Exception:
            pass
        after = _get_widget_value(widget)
        if used_tk_undo and after != before:
            _record_widget_snapshot(widget, force=True)
            return "break"
        history = getattr(widget, "_undo_history", [])
        if len(history) >= 2:
            history.pop()
            widget._undo_history = history
            _set_widget_value(widget, history[-1])
        return "break"

    if _is_entry_like(widget):
        history = getattr(widget, "_undo_history", [])
        current = _get_widget_value(widget)
        if current is None:
            return "break"
        if not history:
            history = [current]
        elif history[-1] != current:
            history.append(current)
        if len(history) < 2:
            return "break"
        history.pop()
        widget._undo_history = history
        _set_widget_value(widget, history[-1])
        return "break"

    return "break"


def _bind_widget_classes(root):
    if getattr(tk, "_hope_pharma_undo_class_bindings_installed", False):
        return
    for class_name in ("Entry", "TEntry", "Text", "TCombobox", "Spinbox", "TSpinbox"):
        root.bind_class(class_name, "<FocusIn>", _remember_focus, add="+")
        root.bind_class(class_name, "<KeyRelease>", _on_key_release, add="+")
    root.bind_class("TCombobox", "<<ComboboxSelected>>", _on_combo_selected, add="+")
    tk._hope_pharma_undo_class_bindings_installed = True


def disable_window_undo(window):
    try:
        window._undo_disabled = True
        overlay = getattr(window, "_undo_overlay", None)
        if overlay and overlay.winfo_exists():
            overlay.destroy()
    except Exception:
        pass


def install_undo_support(window):
    try:
        if getattr(window, "_undo_disabled", False) or getattr(window, "_undo_installed", False):
            return
        if _is_overrideredirect(window):
            return
    except Exception:
        return

    try:
        _bind_widget_classes(window)
        overlay = ttk.Frame(window)
        button = ttk.Button(overlay, text="Undo", width=8, command=lambda w=window: undo_focused_widget(w))
        button.pack(side="right")
        overlay.place(relx=1.0, rely=1.0, x=-12, y=-12, anchor="se")
        overlay.lift()
        window.bind("<Configure>", lambda event, o=overlay: o.lift(), add="+")
        window.bind("<Command-z>", lambda event, w=window: undo_focused_widget(w), add="+")
        window.bind("<Control-z>", lambda event, w=window: undo_focused_widget(w), add="+")
        window._undo_overlay = overlay
        window._undo_installed = True
    except Exception:
        pass


def _schedule_undo_install(window):
    try:
        if getattr(window, "_undo_install_scheduled", False):
            return
        window._undo_install_scheduled = True
        window.after_idle(lambda w=window: install_undo_support(w))
    except Exception:
        pass


def patch_tk_undo():
    if getattr(tk, "_hope_pharma_undo_patched", False):
        return

    original_tk_init = tk.Tk.__init__
    original_toplevel_init = tk.Toplevel.__init__

    def patched_tk_init(self, *args, **kwargs):
        original_tk_init(self, *args, **kwargs)
        _schedule_undo_install(self)

    def patched_toplevel_init(self, *args, **kwargs):
        original_toplevel_init(self, *args, **kwargs)
        _schedule_undo_install(self)

    tk.Tk.__init__ = patched_tk_init
    tk.Toplevel.__init__ = patched_toplevel_init
    tk._hope_pharma_undo_patched = True


patch_tk_undo()
