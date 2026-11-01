"""Blaykie Lab: the DOCX citation workflow.

Open a .docx, find the spot, tap or select it, search your Zotero library,
insert the citation, then export a NEW .docx with the APA 7 reference list.
The original file is never changed.
"""

import json
import os
import shutil
import threading
import traceback

from kivy.app import App
from kivy.clock import Clock
from kivy.core.window import Window
from kivy.metrics import dp, sp
from kivy.uix.behaviors import ButtonBehavior
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from kivy.uix.modalview import ModalView
from kivy.uix.scrollview import ScrollView
from kivy.uix.textinput import TextInput
from kivy.utils import escape_markup, platform

from ui import (
    ACCENT_TEXT, BG, ERROR, LINE, MUTED, SURFACE, SURFACE2, TEXT,
    InputBox, PillButton, ResultCard, col, paint, small_label,
)

Window.clearcolor = col(BG)
Window.softinput_mode = "below_target"

IMPORT_ERRORS = []
try:
    from blaykie_docx import (
        DocxProject, MissingReferences, NotEditable, build_citation_text,
    )
    from zotero_engine import (
        fetch_items_by_keys, get_authors, get_year, search_zotero,
    )
except Exception:                                    # pragma: no cover
    DocxProject = None
    NotEditable = MissingReferences = Exception
    IMPORT_ERRORS.append(traceback.format_exc(limit=3))

Chooser = SharedStorage = ShareSheet = None
api_version = 0
request_permissions = Permission = None
if platform == "android":
    try:
        from android import api_version
        from android.permissions import Permission, request_permissions
    except Exception:
        IMPORT_ERRORS.append("android: " + traceback.format_exc(limit=2))
    try:
        from androidstorage4kivy import Chooser, ShareSheet, SharedStorage
    except Exception:
        IMPORT_ERRORS.append("storage: " + traceback.format_exc(limit=2))

BLOCKS_PER_PAGE = 14
PREVIEW_CHARS = 320


def panel_view(height_hint=0.94):
    """A rounded pop-up with a dark panel inside. Returns (view, panel)."""
    view = ModalView(
        size_hint=(0.96, height_hint), auto_dismiss=False,
        background="", background_color=(0, 0, 0, 0),
        overlay_color=(0, 0, 0, 0.75),
    )
    panel = BoxLayout(orientation="vertical", padding=dp(16), spacing=dp(8))
    paint(panel, SURFACE, [dp(24)] * 4, edge_hex=LINE)
    view.add_widget(panel)
    return view, panel


def wrapped_label(text, color=TEXT, size=14, height=None):
    label = Label(text=text, color=col(color), font_size=sp(size),
                  halign="left", valign="top", size_hint_y=None,
                  height=dp(height or 24))
    label.bind(width=lambda w, v: setattr(w, "text_size", (v, None)))
    if height is None:
        label.bind(texture_size=lambda w, v: setattr(w, "height", v[1] + dp(4)))
    return label


class BlockCard(ButtonBehavior, BoxLayout):
    """One paragraph / heading / table of the document."""

    def __init__(self, text, kind, on_tap, **kwargs):
        super().__init__(orientation="vertical", size_hint_y=None,
                         height=dp(56), padding=[dp(14), dp(10)], **kwargs)
        self.on_tap = on_tap
        heading = kind.startswith("heading")
        paint(self, SURFACE2 if heading else SURFACE, [dp(14)], edge_hex=LINE)
        self.label = Label(
            text=text, halign="left", valign="top", size_hint_y=None,
            height=dp(30), bold=heading, font_size=sp(15 if heading else 14),
            color=col(ACCENT_TEXT if heading else (TEXT if kind == "paragraph" else MUTED)),
        )
        self.label.bind(width=lambda w, v: setattr(w, "text_size", (v, None)),
                        texture_size=self.fit)
        self.add_widget(self.label)

    def fit(self, label, size):
        label.height = size[1]
        self.height = size[1] + dp(20)

    def on_release(self):
        self.on_tap()


class LabApp(App):

    # ------------------------------------------------------------------
    #  start-up
    # ------------------------------------------------------------------

    def build(self):
        self.project = None
        self.visible = []
        self.page = 0
        self.doc_name = ""
        self.shared_output = None
        self.sharer = None
        self.chooser = None
        self.docs_dir = os.path.join(self.user_data_dir, "docs")
        os.makedirs(self.docs_dir, exist_ok=True)
        self.state_path = os.path.join(self.user_data_dir, "state.json")
        self.config_path = os.path.join(self.user_data_dir, "credentials.json")

        self.root_box = BoxLayout(orientation="vertical")
        self.build_settings()
        self.load_credentials()
        self.show_home()
        return self.root_box

    def on_start(self):
        for message in IMPORT_ERRORS:
            self.home_status.text = "Problem: " + message[-300:]
        if platform == "android":
            if Chooser is not None:
                self.chooser = Chooser(self.on_chosen)
                self.sharer = ShareSheet() if ShareSheet else None
            if request_permissions and api_version and api_version < 29:
                request_permissions([Permission.WRITE_EXTERNAL_STORAGE,
                                     Permission.READ_EXTERNAL_STORAGE])

    def on_pause(self):
        return True

    def on_resume(self):
        pass

    # ------------------------------------------------------------------
    #  settings (Zotero details)
    # ------------------------------------------------------------------

    def build_settings(self):
        view, panel = panel_view(0.62)
        panel.add_widget(small_label("Zotero settings", color=TEXT, size=18,
                                     bold=True, height=32))
        panel.add_widget(wrapped_label(
            "Find your User ID and API key at zotero.org/settings/keys "
            "(read access is enough).", MUTED, 12, 44))
        panel.add_widget(small_label("User ID", bold=True))
        id_box = InputBox(hint="Numbers only")
        self.user_id = id_box.input
        panel.add_widget(id_box)
        panel.add_widget(small_label("API Key", bold=True))
        key_box = InputBox(hint="Paste your API key", password=True)
        self.api_key = key_box.input
        panel.add_widget(key_box)
        self.settings_message = small_label("", color=ERROR, size=12, height=22)
        panel.add_widget(self.settings_message)
        row = BoxLayout(size_hint_y=None, height=dp(52), spacing=dp(10))
        cancel = PillButton(text="Cancel", style="ghost")
        cancel.bind(on_release=lambda *a: view.dismiss())
        save = PillButton(text="Save", style="primary")
        save.bind(on_release=self.save_settings)
        row.add_widget(cancel)
        row.add_widget(save)
        panel.add_widget(row)
        self.settings_view = view

    def open_settings(self, *args):
        self.settings_message.text = ""
        self.settings_view.open()

    def save_settings(self, *args):
        user_id, api_key = self.user_id.text.strip(), self.api_key.text.strip()
        if not user_id.isdigit():
            self.settings_message.text = "User ID must be numbers only."
            return
        if not api_key:
            self.settings_message.text = "Please enter your API key."
            return
        try:
            with open(self.config_path, "w", encoding="utf-8") as file:
                json.dump({"user_id": user_id, "api_key": api_key}, file)
        except OSError:
            pass
        self.settings_view.dismiss()

    def load_credentials(self):
        try:
            with open(self.config_path, "r", encoding="utf-8") as file:
                saved = json.load(file)
            self.user_id.text = saved.get("user_id", "")
            self.api_key.text = saved.get("api_key", "")
        except (OSError, ValueError):
            pass

    def credentials(self):
        user_id, api_key = self.user_id.text.strip(), self.api_key.text.strip()
        if not user_id.isdigit() or not api_key:
            return None
        return user_id, api_key

    # ------------------------------------------------------------------
    #  home
    # ------------------------------------------------------------------

    def show_home(self):
        self.root_box.clear_widgets()
        area = BoxLayout(orientation="vertical", padding=[dp(24), dp(48), dp(24), dp(40)],
                         spacing=dp(14))
        title = Label(
            text="[size=34sp][b]Blaykie[/b][/size]\n[size=14sp][color=#A3A9BC]"
                 "Add Zotero citations to your Word document[/color][/size]",
            markup=True, color=col(TEXT), size_hint_y=None, height=dp(110),
            halign="left", valign="middle")
        title.bind(size=lambda w, v: setattr(w, "text_size", v))
        area.add_widget(title)

        open_btn = PillButton(text="Open a .docx", style="primary",
                              size_hint_y=None, height=dp(56))
        open_btn.bind(on_release=self.choose)
        area.add_widget(open_btn)

        saved = self.read_state()
        if saved:
            resume = PillButton(
                text="Continue: " + os.path.basename(saved["path"])[:28],
                style="secondary", size_hint_y=None, height=dp(56))
            resume.bind(on_release=self.resume)
            area.add_widget(resume)

        settings = PillButton(text="Zotero settings", style="ghost",
                              size_hint_y=None, height=dp(52))
        settings.bind(on_release=self.open_settings)
        area.add_widget(settings)

        self.home_status = wrapped_label("", MUTED, 13)
        area.add_widget(self.home_status)
        area.add_widget(Label())
        self.root_box.add_widget(area)

    # ------------------------------------------------------------------
    #  opening a document
    # ------------------------------------------------------------------

    def choose(self, *args):
        try:
            self.chooser.choose_content("*/*")
        except Exception:
            self.home_status.text = "Could not open the file picker."

    def on_chosen(self, uri_list):
        Clock.schedule_once(lambda dt: self.handle_chosen(uri_list))

    def handle_chosen(self, uri_list):
        try:
            if not uri_list:
                return
            path = SharedStorage().copy_from_shared(uri_list[0])
            if not path:
                self.home_status.text = "Could not read that file."
                return
            self.open_project(path)
        except Exception:
            self.home_status.text = "Could not open it: " + traceback.format_exc(limit=1)[-200:]

    def open_project(self, path):
        name = os.path.basename(path)
        if not name.lower().endswith(".docx"):
            self.home_status.text = "Please choose a .docx file."
            return
        target = os.path.join(self.docs_dir, name)
        if os.path.abspath(path) != os.path.abspath(target):
            shutil.copy(path, target)
        try:
            project = DocxProject(target)
        except Exception:
            self.home_status.text = "This file could not be read as a Word document."
            return
        self.project = project
        self.doc_name = name
        self.shared_output = None
        self.save_state()
        self.start_reading()

    def read_state(self):
        try:
            with open(self.state_path, "r", encoding="utf-8") as file:
                state = json.load(file)
            if os.path.exists(state["path"]):
                return state
        except (OSError, ValueError, KeyError):
            pass
        return None

    def save_state(self):
        if not self.project:
            return
        try:
            with open(self.state_path, "w", encoding="utf-8") as file:
                json.dump({"path": self.project.path, "ops": self.project.ops}, file)
        except OSError:
            pass

    def resume(self, *args):
        state = self.read_state()
        if not state:
            return
        try:
            project = DocxProject(state["path"])
            project.load_ops(state["ops"])
        except Exception:
            self.home_status.text = "Could not restore the last session."
            return
        self.project = project
        self.doc_name = os.path.basename(state["path"])
        self.start_reading()

    def start_reading(self):
        project = self.project
        self.visible = [
            i for i in range(project.block_count)
            if project.block_text(i).strip() or project.block_kind(i) == "table"
        ]
        self.page = 0
        self.show_doc()

    # ------------------------------------------------------------------
    #  reading view
    # ------------------------------------------------------------------

    def show_doc(self):
        root = self.root_box
        root.clear_widgets()

        bar = BoxLayout(size_hint_y=None, height=dp(52), spacing=dp(6),
                        padding=[dp(10), dp(8), dp(10), 0])
        for text, handler, style in (
            ("Close", self.close_doc, "ghost"),
            ("Outline", self.open_outline, "secondary"),
            ("Search", self.open_search, "secondary"),
            ("Undo", self.undo, "secondary"),
            ("Export", self.open_export, "primary"),
        ):
            button = PillButton(text=text, style=style, font_size=sp(12))
            button.bind(on_release=handler)
            bar.add_widget(button)
        root.add_widget(bar)

        self.info_label = small_label("", color=MUTED, size=12, height=26)
        info_row = BoxLayout(size_hint_y=None, height=dp(26), padding=[dp(14), 0])
        info_row.add_widget(self.info_label)
        root.add_widget(info_row)

        scroll = ScrollView(bar_width=dp(3))
        self.cards_box = GridLayout(cols=1, spacing=dp(8), size_hint_y=None,
                                    padding=[dp(10), dp(4), dp(10), dp(12)])
        self.cards_box.bind(minimum_height=self.cards_box.setter("height"))
        scroll.add_widget(self.cards_box)
        self.cards_scroll = scroll
        root.add_widget(scroll)

        pager = BoxLayout(size_hint_y=None, height=dp(92), spacing=dp(10),
                          padding=[dp(14), dp(10), dp(14), dp(30)])
        prev_btn = PillButton(text="Previous", style="secondary", font_size=sp(13))
        prev_btn.bind(on_release=lambda *a: self.go_page(self.page - 1))
        self.page_label = Label(text="", color=col(MUTED), font_size=sp(13),
                                size_hint_x=None, width=dp(90))
        next_btn = PillButton(text="Next", style="secondary", font_size=sp(13))
        next_btn.bind(on_release=lambda *a: self.go_page(self.page + 1))
        pager.add_widget(prev_btn)
        pager.add_widget(self.page_label)
        pager.add_widget(next_btn)
        root.add_widget(pager)
        self.render_page()

    def page_count(self):
        return max(1, -(-len(self.visible) // BLOCKS_PER_PAGE))

    def go_page(self, page):
        self.page = min(max(0, page), self.page_count() - 1)
        self.render_page()
        self.cards_scroll.scroll_y = 1

    def goto_block(self, block):
        if block in self.visible:
            position = self.visible.index(block)
        else:
            position = max((i for i, b in enumerate(self.visible) if b <= block), default=0)
        self.go_page(position // BLOCKS_PER_PAGE)

    def render_page(self):
        project = self.project
        self.cards_box.clear_widgets()
        start = self.page * BLOCKS_PER_PAGE
        for block in self.visible[start:start + BLOCKS_PER_PAGE]:
            kind = project.block_kind(block)
            text = project.block_text(block).strip()
            if kind == "table":
                text = "[Table] " + text
            if len(text) > PREVIEW_CHARS:
                text = text[:PREVIEW_CHARS].rstrip() + " ..."
            card = BlockCard(text, kind, lambda b=block: self.open_block(b))
            self.cards_box.add_widget(card)
        self.page_label.text = f"{self.page + 1} / {self.page_count()}"
        cites = len(project.citations())
        self.info_label.text = f"{self.doc_name[:34]}  |  {cites} citation(s) added"

    def close_doc(self, *args):
        self.save_state()
        self.show_home()

    def undo(self, *args):
        if self.project and self.project.undo():
            self.save_state()
            self.render_page()
        else:
            self.info_label.text = "Nothing to undo."

    # ------------------------------------------------------------------
    #  outline and search
    # ------------------------------------------------------------------

    def list_sheet(self, title, rows, on_pick, with_search=None):
        view, panel = panel_view(0.9)
        panel.add_widget(small_label(title, color=TEXT, size=18, bold=True, height=32))
        holder = {}
        if with_search:
            box = InputBox(hint="Type a word or phrase")
            holder["box"] = box
            search_btn = PillButton(text="Search", style="primary",
                                    size_hint=(None, None), size=(dp(96), dp(52)))
            row = BoxLayout(size_hint_y=None, height=dp(52), spacing=dp(8))
            row.add_widget(box)
            row.add_widget(search_btn)
            panel.add_widget(row)
        scroll = ScrollView(bar_width=dp(3))
        grid = GridLayout(cols=1, spacing=dp(6), size_hint_y=None)
        grid.bind(minimum_height=grid.setter("height"))
        scroll.add_widget(grid)
        panel.add_widget(scroll)
        close = PillButton(text="Close", style="ghost", size_hint_y=None, height=dp(48))
        close.bind(on_release=lambda *a: view.dismiss())
        panel.add_widget(close)

        def fill(items):
            grid.clear_widgets()
            if not items:
                grid.add_widget(wrapped_label("Nothing found.", MUTED, 13))
            for label_text, value in items:
                card = BlockCard(label_text, "paragraph",
                                 lambda v=value: (view.dismiss(), on_pick(v)))
                grid.add_widget(card)

        fill(rows)
        if with_search:
            def run(*args):
                fill(with_search(holder["box"].input.text))
            search_btn.bind(on_release=run)
            holder["box"].input.bind(on_text_validate=run)
        view.open()
        return view

    def open_outline(self, *args):
        rows = []
        for block, level, text in self.project.outline():
            rows.append((("   " * (level - 1)) + text[:90], block))
        self.list_sheet("Outline", rows, self.goto_block)

    def open_search(self, *args):
        def run(query):
            hits = self.project.search(query)
            return [(f"...{snippet}...", block) for block, _at, snippet in hits[:60]]
        self.list_sheet("Search the document", [], self.goto_block, with_search=run)

    # ------------------------------------------------------------------
    #  picking the spot
    # ------------------------------------------------------------------

    def open_block(self, block):
        project = self.project
        if not project.block_editable(block):
            self.info_label.text = "Tables and special blocks cannot be edited."
            return
        view, panel = panel_view(0.94)
        panel.add_widget(small_label(f"Paragraph {block}", color=TEXT, size=18,
                                     bold=True, height=30))
        panel.add_widget(wrapped_label(
            "Tap where the citation should go. To replace a placeholder, "
            "long-press the text and select it.", MUTED, 12, 44))
        field = TextInput(
            text=project.block_text(block), readonly=True, multiline=True,
            background_normal="", background_active="",
            background_color=col(SURFACE2), foreground_color=col(TEXT),
            cursor_color=col(ACCENT_TEXT), selection_color=col("#5B4BEA", 0.45),
            font_size=sp(16), padding=[dp(12), dp(12)],
        )
        panel.add_widget(field)
        self.para_view, self.para_field, self.para_block = view, field, block

        message = small_label("", color=ERROR, size=12, height=22)
        self.para_message = message
        panel.add_widget(message)

        row = BoxLayout(size_hint_y=None, height=dp(52), spacing=dp(10))
        delete_btn = PillButton(text="Delete selection", style="secondary", font_size=sp(13))
        delete_btn.bind(on_release=self.delete_selection)
        insert_btn = PillButton(text="Insert citation", style="primary", font_size=sp(13))
        insert_btn.bind(on_release=self.insert_here)
        row.add_widget(delete_btn)
        row.add_widget(insert_btn)
        panel.add_widget(row)
        close = PillButton(text="Close", style="ghost", size_hint_y=None, height=dp(48))
        close.bind(on_release=lambda *a: view.dismiss())
        panel.add_widget(close)
        view.open()

    def selected_range(self):
        field = self.para_field
        length = len(field.text)
        if field.selection_text:
            a, b = sorted((field.selection_from, field.selection_to))
            return max(0, a), min(length, b)
        index = min(max(0, field.cursor_index()), length)
        return index, index

    def delete_selection(self, *args):
        start, end = self.selected_range()
        if start == end:
            self.para_message.text = "Select the text to delete first."
            return
        try:
            self.project.delete_range(self.para_block, start, end)
        except NotEditable as error:
            self.para_message.text = str(error)[:90]
            return
        self.para_field.text = self.project.block_text(self.para_block)
        self.para_message.text = ""
        self.save_state()
        self.render_page()

    def insert_here(self, *args):
        start, end = self.selected_range()
        self.open_citation_sheet(self.para_block, start, end)

    # ------------------------------------------------------------------
    #  choosing the Zotero reference
    # ------------------------------------------------------------------

    def open_citation_sheet(self, block, start, end):
        if not self.credentials():
            self.para_message.text = "Add your Zotero details first (Home > Zotero settings)."
            return
        view, panel = panel_view(0.96)
        self.cite_view = view
        self.cite_target = (block, start, end)
        self.cite_results = []
        self.cite_selected = []
        self.cite_style = "parenthetical"

        title = "Replace the selected text" if start != end else "Insert citation"
        panel.add_widget(small_label(title, color=TEXT, size=17, bold=True, height=28))

        row = BoxLayout(size_hint_y=None, height=dp(52), spacing=dp(8))
        box = InputBox(hint="Search your Zotero library")
        self.cite_box = box
        go = PillButton(text="Search", style="primary", size_hint=(None, None),
                        size=(dp(96), dp(52)))
        go.bind(on_release=self.cite_search)
        box.input.bind(on_text_validate=self.cite_search)
        row.add_widget(box)
        row.add_widget(go)
        panel.add_widget(row)

        self.cite_status = small_label("", color=MUTED, size=12, height=22)
        panel.add_widget(self.cite_status)

        scroll = ScrollView(bar_width=dp(3))
        self.cite_grid = GridLayout(cols=1, spacing=dp(8), size_hint_y=None)
        self.cite_grid.bind(minimum_height=self.cite_grid.setter("height"))
        scroll.add_widget(self.cite_grid)
        panel.add_widget(scroll)

        style_row = BoxLayout(size_hint_y=None, height=dp(44), spacing=dp(8))
        self.btn_paren = PillButton(text="(Author, 2020)", style="primary", font_size=sp(12))
        self.btn_narr = PillButton(text="Author (2020)", style="secondary", font_size=sp(12))
        self.btn_paren.bind(on_release=lambda *a: self.set_style("parenthetical"))
        self.btn_narr.bind(on_release=lambda *a: self.set_style("narrative"))
        style_row.add_widget(self.btn_paren)
        style_row.add_widget(self.btn_narr)
        panel.add_widget(style_row)

        extra = BoxLayout(size_hint_y=None, height=dp(52), spacing=dp(8))
        self.page_box = InputBox(hint="Page, e.g. 15")
        self.prefix_box = InputBox(hint="Prefix: see")
        for widget in (self.page_box, self.prefix_box):
            widget.input.bind(text=lambda *a: self.update_preview())
            extra.add_widget(widget)
        panel.add_widget(extra)

        self.cite_preview = wrapped_label("", ACCENT_TEXT, 13, 40)
        panel.add_widget(self.cite_preview)
        self.cite_message = small_label("", color=ERROR, size=12, height=22)
        panel.add_widget(self.cite_message)

        buttons = BoxLayout(size_hint_y=None, height=dp(52), spacing=dp(10))
        cancel = PillButton(text="Cancel", style="ghost")
        cancel.bind(on_release=lambda *a: view.dismiss())
        insert = PillButton(text="Insert", style="primary")
        insert.bind(on_release=self.confirm_insert)
        buttons.add_widget(cancel)
        buttons.add_widget(insert)
        panel.add_widget(buttons)
        view.open()

    def set_style(self, style):
        self.cite_style = style
        self.btn_paren.style = "primary" if style == "parenthetical" else "secondary"
        self.btn_narr.style = "primary" if style == "narrative" else "secondary"
        self.btn_paren.redraw()
        self.btn_narr.redraw()
        self.update_preview()

    def run_background(self, work, done, failed):
        def target():
            try:
                result = work()
            except Exception as error:
                message = str(error) or "Something went wrong."
                Clock.schedule_once(lambda dt, m=message: failed(m))
                return
            Clock.schedule_once(lambda dt, r=result: done(r))
        threading.Thread(target=target, daemon=True).start()

    def cite_search(self, *args):
        query = self.cite_box.input.text.strip()
        creds = self.credentials()
        if not query:
            self.cite_status.text = "Type something to search."
            return
        self.cite_status.text = "Searching Zotero..."
        self.cite_grid.clear_widgets()
        self.cite_results, self.cite_selected = [], []
        self.run_background(
            lambda: search_zotero(query, creds[0], creds[1]),
            self.show_cite_results,
            lambda message: setattr(self.cite_status, "text", message))

    def show_cite_results(self, results):
        self.cite_results = results
        self.cite_status.text = f"{len(results)} result(s)" if results else "No references found."
        for number, item in enumerate(results):
            data = item.get("data", {})
            authors = get_authors(data)
            if not authors:
                who = "Unknown author"
            elif len(authors) == 1:
                who = authors[0]["last"]
            elif len(authors) == 2:
                who = f"{authors[0]['last']} & {authors[1]['last']}"
            else:
                who = f"{authors[0]['last']} et al."
            title = (data.get("title") or "[No title]")[:100]
            text = (f"[size=14sp][color=#F3F4F8][b]{escape_markup(who)} \u00b7 "
                    f"{escape_markup(get_year(data))}[/b][/color][/size]\n{escape_markup(title)}")
            card = ResultCard(text, lambda sel, n=number: self.toggle_item(n, sel))
            self.cite_grid.add_widget(card)

    def toggle_item(self, number, selected):
        if selected and number not in self.cite_selected:
            self.cite_selected.append(number)
        if not selected and number in self.cite_selected:
            self.cite_selected.remove(number)
        self.update_preview()

    def chosen_items(self):
        return [self.cite_results[n] for n in sorted(self.cite_selected)
                if n < len(self.cite_results)]

    def update_preview(self):
        items = self.chosen_items()
        if not items:
            self.cite_preview.text = "Select one or more references."
            return
        text = build_citation_text(
            items, self.cite_style, self.page_box.input.text.strip(),
            self.prefix_box.input.text.strip())
        self.cite_preview.text = "Will insert:  " + text

    def confirm_insert(self, *args):
        items = self.chosen_items()
        if not items:
            self.cite_message.text = "Select a reference first."
            return
        block, start, end = self.cite_target
        try:
            self.project.insert_citation(
                block, start, end, items, self.cite_style,
                self.page_box.input.text.strip(), self.prefix_box.input.text.strip())
        except NotEditable as error:
            self.cite_message.text = str(error)[:100]
            return
        self.cite_view.dismiss()
        self.para_view.dismiss()
        self.save_state()
        self.render_page()
        self.info_label.text = "Citation added: " + self.project.citations()[-1][2][:50]

    # ------------------------------------------------------------------
    #  export
    # ------------------------------------------------------------------

    def open_export(self, *args):
        project = self.project
        view, panel = panel_view(0.8)
        self.export_view = view
        panel.add_widget(small_label("Export a new .docx", color=TEXT, size=18,
                                     bold=True, height=32))
        lines = [f"Citations added now: {len(project.citations())}",
                 f"Blaykie references in the file: {len(project.present_keys())}"]
        if project.existing_reference_headings() and not project.has_blaykie_bibliography():
            lines.append("Note: the file already has a References heading. "
                         "Blaykie adds its own list at the end.")
        if project.has_tracked_changes:
            lines.append("Note: the file has tracked changes. Citations are inserted as normal text.")
        if project.has_zotero_fields():
            lines.append("Zotero plugin citations are left untouched.")
        panel.add_widget(wrapped_label("\n".join(lines), TEXT, 13, 120))
        self.export_status = wrapped_label("", ACCENT_TEXT, 13, 80)
        panel.add_widget(self.export_status)
        panel.add_widget(Label())

        row = BoxLayout(size_hint_y=None, height=dp(52), spacing=dp(10))
        self.export_btn = PillButton(text="Export now", style="primary")
        self.export_btn.bind(on_release=self.run_export)
        self.share_btn = PillButton(text="Share / open", style="secondary", disabled=True)
        self.share_btn.bind(on_release=self.share)
        row.add_widget(self.export_btn)
        row.add_widget(self.share_btn)
        panel.add_widget(row)
        close = PillButton(text="Close", style="ghost", size_hint_y=None, height=dp(48))
        close.bind(on_release=lambda *a: view.dismiss())
        panel.add_widget(close)
        view.open()

    def run_export(self, *args, allow_missing=False):
        project = self.project
        missing = project.missing_keys()
        if missing and not allow_missing:
            creds = self.credentials()
            if not creds:
                self.export_status.text = ("Some earlier citations need their references. "
                                           "Add your Zotero details first.")
                return
            self.export_status.text = f"Getting {len(missing)} reference(s) from Zotero..."
            self.run_background(
                lambda: fetch_items_by_keys(missing, creds[0], creds[1]),
                self.after_fetch,
                lambda message: setattr(self.export_status, "text", message))
            return
        self.finish_export(allow_missing)

    def after_fetch(self, items):
        self.project.register_items(items)
        left = self.project.missing_keys()
        if left:
            self.export_status.text = (f"{len(left)} reference(s) were not found in your "
                                       "library. They will be left out of the list.")
            self.finish_export(True)
            return
        self.finish_export(False)

    def finish_export(self, allow_missing):
        try:
            stem = os.path.splitext(self.doc_name)[0]
            out = os.path.join(self.user_data_dir, stem + "_Blaykie.docx")
            self.project.export(out, allow_missing=allow_missing)
            text = f"Exported ({os.path.getsize(out) // 1024} KB)."
            if SharedStorage:
                self.shared_output = SharedStorage().copy_to_shared(out)
                text += " Saved to your Documents/Downloads. Tap Share / open to view it."
                self.share_btn.disabled = False
                self.share_btn.redraw()
            self.export_status.text = text
        except Exception:
            self.export_status.text = "Export failed: " + traceback.format_exc(limit=1)[-160:]

    def share(self, *args):
        try:
            self.sharer.share_file(self.shared_output)
        except Exception:
            self.export_status.text = "Could not open the share sheet."


if __name__ == "__main__":
    LabApp().run()
