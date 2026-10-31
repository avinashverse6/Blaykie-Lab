"""Blaykie Lab: phone smoke test for the DOCX workflow.

1. choose a .docx with the Android file picker (no storage permission needed)
2. read it with Blaykie's DOCX core
3. add a TEST citation + reference (dummy data)
4. export a NEW .docx and save it to shared storage
5. open the share sheet so you can open it in Word / Docs / WPS

Every step is written to the log on screen. Copy the log and send it if
anything fails.
"""

import os
import traceback

from kivy.app import App
from kivy.clock import Clock
from kivy.core.clipboard import Clipboard
from kivy.core.window import Window
from kivy.metrics import dp, sp
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView
from kivy.utils import platform

Window.clearcolor = (0.055, 0.063, 0.086, 1)
Window.softinput_mode = "below_target"

IMPORT_ERRORS = []

try:
    from blaykie_docx import DocxProject, NotEditable
except Exception:                                   # pragma: no cover
    DocxProject = None
    NotEditable = Exception
    IMPORT_ERRORS.append("blaykie_docx: " + traceback.format_exc(limit=2))

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
        IMPORT_ERRORS.append("androidstorage4kivy: " + traceback.format_exc(limit=2))

DUMMY_ITEM = {
    "key": "LABTEST1",
    "data": {
        "itemType": "journalArticle",
        "title": "A test reference added by Blaykie Lab",
        "date": "2020",
        "creators": [{"creatorType": "author", "firstName": "Blaykie",
                      "lastName": "Test"}],
        "publicationTitle": "Journal of Smoke Tests",
        "volume": "1", "issue": "1", "pages": "1-2",
    },
}


class LabApp(App):

    def build(self):
        self.project = None
        self.private_output = None
        self.shared_output = None
        self.test_spot = None
        self.sharer = None
        self.lines = []

        root = BoxLayout(orientation="vertical", padding=dp(14), spacing=dp(8))
        root.add_widget(Label(
            text="Blaykie Lab: DOCX test", font_size=sp(20), bold=True,
            size_hint_y=None, height=dp(40)))

        for text, handler in (
            ("1. Choose a .docx file", self.choose),
            ("2. Add a TEST citation", self.add_test),
            ("3. Export a copy", self.export),
            ("4. Share / open the copy", self.share),
            ("Copy log", self.copy_log),
        ):
            button = Button(text=text, size_hint_y=None, height=dp(46),
                            background_normal="", background_color=(0.36, 0.29, 0.92, 1)
                            if text[0].isdigit() else (0.17, 0.2, 0.27, 1))
            button.bind(on_release=handler)
            root.add_widget(button)

        scroll = ScrollView()
        self.log_label = Label(
            text="", font_size=sp(12), halign="left", valign="top",
            size_hint_y=None, color=(0.9, 0.92, 0.97, 1))
        self.log_label.bind(
            width=lambda w, v: setattr(w, "text_size", (v, None)),
            texture_size=lambda w, v: setattr(w, "height", v[1] + dp(20)))
        scroll.add_widget(self.log_label)
        root.add_widget(scroll)
        return root

    def on_start(self):
        self.log(f"platform={platform}  android api={api_version}")
        for message in IMPORT_ERRORS:
            self.log("IMPORT PROBLEM " + message)
        if platform == "android":
            if Chooser is not None:
                self.chooser = Chooser(self.on_chosen)
                self.sharer = ShareSheet() if ShareSheet else None
                self.log("file picker ready")
            if request_permissions and api_version and api_version < 29:
                request_permissions([Permission.WRITE_EXTERNAL_STORAGE,
                                     Permission.READ_EXTERNAL_STORAGE])
        else:
            self.log("Not on Android: the file picker only works on the phone.")

    def on_pause(self):
        return True

    def on_resume(self):
        pass

    # ----- logging -----

    def log(self, message):
        self.lines.append(str(message))
        self.log_label.text = "\n".join(self.lines)
        print("LAB:", message)

    def copy_log(self, *args):
        Clipboard.copy("\n".join(self.lines))
        self.log("(log copied to clipboard)")

    # ----- step 1 -----

    def choose(self, *args):
        try:
            self.chooser.choose_content("*/*")
        except Exception:
            self.log("PICKER ERROR " + traceback.format_exc(limit=3))

    def on_chosen(self, uri_list):
        Clock.schedule_once(lambda dt: self.handle_chosen(uri_list))

    def handle_chosen(self, uri_list):
        try:
            if not uri_list:
                self.log("No file chosen.")
                return
            uri = uri_list[0]
            self.log(f"chosen: {uri}")
            path = SharedStorage().copy_from_shared(uri)
            if not path:
                self.log("Could not copy the file into the app.")
                return
            self.open_docx(path)
        except Exception:
            self.log("CHOOSE ERROR " + traceback.format_exc(limit=4))

    def open_docx(self, path):
        size_kb = os.path.getsize(path) / 1024
        self.log(f"copied into app: {os.path.basename(path)} ({size_kb:.0f} KB)")
        if DocxProject is None:
            self.log("DOCX core is missing.")
            return
        project = DocxProject(path)
        self.project = project
        kinds = {}
        for index in range(project.block_count):
            kind = project.block_kind(index)
            kinds[kind] = kinds.get(kind, 0) + 1
        self.log(f"blocks: {project.block_count}  {kinds}")
        outline = project.outline()
        self.log(f"headings found: {len(outline)}")
        for index, level, text in outline[:6]:
            self.log(f"   [{index}] H{level}: {text[:50]}")
        self.log(f"tracked changes: {project.has_tracked_changes}   "
                 f"Zotero plugin fields: {project.has_zotero_fields()}   "
                 f"existing References heading: {bool(project.existing_reference_headings())}")
        hits = project.search("the")
        self.log(f'search test: "the" found {len(hits)} times')

        self.test_spot = None
        for index in range(project.block_count):
            text = project.block_text(index)
            if project.block_editable(index) and len(text) >= 60:
                self.test_spot = (index, len(text.rstrip(".")))
                break
        if self.test_spot:
            self.log(f"test spot: paragraph {self.test_spot[0]}, before its last full stop")
        else:
            self.log("No suitable paragraph found for the test.")

    # ----- step 2 -----

    def add_test(self, *args):
        try:
            if not self.project or not self.test_spot:
                self.log("Choose a document first (step 1).")
                return
            block, position = self.test_spot
            number = self.project.insert_citation(block, position, position, [DUMMY_ITEM])
            self.log(f"citation #{number} added: {self.project.citations()[-1][2]}")
            self.log("text now: ..." + self.project.block_text(block)[max(0, position - 40):position + 30])
            self.test_spot = (block, position)
        except NotEditable as error:
            self.log(f"Not allowed here: {error}")
        except Exception:
            self.log("ADD ERROR " + traceback.format_exc(limit=4))

    # ----- step 3 -----

    def export(self, *args):
        try:
            if not self.project:
                self.log("Choose a document first (step 1).")
                return
            out = os.path.join(self.user_data_dir, "Blaykie_test_output.docx")
            self.project.export(out)
            self.private_output = out
            self.log(f"exported in app: {os.path.getsize(out) / 1024:.0f} KB")
            if SharedStorage:
                self.shared_output = SharedStorage().copy_to_shared(out)
                self.log(f"saved to shared storage: {self.shared_output}")
        except Exception:
            self.log("EXPORT ERROR " + traceback.format_exc(limit=4))

    # ----- step 4 -----

    def share(self, *args):
        try:
            if not self.shared_output:
                self.log("Export first (step 3).")
                return
            self.sharer.share_file(self.shared_output)
            self.log("share sheet opened")
        except Exception:
            self.log("SHARE ERROR " + traceback.format_exc(limit=4))


if __name__ == "__main__":
    LabApp().run()
