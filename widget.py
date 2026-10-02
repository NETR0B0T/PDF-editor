import sys
import os
import tempfile
import pymupdf
import pythoncom
import win32com.client

from PySide6.QtWidgets import QApplication, QFileDialog, QListWidgetItem, QAbstractItemView, QSizePolicy, QMessageBox
from PySide6.QtUiTools import QUiLoader
from PySide6.QtGui import QIcon, QPixmap, QImage, QShortcut, QKeySequence
from PySide6.QtCore import Qt, QSize, QThread, Signal, QObject, QEvent, QFile

def get_resource_path(relative_path):
    """ Visszaadja a fájl útvonalát, működik fejlesztés alatt és lefordított .exe-ként is. """
    if hasattr(sys, '_MEIPASS'):
        # Ha exe-ként fut, a PyInstaller ideiglenes mappájában keresi
        return os.path.join(sys._MEIPASS, relative_path)
    # Ha simán futtatod, a jelenlegi mappában keresi
    return os.path.join(os.path.abspath("."), relative_path)

# ==========================================
# 1. HÁTTÉRSZÁL (KIBŐVÍTVE OFFICE KONVERZIÓVAL)
# ==========================================
class PDFLoaderThread(QThread):
    page_loaded = Signal(QImage, str, int)

    def __init__(self, file_paths):
        super().__init__()
        self.file_paths = file_paths

    def run(self):
        # COM (Windows folyamatok) inicializálása a háttérszálban - KÖTELEZŐ!
        pythoncom.CoInitialize()

        for file_path in self.file_paths:
            try:
                # Alapból feltételezzük, hogy a feldolgozandó fájl maga a bemenet
                process_path = file_path
                ext = os.path.splitext(file_path)[1].lower()

                # Ha Word vagy PowerPoint fájl, konvertáljuk PDF-re a háttérben
                if ext in ['.doc', '.docx', '.ppt', '.pptx']:
                    # Abszolút útvonalak kellenek a Windows Office API-nak
                    abs_input = os.path.abspath(file_path)

                    # Generálunk egy ideiglenes PDF útvonalat a Temp mappába
                    base_name = os.path.basename(file_path) + ".pdf"
                    temp_pdf = os.path.abspath(os.path.join(tempfile.gettempdir(), base_name))

                    if ext in ['.doc', '.docx']:
                        word = win32com.client.Dispatch("Word.Application")
                        doc = word.Documents.Open(abs_input)
                        doc.SaveAs(temp_pdf, FileFormat=17) # 17 = wdFormatPDF
                        doc.Close()

                    elif ext in ['.ppt', '.pptx']:
                        ppt = win32com.client.Dispatch("PowerPoint.Application")
                        presentation = ppt.Presentations.Open(abs_input, WithWindow=False)
                        presentation.SaveAs(temp_pdf, 32) # 32 = ppSaveAsPDF
                        presentation.Close()

                    # Sikeres konverzió után innentől a Temp mappás PDF-et dolgozzuk fel!
                    process_path = temp_pdf

                # Innentől minden megy a régiben (csak a process_path-t használjuk)
                doc = pymupdf.open(process_path)
                for page_num in range(len(doc)):
                    page = doc.load_page(page_num)

                    zoom = pymupdf.Matrix(0.2, 0.2)
                    pix = page.get_pixmap(matrix=zoom)

                    qimg = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)

                    # A háttérbe a process_path (esetlegesen ideiglenes PDF) útvonalát mentjük,
                    # így az előnézet és a végső mentés sima PDF-ként tudja kezelni!
                    self.page_loaded.emit(qimg.copy(), process_path, page_num)

                doc.close()
            except Exception as e:
                print(f"Hiba a fájl feldolgozásakor: {file_path} - {e}")

        # Szál lezárása
        pythoncom.CoUninitialize()


# ==========================================
# 2. DRAG & DROP ESEMÉNYSZŰRŐ (KIBŐVÍTVE)
# ==========================================
class DropFilter(QObject):
    def __init__(self, load_callback):
        super().__init__()
        self.load_callback = load_callback

    def eventFilter(self, watched, event):
        if event.type() == QEvent.Type.DragEnter:
            if event.mimeData().hasUrls():
                event.accept()
                return True

        elif event.type() == QEvent.Type.Drop:
            if event.mimeData().hasUrls():
                urls = event.mimeData().urls()
                # Kiterjesztettük a szűrőt Word és PPT formátumokra
                allowed_exts = ('.pdf', '.doc', '.docx', '.ppt', '.pptx')
                valid_paths = [u.toLocalFile() for u in urls if u.toLocalFile().lower().endswith(allowed_exts)]

                if valid_paths:
                    self.load_callback(valid_paths)
                    event.accept()
                    return True

        return super().eventFilter(watched, event)


# ==========================================
# 3. FŐ ALKALMAZÁS
# ==========================================
class PDFEditorApp:
    def __init__(self):
        loader = QUiLoader()
        # A betöltés mostantól a segédfüggvényen keresztül történik:
        ui_path = get_resource_path("mainwindow.ui")

        ui_file = QFile(ui_path)
        if not ui_file.open(QFile.ReadOnly):
            print(f"HIBA: Nem sikerült megnyitni a UI fájlt: {ui_path}")
            sys.exit(-1)

        loader = QUiLoader()
        self.window = loader.load(ui_file)
        ui_file.close()

        self.window.actionPDF_Hozzaadasa.triggered.connect(self.add_pdf_dialog)
        self.window.actionMentes.triggered.connect(self.save_pdf)
        # Közvetlenül a menüsávhoz adunk egy kattintható "Súgó" gombot
        sugo_gomb = self.window.menuBar().addAction("Súgó")
        sugo_gomb.triggered.connect(self.show_help)

        # Ha beállítottad az előző lépésben a kijelöltek mentését, ne felejtsd el összekötni:
        if hasattr(self.window, 'actionKijeloltek_mentese'):
            self.window.actionKijeloltek_mentese.triggered.connect(self.save_selected_pdf)

        self.window.listWidget.setIconSize(QSize(150, 200))
        self.window.listWidget.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.window.listWidget.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)

        self.delete_shortcut = QShortcut(QKeySequence(Qt.Key_Delete), self.window.listWidget)
        self.delete_shortcut.activated.connect(self.delete_pages)

        self.drop_filter = DropFilter(self.start_loading)
        self.window.listWidget.installEventFilter(self.drop_filter)

        self.active_threads = []

        # Előnézet beállításai
        self.window.lblPreview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.window.lblPreview.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Ignored)
        self.window.lblPreview.setText("Válassz ki egy lapot az előnézethez.")
        self.window.listWidget.itemSelectionChanged.connect(self.update_preview)

    def show(self):
        self.window.show()

    def update_preview(self):
        selected_items = self.window.listWidget.selectedItems()
        if not selected_items:
            self.window.lblPreview.clear()
            self.window.lblPreview.setText("Válassz ki egy lapot az előnézethez.")
            return

        item = selected_items[0]
        file_path, page_num = item.data(Qt.UserRole)

        try:
            doc = pymupdf.open(file_path)
            page = doc.load_page(page_num)
            zoom = pymupdf.Matrix(1.5, 1.5)
            pix = page.get_pixmap(matrix=zoom)

            qimg = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
            pixmap = QPixmap.fromImage(qimg)

            scaled_pixmap = pixmap.scaled(
                self.window.lblPreview.size(),
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation
            )

            self.window.lblPreview.setPixmap(scaled_pixmap)
            doc.close()

        except Exception as e:
            self.window.lblPreview.setText(f"Előnézet hiba: {e}")

    def delete_pages(self):
        for item in self.window.listWidget.selectedItems():
            self.window.listWidget.takeItem(self.window.listWidget.row(item))

    def add_pdf_dialog(self):
        # Bővített szűrő a fájlválasztó ablakban
        file_paths, _ = QFileDialog.getOpenFileNames(
            self.window, "Válassz fájl(oka)t", "",
            "Dokumentumok (*.pdf *.doc *.docx *.ppt *.pptx);;PDF fájlok (*.pdf);;Office fájlok (*.doc *.docx *.ppt *.pptx)"
        )
        if file_paths:
            self.start_loading(file_paths)

    def start_loading(self, file_paths):
        thread = PDFLoaderThread(file_paths)
        thread.page_loaded.connect(self.add_thumbnail_to_list)
        thread.finished.connect(lambda t=thread: self.active_threads.remove(t))
        self.active_threads.append(thread)
        thread.start()

    def add_thumbnail_to_list(self, qimg, file_path, page_num):
        icon = QIcon(QPixmap.fromImage(qimg))
        item = QListWidgetItem(icon, f"Oldal {page_num + 1}")
        item.setData(Qt.UserRole, (file_path, page_num))
        self.window.listWidget.addItem(item)

    def save_pdf(self):
        save_path, _ = QFileDialog.getSaveFileName(
            self.window, "Mentés másként", "", "PDF fájlok (*.pdf)"
        )
        if not save_path:
            return

        self._execute_save(save_path, only_selected=False)

    def save_selected_pdf(self):
        if not self.window.listWidget.selectedItems():
            return

        save_path, _ = QFileDialog.getSaveFileName(
            self.window, "Kijelölt lapok mentése", "", "PDF fájlok (*.pdf)"
        )
        if not save_path:
            return

        self._execute_save(save_path, only_selected=True)

    # Közös mentőfüggvény a kódismétlés elkerülésére
    def _execute_save(self, save_path, only_selected):
        new_doc = pymupdf.open()
        opened_docs = {}

        for i in range(self.window.listWidget.count()):
            item = self.window.listWidget.item(i)

            if only_selected and not item.isSelected():
                continue

            source_path, original_page_num = item.data(Qt.UserRole)

            if source_path not in opened_docs:
                opened_docs[source_path] = pymupdf.open(source_path)

            src_doc = opened_docs[source_path]
            new_doc.insert_pdf(src_doc, from_page=original_page_num, to_page=original_page_num)

        new_doc.save(save_path)
        new_doc.close()

        for doc in opened_docs.values():
            doc.close()
    def show_help(self):
        # Létrehozunk egy információs popup ablakot
        sugo_szoveg = (
            "Kompakt PDF Szerkesztő - Használati Útmutató\n\n"
            "• PDF, Word és PPT fájlok betöltése: Használd a Fájl menüt, vagy "
            "egyszerűen húzd be a fájlokat az ablakba (Drag & Drop).\n\n"
            "• Sorrend cseréje: Fogd meg a lapokat az egérrel, és húzd őket az új helyükre.\n\n"
            "• Lapok törlése: Jelöld ki a törölni kívánt lapokat, és nyomd meg a 'Delete' gombot.\n\n"
            "• Előnézet: Kattints bármelyik lapra, hogy jobb oldalon nagyban lásd a tartalmát.\n\n"
            "• Mentés: A Fájl menüben elmentheted a teljes új dokumentumot, "
            "vagy külön csak a kijelölt lapokat."
        )

        # A QMessageBox.information megjeleníti az ablakot egy (i) ikonnal
        QMessageBox.information(
            self.window,          # Szülő ablak (hogy középen jelenjen meg)
            "Súgó és Funkciók",   # Az ablak címe
            sugo_szoveg           # A tartalom
        )

if __name__ == "__main__":
    app = QApplication(sys.argv)
    editor = PDFEditorApp()
    editor.show()
    sys.exit(app.exec())
