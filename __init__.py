from datetime import datetime
from PyQt5.QtWidgets import *
from qgis.core import *
from qgis.gui import *
from qgis.PyQt.QtGui import QColor, QIcon
from qgis.PyQt.QtCore import QSize, Qt
from qgis.PyQt.QtXml import QDomDocument
import os, shutil, re, math
from . import xlsxwriter
from collections import Counter


'''
Provide the entry point for the qgis plugin.
'''
def classFactory(iface):
    return GeneratePresentation(iface)

'''
Turn a QColor into a color string that TikZ can understand
'''
def color_to_tikz(c):
    return f'{{rgb,255:red,{c.red()};green,{c.green()};blue,{c.blue()}}}'

class Task:
    def __init__(self, run, effort=1, name=''):
        def callback(data, resolve, reject):
            try:
                run(data, resolve, reject)
            except Exception as e:
                raise e
                reject(e)
        self.run = callback
        self.effort = effort
        self.name = name

    @staticmethod
    def synchronous(run, effort=1, name=''):
        def callback(data, resolve, reject):
            try:
                run(data)
                resolve()
            except Exception as e:
                raise e
                reject(e)

        return Task(callback, effort, name)


class TaskQueue:
    IDLE = 0
    RUNNING = 1
    ABORTED = 2

    def __init__(self):
        self.tasks = []
        self.total_effort = 0
        self.progress = 0
        self.status = TaskQueue.IDLE
        self.data = dotdict()

        def noop(*args):
            pass

        self.on_task_complete = noop
        self.on_error = noop

    def notify(self):
        total = self.total_effort
        progress = self.progress / total if total > 0 else 0
        self.on_task_complete(progress)

    def handle_error(self, e):
        self.on_error(e)
        self.abort()

    def add_task(self, run, effort=1, name=''):
        task = Task.synchronous(run, effort, name)
        self.tasks.append(task)
        self.total_effort += effort
        return task

    def add_async_task(self, run, effort=1, name=''):
        task = Task(run, effort, name)
        self.tasks.append(task)
        self.total_effort += effort
        return task

    def update_effort(self, task, effort):
        self.total_effort += effort - task.effort
        task.effort = effort


    def next(self):
        if self.status == TaskQueue.ABORTED:
            return

        if len(self.tasks) == 0:
            self.status = TaskQueue.IDLE
            return

        task = self.tasks.pop(0)
        def callback(*args):
            self.progress += task.effort
            self.notify()
            self.next()

        task.run(self.data, callback, self.handle_error)

    def start(self):
        self.status = TaskQueue.RUNNING
        self.next()

    def abort(self):
        self.status = TaskQueue.ABORTED


'''
Turn a dictionary into one whose attributes can be accessed using dot
notation. Example:

  person = { 'age': 31 }
  print(person['age'])    # 31
  print(person.age)       # 31
'''
class dotdict(dict):
    __getattr__ = dict.get
    __setattr__ = dict.__setitem__
    __delattr__ = dict.__delitem__

'''
QWiget to let the user select a directory. Consists of a QLabel to display
the currently selected directory and a button which opens a selection dialog.
'''
class SelectDirectoryWidget(QWidget):
    def __init__(self, default=''):
        super().__init__()

        # self.path contains the currently selected directory.
        self.path = default

        self.label = QLabel(self.path)
        self.label.setLineWidth(1)
        self.label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)

        button_text = 'auswählen'
        self.button = QPushButton(button_text)
        self.button.clicked.connect(self.selectDirectory)
        width = self.button.fontMetrics().boundingRect(button_text).width() + 7
        self.button.setMaximumWidth(width)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.label)
        layout.addWidget(self.button)

    def selectDirectory(self):
        self.path = QFileDialog.getExistingDirectory(None, 'Ordner auswählen')
        self.label.setText(self.path)


'''
QDialog to ask the user for metadata and layers to evaluate.
'''
class EvaluationDialog(QDialog):

    '''
    * on_accept: Callback function once the dialog is accepted (user clicks OK).
    * data: Dictionary to store the user input. It should already contain the
      following attributes:
      - title: Title of the dialog.
      - destination: Default value for directory selection.
    * text_fields: Text metadata the user should be asked to input. For example,
        { 'age': { 'label': 'Alter', 'value': 31   } }
      would ask the user to enter their age, with a default value of 31. The
      users answer will be written to data.age.
    * layer_fields: Layers the user should be asked to choose. Example:
      {
        'polygons': {
          'label': 'Polygone',
          'required': ['Ort', 'Kreis', 'Bundesland'],
          'renderer': 'categorizedSymbol'
        }
      }
      This would prompt the reader to select a layer with the fields 'Ort',
      'Kreis', 'Bundesland' and a QgsCategorizedSymbolRenderer.
    '''
    def __init__(self, on_accept, data, text_fields={}, layer_fields={}):
        super().__init__()
        self.setWindowTitle(data.title)

        self.on_accept = on_accept
        layout = QFormLayout()

        self.data = data
        self.text_fields = {}
        self.layer_fields = {}
        self.feature_fields = {}

        for key, field in text_fields.items():
            input = QLineEdit(self)
            input.setText(field['value'])
            self.text_fields[key] = input
            layout.addRow(field['label'], input)

        for key, field in layer_fields.items():
            input = QgsMapLayerComboBox(self)
            if 'required' in field:
                renderer = field['renderer'] if 'renderer' in field else ''
                input.setExceptedLayerList(EvaluationDialog.get_exceptions(field['required'], renderer, False))
            else:
                input.setAllowEmptyLayer(True, 'kein Hintergrund')
            if 'default' in field and field['default']:
                input.setLayer(field['default'])

            self.layer_fields[key] = input
            layout.addRow(field['label'], input)

        self.directoryChooser = SelectDirectoryWidget(data.destination)
        layout.addRow('Zielordner:', self.directoryChooser)

        QBtn = (
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )

        buttonBox = QDialogButtonBox(QBtn)
        buttonBox.accepted.connect(self.callback)
        buttonBox.rejected.connect(self.reject)
        layout.addRow(buttonBox)

        self.setLayout(layout)
        self.setFixedWidth(600)

        # show the window
        self.show()

    def callback(self):
        if not self.directoryChooser.path:
            QMessageBox.critical(self, "Error", 'Kein gültiger Zielordner ausgewählt.')
            return

        for key, field in self.text_fields.items():
            self.data[key] = field.text()

        for key, field in self.layer_fields.items():
            self.data[key] = field.currentLayer()
        self.data.destination = self.directoryChooser.path

        self.on_accept(self.data)
        self.accept()

    '''
    Determine which layers do NOT conform to the given fields and renderer
    specification. If the layer is not optional and none conforms to the
    specifications, display an error.
    '''
    @staticmethod
    def get_exceptions(required_fields, renderer, optional):
        def disallowed(layer):
            if layer.type() != QgsMapLayer.VectorLayer:
                return True

            if renderer and layer.renderer().type() != renderer:
                return True

            fields = layer.fields().names()
            for field in required_fields:
                if field not in fields:
                    return True
            return False
        layers = list(QgsProject.instance().mapLayers().values())
        filtered = list(filter(disallowed, layers))

        if not optional and len(layers) == len(filtered):
            required = ', '.join(['"' + f + '"' for f in required_fields])
            r = 'Renderer "' + renderer + '" und' if renderer else ''
            raise RuntimeError(f'Kein Layer mit {r} den folgenden Attributfeldern gefunden: ' + required + '.')

        return filtered

'''
Unified way to store data in a table and export it to LaTeX code or to an Excel
file. Tuples as cell entries indicate a quantity together with a unit. Example:

# Create table with blue table head:
my_table = Table()
my_table.add_row(['Name', 'Alter', ('Gewicht', '')], Table.Highlight.PRIMARY)

# Fill the table
my_table.add_row(['Andreas', 31, (80, 'kg')])
my_table.add_row(['Berta', 56, (61, 'kg')])

# Convert the table to LaTeX with column alignment left, right, right and LaTeX
# command to handle the color.
my_table.to_latex('lrr', 'coloredbullet')

# Convert the table to Excel with column widths 20, 5, 5, 2 (for the 'kg' column)
my_table.to_xlsx(handle_to_xlsx_file, [20, 5, 5, 2])
'''
class Table:
    class Highlight(Enum):
        NONE = 0
        PRIMARY = 1
        SECONDARY = 2

    offset = 0

    def __init__(self):
        self.row_highlight_primary = []
        self.row_highlight_secondary = []
        self.rows = []
        self.row_colors = {}

    '''
    Convert all cells containing a number x to the tuple (x, unit)
    '''
    def numbers_to_unit(self, unit):
        for row in self.rows:
            for i in range(len(row)):
                if isinstance(row[i], int) or isinstance(row[i], float):
                    row[i] = (row[i], unit)

    def set_row_color(self, i, c):
        self.row_colors[i] = c

    '''
    Add a row to the table. highlight may be
    * Table.Highlight.NONE: No highlighting.
    * Table.Highlight.PRIMARY: For table headings.
    * Table.Highlight.SECONDARY: For special intermediate columns.
    * An instance of QColor: The row will be decorated with the corresponding
      color.
    * A tuple (s, t): In the LaTeX export, the row will be decorated with s. In
      the Excel export, the row will be decorated with t.
    '''
    def add_row(self, row, highlight=Highlight.NONE):
        if highlight == Table.Highlight.PRIMARY:
            self.row_highlight_primary.append(len(self.rows))
        elif highlight == Table.Highlight.SECONDARY:
            self.row_highlight_secondary.append(len(self.rows))
        elif isinstance(highlight, QColor) or isinstance(highlight, tuple):
            self.row_colors[len(self.rows)] = highlight
        self.rows.append(row)

    def set_row(self, i, row):
        self.rows[i] = row

    def set_cell(self, i, j, value):
        if i < len(self.rows) and j < len(self.rows[i]):
            self.rows[i][j] = value

    @staticmethod
    def pad_and_add_units(row):
        result = []
        for x in row:
            if isinstance(x, tuple):
                number, unit = x
                if unit == '%':
                    unit = '\\%'
                s = str(number) + '~' + unit
            elif x is None:
                s = ''
            else:
                s = str(x)
            result.append(s)
        return ' & '.join(result)

    def to_latex(self, colspec, color_command):
        primary = ','.join([str(i + 1) for i in self.row_highlight_primary])
        secondary = ','.join([str(i + 1) for i in self.row_highlight_secondary])
        result = '\\begin{tblr}{width=\\textwidth,colspec={' + colspec + '}'
        if len(primary) > 0:
            result += ',row{' + primary + '}={3.5ex,f,bg=dnpblue,fg=white,font=\\bfseries}'
        if len(secondary) > 0:
            result += ',row{' + secondary + '}={3.5ex,f,font=\\bfseries,bg=dnplightblue,fg=black}'
        result += '}'

        for (i, row) in enumerate(self.rows):
            result += '\n    '
            if i in self.row_colors:
                color = self.row_colors[i]
                if isinstance(color, QColor):
                    result += f'\\{color_command}{{{color_to_tikz(color)}}} '
                elif isinstance(color, tuple):
                    result += color[0]

            result += ' & '
            result += Table.pad_and_add_units(row)
            result += ' \\\\'

        result += '\n\\end{tblr}'

        return result

    @staticmethod
    def add_units(row):
        result = []
        for x in row:
            if isinstance(x, tuple):
                number, unit = x
                result.append(number)
                result.append(unit)
            elif x is None:
                result.append('')
                result.append('')
            else:
                result.append(x)
        return result

    def to_xlsx(self, workbook, column_widths=[], worksheet=None):
        offset = Table.offset

        if not worksheet:
            worksheet = workbook.add_worksheet()

        # set column widths
        for i, w in enumerate(column_widths):
            worksheet.set_column(i, i, w)

        primary = workbook.add_format()
        primary.set_bold()
        primary.set_bg_color('#001aae') # DNP blue
        primary.set_font_color('white')

        secondary = workbook.add_format()
        secondary.set_bold()
        secondary.set_bg_color('#dde2ff')

        bg_white = workbook.add_format()
        bg_white.set_bg_color('white')

        for i in range(0, 100):
            for j in range(0, 20):
                worksheet.write(i + offset, j, '', bg_white)

        for (i, row) in enumerate(self.rows):
            fmt = bg_white
            if i in self.row_highlight_secondary:
                fmt = secondary
            if i in self.row_highlight_primary:
                fmt = primary

            if i in self.row_colors:
                color = self.row_colors[i]
                if isinstance(color, QColor):
                    color_fmt = workbook.add_format()
                    color_fmt.set_bg_color(self.row_colors[i].name())
                    worksheet.write(i + offset, 0, '', color_fmt)
                elif isinstance(color, tuple):
                    worksheet.write(i + offset, 0, color[1], fmt)
            else:
                worksheet.write(i + offset, 0, '', fmt)

            worksheet.write_row(i + offset, 1, Table.add_units(row), fmt)

        Table.offset += len(self.rows) + 1


class SymbologyCategory:
    def __init__(self, token, color, label, value):
        self.token = token
        self.color = color
        self.label = label
        self.value = value

    @staticmethod
    def extract_symbology_categories(layer, field, columns):
        result = []
        for c in layer.renderer().categories():
            token = c.value()
            color = c.symbol().color()
            label = c.label()
            match = re.search(r'\(.*\) (.*)', label)
            if match:
                label = match.group(1)

            condition = f'"{field}" = \'{token}\''
            value = [GeneratePresentation.filtered_column_sum(layer, condition, column) for column in columns]
            result.append(SymbologyCategory(token, color, label, value))

        return result


class RectangleMapTool(QgsMapTool):
    def __init__(self, canvas, action):
        self.canvas = canvas
        self.action = action
        QgsMapTool.__init__(self, self.canvas)
        self.rubberBand = QgsRubberBand(self.canvas, QgsWkbTypes.PolygonGeometry)
        self.rubberBand.setStrokeColor(Qt.red)
        self.rubberBand.setWidth(1)
        self.reset()

    def reset(self):
        self.startPoint = self.endPoint = None
        self.isEmittingPoint = False
        self.rubberBand.reset(QgsWkbTypes.PolygonGeometry)

    def canvasPressEvent(self, e):
        self.startPoint = self.toMapCoordinates(e.pos())
        self.endPoint = self.startPoint
        self.isEmittingPoint = True
        self.showRect(self.startPoint, self.endPoint)


    def canvasReleaseEvent(self, e):
        self.isEmittingPoint = False
        r = self.rectangle()

        if r is not None:
            self.reset()
            self.action(r)

    def canvasMoveEvent(self, e):
        if not self.isEmittingPoint:
            return

        self.endPoint = self.toMapCoordinates(e.pos())
        self.showRect(self.startPoint, self.endPoint)


    def showRect(self, startPoint, endPoint):
        self.rubberBand.reset(QgsWkbTypes.PolygonGeometry)
        if startPoint.x() == endPoint.x() or startPoint.y() == endPoint.y():
            return

        point1 = QgsPointXY(startPoint.x(), startPoint.y())
        point2 = QgsPointXY(startPoint.x(), endPoint.y())
        point3 = QgsPointXY(endPoint.x(), endPoint.y())
        point4 = QgsPointXY(endPoint.x(), startPoint.y())

        self.rubberBand.addPoint(point1, False)
        self.rubberBand.addPoint(point2, False)
        self.rubberBand.addPoint(point3, False)
        self.rubberBand.addPoint(point4, True)    # true to update canvas
        self.rubberBand.show()

    def rectangle(self):
        if self.startPoint is None or self.endPoint is None:
            return None
        elif (self.startPoint.x() == self.endPoint.x() or \
            self.startPoint.y() == self.endPoint.y()):
            return None
        return QgsRectangle(self.startPoint, self.endPoint)

    def deactivate(self):
        QgsMapTool.deactivate(self)
        self.reset()
        self.deactivated.emit()


class GeneratePresentation:
    def __init__(self, iface):
        self.iface = iface
        self.image_width = 1150
        self.image_height = 800
        self.destination_directory = os.path.expanduser("~")
        self.dir_path = os.path.dirname(os.path.realpath(__file__))
        self.progress = None

    def initGui(self):
        presIcon = QIcon(os.path.join(self.dir_path, 'file-easel.png'))
        cameraIcon = QIcon(os.path.join(self.dir_path, 'image.png'))
        rectIcon = QIcon(os.path.join(self.dir_path, 'rectangle.png'))

        menu = QMenu()

        self.trenches_action = QAction(presIcon, 'Präsentation zu Adressen und Trenches erzeugen', menu)
        self.trenches_action.triggered.connect(self.evaluate_trenches)
        menu.addAction(self.trenches_action)
        self.iface.registerMainWindowAction(self.trenches_action, 'Ctrl+Alt+U')

        self.surfaces_action = QAction(presIcon, 'Präsentation zu Oberflächenanalyse erzeugen', menu)
        self.surfaces_action.triggered.connect(self.evaluate_surfaces)
        menu.addAction(self.surfaces_action)
        self.iface.registerMainWindowAction(self.surfaces_action, 'Ctrl+Alt+O')

        self.make_pic_action = QAction(cameraIcon, 'Bild mit Polygonmaßen', menu)
        self.make_pic_action.triggered.connect(self.attempt(self.make_pic_user))
        menu.addAction(self.make_pic_action)
        self.iface.registerMainWindowAction(self.make_pic_action, None)

        self.select_rectangle_action = QAction(rectIcon, 'Bild mit benutzerdefinierten Maßen', menu)
        self.select_rectangle_action.triggered.connect(self.select_rectangle)
        menu.addAction(self.select_rectangle_action)
        self.iface.registerMainWindowAction(self.select_rectangle_action, None)

        toolButton = QToolButton()
        toolButton.setMenu(menu)
        toolButton.setDefaultAction(QAction(presIcon, 'Auswertungstools'))
        toolButton.setPopupMode(QToolButton.InstantPopup)

        self.tool_action = self.iface.addToolBarWidget(toolButton)

    def unload(self):
        self.iface.removeToolBarIcon(self.tool_action)
        del self.tool_action
        self.iface.unregisterMainWindowAction(self.trenches_action)
        del self.trenches_action
        self.iface.unregisterMainWindowAction(self.surfaces_action)
        del self.surfaces_action
        self.iface.unregisterMainWindowAction(self.make_pic_action)
        del self.make_pic_action
        self.iface.unregisterMainWindowAction(self.select_rectangle_action)
        del self.select_rectangle_action

    def attempt(self, fn):
        def inner(*args):
            try:
                fn(*args)
            except Exception as e:
                self.iface.messageBar().clearWidgets()
                self.iface.messageBar().pushMessage("Error", str(e), level=Qgis.Critical)
        return inner

    def select_rectangle(self):
        rectangleTool = RectangleMapTool(self.iface.mapCanvas(), self.make_pic_user)
        self.iface.mapCanvas().setMapTool(rectangleTool)

    def init_progress_bar(self, maximum):
        message_bar = self.iface.messageBar()
        message_bar.clearWidgets()
        progressMessageBar = message_bar.createMessage("Präsentation wird generiert ...")
        self.progress = QProgressBar()
        self.progress.setMaximum(maximum)
        self.progress.setAlignment(Qt.AlignLeft|Qt.AlignVCenter)
        progressMessageBar.layout().addWidget(self.progress)
        message_bar.pushWidget(progressMessageBar, Qgis.Info)

    def print_error(self, e):
        self.iface.messageBar().clearWidgets()
        self.iface.messageBar().pushMessage("Error", str(e), level=Qgis.Critical)

    def show_success(self, data):
        dst = data.destination
        self.iface.messageBar().clearWidgets()
        self.iface.messageBar().pushMessage(
            "Success",
            "Presentation prepared in <a href=\"file:///" + dst + "\">" + dst + "</a>.",
            level=Qgis.MessageLevel.Success,
            duration=15
        )

    def set_progress(self, value):
        if not self.progress:
            return
        self.progress.setValue(round(value * 100))

    def increment_progess(self, increment=1):
        if not self.progress:
            return
        self.progress.setValue(self.progress.value() + increment)

    @staticmethod
    def require_layer_gracious(key):
        layers = QgsProject.instance().mapLayers()
        for layer in layers.values():
            if key in layer.name():
                return layer

        return None

    @staticmethod
    def require_layer(key):
        layers = QgsProject.instance().mapLayers()
        result = None
        for layer in layers.values():
            if key in layer.name():
                if result:
                    raise RuntimeError('Multiple layers containing the key "' + key + '" found! Make sure that there is only one.')
                else:
                    result = layer

        if not result:
            raise RuntimeError('No layer containing the key "' + key + '" found!')

        return result

    def copy_template(self, subfolder, destination):
        source = os.path.join(self.dir_path, "template", subfolder)
        shutil.copytree(source, destination, dirs_exist_ok=True)

    @staticmethod
    def add_rule(root_rule, expression, color, stroke_color = None, width = None):
        rule = root_rule.children()[0].clone()
        rule.setFilterExpression(expression)

        symbol = rule.symbol()
        symbol.setColor(color)

        marker = symbol.symbolLayer(0)
        if marker and marker.type() == Qgis.SymbolType.Marker:
            marker.setStrokeColor(stroke_color if stroke_color else color)
            marker.setStrokeWidth(width)
        elif width:
            symbol.setWidth(width)

        root_rule.appendChild(rule)

    @staticmethod
    def style_layer(layer, rules):
        layer_ = layer.clone()
        symbol = QgsSymbol.defaultSymbol(layer_.geometryType())
        renderer = QgsRuleBasedRenderer(symbol)

        root_rule = renderer.rootRule()
        for (expression, color, stroke_color, width) in rules:
            GeneratePresentation.add_rule(root_rule, expression, color, stroke_color, width)
        root_rule.removeChildAt(0)
        layer_.setRenderer(renderer)
        layer_.triggerRepaint()
        return layer_

    @staticmethod
    def adjust_line_spacing(layer, rule_label):
        renderer = layer.renderer()
        root = renderer.rootRule()
        for rule in root.children():
            if rule.label() == rule_label:
                symbol = rule.symbol()
                layers = symbol.symbolLayers()
                if len(layers) == 0 or not isinstance(layers[0], QgsLinePatternFillSymbolLayer):
                    continue
                layers[0].setDistance(layers[0].distance() * 2)

    def make_pic_user(self, extent=None):
        default_file = os.path.join(self.destination_directory, "Karten", "map.pdf")
        destination = QFileDialog.getSaveFileName(
            None, "Save currently checked layers as PDF",
            default_file, "Portable Document Format (*.pdf)"
        )

        if not extent:
            extent = self.calculate_extent([self.iface.activeLayer().boundingBoxOfSelected()])

        if destination and destination[0]:
            self.make_pic_pdf(self.iface.mapCanvas().layers(), destination[0], extent)
            self.iface.messageBar().pushMessage(
                "Success",
                "Picture saved to <a href=\"file:///" + destination[0] + "\">" + destination[0] + "</a>.",
                level=Qgis.MessageLevel.Success,
                duration=15
            )

    def make_pic_pdf(self, layers, destination, extent=None, zoom_factor=4):
        project = QgsProject.instance()
        layout = QgsPrintLayout(project)
        layout.initializeDefaults()

        width = self.image_width / zoom_factor
        height = self.image_height / zoom_factor
        size = QgsLayoutSize(width, height)
        pc = layout.pageCollection()
        pc.pages()[0].setPageSize(size)

        if not extent:
            extent = self.calculate_extent()

        map = QgsLayoutItemMap(layout)
        map.setRect(0, 0, width, height)
        map.zoomToExtent(extent)
        map.setExtent(extent)
        map.setLayers(layers)
        map.setBackgroundColor(QColor(255, 255, 255, 0))
        layout.addLayoutItem(map)

        exporter = QgsLayoutExporter(layout)
        exporter.exportToPdf(destination, QgsLayoutExporter.PdfExportSettings())

    def make_pic_png(self, layers, destination):
        settings = QgsMapSettings()
        settings.setLayers(layers)
        settings.setBackgroundColor(QColor(255, 255, 255))
        settings.setOutputSize(QSize(self.image_width, self.image_height))
        settings.setExtent(self.calculate_extent())
        render = QgsMapRendererParallelJob(settings)

        def finished():
            img = render.renderedImage()
            img.save(destination, "png")
        render.finished.connect(finished)

        # Start the rendering
        render.start()

    @staticmethod
    def features_within_selection(layer, selection):
        ids = set()
        for polygon in selection:
            request = QgsFeatureRequest().setDistanceWithin(polygon.geometry(), 0).setFlags(QgsFeatureRequest.NoGeometry).setSubsetOfAttributes([])
            ids.update([f.id() for f in layer.getFeatures(request)])
        copy = layer.materialize(QgsFeatureRequest().setFilterFids(list(ids)))
        copy.setRenderer(layer.renderer().clone())
        return copy

    @staticmethod
    def filtered_column_sum(layer, condition, column):
        values = QgsVectorLayerUtils.getValues(layer, f'CASE WHEN {condition} THEN {column} ELSE 0 END')[0]
        return sum([x if x else 0 for x in values])

    @staticmethod
    def filtered_length_sum(layer, condition):
        return math.ceil(GeneratePresentation.filtered_column_sum(layer, condition, '$length'))

    @staticmethod
    def calculate_address_statistics(data):
        layer = data.addresses
        destination = data.destination
        columns = ['1', '"Total Kunde"', '"Total DNP"', '1']
        categories = SymbologyCategory.extract_symbology_categories(layer, 'Pruefung', columns)

        special_tokens = ['n', 'o']
        normal_categories = list(filter(lambda c: c.token not in special_tokens, categories))
        special_categories = list(filter(lambda c: c.token in special_tokens, categories))

        table = Table()
        table.add_row([
            'Adresskulisse', '', '', '', ''
        ], Table.Highlight.PRIMARY)
        table.add_row([
            '', 'Adressen', 'Einheiten \\Kunde', 'Einheiten DNP', 'Differenz'
        ], Table.Highlight.PRIMARY)

        for c in normal_categories:
            c.value[3] = c.value[2] - c.value[1]
            table.add_row([c.label] + c.value, c.color)

        totalResult = []
        for column in range(len(columns)):
            totalResult.append(sum([c.value[column] for c in normal_categories]))
        table.add_row(['Gesamt'] + totalResult, Table.Highlight.SECONDARY)

        # if present, write results for special categories
        if len(special_categories) > 0:
            table.add_row(['', '', '', '', ''])
            for c in special_categories:
                table.add_row([c.label, c.value[0], '', '', ''], c.color)

        with open(os.path.join(destination, "Praesentation", "AdressStatistik.tex"), "w") as f:
            f.write('\\newcommand\\adressStatistik{' + table.to_latex('l@{}l|rrrr', 'colordot') + '}')

        table.set_cell(1, 2, 'Einheiten Kunde')
        total_offset = len(normal_categories) + 2
        table.set_cell(total_offset, 1, f'=SUM(C3:C{total_offset})')
        table.set_cell(total_offset, 2, f'=SUM(D3:D{total_offset})')
        table.set_cell(total_offset, 3, f'=SUM(E3:E{total_offset})')
        table.set_cell(total_offset, 4, f'=SUM(F3:F{total_offset})')

        for i in range(3, 3 + len(normal_categories)):
            table.set_cell(i - 1, 4, f'=E{i}-D{i}')

        workbook = xlsxwriter.Workbook(os.path.join(destination, "Adressauswertung.xlsx"))
        table.to_xlsx(workbook, [2, 25, 15, 15, 15, 15])
        workbook.close()

    @staticmethod
    def calculate_trench_lengths(data):
        layer = data.trenches
        destination = data.destination

        conditions = ['"Belag" = \'a\'', '"Belag" = \'t\'', '"Belag" = \'g\'', '"Belag" = \'m\'',  '"Belag" = \'k\'']
        columns = ['1', '0', '"In_Strasse"', '"Handschachtung"', '"Privatweg"']

        result = []
        for condition in conditions:
            row = []
            for column in columns:
                row.append(GeneratePresentation.filtered_length_sum(layer, f'{condition} and {column}'))
            result.append(row)

        offener_tiefbau = [sum([row[col] for row in result]) for col in range(len(columns))]
        rohrpressung = GeneratePresentation.filtered_length_sum(layer, '"Belag" = \'c\' and "Verfahren" = \'r\'')
        rohrpressung_privat = GeneratePresentation.filtered_length_sum(layer, '"Belag" = \'c\' and "Verfahren" = \'r\' and "Privatweg"')
        spuelbohrung = GeneratePresentation.filtered_length_sum(layer, '"Belag" = \'c\' and "Verfahren" = \'h\'')
        spuelbohrung_privat = GeneratePresentation.filtered_length_sum(layer, '"Belag" = \'c\' and "Verfahren" = \'h\' and "Privatweg"')
        geschlossener_tiefbau = [rohrpressung + spuelbohrung, None, None, None, rohrpressung_privat + spuelbohrung_privat]

        special_crossings = Counter(filter(None, QgsVectorLayerUtils.getValues(layer, '"Sonderquerung"')[0]))

        total = offener_tiefbau[0] + geschlossener_tiefbau[0]
        trench_table = Table()
        trench_table.add_row([
            'Tiefbau gesamt', total, 0, ('im Straßenkörper', ''), ('mit Handschachtung', ''), ('in Privatweg', '')
        ], Table.Highlight.PRIMARY)
        trench_table.add_row(['Offener Tiefbau'] + offener_tiefbau, Table.Highlight.SECONDARY)

        trench_table.add_row(['Asphalt'] + result[0], QColor('#db1e2a'))
        trench_table.add_row(['Pflaster'] + result[1], QColor('#487bb6'))
        trench_table.add_row(['Unbefestigt'] + result[2], QColor('#54b04a'))
        trench_table.add_row(['Mosaikpflaster'] + result[3], QColor('#873bde'))
        trench_table.add_row(['Kopfsteinpflaster'] + result[4], QColor('#00b0f0'))

        trench_table.add_row(['Geschlossener Tiefbau'] + geschlossener_tiefbau, Table.Highlight.SECONDARY)
        trench_table.add_row(['Rohrpressung', rohrpressung, None, None, None, rohrpressung_privat], QColor('#ffba0b'))
        trench_table.add_row(['Spülbohrung',  spuelbohrung, None, None, None, spuelbohrung_privat], QColor('#01ffe1'))

        for row in trench_table.rows:
            row[2] = (round(row[1] * 100 / total), '%')

        if special_crossings.total() == 0:
            trench_table.add_row(['Sonderquerungen', None, None, None, None, None], Table.Highlight.SECONDARY)
            trench_table.add_row(['keine', None, None, None, None, None])
        else:
            trench_table.add_row(['Sonderquerungen', (special_crossings.total(), 'St.'), None, None, None, None], Table.Highlight.SECONDARY)
            for (crossing, count) in special_crossings.most_common():
                trench_table.add_row([crossing, (count, 'St.'), None, None, None, None])

        trench_table.numbers_to_unit('m')

        with open(os.path.join(destination, "Praesentation", "TrenchStatistik.tex"), "w") as f:
            f.write('\\newcommand\\trenchStatistik{' + trench_table.to_latex('l@{}l|rr|rrr', 'colorrule') + '}')

        trench_table.set_cell(1, 1, (f'=SUM(C3:C7)', 'm'))
        trench_table.set_cell(1, 3, (f'=SUM(G3:G7)', 'm'))
        trench_table.set_cell(1, 4, (f'=SUM(I3:I7)', 'm'))
        trench_table.set_cell(1, 5, (f'=SUM(K3:K7)', 'm'))

        trench_table.set_cell(7, 1, ('=SUM(C9:C10)', 'm'))
        trench_table.set_cell(7, 5, ('=SUM(K9:K10)', 'm'))
        trench_table.set_cell(0, 1, ('=C2+C8', 'm'))
        if special_crossings.total() > 0:
            trench_table.set_cell(10, 1, (f'=SUM(C12:C{12+len(special_crossings)})', 'St.'))

        workbook = xlsxwriter.Workbook(os.path.join(destination, "Trenches.xlsx"))
        trench_table.to_xlsx(workbook, [5, 25, 15, 2, 10, 2, 15, 2, 15, 2, 15, 2])
        workbook.close()

    def calculate_extent(self, references = []):
        if len(references) == 0:
            references = [self.iface.mapCanvas().extent()]

        xmin = min([r.xMinimum() for r in references])
        ymin = min([r.yMinimum() for r in references])
        xmax = min([r.xMaximum() for r in references])
        ymax = min([r.yMaximum() for r in references])

        width = xmax - xmin
        height = ymax - ymin
        ratio = self.image_width / self.image_height

        if (width / height) > ratio:
            # extent is very wide, need to pad on the top and bottom
            desired_height = width / ratio
            diff = (desired_height - height) / 2
            extent = QgsRectangle(xmin, ymin - diff, xmax, ymax + diff)
        else:
            # extent is very high, need to pad on the left and right
            desired_width = height * ratio
            diff = (desired_width - width) / 2
            extent = QgsRectangle(xmin - diff, ymin, xmax + diff, ymax)

        extent.scale(1.2)
        return extent

    def rectangle_around_point(self, point, width=100, height=100):
        rect = QgsRectangle(
            point.x() - width/2,
            point.y() - height/2,
            point.x() + width/2,
            point.y() + height/2
        )
        return self.calculate_extent([rect])

    def process_points_of_interest(self, data, layers):
        points = data.points_of_interest
        extent = data.extent
        dst = data.destination

        max_id = max([point["Punkt_ID"] for point in points] + [6])
        x_coords = [0] * max_id
        y_coords = [0] * max_id
        for point in points:
            geometry = point.geometry()
            if geometry.type() != Qgis.GeometryType.Point:
                continue

            pt = geometry.asPoint()
            id = int(point["Punkt_ID"])
            x_coords[id-1] = (pt.x() - extent.xMinimum()) / extent.width()
            y_coords[id-1] = 1 - (pt.y() - extent.yMinimum()) / extent.height()

            rect = self.rectangle_around_point(pt)
            path = os.path.join(dst, "Bilder", f"fotopunkt{id}.pdf")
            self.make_pic_pdf(layers, path, rect, zoom_factor=20)

        with open(os.path.join(dst, "Praesentation", "PointsOfInterest.tex"), "w") as f:
            x_coords_str = ''.join(['{' + str(x) + '}' for x in x_coords])
            y_coords_str = ''.join(['{' + str(y) + '}' for y in y_coords])
            f.write('\\storedata\\xcoords{' + x_coords_str + '}\n')
            f.write('\\storedata\\ycoords{' + y_coords_str + '}')

    @staticmethod
    def get_selection_fields(layer, fields):
        if layer.type() != QgsMapLayer.VectorLayer:
            raise RuntimeError('Kein Vektor-Layer ausgewählt.')

        features = list(layer.getSelectedFeatures())
        if len(features) == 0:
            raise RuntimeError('Keine Polygone im aktiven Layer ausgewählt.')

        layer_field_names = layer.fields().names()
        for name in fields:
            if name not in layer_field_names:
                raise RuntimeError(f'Die ausgewählten Polygone haben kein Attributfeld "{name}".')

        return features

    def write_metadata(data):
        with open(os.path.join(data.destination, "Praesentation", "Commands.tex"), "a") as f:
            f.write('\n\n% Ort\n\\newcommand{\\Ort}{' + data.ort + '}\n')
            f.write('\n% Landkreis\n\\newcommand{\\Kreis}{' + data.kreis + '}\n')
            f.write('\n% Bundesland\n\\newcommand{\\Land}{' + data.land + '}\n')
            f.write('\n% Abgabedatum\n\\newcommand{\\Abgabedatum}{' + data.datum + '}\n')
            f.write('\n% Kunde (z.B. GF+ oder FED)\n\\newcommand{\\Kunde}{' + data.kunde + '}\n')

    def evaluate_trenches(self, *args):
        self.progress = None

        q = TaskQueue()
        q.on_task_complete = self.set_progress
        q.on_error = self.print_error

        q.add_async_task(self.show_trenches_dialog, name='Show trenches dialog')

        def copy_template(data):
            Table.offset = 0
            self.init_progress_bar(100)
            data.poi = GeneratePresentation.features_within_selection(data.poi, data.selection)
            data.addresses = GeneratePresentation.features_within_selection(data.addresses, data.selection)
            data.trenches = GeneratePresentation.features_within_selection(data.trenches, data.selection)
            data.polygons = GeneratePresentation.features_within_selection(data.polygons, data.selection)

            data.points_of_interest = list(data.poi.getFeatures())
            q.update_effort(poi_task, 1 + len(data.points_of_interest))
            data.maps_dir = os.path.join(data.destination, "Karten")

            self.destination_directory = data.destination
            self.copy_template("common", data.destination)
            self.copy_template("address_and_trenches", data.destination)

        q.add_task(copy_template, name='Copy template')
        q.add_task(GeneratePresentation.write_metadata, name='Write metadata')
        q.add_task(GeneratePresentation.calculate_address_statistics, name='Calculate address statistics')
        q.add_task(GeneratePresentation.calculate_trench_lengths, name='Calculate trench lengths')

        poi_task = q.add_task(
            lambda data: self.process_points_of_interest(data, [data.addresses, data.trenches, data.background]),
            name='Process points of interest'
        )

        def make_title_pic(data):
            titlepic_path = os.path.join(data.destination, "Bilder", "titelbild.pdf")
            layers = [data.poi, data.addresses, data.polygons, data.background]
            self.make_pic_pdf(layers, titlepic_path, data.extent)
        q.add_task(make_title_pic, name='Make title pic')

        def make_address_map(data):
            address_check_path = os.path.join(data.maps_dir, "adresscheck.pdf")
            self.make_pic_pdf([data.addresses, data.polygons, data.background], address_check_path, data.extent)
        q.add_task(make_address_map, name='Print address map')

        def make_hp_distribution(data):
            hp_distribution_path = os.path.join(data.maps_dir, "hp-verteilung.pdf")
            color1 = QColor(72, 123, 182)
            color2 = QColor(228, 187, 114)
            color3 = QColor(84, 174, 74)
            hp_distribution = GeneratePresentation.style_layer(data.addresses, [
                ('"Total DNP" > 12', color1, color1.darker(), 0.3),
                ('"Total DNP" > 2 and "Total DNP" <= 12', color2, color2.darker(), 0.3),
                ('"Total DNP" <= 2 and "Total DNP" is not null', color3, color3.darker(), 0.3)
            ])
            self.make_pic_pdf([hp_distribution, data.polygons, data.background], hp_distribution_path, data.extent)
            self.increment_progess()
        q.add_task(make_hp_distribution, name='Print HP distribution')

        def make_trenches_map(data):
            trenches_path = os.path.join(data.maps_dir, "trenches.pdf")
            self.make_pic_pdf([data.trenches, data.polygons, data.background], trenches_path, data.extent)
        q.add_task(make_trenches_map, name='Print trenches map')

        def make_trench_detail_maps(data):
            maps_dir = data.maps_dir
            by_hands_path = os.path.join(maps_dir, "trenches-handschachtung.pdf")
            by_hands = GeneratePresentation.style_layer(data.trenches, [
                ('"Handschachtung" = false', QColor('black'), None, 0.3),
                ('"Handschachtung" = true', QColor('#54b04a'), None, 0.7)
            ])
            self.make_pic_pdf([by_hands, data.polygons, data.background], by_hands_path, data.extent)

            by_streets_path = os.path.join(maps_dir, "trenches-strassenkoerper.pdf")
            by_streets = GeneratePresentation.style_layer(data.trenches, [
                ('"In_Strasse" = false', QColor('black'), None, 0.3),
                ('"In_Strasse" = true', QColor('#db1e2a'), None, 0.7)
            ])
            self.make_pic_pdf([by_streets, data.polygons, data.background], by_streets_path, data.extent)

            by_private_path = os.path.join(maps_dir, "trenches-privatweg.pdf")
            by_private = GeneratePresentation.style_layer(data.trenches, [
                ('"Privatweg" = false', QColor('black'), None, 0.3),
                ('"Privatweg" = true', QColor('#487bb6'), None, 0.7)
            ])
            self.make_pic_pdf([by_private, data.polygons, data.background], by_private_path, data.extent)
        q.add_task(make_trench_detail_maps, name='Print trenches maps')

        q.add_task(self.show_success, name='Show success')

        q.start()

    def show_trenches_dialog(self, data, resolve, reject):
        self.iface.messageBar().clearWidgets()
        osm = GeneratePresentation.require_layer_gracious('OpenStreetMap')

        layer = self.iface.activeLayer()
        selection = GeneratePresentation.get_selection_fields(layer, ['Name DNP', 'Kreis', 'Bundesland'])

        ort = selection[0]["Name DNP"]
        kreis = selection[0]["Kreis"]
        land = selection[0]["Bundesland"]
        datum = datetime.today().strftime('%d.%m.%Y')

        data.extent = self.calculate_extent([layer.boundingBoxOfSelected()])
        data.destination = self.destination_directory
        data.selection = selection
        data.title = 'Adressen und Trenches auswerten'

        self.dialog = EvaluationDialog(
            resolve,
            data,
            {
                'ort': { 'label': 'Ort:', 'value': ort },
                'kreis': { 'label': 'Kreis:', 'value': kreis },
                'land': { 'label': 'Bundesland:', 'value': land },
                'datum': { 'label': 'Abgabedatum:', 'value': datum },
                'kunde': { 'label': 'Kunde:', 'value': '' },
            },
            {
                'poi': { 'label': 'Fotopunkt:', 'required': ['Punkt_ID'] },
                'addresses': {
                    'label': 'Adressen:',
                    'required': ['Total Kunde', 'Total DNP'],
                    'renderer': 'categorizedSymbol'
                },
                'trenches': {
                    'label': 'Trenches:',
                    'required': ['Belag', 'In_Strasse', 'Handschachtung', 'Privatweg', 'Verfahren'],
                    'renderer': 'RuleRenderer'
                },
                'polygons': { 'label': 'Polygone:', 'required': ['Name DNP', 'Kreis', 'Bundesland'] },
                'background': { 'label': 'Hintergrund:', 'default': osm },
            }
        )

    @staticmethod
    def remove_layer_attributes(layer, field_names):
        attributes = [layer.fields().indexFromName(name) for name in field_names]
        attributes = list(filter(lambda i: i > -1, attributes))
        layer.dataProvider().deleteAttributes(attributes)
        layer.updateFields()

    @staticmethod
    def export_layer(layer, name, path, mode='w'):
        context = QgsProject.instance().transformContext()
        options = QgsVectorFileWriter.SaveVectorOptions()

        if mode == 'a':
            # do not overwrite the file but append to it
            options.actionOnExistingFile = QgsVectorFileWriter.CreateOrOverwriteLayer

        options.layerName = name
        options.fileEncoding = layer.dataProvider().encoding()
        options.driverName = "GPKG"
        QgsVectorFileWriter.writeAsVectorFormatV3(layer, path, context, options)
        doc = QDomDocument();
        readWriteContext = context = QgsReadWriteContext()
        layer.exportNamedStyle(doc);
        gpkg_layer = QgsVectorLayer(f"{path}|layername={name}", name, "ogr")
        gpkg_layer.importNamedStyle(doc)
        gpkg_layer.saveStyleToDatabase(name, "", True, "")

    def export_surfaces_gpkg(self, data):
        path = os.path.join(data.destination, data.ort + '_Oberflächenanalyse_vorschlag.gpkg')
        GeneratePresentation.export_layer(data.poi, 'Fotopunkt', path)

        # before exporting the surfaces, remove the "noch zu klassifizieren" rule:
        renderer = data.surfaces.renderer()
        root = renderer.rootRule()
        for rule in root.children():
            if rule.label() == 'noch zu klassifizieren':
                root.removeChild(rule)

        GeneratePresentation.remove_layer_attributes(data.surfaces, ['Typ', 'Area', 'Polygon'])
        GeneratePresentation.export_layer(data.surfaces, 'Oberflächen', path, 'a')

    def export_trenches(self, data):
        path = os.path.join(data.destination, data.ort + '_vorschlag.gpkg')
        GeneratePresentation.export_layer(data.poi, 'Fotopunkt', path)

        renderer = data.addresses.renderer()
        for category in renderer.categories():
            if 'Prüfung ausstehend' in category.label():
                index = renderer.categoryIndexForLabel(category.label())
                if index >= 0:
                    renderer.deleteCategory(index)
        GeneratePresentation.remove_layer_attributes(data.addresses, ['Polygon ID', 'Nicht sichtbar'])
        # TODO: There's more to clean up
        GeneratePresentation.export_layer(data.addresses, 'Adressen', path, 'a')

        renderer = data.trenches.renderer()
        root = renderer.rootRule()
        for rule in root.children():
            if rule.label() == 'Nicht klassifiziert':
                root.removeChild(rule)
        GeneratePresentation.export_layer(data.trenches, 'Trenches', path, 'a')


    def show_surfaces_dialog(self, data, resolve, reject):
        self.iface.messageBar().clearWidgets()
        osm = GeneratePresentation.require_layer_gracious('OpenStreetMap')

        layer = self.iface.activeLayer()
        selection = GeneratePresentation.get_selection_fields(layer, ['Name DNP', 'Kreis', 'Bundesland', 'Strassenmeter'])

        ort = selection[0]["Name DNP"]
        kreis = selection[0]["Kreis"]
        land = selection[0]["Bundesland"]
        datum = datetime.today().strftime('%d.%m.%Y')

        data.extent = self.calculate_extent([layer.boundingBoxOfSelected()])
        data.destination = self.destination_directory
        data.selection = selection
        data.title = 'Oberflächenklassifikation auswerten'

        self.dialog = EvaluationDialog(
            resolve,
            data,
            {
                'ort': { 'label': 'Ort:', 'value': ort },
                'kreis': { 'label': 'Kreis:', 'value': kreis },
                'land': { 'label': 'Bundesland:', 'value': land },
                'datum': { 'label': 'Abgabedatum:', 'value': datum },
                'kunde': { 'label': 'Kunde:', 'value': '' },
                'number_special': { 'label': 'Anzahl Sonderquerungen:', 'value': '0' },
            },
            {
                'poi': { 'label': 'Fotopunkt:', 'required': ['Punkt_ID'] },
                'surfaces': {
                    'label': 'Oberflächenklassifikation:',
                    'required': ['Belag', 'Typ']
                },
                'polygons': { 'label': 'Polygone:', 'required': ['Name DNP', 'Kreis', 'Bundesland', 'Strassenmeter'] },
                'background': { 'label': 'Hintergrund:', 'default': osm },
            }
        )

    @staticmethod
    def calculate_surface_statistics(data):
        layer = data.surfaces
        area_to_length = '$area / (CASE WHEN "Typ" = \'b\' THEN 1.281 ELSE 5.787 END)'

        class CategoryGroup:
            surface_types = []

            def __init__(self, title):
                self.table = Table()
                self.total_meters = 0
                self.total_numbers = 0

                self.table.add_row(
                    [title, ('Länge', ''), ('Anteil', '')],
                    Table.Highlight.PRIMARY
                )

            def add_surface_type(self, condition, label, color):
                meters = GeneratePresentation.filtered_column_sum(layer, condition, area_to_length)
                meters = math.ceil(meters)
                self.total_meters += meters

                color = color.lighter() # create a pseudo-transparency effect
                self.table.add_row([label, meters, None], color)
                CategoryGroup.surface_types.append((label, color))

            def add_total(self, label='Gesamt'):
                if self.total_meters > 0:
                    for row in self.table.rows[1:]:
                        row[2] = round(row[1] * 100 / self.total_meters)

                self.table.add_row(
                    [label, self.total_meters, 100],
                    Table.Highlight.SECONDARY
                )
                return self.total_meters

            def cleanup(self):
                for row in self.table.rows:
                    if isinstance(row[1], int) or isinstance(row[1], float):
                        row[1] = (row[1], 'm')
                    if isinstance(row[2], int) or isinstance(row[2], float):
                        row[2] = (row[2], '%')

            def to_latex(self):
                colspec = 'l@{}X[l]rr'
                return self.table.to_latex(colspec, 'colorsquare')

        sidewalk = CategoryGroup('Oberflächen Bürgersteig')
        sidewalk.add_surface_type('"Belag" = \'a\' OR ("Belag" = \'sa\' AND "Typ" = \'b\')', 'Asphalt', QColor('#fa182a'))
        sidewalk.add_surface_type('"Belag" = \'b\' OR ("Belag" = \'sb\' AND "Typ" = \'b\')', 'Betonplatten', QColor('#90000c'))
        sidewalk.add_surface_type('"Belag" = \'t\' OR ("Belag" = \'st\' AND "Typ" = \'b\')', 'Pflaster', QColor('#100bb3'))
        sidewalk.add_surface_type('"Belag" = \'g\' OR ("Belag" = \'sg\' AND "Typ" = \'b\')', 'Unbefestigt', QColor('#28c028'))
        sidewalk.add_surface_type('"Belag" = \'v\' OR ("Belag" = \'sv\' AND "Typ" = \'b\')', 'Verdichtet/Schotter', QColor('#066c06'))
        sidewalk.add_surface_type('"Belag" = \'m\' OR ("Belag" = \'sm\' AND "Typ" = \'b\')', 'Kopfsteinpflaster', QColor('#ff7f00'))
        sidewalk.add_surface_type('"Belag" = \'n\' OR ("Belag" = \'sn\' AND "Typ" = \'b\')', 'kein Bürgersteig', QColor('#959595'))
        sidewalk_total = sidewalk.add_total()
        sidewalk.cleanup()

        street = CategoryGroup('Oberflächen Straße')
        street.add_surface_type('"Belag" = \'sa\' AND "Typ" = \'s\'', 'Asphaltierte Straße', QColor('#ebd407'))
        street.add_surface_type('"Belag" = \'sb\' AND "Typ" = \'s\'', 'Straße mit Betonplatten', QColor('#948306'))
        street.add_surface_type('"Belag" = \'st\' AND "Typ" = \'s\'', 'Gepflasterte Straße', QColor('#2fffee'))
        street.add_surface_type('"Belag" = \'sg\' AND "Typ" = \'s\'', 'Unbefestigte Straße', QColor('#becf50'))
        street.add_surface_type('"Belag" = \'sm\' AND "Typ" = \'s\'', 'Straße mit Kopfsteinpflaster', QColor('#87650f'))
        street_total = street.add_total()
        polygons_total = round(sum([f['Strassenmeter'] for f in data.selection]))
        street.table.add_row(['nicht BIS-geeignet', polygons_total - street_total, None], ('$\\times$', '✖'))
        street.cleanup()

        special = CategoryGroup('Sonderpositionen')
        handschachtung = GeneratePresentation.filtered_column_sum(
                layer, '"Handschachtung"', area_to_length)
        special.table.add_row(['Handschachtung', round(handschachtung), None], ('\\hatchedsquare', '▨'))
        traglast = GeneratePresentation.filtered_column_sum(
                layer, '"Typ" = \'b\' AND "Belag" LIKE \'s%\'', area_to_length)
        special.table.add_row(['Bürgersteig mit besonderer Traglastanforderung', round(traglast), None])
        special.cleanup()

        special_crossing = CategoryGroup(f'Sonderquerungen ({data.number_special} St.)')
        special_crossing.add_surface_type('"Belag" = \'x\' OR ("Belag" = \'sx\' AND "Typ" = \'b\')', 'Sonderquerung Bürgersteig', QColor('#9a50cf'))
        special_crossing.add_surface_type('"Belag" = \'sx\' AND "Typ" = \'s\'', 'Sonderquerung Straße', QColor('#8300d4'))
        special_crossing.cleanup()

        summary = CategoryGroup('Gesamtoberfläche')
        summary.table.add_row(['Bürgersteig', sidewalk_total, 0])
        summary.table.add_row(['Straße', street_total, 0])
        summary.total_meters = sidewalk_total + street_total
        summary.add_total()
        summary.cleanup()

        with open(os.path.join(data.destination, "Praesentation", "OberflaechenStatistik.tex"), "w") as f:
            f.write('\\newcommand\\surfacetypes{')
            for (label, color) in CategoryGroup.surface_types:
                f.write('\\item[\\colorsquare{' + color_to_tikz(color) + '}] ' + label + '\n')
            f.write('\\item[\\hatchedsquare] Handschachtung \n')
            f.write('}\n')

            f.write('\\newcommand\\oberflaechenBuergersteig{' + sidewalk.to_latex() + '}\n')
            f.write('\\newcommand\\oberflaechenStrasse{' + street.to_latex() + '}\n')
            f.write('\\newcommand\\oberflaechenSonderposition{' + special.to_latex() + '}\n')
            f.write('\\newcommand\\oberflaechenSonderquerung{' + special_crossing.to_latex() + '}\n')
            f.write('\\newcommand\\oberflaechenGesamt{' + summary.to_latex() + '}\n')

            per_meter = 0.21
            total_price = polygons_total * per_meter
            f.write('\n\\newcommand\\preisliste{')
            f.write(
                '\\item[$\\bullet$] Analysierte Straßenmeter mit einem Preis von ${}~\\text{{\\euro}}$ pro Meter\n'.format(per_meter)
            )
            f.write(
                '\\item[$\\bullet$] Kosten für \\Ort: ${}~\\text{{m}} \\times {}~\\text{{\\euro{{}} pro m}} = {}~\\text{{\\euro}}$\n'.format(polygons_total, per_meter, total_price)
            )
            f.write('\\begin{itemize}\n')
            for polygon in data.selection:
                f.write('   \\item[$\\bullet$] \\Ort{} -- ' + polygon['Name DNP'] + ': ' + str(round(polygon['Strassenmeter'])) + '~m\n')

            f.write('\\end{itemize}\n')
            f.write('%\\item[$\\bullet$] Nicht rechnungsrelevant: $XXX~\\text{m} \\times 0.21~\\text{\\euro{} pro m} = XXX~\\text{\\euro}$\n')
            f.write('}\n')

        Table.offset = 0
        workbook = xlsxwriter.Workbook(os.path.join(data.destination, "Oberflaechenanalyse.xlsx"))
        worksheet = workbook.add_worksheet()
        sidewalk.table.to_xlsx(workbook, [2, 50, 10, 2, 10, 2, 10], worksheet=worksheet)
        street.table.to_xlsx(workbook, worksheet=worksheet)
        special.table.to_xlsx(workbook, worksheet=worksheet)
        special_crossing.table.to_xlsx(workbook, worksheet=worksheet)
        summary.table.to_xlsx(workbook, worksheet=worksheet)
        workbook.close()

    def evaluate_surfaces(self, *args):
        self.progress = None

        q = TaskQueue()
        q.on_task_complete = self.set_progress
        q.on_error = self.print_error

        q.add_async_task(self.show_surfaces_dialog, name='Show surfaces dialog')

        def copy_template(data):
            self.init_progress_bar(100)

            data.poi = GeneratePresentation.features_within_selection(data.poi, data.selection)
            data.surfaces = GeneratePresentation.features_within_selection(data.surfaces, data.selection)
            data.polygons = GeneratePresentation.features_within_selection(data.polygons, data.selection)

            data.points_of_interest = list(data.poi.getFeatures())
            q.update_effort(poi_task, 1 + len(data.points_of_interest))

            self.destination_directory = data.destination
            self.copy_template("common", data.destination)
            self.copy_template("surface_classification", data.destination)

        q.add_task(copy_template, name='Copy template')
        q.add_task(GeneratePresentation.write_metadata, name='Write metadata')
        q.add_task(GeneratePresentation.calculate_surface_statistics, name='Calculate surface statistics')

        poi_task = q.add_task(
            lambda data: self.process_points_of_interest(data, [data.surfaces, data.background]),
            name='Process points of interest'
        )

        def make_title_pic(data):
            titlepic_path = os.path.join(data.destination, "Bilder", "titelbild.pdf")
            layers = [data.poi, data.surfaces, data.polygons, data.background]
            self.make_pic_pdf(layers, titlepic_path, data.extent)
        q.add_task(make_title_pic, name='Make title pic')

        def make_map(data):
            map_path = os.path.join(data.destination, "Karten", "karte.pdf")
            layers = [data.surfaces, data.polygons, data.background]
            self.make_pic_pdf(layers, map_path, data.extent)
        q.add_task(make_map, name='Print map')

        q.add_task(self.show_success, name='Show success')
        q.start()
