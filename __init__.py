from PyQt5.QtWidgets import QAction, QFileDialog, QProgressBar, QMenu, QToolButton
from qgis.core import *
from qgis.gui import *
from qgis.PyQt.QtGui import QColor, QIcon
from qgis.PyQt.QtCore import QSize, Qt
import os, shutil, re, math
from . import xlsxwriter
from collections import Counter


def classFactory(iface):
    return GeneratePresentation(iface)

def color_to_tikz(c):
    return f'{{rgb,255:red,{c.red()};green,{c.green()};blue,{c.blue()}}}'

class Table:
    class Highlight(Enum):
        NONE = 0
        PRIMARY = 1
        SECONDARY = 2

    def __init__(self):
        self.row_highlight_primary = []
        self.row_highlight_secondary = []
        self.rows = []
        self.row_colors = {}

    def numbers_to_unit(self, unit):
        for row in self.rows:
            for i in range(len(row)):
                if isinstance(row[i], int) or isinstance(row[i], float):
                    row[i] = (row[i], unit)

    def set_row_color(self, i, c):
        self.row_colors[i] = c

    def add_row(self, row, highlight=0):
        if highlight == Table.Highlight.PRIMARY:
            self.row_highlight_primary.append(len(self.rows))
        elif highlight == Table.Highlight.SECONDARY:
            self.row_highlight_secondary.append(len(self.rows))
        elif isinstance(highlight, QColor):
            self.row_colors[len(self.rows)] = highlight
        self.rows.append(row)
    
    def set_row(self, i, row):
        self.rows[i] = row

    def set_cell(self, i, j, value):
        if i < len(self.rows) and j < len(self.rows[i]):
            self.rows[i][j] = value
    
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
        result = '\\begin{tblr}{width=\\textwidth,colspec={' + colspec + '},row{'
        result += primary + '}={bg=dnpblue,fg=white,font=\\bfseries},row{'
        result += secondary + '}={font=\\bfseries,bg=dnplightblue,fg=black}}'

        for (i, row) in enumerate(self.rows):
            result += '\n    '
            if i in self.row_colors:
                color = self.row_colors[i]
                result += f'\\{color_command}{{{color_to_tikz(color)}}} '
            
            result += ' & '
            result += Table.pad_and_add_units(row)
            result += ' \\\\'
        
        result += '\n\\end{tblr}'
        
        return result

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

    def to_xlsx(self, workbook, column_widths=[], worksheet=None, offset=0):
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
                color_fmt = workbook.add_format()
                color_fmt.set_bg_color(self.row_colors[i].name())
                worksheet.write(i + offset, 0, '', color_fmt)
            else:
                worksheet.write(i + offset, 0, '', fmt)

            worksheet.write_row(i + offset, 1, Table.add_units(row), fmt)


class SymbologyCategory:
    def __init__(self, token, color, label, value):
        self.token = token
        self.color = color
        self.label = label
        self.value = value
    
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

    def initGui(self):
        presIcon = QIcon(os.path.join(self.dir_path, 'file-easel.png'))
        cameraIcon = QIcon(os.path.join(self.dir_path, 'image.png'))
        rectIcon = QIcon(os.path.join(self.dir_path, 'rectangle.png'))

        menu = QMenu()

        init_template_action = QAction(presIcon, 'Präsentation zu Adressen und Trenches erzeugen', menu)
        init_template_action.triggered.connect(self.attempt_instantiate_address_trench_template)
        menu.addAction(init_template_action)

        surfaces_action = QAction(presIcon, 'Präsentation zu Oberflächenanalyse erzeugen', menu)
        surfaces_action.triggered.connect(self.attempt_instantiate_surface_template)
        menu.addAction(surfaces_action)

        make_pic_action = QAction(cameraIcon, 'Bild mit Polygonmaßen', menu)
        make_pic_action.triggered.connect(self.attempt_make_pic_user)
        menu.addAction(make_pic_action)

        select_rectangle_action = QAction(rectIcon, 'Bild mit benutzerdefinierten Maßen', menu)
        select_rectangle_action.triggered.connect(self.select_rectangle)
        menu.addAction(select_rectangle_action)

        toolButton = QToolButton()
        toolButton.setMenu(menu)
        toolButton.setDefaultAction(QAction(presIcon, 'Auswertungstools'))
        toolButton.setPopupMode(QToolButton.InstantPopup)

        self.tool_action = self.iface.addToolBarWidget(toolButton)

    def unload(self):
        self.iface.removeToolBarIcon(self.tool_action)
        del self.tool_action
    
    def select_rectangle(self):
        rectangleTool = RectangleMapTool(self.iface.mapCanvas(), self.attempt_make_pic_user)
        self.iface.mapCanvas().setMapTool(rectangleTool)
    
    def init_progress_bar(self, maximum):
        message_bar = self.iface.messageBar()
        message_bar.clearWidgets()
        progressMessageBar = message_bar.createMessage("Preparing presentation ...")
        self.progress = QProgressBar()
        self.progress.setMaximum(maximum)
        self.progress.setAlignment(Qt.AlignLeft|Qt.AlignVCenter)
        progressMessageBar.layout().addWidget(self.progress)
        message_bar.pushWidget(progressMessageBar, Qgis.Info)

    def increment_progess(self, increment=1):
        if not self.progress:
            return
        self.progress.setValue(self.progress.value() + increment)

    def require_layer_gracious(key):
        layers = QgsProject.instance().mapLayers()
        for layer in layers.values():
            if key in layer.name():
                return layer
        
        return None

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

    def make_pic_user(self, extent=None):
        default_file = os.path.join(self.destination_directory, "Karten", "map.pdf")
        destination = QFileDialog.getSaveFileName(
            None, "Save currently checked layers as PDF",
            default_file, "Portable Document Format (*.pdf)"
        )

        if destination and destination[0]:
            self.make_pic_pdf(self.iface.mapCanvas().layers(), destination[0], extent)
            self.iface.messageBar().pushMessage(
                "Success",
                "Picture saved to <a href=\"file:///" + destination[0] + "\">" + destination[0] + "</a>.",
                level=Qgis.MessageLevel.Success,
                duration=15
            )

    def attempt_make_pic_user(self, rect=None):
        try:
            self.make_pic_user(rect)
        except Exception as e:
            self.iface.messageBar().clearWidgets()
            self.iface.messageBar().pushMessage("Error", str(e), level=Qgis.Critical)

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
            addresses = GeneratePresentation.require_layer_gracious("Adressen")
            trenches = GeneratePresentation.require_layer_gracious("Trenches")
            oberflaechen = GeneratePresentation.require_layer_gracious("Oberflächenanalyse")
            references = [l.extent() for l in filter(None, [addresses, trenches, oberflaechen])]
            extent = self.calculate_extent(references)

        map = QgsLayoutItemMap(layout)
        map.setRect(0, 0, width, height)
        map.zoomToExtent(extent)
        map.setExtent(extent)
        map.setLayers(layers)
        map.setBackgroundColor(QColor(255, 255, 255, 0))
        layout.addLayoutItem(map)

        # scaleBar = QgsLayoutItemScaleBar(layout)
        # scaleBar.setLinkedMap(map)
        # scaleBar.applyDefaultSize()
        # layout.addLayoutItem(scaleBar)

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

    def filtered_column_sum(layer, condition, column):
        values = QgsVectorLayerUtils.getValues(layer, f'CASE WHEN {condition} THEN {column} ELSE 0 END')[0]
        return sum([x if x else 0 for x in values])

    def filtered_length_sum(layer, condition):
        return math.ceil(GeneratePresentation.filtered_column_sum(layer, condition, '$length'))

    def calculate_address_statistics(layer, destination):
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

    def calculate_trench_lengths(layer, destination):
        conditions = ['"Belag" = \'a\'', '"Belag" = \'t\'', '"Belag" = \'g\'', '"Belag" = \'m\'',  '"Belag" = \'k\'']
        columns = ['true', '"In_Strasse"', '"Handschachtung"', '"Privatweg"']

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
        geschlossener_tiefbau = [rohrpressung + spuelbohrung, None, None, rohrpressung_privat + spuelbohrung_privat]

        special_crossings = Counter(filter(None, QgsVectorLayerUtils.getValues(layer, '"Sonderquerung"')[0]))
        
        trench_table = Table()
        trench_table.add_row([
            'Tiefbau gesamt', (offener_tiefbau[0] + geschlossener_tiefbau[0], 'm'), None, None, None
        ], Table.Highlight.PRIMARY)
        trench_table.add_row([
            '', None, ('im Straßenkörper', ''), ('mit Handschachtung', ''), ('in Privatweg', '')
        ], Table.Highlight.PRIMARY)
        trench_table.add_row(['Offener Tiefbau'] + offener_tiefbau, Table.Highlight.SECONDARY)

        trench_table.add_row(['Asphalt'] + result[0], QColor('#db1e2a'))
        trench_table.add_row(['Pflaster'] + result[1], QColor('#487bb6'))
        trench_table.add_row(['Unbefestigt'] + result[2], QColor('#54b04a'))
        trench_table.add_row(['Mosaikpflaster'] + result[3], QColor('#873bde'))
        trench_table.add_row(['Kopfsteinpflaster'] + result[4], QColor('#00b0f0'))

        trench_table.add_row(['Geschlossener Tiefbau'] + geschlossener_tiefbau, Table.Highlight.SECONDARY)
        trench_table.add_row(['Rohrpressung', rohrpressung, None, None, rohrpressung_privat], QColor('#ffba0b'))
        trench_table.add_row(['Spülbohrung',  spuelbohrung, None, None, spuelbohrung_privat], QColor('#01ffe1'))

        if special_crossings.total() == 0:
            trench_table.add_row(['Sonderquerungen', None, None, None, None], Table.Highlight.SECONDARY)
            trench_table.add_row(['keine', None, None, None, None])
        else:
            trench_table.add_row(['Sonderquerungen', (special_crossings.total(), 'St.'), None, None, None], Table.Highlight.SECONDARY)
            for (crossing, count) in special_crossings.most_common():
                trench_table.add_row([crossing, (count, 'St.'), None, None, None])

        trench_table.numbers_to_unit('m')

        with open(os.path.join(destination, "Praesentation", "TrenchStatistik.tex"), "w") as f:
            f.write('\\newcommand\\trenchStatistik{' + trench_table.to_latex('l@{}lrrrr', 'colorrule') + '}')

        for (x, letter) in enumerate(['C', 'E', 'G', 'I']):
            trench_table.set_cell(2, x + 1, (f'=SUM({letter}4:{letter}8)', 'm'))
        
        trench_table.set_cell(8, 1, ('=SUM(C10:C11)', 'm'))
        trench_table.set_cell(8, 4, ('=SUM(I10:I11)', 'm'))
        trench_table.set_cell(0, 1, ('=C3+C9', 'm'))
        if special_crossings.total() > 0:
            trench_table.set_cell(11, 1, (f'=SUM(C13:C{12+len(special_crossings)})', 'St.'))

        workbook = xlsxwriter.Workbook(os.path.join(destination, "Trenches.xlsx"))
        trench_table.to_xlsx(workbook, [5, 25, 15, 2, 15, 2, 15, 2, 15, 2])
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

    def process_points_of_interest(self, points, layers, extent, destination):
        max_id = max([point["Punkt_ID"] for point in points])
        x_coords = [0] * max_id
        y_coords = [0] * max_id
        for point in points:
            geometry = point.geometry()
            if geometry.type() != QgsWkbTypes.PointGeometry:
                continue

            pt = geometry.asPoint()
            id = int(point["Punkt_ID"])
            x_coords[id-1] = (pt.x() - extent.xMinimum()) / extent.width()
            y_coords[id-1] = 1 - (pt.y() - extent.yMinimum()) / extent.height()

            rect = self.rectangle_around_point(pt)
            path = os.path.join(destination, "Bilder", f"fotopunkt{id}.pdf")
            self.make_pic_pdf(layers, path, rect, zoom_factor=20)
            self.increment_progess()

        with open(os.path.join(destination, "Praesentation", "PointsOfInterest.tex"), "w") as f:
            x_coords_str = ''.join(['{' + str(x) + '}' for x in x_coords])
            y_coords_str = ''.join(['{' + str(y) + '}' for y in y_coords])
            f.write('\\storedata\\xcoords{' + x_coords_str + '}\n')
            f.write('\\storedata\\ycoords{' + y_coords_str + '}')

    def instantiate_address_trench_template(self):
        fotopunkt = GeneratePresentation.require_layer('Fotopunkt')
        trenches = GeneratePresentation.require_layer('Trenches')
        addresses = GeneratePresentation.require_layer('Adressen')
        polygons = GeneratePresentation.require_layer('Polygone')
        osm = GeneratePresentation.require_layer('OpenStreetMap')

        destination = QFileDialog.getExistingDirectory(None, 'Select Destination')
        if not destination:
            return
        self.destination_directory = destination
        maps_dir = os.path.join(destination, "Karten")

        points_of_interest = list(fotopunkt.getFeatures())

        self.init_progress_bar(11 + len(points_of_interest))
        self.copy_template("common", destination)
        self.copy_template("address_and_trenches", destination)
        self.increment_progess()

        GeneratePresentation.calculate_address_statistics(addresses, destination)
        self.increment_progess()

        GeneratePresentation.calculate_trench_lengths(trenches, destination)
        self.increment_progess()

        extent = self.calculate_extent([addresses.extent(), trenches.extent()])
        self.process_points_of_interest(points_of_interest, [addresses, trenches, osm], extent, destination)
        self.increment_progess()

        titlepic_path = os.path.join(destination, "Bilder", "titelbild.pdf")
        self.make_pic_pdf([fotopunkt, addresses, polygons, osm], titlepic_path)
        self.increment_progess()

        address_check_path = os.path.join(maps_dir, "adresscheck.pdf")
        self.make_pic_pdf([addresses, polygons, osm], address_check_path)
        self.increment_progess()

        hp_distribution_path = os.path.join(maps_dir, "hp-verteilung.pdf")
        color1 = QColor(72, 123, 182)
        color2 = QColor(228, 187, 114)
        color3 = QColor(84, 174, 74)
        hp_distribution = GeneratePresentation.style_layer(addresses, [
            ('"Total DNP" > 12', color1, color1.darker(), 0.3),
            ('"Total DNP" > 2 and "Total DNP" <= 12', color2, color2.darker(), 0.3),
            ('"Total DNP" <= 2 and "Total DNP" is not null', color3, color3.darker(), 0.3)
        ])
        self.make_pic_pdf([hp_distribution, polygons, osm], hp_distribution_path)
        self.increment_progess()

        trenches_path = os.path.join(maps_dir, "trenches.pdf")
        self.make_pic_pdf([trenches, polygons, osm], trenches_path)
        self.increment_progess()

        by_hands_path = os.path.join(maps_dir, "trenches-handschachtung.pdf")
        by_hands = GeneratePresentation.style_layer(trenches, [
            ('"Handschachtung" = false', QColor('black'), None, 0.3),
            ('"Handschachtung" = true', QColor('#54b04a'), None, 0.7)
        ])
        self.make_pic_pdf([by_hands, polygons, osm], by_hands_path)
        self.increment_progess()

        by_streets_path = os.path.join(maps_dir, "trenches-strassenkoerper.pdf")
        by_streets = GeneratePresentation.style_layer(trenches, [
            ('"In_Strasse" = false', QColor('black'), None, 0.3),
            ('"In_Strasse" = true', QColor('#db1e2a'), None, 0.7)
        ])
        self.make_pic_pdf([by_streets, polygons, osm], by_streets_path)
        self.increment_progess()

        by_private_path = os.path.join(maps_dir, "trenches-privatweg.pdf")
        by_private = GeneratePresentation.style_layer(trenches, [
            ('"Privatweg" = false', QColor('black'), None, 0.3),
            ('"Privatweg" = true', QColor('#487bb6'), None, 0.7)
        ])
        self.make_pic_pdf([by_private, polygons, osm], by_private_path)
        self.increment_progess()

        self.iface.messageBar().clearWidgets()
        self.iface.messageBar().pushMessage(
            "Success",
            "Presentation prepared in <a href=\"file:///" + destination + "\">" + destination + "</a>.",
            level=Qgis.MessageLevel.Success,
            duration=15
        )

    def attempt_instantiate_address_trench_template(self):
        try:
            self.instantiate_address_trench_template()
        except Exception as e:
            self.iface.messageBar().clearWidgets()
            self.iface.messageBar().pushMessage("Error", str(e), level=Qgis.Critical)

    def instantiate_surface_template(self):
        fotopunkt = GeneratePresentation.require_layer('Fotopunkt')
        oberflaechen = GeneratePresentation.require_layer('Oberflächenanalyse')
        polygons = GeneratePresentation.require_layer('Polygone')
        osm = GeneratePresentation.require_layer('OpenStreetMap')

        destination = QFileDialog.getExistingDirectory(None, 'Select Destination')
        if not destination:
            return
        self.destination_directory = destination

        points_of_interest = list(fotopunkt.getFeatures())
        self.init_progress_bar(5 + len(points_of_interest))
        self.copy_template("common", destination)
        self.copy_template("surface_classification", destination)
        self.increment_progess()

        GeneratePresentation.calculate_surface_statistics(oberflaechen, destination)
        self.increment_progess()

        extent = self.calculate_extent([oberflaechen.extent()])
        self.process_points_of_interest(points_of_interest, [oberflaechen, osm], extent, destination)
        self.increment_progess()

        titlepic_path = os.path.join(destination, "Bilder", "titelbild.pdf")
        self.make_pic_pdf([fotopunkt, oberflaechen, polygons, osm], titlepic_path)
        self.increment_progess()

        map_path = os.path.join(destination, "Karten", "karte.pdf")
        self.make_pic_pdf([oberflaechen, polygons, osm], map_path)
        self.increment_progess()

        self.iface.messageBar().clearWidgets()
        self.iface.messageBar().pushMessage(
            "Success",
            "Presentation prepared in <a href=\"file:///" + destination + "\">" + destination + "</a>.",
            level=Qgis.MessageLevel.Success,
            duration=15
        )

    def attempt_instantiate_surface_template(self):
        try:
            self.instantiate_surface_template()
        except Exception as e:
            self.iface.messageBar().clearWidgets()
            self.iface.messageBar().pushMessage("Error", str(e), level=Qgis.Critical)

    def calculate_surface_statistics(layer, destination):
        categories = SymbologyCategory.extract_symbology_categories(
            layer,
            'Belag',
            ['$area / (CASE WHEN "Typ" = \'b\' THEN 1.281 ELSE 5.787 END)', '1']
        )
        cats = {}
        for c in categories:
            cats[c.token] = c
        
        class CategoryGroup:
            def __init__(self, title, categories=[]):
                self.table = Table()
                self.total_meters = 0
                self.total_numbers = 0

                self.table.add_row(
                    [title, ('Länge', ''), ('Anteil', '')],
                    Table.Highlight.PRIMARY
                )

                for c in categories:
                    self.add_category(c)
            
            def add_category(self, c):
                value = c.value[0]
                self.total_meters += value
                self.table.add_row([c.label, value, 0], c.color)
            
            def add_total(self, label='Gesamt'):
                for row in self.table.rows[1:]:
                    row[2] = row[1] * 100 / self.total_meters

                self.table.add_row(
                    [label, self.total_meters, 100],
                    Table.Highlight.SECONDARY
                )
                return self.total_meters
            
            def cleanup(self):
                for row in self.table.rows:
                    if isinstance(row[1], int) or isinstance(row[1], float):
                        row[1] = (round(row[1]), 'm')
                    if isinstance(row[2], int) or isinstance(row[2], float):
                        row[2] = (round(row[2]), '%')
            
            def to_latex(self):
                colspec = 'l@{}X[l]rr'
                return self.table.to_latex(colspec, 'colorsquare')

        street_types = ['a', 't', 'g', 'v', 'm', 'n']
        street_categories = [cats[token] for token in street_types]
        sidewalk = CategoryGroup('Oberflächen Bürgersteig', street_categories)
        sidewalk_total = sidewalk.add_total()
        sidewalk.cleanup()

        street_types = ['sa', 'st', 'sg', 'sm', 'sn']
        street_categories = [cats[token] for token in street_types]
        street = CategoryGroup('Oberflächen Straße', street_categories)
        street_total = street.add_total()
        street.cleanup()

        special_types = ['x', 'sx']
        special_categories = [cats[token] for token in special_types]
        special = CategoryGroup('Sonderquerungen', special_categories)
        special.add_total()
        special.cleanup()
        number_special = cats['sx'].value[1]

        summary = CategoryGroup('Gesamtoberfläche')
        summary.table.add_row(['Bürgersteig', sidewalk_total, 0])
        summary.table.add_row(['Straße', street_total, 0])
        summary.total_meters = sidewalk_total + street_total
        summary.add_total()
        summary.cleanup()

        with open(os.path.join(destination, "Praesentation", "OberflaechenStatistik.tex"), "w") as f:
            f.write('\\newcommand\\surfacetypes{')
            for c in categories:
                f.write('\\item[\\colorsquare{' + color_to_tikz(c.color) + '}] ' + c.label + '\n')
            f.write('}\n')

            f.write('\\newcommand\\oberflaechenBuergersteig{' + sidewalk.to_latex() + '}\n')
            f.write('\\newcommand\\oberflaechenStrasse{' + street.to_latex() + '}\n')
            f.write('\\newcommand\\oberflaechenSonderquerung{' + special.to_latex() + '}\n')
            f.write('\\newcommand\\oberflaechenGesamt{' + summary.to_latex() + '}\n')

        workbook = xlsxwriter.Workbook(os.path.join(destination, "Oberflaechenanalyse.xlsx"))
        worksheet = workbook.add_worksheet()
        sidewalk.table.to_xlsx(workbook, [2, 30, 10, 2, 10, 2, 10], worksheet=worksheet)
        street.table.to_xlsx(workbook, worksheet=worksheet, offset=9)
        special.table.to_xlsx(workbook, worksheet=worksheet, offset=17)
        summary.table.to_xlsx(workbook, worksheet=worksheet, offset=22)
        workbook.close()