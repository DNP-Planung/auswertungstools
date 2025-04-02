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


class SymbologyCategory():
    def __init__(self, token, color, label, value):
        self.token = token
        self.color = color
        self.label = label
        self.value = value
    
    def extract_symbology_categories(layer, field, columns):
        result = []
        for c in layer.renderer().categories():
            token = c.value()
            color = c.symbol().color().name()[1:]
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
        self.zoom_factor = 4
        self.destination_directory = os.path.expanduser("~")
        self.dir_path = os.path.dirname(os.path.realpath(__file__))

    def initGui(self):
        presIcon = QIcon(os.path.join(self.dir_path, 'file-easel.png'))
        cameraIcon = QIcon(os.path.join(self.dir_path, 'camera.png'))
        rectIcon = QIcon(os.path.join(self.dir_path, 'square.png'))

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
        toolButton.setDefaultAction(QAction(presIcon, 'Präsentation erzeugen'))
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

    def increment_progess(self):
        if not self.progress:
            return
        self.progress.setValue(self.progress.value() + 1)

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

    def get_feature_coords(feature, extent):
        geometry = feature.geometry()
        if geometry.type() != QgsWkbTypes.PointGeometry:
            return None
        
        x = (geometry.asPoint()[0] - extent.xMinimum()) / extent.width()
        y = 1 - (geometry.asPoint()[1] - extent.yMinimum()) / extent.height()
        return [x, y]

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
        except RuntimeError as e:
            self.iface.messageBar().pushMessage("Error", str(e), level=Qgis.Critical)

    def make_pic_pdf(self, layers, destination, extent=None):
        project = QgsProject.instance()
        layout = QgsPrintLayout(project)
        layout.initializeDefaults()

        width = self.image_width / self.zoom_factor
        height = self.image_height / self.zoom_factor
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
        columns = ['1', '"Total Kunde"', '"Total DNP"', '"Total DNP" - "Total Kunde"']
        categories = SymbologyCategory.extract_symbology_categories(layer, 'Pruefung', columns)

        special_tokens = ['n', 'o']
        normal_categories = list(filter(lambda c: c.token not in special_tokens, categories))
        special_categories = list(filter(lambda c: c.token in special_tokens, categories))

        totalResult = []
        for column in range(len(columns)):
            totalResult.append(sum([c.value[column] for c in normal_categories]))

        with open(os.path.join(destination, "Praesentation", "AdressStatistik.tex"), "w") as f:
            # define colors
            for c in normal_categories + special_categories:
                f.write('\\definecolor{{addresscolor{}}}{{HTML}}{{{}}}\n'.format(c.token, c.color))

            # write table head
            f.write('\\newcommand\\adressStatistik{\n')
            f.write('\\begin{tblr}{colspec={l@{}l|rrrr},row{1,2}={bg=dnpblue,fg=white,font=\\bfseries},row{' + str(len(normal_categories) + 3) + '}={font=\\bfseries,bg=dnplightblue,fg=black}}\n')
            f.write('& Adresskulisse &&&& \\\\\n')
            f.write('&& Adressen & Einheiten \\Kunde & Einheiten DNP & Differenz \\\\\n')

            # write normal results
            for c in normal_categories:
                cells = ' & '.join([str(x) for x in c.value])
                f.write('\\colordot{{addresscolor{}}} & {} & {} \\\\\n'.format(c.token, c.label, cells))
            
            # write total
            cells = ' & '.join([str(x) for x in totalResult])
            f.write('\hline\n & Gesamt & {} \\\\\n'.format(cells))

            # if present, write results for special categories
            if len(special_categories) > 0:
                f.write('&&&&& \\\\\n')
                for c in special_categories:
                    f.write('\\colordot{{addresscolor{}}} & {} & {} &&& \\\\\n'.format(c.token, c.label, c.value[0]))
            f.write('\\end{tblr}}')

        # Create a workbook and add a worksheet.
        workbook = xlsxwriter.Workbook(os.path.join(destination, "Adressauswertung.xlsx"))
        worksheet = workbook.add_worksheet()

        highlight = workbook.add_format()
        highlight.set_bold()
        highlight.set_bg_color('#001aae') # DNP blue
        highlight.set_font_color('white')
        highlight.set_align('right')

        highlight_heading = workbook.add_format()
        highlight_heading.set_bold()
        highlight_heading.set_bg_color('#001aae') # DNP blue
        highlight_heading.set_font_color('white')
        highlight_heading.set_font_size(13)

        bg_white = workbook.add_format()
        bg_white.set_bg_color('white')

        bg_gray = workbook.add_format()
        bg_gray.set_bg_color('#dde2ff')

        border_top = workbook.add_format()
        border_top.set_top()
        border_top.set_bold()
        border_top.set_bg_color('#dde2ff')

        # set column width
        worksheet.set_column(0, 0, 2)
        worksheet.set_column(1, 1, 25)
        worksheet.set_column(2, 5, 15)

        for i in range(0, 100):
            for j in range(0, 20):
                worksheet.write(i, j, "", bg_white)

        for i in range(0, 2):
            for j in range(0, 6):
                worksheet.write(i, j, "", highlight)

        worksheet.write(0, 1, "Adresskulisse", highlight_heading)
        worksheet.write(1, 2, "Adressen", highlight)
        worksheet.write(1, 3, "Einheiten Kunde", highlight)
        worksheet.write(1, 4, "Einheiten DNP", highlight)
        worksheet.write(1, 5, "Differenz", highlight)

        for (i, c) in enumerate(normal_categories):
            color_fmt = workbook.add_format()
            color_fmt.set_bg_color('#' + c.color)
            worksheet.write(i + 2, 0, '', color_fmt)
            worksheet.write_row(i + 2, 1, [c.label] + c.value[:-1] + [f'=E{i+3}-D{i+3}'], bg_white)
        
        offset = len(normal_categories) + 2
        worksheet.write_row(offset, 0, [
            '', 'Gesamt', f'=SUM(C3:C{offset})', f'=SUM(D3:D{offset})', f'=SUM(E3:E{offset})', f'=SUM(F3:F{offset})'
        ], border_top)
        
        if len(special_categories) > 0:
            for (i, c) in enumerate(special_categories):
                color_fmt = workbook.add_format()
                color_fmt.set_bg_color('#' + c.color)
                worksheet.write(offset + 2 + i, 0, '', color_fmt)
                worksheet.write_row(offset + 2 + i, 1, [c.label, c.value[0]], bg_white)

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
        
        offenerTiefbau = []
        for column in range(len(columns)):
            offenerTiefbau.append(sum([row[column] for row in result]))
        result = [offenerTiefbau] + result

        rohrpressung = GeneratePresentation.filtered_length_sum(layer, '"Belag" = \'c\' and "Verfahren" = \'r\'')
        rohrpressung_privat = GeneratePresentation.filtered_length_sum(layer, '"Belag" = \'c\' and "Verfahren" = \'r\' and "Privatweg"')
        spuelbohrung = GeneratePresentation.filtered_length_sum(layer, '"Belag" = \'c\' and "Verfahren" = \'h\'')
        spuelbohrung_privat = GeneratePresentation.filtered_length_sum(layer, '"Belag" = \'c\' and "Verfahren" = \'h\' and "Privatweg"')
        geschlossener_tiefbau = [rohrpressung + spuelbohrung, rohrpressung_privat + spuelbohrung_privat]

        result_strings = [str(offenerTiefbau[0] + geschlossener_tiefbau[0])]
        result_strings += ['~m & '.join([str(x) for x in row]) + '~m' for row in result]
        result_strings.append(f'{geschlossener_tiefbau[0]}~m &&& {geschlossener_tiefbau[1]}~m')
        result_strings.append(f'{rohrpressung}~m &&& {rohrpressung_privat}~m')
        result_strings.append(f'{spuelbohrung}~m &&& {spuelbohrung_privat}~m')

        special_crossings = Counter(filter(None, QgsVectorLayerUtils.getValues(layer, '"Sonderquerung"')[0]))
        if special_crossings.total() == 0:
            result_strings.append('')
        else:
            result_strings.append(str(special_crossings.total()) + '~St.')

        with open(os.path.join(destination, "Praesentation", "TrenchStatistik.tex"), "w") as f:
            f.write('''\\newcommand\\trenchStatistik{{
\\begin{{tblr}}{{
    colspec={{l@{{}}lrrrr}},
    row{{1,2}}={{bg=dnpblue,fg=white,font=\\bfseries}},
    row{{3,9,12}}={{bg=dnplightblue,fg=black,font=\\bfseries}}
}}
    & Tiefbau gesamt &{0}~m &&& \\\\
    &&& im Straßenkörper & mit Handschachtung & in Privatweg \\\\
    & Offener Tiefbau								    & {1} \\\\
    \\colorrule{{trenchred}} 		& Asphalt 			& {2} \\\\
    \\colorrule{{trenchblue}} 		& Pflaster 			& {3} \\\\
    \\colorrule{{trenchgreen}} 	    & Unbefestigt		& {4} \\\\
    \\colorrule{{trenchpurple}} 	& Mosaikpflaster	& {5} \\\\
    \\colorrule{{trenchlightblue}}  & Kopfsteinpflaster	& {6} \\\\
    & Geschlossener Tiefbau                             & {7} \\\\
    \\colorrule{{trenchorange}} 	& Rohrpressung 		& {8} \\\\
    \\colorrule{{trenchspuelbohrung}} & Spülbohrung 	& {9} \\\\
    & Sonderquerungen & {10} &&& \\\\
'''.format(*result_strings))

            if special_crossings.total() == 0:
                f.write('& keine &&&& \\\\\n')
            else:
                for (crossing, count) in special_crossings.most_common():
                    f.write(f'& {crossing} & {count}~St. &&& \\\\\n')

            f.write('\\end{tblr}}')

        workbook = xlsxwriter.Workbook(os.path.join(destination, "Trenches.xlsx"))
        worksheet = workbook.add_worksheet()

        highlight = workbook.add_format()
        highlight.set_bg_color('#001aae') # DNP blue
        highlight.set_font_color('white')

        highlight_heading = workbook.add_format()
        highlight_heading.set_bg_color('#001aae') # DNP blue
        highlight_heading.set_font_color('white')
        highlight_heading.set_bold()
        highlight_heading.set_font_size(14)

        bg_white = workbook.add_format()
        bg_white.set_bg_color('white')

        bg_gray = workbook.add_format()
        bg_gray.set_bg_color('#dde2ff')
        bg_gray.set_bold()
        bg_gray.set_font_size(12)

        border_top = workbook.add_format()
        border_top.set_top()
        border_top.set_bg_color('#dde2ff')

        border_top_right = workbook.add_format()
        border_top_right.set_top()
        border_top_right.set_right()
        border_top_right.set_bg_color('#dde2ff')
        border_top_right.set_bold()

        border_right = workbook.add_format()
        border_right.set_right()
        border_right.set_bg_color('white')
        border_right.set_bold()

        border_right_gray = workbook.add_format()
        border_right_gray.set_right()
        border_right_gray.set_bg_color('#dde2ff')
        border_right_gray.set_bold()

        # set column widths
        worksheet.set_column(0, 0, 25)
        worksheet.set_column(1, 8, 15)
        for i in [2, 4, 6, 8]:
            worksheet.set_column(i, i, 2)

        # set background
        for i in range(0, 100):
            for j in range(0, 20):
                worksheet.write(i, j, "", bg_white)

        for i in range(0, 2):
            for j in range(0, 9):
                worksheet.write(i, j, "", highlight)

        for i in [2, 9, 13]:
            for j in range(0, 9):
                worksheet.write(i, j, "", bg_gray)

        # table headings
        worksheet.write_row(0, 0, ["Tiefbau gesamt", "=B3+B10", "m"], highlight_heading)
        worksheet.write(1, 3, "im Straßenkörper", highlight)
        worksheet.write(1, 5, "mit Handschachtung", highlight)
        worksheet.write(1, 7, "in Privatweg", highlight)
        worksheet.write(2, 0, "Offener Tiefbau", bg_gray)
        worksheet.write(3, 0, "Asphalt", bg_white)
        worksheet.write(4, 0, "Pflaster", bg_white)
        worksheet.write(5, 0, "Unbefestigt", bg_white)
        worksheet.write(6, 0, "Mosaikpflaster", bg_white)
        worksheet.write(7, 0, "Kopfsteinpflaster", bg_white)
        worksheet.write(9, 0, "Geschlossener Tiefbau", bg_gray)
        worksheet.write(10, 0, "Rohrpressung", bg_white)
        worksheet.write(11, 0, "Spülbohrung", bg_white)
        worksheet.write(13, 0, "Sonderquerungen", bg_gray)

        # offener tiefbau
        worksheet.write_row(2, 1, [p for c in 'BDFH' for p in [f'=SUM({c}4:{c}8)', 'm']], bg_gray)
        for (x, row) in enumerate(result[1:]):
            worksheet.write_row(x + 3, 1, [p for q in row for p in [q, 'm']], bg_white)
        
        # geschlossener Tiefbau
        worksheet.write_row(9, 1, ["=SUM(B11:B12)", "m"], bg_gray)
        worksheet.write_row(9, 7, ["=SUM(H11:H12)", "m"], bg_gray)
        worksheet.write_row(10, 1, [rohrpressung, "m"], bg_white)
        worksheet.write_row(10, 7, [rohrpressung_privat, "m"], bg_white)
        worksheet.write_row(11, 1, [spuelbohrung, "m"], bg_white)
        worksheet.write_row(11, 7, [spuelbohrung_privat, "m"], bg_white)

        if special_crossings.total() == 0:
            worksheet.write(14, 0, 'keine', bg_white)
        else:
            worksheet.write(13, 1, special_crossings.total(), bg_gray)
            worksheet.write(13, 2, 'St.', bg_gray)
            for (i, (crossing, count)) in enumerate(special_crossings.most_common()):
                worksheet.write_row(14 + i, 0, [crossing, count, 'St.'], bg_white)

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

    def write_poi_file(points, extent, destination):
        max_id = max([point["Punkt_ID"] for point in points])
        x_coords = [0] * max_id
        y_coords = [0] * max_id
        for point in points:
            coords = GeneratePresentation.get_feature_coords(point, extent)
            id = int(point["Punkt_ID"])
            x_coords[id-1] = coords[0]
            y_coords[id-1] = coords[1]

        with open(destination, "w") as f:
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
        images_dir = os.path.join(destination, "Karten")

        self.init_progress_bar(11)
        self.copy_template("common", destination)
        self.copy_template("address_and_trenches", destination)
        self.increment_progess()

        GeneratePresentation.calculate_address_statistics(addresses, destination)
        self.increment_progess()

        GeneratePresentation.calculate_trench_lengths(trenches, destination)
        self.increment_progess()

        poi_file = os.path.join(destination, "Praesentation", "PointsOfInterest.tex")
        extent = self.calculate_extent([addresses.extent(), trenches.extent()])
        GeneratePresentation.write_poi_file(list(fotopunkt.getFeatures()), extent, poi_file)
        self.increment_progess()

        titlepic_path = os.path.join(images_dir, "titelbild.pdf")
        self.make_pic_pdf([fotopunkt, addresses, polygons, osm], titlepic_path)
        self.increment_progess()

        address_check_path = os.path.join(images_dir, "adresscheck.pdf")
        self.make_pic_pdf([addresses, polygons, osm], address_check_path)
        self.increment_progess()

        hp_distribution_path = os.path.join(images_dir, "hp-verteilung.pdf")
        hp_distribution = GeneratePresentation.style_layer(addresses, [
            ('"Total DNP" > 12', QColor(72, 123, 182), QColor(60, 100, 160), 0.3),
            ('"Total DNP" > 2 and "Total DNP" <= 12', QColor(228, 187, 114), QColor(190, 160, 90), 0.3),
            ('"Total DNP" <= 2 and "Total DNP" is not null', QColor(84, 174, 74), QColor(70, 150, 60), 0.3)
        ])
        self.make_pic_pdf([hp_distribution, polygons, osm], hp_distribution_path)
        self.increment_progess()

        trenches_path = os.path.join(images_dir, "trenches.pdf")
        self.make_pic_pdf([trenches, polygons, osm], trenches_path)
        self.increment_progess()

        by_hands_path = os.path.join(images_dir, "trenches-handschachtung.pdf")
        by_hands = GeneratePresentation.style_layer(trenches, [
            ('"Handschachtung" = false', QColor('black'), None, 0.3),
            ('"Handschachtung" = true', QColor('#54b04a'), None, 0.7)
        ])
        self.make_pic_pdf([by_hands, polygons, osm], by_hands_path)
        self.increment_progess()

        by_streets_path = os.path.join(images_dir, "trenches-strassenkoerper.pdf")
        by_streets = GeneratePresentation.style_layer(trenches, [
            ('"In_Strasse" = false', QColor('black'), None, 0.3),
            ('"In_Strasse" = true', QColor('#db1e2a'), None, 0.7)
        ])
        self.make_pic_pdf([by_streets, polygons, osm], by_streets_path)
        self.increment_progess()

        by_private_path = os.path.join(images_dir, "trenches-privatweg.pdf")
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
        except RuntimeError as e:
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
        images_dir = os.path.join(destination, "Karten")

        self.init_progress_bar(5)
        self.copy_template("common", destination)
        self.copy_template("surface_classification", destination)
        self.increment_progess()

        categories = SymbologyCategory.extract_symbology_categories(oberflaechen, 'Belag', ['$area', '1'])
        GeneratePresentation.calculate_surface_statistics(categories, destination)
        self.increment_progess()

        poi_file = os.path.join(destination, "Praesentation", "PointsOfInterest.tex")
        extent = self.calculate_extent([oberflaechen.extent()])
        GeneratePresentation.write_poi_file(list(fotopunkt.getFeatures()), extent, poi_file)
        self.increment_progess()

        titlepic_path = os.path.join(images_dir, "titelbild.pdf")
        self.make_pic_pdf([fotopunkt, oberflaechen, polygons, osm], titlepic_path)
        self.increment_progess()

        map_path = os.path.join(images_dir, "karte.pdf")
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
        except RuntimeError as e:
            self.iface.messageBar().pushMessage("Error", str(e), level=Qgis.Critical)

    def calculate_surface_statistics(categories, destination):
        cats = {}
        for c in categories:
            cats[c.token] = c
        
        class CategoryGroup:
            def __init__(self, title, factor, with_numbers, tokens):
                self.title = title
                self.factor = factor
                self.with_numbers = with_numbers
                self.rows = []
                self.total_row_index = -1

                for token in tokens:
                    self.add_category(cats[token])
            
            def add_category(self, c, count_percentage = True):
                self.rows.append([
                    '\\colorsquare{' + c.token + '}\\hspace{5pt}',
                    c.label,
                    c.value[0] / self.factor,
                    0,
                    c.value[1],
                    count_percentage
                ])
            
            def add_total(self):
                self.total_row_index = len(self.rows)
                self.rows.append([
                    '', 'Gesamt',
                    sum([r[2] if r[5] else 0 for r in self.rows]),
                    100,
                    sum([r[4] for r in self.rows])
                ])
            
            def calculate_percentages(self):
                if self.total_row_index < 0:
                    return
                
                total = self.rows[self.total_row_index][2]
                print(self.rows[0:self.total_row_index], total)
                for row in self.rows[0:self.total_row_index]:
                    row[3] = row[2] * 100 / total
            
            def to_string(self):
                result = '\\begin{tblr}{width=\\textwidth,colspec={l@{}Xrrr},row{1}={bg=dnpblue,fg=white,font=\\bfseries},row{' + str(self.total_row_index + 2) + '}={font=\\bfseries}}\n'
                result += ' & ' + self.title + ' & Länge & Anteil & Anzahl \\\\\n'

                self.calculate_percentages()

                for row in self.rows:
                    if row[5]:
                        # counts into percentage
                        row[3] = str(round(row[3])) + '~%'
                    else:
                        row[3] = ''
                    if not self.with_numbers:
                        row[4] = ''

                    result += f'{row[0]} & {row[1]} & {round(row[2])}~m & {row[3]} & {row[4]} \\\\\n'
                
                result += '\\end{tblr}\n'
                return result

        sidewalk_types = ['a', 't', 'g', 'v', 'm', 'n']
        sidewalk_factor = 1.281
        sidewalk = CategoryGroup('Oberflächen Bürgersteig', sidewalk_factor, False, sidewalk_types)
        sidewalk.add_total()
        print(sidewalk.to_string())

        street_types = ['sa', 'st', 'sg', 'sm']
        street_factor = 5.787
        street = CategoryGroup('Oberflächen Straße', street_factor, True, street_types)
        street.add_category(cats['sn'], False)
        street.add_total()
        print(street.to_string())

        return

        special_types = ['x', 'sx']
        categories = [sidewalk_types, street_types, special_types]
        category_names = ['Bürgersteig', 'Straße', 'Sonderquerung']
        area_to_length_factors = [1.281, 5.787, 1]

        with open(os.path.join(destination, "Praesentation", "OberflaechenStatistik.tex"), "w") as f:
            # define colors
            for c in categories:
                f.write('\\definecolor{{surface{}}}{{HTML}}{{{}}}\n'.format(c.token, c.color))

            f.write('\\newcommand\\surfacetypes{')
            for c in categories:
                f.write(f'\\item[\\colorsquare{{{c.token}}}] {c.label}\n')
            f.write('}\n')

        for (category, name, factor) in zip(categories, category_names, area_to_length_factors):
            lengths = []
            numbers = []

            for type in category:
                query = f'CASE WHEN "Belag" = \'{type}\' THEN $area ELSE 0 END'
                values = list(filter(lambda x: x and x > 0, QgsVectorLayerUtils.getValues(layer, query)[0]))
                lengths.append(sum(values) / factor)
                numbers.append(len(values))
            
            total = sum(lengths)
            if not total:
                total = 1
            for (i, type) in enumerate(category):
                print('{} {} {} {}'.format(
                    type, round(lengths[i]), (lengths[i] * 100 / total), numbers[i]
                ))
            print('{}\t\t{}\t\t{}%\n\n'.format(name, round(total), 100))