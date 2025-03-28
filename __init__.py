from PyQt5.QtWidgets import QAction, QFileDialog, QProgressBar
from qgis.core import *
from qgis.PyQt.QtGui import QColor, QIcon
from qgis.PyQt.QtCore import QSize, Qt
import os, shutil, math
from . import xlsxwriter


def classFactory(iface):
    return GeneratePresentation(iface)


class GeneratePresentation:
    def __init__(self, iface):
        self.iface = iface
        self.image_width = 1150
        self.image_height = 800
        self.zoom_factor = 4
        self.destination_directory = os.path.expanduser("~")

    def initGui(self):
        self.dir_path = os.path.dirname(os.path.realpath(__file__))

        presIcon = QIcon(os.path.join(self.dir_path, 'file-easel.png'))
        self.init_template_action = QAction(presIcon, 'Prepare Presentation', self.iface.mainWindow())
        self.init_template_action.triggered.connect(self.attempt_instantiate_template)
        self.iface.addToolBarIcon(self.init_template_action)

        cameraIcon = QIcon(os.path.join(self.dir_path, 'camera.png'))
        self.make_pic_action = QAction(cameraIcon, "Take picture", self.iface.mainWindow())
        self.make_pic_action.triggered.connect(self.attempt_make_pic_user)
        self.iface.addToolBarIcon(self.make_pic_action)

    def unload(self):
        self.iface.removeToolBarIcon(self.init_template_action)
        del self.init_template_action
        self.iface.removeToolBarIcon(self.make_pic_action)
        del self.make_pic_action

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

    def find_layer(key):
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

    def copy_template(self, destination):
        source = os.path.join(self.dir_path, "template")
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

    def make_pic_user(self):
        default_file = os.path.join(self.destination_directory, "Bilder", "map.pdf")
        destination = QFileDialog.getSaveFileName(
            None, "Save currently checked layers as PDF",
            default_file, "Portable Document Format (*.pdf)"
        )
        if destination and destination[0]:
            self.make_pic_pdf(self.iface.mapCanvas().layers(), destination[0])

        self.iface.messageBar().pushMessage(
            "Success",
            "Picture saved to <a href=\"file:///" + destination[0] + "\">" + destination[0] + "</a>.",
            level=Qgis.MessageLevel.Success,
            duration=15
        )

    def attempt_make_pic_user(self):
        try:
            self.make_pic_user()
        except RuntimeError as e:
            self.iface.messageBar().pushMessage("Error", str(e), level=Qgis.Critical)

    def make_pic_pdf(self, layers, destination):
        project = QgsProject.instance()
        layout = QgsPrintLayout(project)
        layout.initializeDefaults()

        width = self.image_width / self.zoom_factor
        height = self.image_height / self.zoom_factor
        size = QgsLayoutSize(width, height)
        pc = layout.pageCollection()
        pc.pages()[0].setPageSize(size)

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

    def filtered_column_sum(layer, condition, column):
        return sum(QgsVectorLayerUtils.getValues(layer, f'CASE WHEN {condition} THEN {column} ELSE 0 END')[0])

    def calculate_address_statistics(layer, destination):
        conditions = ['"Pruefung" = \'k\'', '"Pruefung" = \'d\'', '"Pruefung" = \'v\'', '"Pruefung" = \'h\'']
        columns = ['1', '"Total Kunde"', '"Total DNP"', '"Total DNP" - "Total Kunde"']
        result = []

        for condition in conditions:
            row = []
            for column in columns:
                row.append(GeneratePresentation.filtered_column_sum(layer, condition, column))
            result.append(row)
        
        totalRow = []
        for column in range(len(columns)):
            totalRow.append(sum([row[column] for row in result]))
        result.append(totalRow)
        
        optimiert = GeneratePresentation.filtered_column_sum(layer, '"Pruefung" = \'o\'', '1')
        nicht_vorhanden = GeneratePresentation.filtered_column_sum(layer, '"Pruefung" = \'n\'', '1')

        result_strings = [' & '.join([str(x) for x in row]) for row in result]
        result_strings.append(str(optimiert))
        result_strings.append(str(nicht_vorhanden))

        with open(os.path.join(destination, "Praesentation", "AdressStatistik.tex"), "w") as f:
            f.write('''\\newcommand\\adressStatistik{{
            \\begin{{tblr}}{{colspec={{l@{{}}l|rrrr}},row{{1,2}}={{bg=dnpblue,fg=white,font=\\bfseries}},row{{3,5,7}}={{bg=dnplightblue,fg=black}},row{{7}}={{font=\\bfseries}}}}
                & Adresskulisse &&&& \\\\
                && Adressen & Einheiten \Kunde & Einheiten DNP & Differenz \\\\
                \\colordot{{addressgreen}} & Adresse ohne Lage-Korrektur & {0} \\\\
                \\colordot{{addressyellow}} & Adressdaten angepasst	  & {1} \\\\
                \\colordot{{addressorange}} & Adresse verschoben 		  & {2} \\\\
                \\colordot{{addressblue}} & Adresse hinzugefügt 		  & {3} \\\\\\hline
                & Gesamt & {4} \\\\
                &&&&& \\\\
                \\colordot{{addressblack}} & Adresse optimiert & {5} &&& \\\\
                \\colordot{{addresspink}} & Adresse nicht vorhanden & {6} &&&
            \\end{{tblr}}
            }}'''.format(*result_strings))

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

        # set column width
        worksheet.set_column(0, 0, 25)
        worksheet.set_column(1, 4, 15)

        for i in range(0, 100):
            for j in range(0, 20):
                worksheet.write(i, j, "", bg_white)

        for i in range(0, 2):
            for j in range(0, 5):
                worksheet.write(i, j, "", highlight)

        worksheet.write(0, 0, "Adresskulisse", highlight_heading)
        worksheet.write(1, 1, "Adressen", highlight)
        worksheet.write(1, 2, "Einheiten Kunde", highlight)
        worksheet.write(1, 3, "Einheiten DNP", highlight)
        worksheet.write(1, 4, "Differenz", highlight)
        worksheet.write(2, 0, "Adresse ohne Lage-Korrektur", border_right_gray)
        worksheet.write(3, 0, "Adressdaten angepasst", border_right)
        worksheet.write(4, 0, "Adresse verschoben", border_right_gray)
        worksheet.write(5, 0, "Adresse hinzugefügt", border_right)
        worksheet.write(6, 0, "Gesamt", border_top_right)
        worksheet.write(7, 0, "", border_right)
        worksheet.write(8, 0, "", border_right)
        worksheet.write(9, 0, "Adresse optimiert", border_right_gray)
        worksheet.write(10, 0, "Adresse nicht vorhanden", border_right)

        worksheet.write_row(2, 1, result[0], bg_gray)
        worksheet.write_row(3, 1, result[1], bg_white)
        worksheet.write_row(4, 1, result[2], bg_gray)
        worksheet.write_row(5, 1, result[3], bg_white)

        worksheet.write(2, 4, "=D3-C3", bg_gray)
        worksheet.write(3, 4, "=D4-C4", bg_white)
        worksheet.write(4, 4, "=D5-C5", bg_gray)
        worksheet.write(5, 4, "=D6-C6", bg_white)

        worksheet.write(6, 1, "=SUM(B3:B6)", border_top)
        worksheet.write(6, 2, "=SUM(C3:C6)", border_top)
        worksheet.write(6, 3, "=SUM(D3:D6)", border_top)
        worksheet.write(6, 4, "=SUM(E3:E6)", border_top)

        worksheet.write(9, 1, optimiert, bg_gray)
        worksheet.write(10, 1, nicht_vorhanden, bg_white)

        workbook.close()

    def filtered_length_sum(layer, condition):
        return math.ceil(sum(QgsVectorLayerUtils.getValues(layer, f'CASE WHEN {condition} THEN $length ELSE 0 END')[0]))

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
            \\end{{tblr}}
            }}'''.format(*result_strings))

        # Create a workbook and add a worksheet.
        workbook = xlsxwriter.Workbook(os.path.join(destination, "Trenches.xlsx"))
        worksheet = workbook.add_worksheet()
        worksheet.write(0, 0, 'TODO')
        workbook.close()


    def attempt_instantiate_template(self):
        try:
            self.instantiate_template()
        except RuntimeError as e:
            self.iface.messageBar().pushMessage("Error", str(e), level=Qgis.Critical)

    def calculate_extent(self):
        reference1 = GeneratePresentation.find_layer("Adressen").extent()
        reference2 = GeneratePresentation.find_layer("Trenches").extent()
        xmin = min(reference1.xMinimum(), reference2.xMinimum())
        ymin = min(reference1.yMinimum(), reference2.yMinimum())
        xmax = max(reference1.xMaximum(), reference2.xMaximum())
        ymax = max(reference1.yMaximum(), reference2.yMaximum())
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
        numbers = {'1': 'one', '2': 'two', '3': 'three', '4': 'four', '5': 'five', '6': 'six'}
        with open(destination, "w") as f:
            for point in points:
                coords = GeneratePresentation.get_feature_coords(point, extent)
                id = str(point["Punkt_ID"])
                if not coords or id not in numbers:
                    continue
                f.write("\\newcommand\\pointofinterest" + numbers[id] + "X{" + str(coords[0]) + "}\n")
                f.write("\\newcommand\\pointofinterest" + numbers[id] + "Y{" + str(coords[1]) + "}\n")
                del numbers[id]
        
        keys = numbers.keys()
        if len(keys) > 0:
            raise RuntimeError('No point with Point_ID=' + next(iter(keys)) + ' found. The LaTeX presentation will not compile.')

    def instantiate_template(self):
        fotopunkt = GeneratePresentation.find_layer('Fotopunkt')
        trenches = GeneratePresentation.find_layer('Trenches')
        addresses = GeneratePresentation.find_layer('Adressen')
        polygons = GeneratePresentation.find_layer('Polygone')
        osm = GeneratePresentation.find_layer('OpenStreetMap')

        destination = QFileDialog.getExistingDirectory(None, 'Select Destination')
        if not destination:
            return
        self.destination_directory = destination
        images_dir = os.path.join(destination, "Karten")

        self.init_progress_bar(11)
        self.copy_template(destination)
        self.increment_progess()

        GeneratePresentation.calculate_address_statistics(addresses, destination)
        self.increment_progess()

        GeneratePresentation.calculate_trench_lengths(trenches, destination)
        self.increment_progess()

        poi_file = os.path.join(destination, "Praesentation", "PointsOfInterest.tex")
        extent = self.calculate_extent()
        GeneratePresentation.write_poi_file(fotopunkt.getFeatures(), extent, poi_file)
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
            ('"Total DNP" <= 2', QColor(84, 174, 74), QColor(70, 150, 60), 0.3)
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

    def attempt_instantiate_template(self):
        try:
            self.instantiate_template()
        except RuntimeError as e:
            self.iface.messageBar().pushMessage("Error", str(e), level=Qgis.Critical)