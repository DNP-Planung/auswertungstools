from PyQt5.QtWidgets import QAction, QFileDialog, QProgressBar
from qgis.core import *
from qgis.PyQt.QtGui import QColor, QIcon
from qgis.PyQt.QtCore import QSize, Qt
import os, shutil, math


def classFactory(iface):
    return GeneratePresentation(iface)


class GeneratePresentation:
    def __init__(self, iface):
        self.iface = iface
        self.image_width = 1150
        self.image_height = 800
        self.zoom_factor = 5
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
        columns = ['"Total Kunde"', '"Total DNP"', '"Total DNP" - "Total Kunde"']
        result = []

        for condition in conditions:
            row = [GeneratePresentation.filtered_column_sum(layer, condition, '1')]
            for column in columns:
                row.append(GeneratePresentation.filtered_column_sum(layer, condition, column))
            result.append(row)
        
        totalRow = []
        for column in range(len(columns) + 1):
            totalRow.append(sum([row[column] for row in result]))
        result.append(totalRow)

        result_strings = [' & '.join([str(x) for x in row]) for row in result]
        result_strings.append(str(GeneratePresentation.filtered_column_sum(layer, '"Pruefung" = \'o\'', '1')))
        result_strings.append(str(GeneratePresentation.filtered_column_sum(layer, '"Pruefung" = \'n\'', '1')))

        with open(destination, "w") as f:
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

    def filtered_length_sum(layer, condition):
        return math.ceil(sum(QgsVectorLayerUtils.getValues(layer, f'CASE WHEN {condition} THEN $length ELSE 0 END')[0]))

    def calculate_trench_lengths(layer, destination):
        conditions = ['"Belag" = \'a\'', '"Belag" = \'t\'', '"Belag" = \'g\'', '"Belag" = \'m\'',  '"Belag" = \'k\'']
        columns = ['true', '"In_Strasse"', '"Handschachtung"', '"Privatweg"']

        result = []
        for condition in conditions:
            row = []
            for column in columns:
                row.append(GeneratePresentation.filtered_column_sum(layer, f'{condition} and {column}'))
            result.append(row)
        
        totalRow = []
        for column in range(len(columns) + 1):
            totalRow.append(sum([row[column] for row in result]))
        result.append(totalRow)
        print(result)

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

        address_statistics_path = os.path.join(destination, "Praesentation", "AdressStatistik.tex")
        GeneratePresentation.calculate_address_statistics(addresses, address_statistics_path)
        self.increment_progess()

        trench_statistics_path = os.path.join(destination, "Praesentation", "TrenchStatistik.tex")
        GeneratePresentation.calculate_trench_lengths(trenches, trench_statistics_path)
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
            ('"Handschachtung" = false', QColor('black'), None, None),
            ('"Handschachtung" = true', QColor('#54b04a'), None, 0.5)
        ])
        self.make_pic_pdf([by_hands, polygons, osm], by_hands_path)
        self.increment_progess()

        by_streets_path = os.path.join(images_dir, "trenches-strassenkoerper.pdf")
        by_streets = GeneratePresentation.style_layer(trenches, [
            ('"In_Strasse" = false', QColor('black'), None, None),
            ('"In_Strasse" = true', QColor('#db1e2a'), None, 0.5)
        ])
        self.make_pic_pdf([by_streets, polygons, osm], by_streets_path)
        self.increment_progess()

        by_private_path = os.path.join(images_dir, "trenches-privatweg.pdf")
        by_private = GeneratePresentation.style_layer(trenches, [
            ('"Privatweg" = false', QColor('black'), None, None),
            ('"Privatweg" = true', QColor('#487bb6'), None, 0.5)
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

        address_statistics_path = os.path.join(destination, "Praesentation", "AdressStatistik.tex")
        GeneratePresentation.calculate_address_statistics(addresses, address_statistics_path)
        self.increment_progess()

        trenches_statistics_path = os.path.join(destination, "Praesentation", "TrenchStatistik.tex")
        GeneratePresentation.calculate_trench_lengths(trenches, trenches_statistics_path)
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
            ('"Handschachtung" = false', QColor('black'), None, None),
            ('"Handschachtung" = true', QColor('#54b04a'), None, 0.5)
        ])
        self.make_pic_pdf([by_hands, polygons, osm], by_hands_path)
        self.increment_progess()

        by_streets_path = os.path.join(images_dir, "trenches-strassenkoerper.pdf")
        by_streets = GeneratePresentation.style_layer(trenches, [
            ('"In_Strasse" = false', QColor('black'), None, None),
            ('"In_Strasse" = true', QColor('#db1e2a'), None, 0.5)
        ])
        self.make_pic_pdf([by_streets, polygons, osm], by_streets_path)
        self.increment_progess()

        by_private_path = os.path.join(images_dir, "trenches-privatweg.pdf")
        by_private = GeneratePresentation.style_layer(trenches, [
            ('"Privatweg" = false', QColor('black'), None, None),
            ('"Privatweg" = true', QColor('#487bb6'), None, 0.5)
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