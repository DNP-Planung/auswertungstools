from PyQt5.QtWidgets import QAction, QFileDialog, QErrorMessage
from qgis.core import QgsProject, QgsWkbTypes, QgsMapRendererParallelJob, QgsMapSettings, QgsRectangle, Qgis
from qgis.PyQt.QtGui import QColor
from qgis.PyQt.QtCore import QSize
import os, shutil


def classFactory(iface):
    return GeneratePresentation(iface)


class GeneratePresentation:
    def __init__(self, iface):
        self.iface = iface
        self.image_width = 1150
        self.image_height = 800
        self.destination_directory = os.path.expanduser("~")

    def initGui(self):
        self.init_template_action = QAction('Pres', self.iface.mainWindow())
        self.init_template_action.triggered.connect(self.attempt_instantiate_template)
        self.iface.addToolBarIcon(self.init_template_action)
        self.make_pic_action = QAction('Pic', self.iface.mainWindow())
        self.make_pic_action.triggered.connect(self.attempt_make_pic_user)
        self.iface.addToolBarIcon(self.make_pic_action)

    def unload(self):
        self.iface.removeToolBarIcon(self.init_template_action)
        del self.init_template_action
        self.iface.removeToolBarIcon(self.make_pic_action)
        del self.make_pic_action

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

    def copy_template(destination):
        dir_path = os.path.dirname(os.path.realpath(__file__))
        source = os.path.join(dir_path, "template")
        shutil.copytree(source, destination, dirs_exist_ok=True)

    def get_feature_coords(feature, extent):
        geometry = feature.geometry()
        if geometry.type() != QgsWkbTypes.PointGeometry:
            return None
        
        x = (geometry.asPoint()[0] - extent.xMinimum()) / extent.width()
        y = 1 - (geometry.asPoint()[1] - extent.yMinimum()) / extent.height()
        return [x, y]
    
    def make_pic_user(self):
        default_file = os.path.join(self.destination_directory, "Bilder", "map.png")
        destination = QFileDialog.getSaveFileName(
            None, "Save currently checked layers as PNG",
            default_file, "Portable Network Graphics (*.png)"
        )
        if destination and destination[0]:
            self.make_pic(self.iface.mapCanvas().layers(), destination[0])

    def attempt_make_pic_user(self):
        try:
            self.make_pic_user()
        except RuntimeError as e:
            self.iface.messageBar().pushMessage("Error", str(e), level=Qgis.Critical)

    def make_pic(self, layers, destination):
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
        points = GeneratePresentation.find_layer("Fotopunkt")
        destination = QFileDialog.getExistingDirectory(None, 'Select Destination')
        if not destination:
            return
        self.destination_directory = destination
        GeneratePresentation.copy_template(destination)

        extent = self.calculate_extent()
        poi_file = os.path.join(destination, "Präsentation", "PointsOfInterest.tex")
        GeneratePresentation.write_poi_file(points.getFeatures(), extent, poi_file)

    def attempt_instantiate_template(self):
        try:
            self.instantiate_template()
        except RuntimeError as e:
            self.iface.messageBar().pushMessage("Error", str(e), level=Qgis.Critical)