# encoding: utf-8

"""Smart Art Parts Objects"""

from pptx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from pptx.opc.package import XmlPart
from pptx.util import lazyproperty


class SmartArtDrawingPart(XmlPart):
    """A Smart Art Drawing Part
    Corresponds to parts having partnames matching ppt/diagrams/drawing[1-9][0-9]*.xml
    """
    partname_template = "/ppt/diagrams/drawing%d.xml"

    @classmethod
    def new(cls, package, xml_blob):
        drawing_part = cls.load(
            package.next_partname(cls.partname_template),
            CT.DML_DIAGRAM_DRAWING,
            package,
            xml_blob)
        return drawing_part


class SmartArtColorsPart(XmlPart):
    """A Smart Art Colors Part
    Corresponds to parts having partnames matching ppt/diagrams/colors[1-9][0-9]*.xml
    """
    partname_template = "/ppt/diagrams/colors%d.xml"

    @classmethod
    def new(cls, package, xml_blob):
        colors_part = cls.load(
            package.next_partname(cls.partname_template),
            CT.DML_DIAGRAM_COLORS,
            package,
            xml_blob)
        return colors_part

class SmartArtDataPart(XmlPart):
    """A Smart Art Data Part
    Corresponds to parts having partnames matching ppt/diagrams/data[1-9][0-9]*.xml
    """
    partname_template = "/ppt/diagrams/data%d.xml"
    def __init__(self, partname, content_type, package, element):
        super(SmartArtDataPart, self).__init__(partname, content_type, package, element)


    @classmethod
    def new(cls, package, data_xml_blob, drawing_xml_blog):
        data_part = cls.load(
            package.next_partname(cls.partname_template),
            CT.DML_DIAGRAM_DATA,
            package,
            data_xml_blob)

        drawing_part = SmartArtDrawingPart.new(package, drawing_xml_blog)
        return data_part, drawing_part

    @property
    def ext_list(self):
        return self._element.extLst.ext_lst

    def set_drawing_rId(self, rId):
       self.ext_list[0].dataModelExt.relId = rId

class SmartArtLayoutPart(XmlPart):
    """A Smart Art Layout Part
    Corresponds to parts having partnames matching ppt/diagrams/layout[1-9][0-9]*.xml
    """
    partname_template = "/ppt/diagrams/layout%d.xml"

    @classmethod
    def new(cls, package, xml_blob):
        layout_part = cls.load(
            package.next_partname(cls.partname_template),
            CT.DML_DIAGRAM_LAYOUT,
            package,
            xml_blob)
        return layout_part


class SmartArtQuickStylePart(XmlPart):
    """A Smart Art Quick Style Part
    Corresponds to parts having partnames matching ppt/diagrams/quickStyle[1-9][0-9]*.xml
    """
    partname_template = "/ppt/diagrams/quickStyle%d.xml"

    @classmethod
    def new(cls, package, xml_blob):
        style_part = cls.load(
            package.next_partname(cls.partname_template),
            CT.DML_DIAGRAM_STYLE,
            package,
            xml_blob)
        return style_part

