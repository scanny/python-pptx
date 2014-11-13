# encoding: utf-8

"""
ImagePart and related objects.
"""

import hashlib
import os

try:
    from PIL import Image as PIL_Image
except ImportError:
    import Image as PIL_Image

from StringIO import StringIO

from pptx.opc.package import Part
from pptx.opc.spec import image_content_types
from pptx.util import lazyproperty, Px


class ImagePart(Part):
    """
    An image part, generally having a partname matching the regex
    ``ppt/media/image[1-9][0-9]*.*``.
    """
    def __init__(self, partname, content_type, blob, package, filename=None):
        super(ImagePart, self).__init__(
            partname, content_type, blob, package
        )
        self._filename = filename

    @classmethod
    def new(cls, package, image):
        """
        Return a new |ImagePart| instance containing *image*, which is an
        |Image| object.
        """
        partname = package.next_image_partname(image.ext)
        return cls(
            partname, image.content_type, image.blob, package, image.filename
        )

    @property
    def ext(self):
        """
        Return file extension for this image e.g. ``'png'``.
        """
        return self.partname.ext

    @classmethod
    def load(cls, partname, content_type, blob, package):
        return cls(partname, content_type, blob, package)

    def scale(self, width, height):
        """
        Return scaled image dimensions based on supplied parameters. If
        *width* and *height* are both |None|, the native image size is
        returned. If neither *width* nor *height* is |None|, their values are
        returned unchanged. If a value is provided for either *width* or
        *height* and the other is |None|, the dimensions are scaled,
        preserving the image's aspect ratio.
        """
        native_width_px, native_height_px = self._size
        native_width = Px(native_width_px)
        native_height = Px(native_height_px)

        if width is None and height is None:
            width = native_width
            height = native_height
        elif width is None:
            scaling_factor = float(height) / float(native_height)
            width = int(round(native_width * scaling_factor))
        elif height is None:
            scaling_factor = float(width) / float(native_width)
            height = int(round(native_height * scaling_factor))
        return width, height

    @lazyproperty
    def sha1(self):
        """
        The SHA1 hash digest for the image binary of this image part, like:
        ``'1be010ea47803b00e140b852765cdf84f491da47'``.
        """
        return hashlib.sha1(self._blob).hexdigest()

    @property
    def _desc(self):
        """
        Return filename associated with this image, either the filename of
        the original image file the image was created with or a synthetic
        name of the form ``image.ext`` where ``ext`` is appropriate to the
        image file format, e.g. ``'jpg'``.
        """
        # return generic filename if original filename is unknown
        if self._filename is None:
            return 'image.%s' % self.ext
        return self._filename

    @property
    def _size(self):
        """
        Return *width*, *height* tuple representing native dimensions of
        image in pixels.
        """
        image = Image.from_blob(self.blob)
        return image.size


class Image(object):
    """
    Immutable value object representing an image such as a JPEG, PNG, or GIF.
    """
    def __init__(self, blob, filename):
        super(Image, self).__init__()
        self._blob = blob
        self._filename = filename

    @classmethod
    def from_blob(cls, blob, filename=None):
        """
        Return a new |Image| object loaded from the image binary in *blob*.
        """
        return cls(blob, filename)

    @classmethod
    def from_file(cls, image_file):
        """
        Return a new |Image| object loaded from *image_file*, which can be
        either a path (string) or a file-like object.
        """
        if isinstance(image_file, basestring):
            # treat image_file as a path
            with open(image_file, 'rb') as f:
                blob = f.read()
            filename = os.path.basename(image_file)
        else:
            # assume image_file is a file-like object
            image_file.seek(0)
            blob = image_file.read()
            filename = None

        return cls.from_blob(blob, filename)

    @property
    def blob(self):
        """
        The binary image bytestream of this image.
        """
        return self._blob

    @lazyproperty
    def content_type(self):
        """
        MIME-type of this image, e.g. ``'image/jpeg'``.
        """
        return image_content_types[self.ext]

    @lazyproperty
    def ext(self):
        """
        Canonical file extension for this image e.g. ``'png'``. The returned
        extension is all lowercase and is the canonical extension for the
        content type of this image, regardless of what extension may have
        been used in its filename, if any.
        """
        ext_map = {
            'BMP': 'bmp', 'GIF': 'gif', 'JPEG': 'jpg', 'PNG': 'png',
            'TIFF': 'tiff', 'WMF': 'wmf'
        }
        format = self._format
        if format not in ext_map:
            tmpl = "unsupported image format, expected one of: %s, got '%s'"
            raise ValueError(tmpl % (ext_map.keys(), format))
        return ext_map[format]

    @property
    def filename(self):
        """
        The filename from the path from which this image was loaded, if
        loaded from the filesystem. |None| if no filename was used in
        loading, such as when loaded from an in-memory stream.
        """
        return self._filename

    @lazyproperty
    def sha1(self):
        """
        SHA1 hash digest of the image blob
        """
        return hashlib.sha1(self._blob).hexdigest()

    @lazyproperty
    def size(self):
        """
        A (width, height) 2-tuple specifying the dimensions of this image in
        pixels.
        """
        return self._pil_props[1]

    @property
    def _format(self):
        """
        The PIL Image format of this image, e.g. 'PNG'.
        """
        return self._pil_props[0]

    @lazyproperty
    def _pil_props(self):
        """
        A tuple containing useful image properties extracted from this image
        using Pillow (Python Imaging Library, or 'PIL').
        """
        stream = StringIO(self._blob)
        pil_image = PIL_Image.open(stream)
        format = pil_image.format
        width_px, height_px = pil_image.size
        stream.close()
        return (format, (width_px, height_px))
