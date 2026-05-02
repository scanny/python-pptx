"""Provides mappings that embody aspects of the Open XML spec ISO/IEC 29500."""

from pptx.opc.constants import CONTENT_TYPE as CT

default_content_types = (
    ("aiff", CT.AUDIO_AIFF),
    ("bin", CT.PML_PRINTER_SETTINGS),
    ("bin", CT.SML_PRINTER_SETTINGS),
    ("bin", CT.WML_PRINTER_SETTINGS),
    ("bmp", CT.BMP),
    ("emf", CT.X_EMF),
    ("fntdata", CT.X_FONTDATA),
    ("gif", CT.GIF),
    ("jpe", CT.JPEG),
    ("jpeg", CT.JPEG),
    ("jpg", CT.JPEG),
    ("m4a", CT.AUDIO_MP4),
    ("mid", CT.AUDIO_MIDI),
    ("midi", CT.AUDIO_MIDI),
    ("mov", CT.MOV),
    ("mp3", CT.AUDIO_MPEG),
    ("mp4", CT.MP4),
    ("mpg", CT.MPG),
    ("ogg", CT.AUDIO_OGG),
    ("png", CT.PNG),
    ("rels", CT.OPC_RELATIONSHIPS),
    ("tif", CT.TIFF),
    ("tiff", CT.TIFF),
    ("vid", CT.VIDEO),
    ("wav", CT.AUDIO_WAV),
    ("wdp", CT.MS_PHOTO),
    ("wma", CT.AUDIO_X_MS_WMA),
    ("wmf", CT.X_WMF),
    ("wmv", CT.WMV),
    ("xlsx", CT.SML_SHEET),
    ("xml", CT.XML),
)


image_content_types = {
    "bmp": CT.BMP,
    "emf": CT.X_EMF,
    "gif": CT.GIF,
    "jpe": CT.JPEG,
    "jpeg": CT.JPEG,
    "jpg": CT.JPEG,
    "png": CT.PNG,
    "tif": CT.TIFF,
    "tiff": CT.TIFF,
    "wdp": CT.MS_PHOTO,
    "wmf": CT.X_WMF,
}
