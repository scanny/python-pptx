# encoding: utf-8

""" Bullet-related objects such as BulletFont and BulletChar."""

from __future__ import absolute_import, division, print_function, unicode_literals

from pptx.enum.text import AUTO_NUMBER_SCHEME
from pptx.oxml.text import (
    CT_TextBulletAutoNumber,
    CT_TextNoBullet, 
    CT_TextBlipBullet, 
    CT_TextCharBullet, 
    CT_TextBulletColorFollowText,
    CT_TextBulletSizeFollowText, 
    CT_TextBulletSizePercent, 
    CT_TextBulletSizePoints,
    CT_TextBulletTypefaceFollowText,
    CT_TextFont,
)
from pptx.oxml.dml.color import CT_Color
from pptx.dml.color import ColorFormat, _SchemeColor, RGBColor
from pptx.util import Pt, Centipoints, Emu

class TextBullet(object):
    """
    Provides acess to the current Text Bullet properties and provies
    methods to change the type
    """

    def __init__(self, parent, bullet_obj):
        super(TextBullet, self).__init__()
        self._parent = parent
        self._bullet = bullet_obj
    
    @classmethod
    def from_parent(cls, parent):
        """
        Return |TextBullet| object
        """

        bullet_elm = parent.eg_textBullet
        bullet = _Bullet(bullet_elm)
        text_bullet = cls(parent, bullet)
        return text_bullet

    def no_bullet(self):
        """
        Sets the Bullet to NoBullet and omits the character
        """
        noBullet = self._parent.get_or_change_to_buNone()
        self._bullet = _NoBullet(noBullet)

    def auto_number(self, char_type=AUTO_NUMBER_SCHEME.ARABIC_PERIOD, start_at=None):
        """
        Sets the Bullet to Auto Number
        """
        autoNum = self._parent.get_or_change_to_buAutoNum()
        self._bullet = _AutoNumBullet(autoNum)
        self._bullet.char_type = char_type
        if start_at:
            self._bullet.start_at = start_at

    def character(self, char="•"):
        """
        Sets the Bullet to a specific character that is passed
        """
        buChar = self._parent.get_or_change_to_buChar()
        self._bullet = _CharBullet(buChar)
        self.char = char

    @property
    def char_type(self):
        """ Format of the AutoNumber Bullets """
        return self._bullet.char_type

    @char_type.setter
    def char_type(self, value):
        """ Set's the Format of the AutoNumber Bullets """
        self._bullet.char_type = value

    @property
    def start_at(self):
        """ Starting Value for AutoNumber Bullets """
        return self._bullet.start_at
    
    @start_at.setter
    def start_at(self, value):
        """ Sets starting Value for AutoNumber Bullets """
        self._bullet.start_at = value
    
    @property
    def char(self):
        """ String used as the bullet for a CharBullet"""
        return self._bullet.character

    @char.setter
    def char(self, value):
        """ String used as the bullet for a CharBullet"""
        self._bullet.char = value


    @property
    def type(self):
        """ Return a string type """
        return self._bullet.type
    

class _Bullet(object):
    """
    Object factory for TextBullet objects
    """

    def __new__(cls, xBullet):
        if isinstance(xBullet, CT_TextNoBullet):
            bullet_cls = _NoBullet
        elif isinstance(xBullet, CT_TextBulletAutoNumber):
            bullet_cls = _AutoNumBullet
        elif isinstance(xBullet, CT_TextCharBullet):
            bullet_cls = _CharBullet
        elif isinstance(xBullet, CT_TextBlipBullet):
            bullet_cls = _PictureBullet
        else:
            bullet_cls = _NoBullet
        
        return super(_Bullet, cls).__new__(bullet_cls)

    @property
    def char(self):
        """ Raise TypeError for types that do not override this property"""
        tmpl = "TextBullet type %s has no char property. call .character() first"
        raise TypeError(tmpl % self.__class__.__name__)

    @char.setter
    def char(self, value):
        """ Raise TypeError for types that do not override this property"""
        tmpl = "TextBullet type %s has no char property. call .character() first"
        raise TypeError(tmpl % self.__class__.__name__)

    @property
    def char_type(self):
        """ Raise TypeError for types that do not override this property"""
        tmpl = "TextBullet type %s has no char_type property. call .auto_number() first"
        raise TypeError(tmpl % self.__class__.__name__)

    @char_type.setter
    def char_type(self, value):
        """ Raise TypeError for types that do not override this property"""
        tmpl = "TextBullet type %s has no char_type property. call .auto_number() first"
        raise TypeError(tmpl % self.__class__.__name__)

    @property
    def start_at(self):
        """ Raise TypeError for types that do not override this property"""
        tmpl = "TextBullet type %s has no start_at property. call .auto_number() first"
        raise TypeError(tmpl % self.__class__.__name__)

    @start_at.setter
    def start_at(self, value):
        """ Raise TypeError for types that do not override this property"""
        tmpl = "TextBullet type %s has no start_at property. call .auto_number() first"
        raise TypeError(tmpl % self.__class__.__name__)

    @property
    def type(self):
        tmpl = ".type property must be implmented on %s"
        raise NotImplementedError(tmpl % self.__class__.__name__)


class _NoBullet(_Bullet):
    @property
    def type(self):
        return "NoBullet"

class _AutoNumBullet(_Bullet):
    def __init__(self, autoNumBullet):
        super(_AutoNumBullet, self).__init__()
        self._autoNumBullet = autoNumBullet

    @property
    def type(self):
        return "AutoNumBullet"

    @property
    def char_type(self):
        return self._autoNumBullet.char_type
    
    @char_type.setter
    def char_type(self, value):
        self._autoNumBullet.char_type = value

class _CharBullet(_Bullet):
    def __init__(self, charBullet):
        super(_CharBullet, self).__init__()
        self._charBullet = charBullet

    @property
    def type(self):
        return "CharBullet"

    @property
    def char(self):
        """
        String Character used as the bullet
        """
        return self._charBullet.char
    
    @char.setter
    def char(self, value):
        self._charBullet.char = str(value)

class _PictureBullet(_Bullet):
    """ Picture Bullets are not fully implemented at this time """
    @property
    def type(self):
        return "PictureBullet"



class BulletColor(object):
    """
    Provides access to the bullet color options
    """
    def __init__(self, parent, bullet_color_obj):
        super(BulletColor, self).__init__()
        self._parent = parent
        self._bullet_color = bullet_color_obj
    
    @classmethod
    def from_parent(cls, parent):
        """
        Return |BulletColor| object
        """
        bullet_color_elm = parent.eg_textBulletColor
        bullet_color = _BulletColor(bullet_color_elm)
        text_bullet_color = cls(parent, bullet_color)
        return text_bullet_color

    def follow_text(self):
        """
        Sets the BulletColor to _BulletColorFollowText
        """
        follow_text = self._parent.get_or_change_to_buClrTx()
        self._bullet_color = _BulletColorFollowText(follow_text)

    def set_color(self, color):
        """
        Sets the BulletColor to _BulletColorSpecific and sets the color
        """
        bullet_color = self._parent.get_or_change_to_buClr()
        self._bullet_color = _BulletColorSpecific(bullet_color)
        self.color = color

    @property
    def type(self):
        """ Return a string type """
        return self._bullet_color.type
    
    @property
    def color(self):
        """ Return the |ColorFormat| object """
        return self._bullet_color.color
    
    @color.setter
    def color(self, value):
        """
        Set the value of the color as either an RGB or Theme Color
        """
        if not isinstance(self._bullet_color, _BulletColorSpecific):
            raise TypeError("BulletColor is not of type BulletColorSpecific")
        color_obj = self._bullet_color.color
        if isinstance(value, RGBColor):
            color_obj.rgb = value
        elif isinstance(value, _SchemeColor):
            color_obj.theme_color = value
        else:
            raise TypeError("Provided color value is incorrect type")
     

class _BulletColor(object):
    """
    Object factory for BulletColor objects
    """

    def __new__(cls, xBulletColor):
        if isinstance(xBulletColor, CT_TextBulletColorFollowText):
            bullet_color_cls = _BulletColorFollowText
        elif isinstance(xBulletColor, CT_Color):
            bullet_color_cls = _BulletColorSpecific
        else:
            bullet_color_cls = _BulletColor
        
        return super(_BulletColor, cls).__new__(bullet_color_cls)
        
    @property
    def type(self):
        return "NoBulletColor"

    @property
    def color(self):
        """Raise TypeError for types that do not override this property."""
        tmpl = (
            "BulletColor type %s has no color, call .set_color() first"
        )
        raise TypeError(tmpl % self.__class__.__name__)

    @color.setter
    def color(self, value):
        """Raise TypeError for types that do not override this property."""
        tmpl = (
            "BulletColor type %s has no color, call .set_color() first"
        )
        raise TypeError(tmpl % self.__class__.__name__)


class _BulletColorFollowText(_BulletColor):
    """
    Designates that the Bullet Color will match the accompanying paragraph text.
    """
    @property
    def type(self):
        return "BulletColorFollowText"

class _BulletColorSpecific(_BulletColor):
    """
    """
    def __init__(self, bullet_color):
        super(_BulletColorSpecific, self).__init__()
        self._bullet_color = bullet_color

    @property
    def type(self):
        return "BulletColorSpecific"

    @property
    def color(self):
        return ColorFormat.from_colorchoice_parent(self._bullet_color)



class BulletSize(object):
    """
    Provides access to the bullet size options
    """
    def __init__(self, parent, bullet_size_obj):
        super(BulletSize, self).__init__()
        self._parent = parent
        self._bullet_size = bullet_size_obj
    
    @classmethod
    def from_parent(cls, parent):
        """
        Return |BulletSize| object
        """
        bullet_size_elm = parent.eg_textBulletSize
        bullet_size = _BulletSize(bullet_size_elm)
        text_bullet_size = cls(parent, bullet_size)
        return text_bullet_size

    def follow_text(self):
        """
        Sets the BulletSize to _BulletSizeFollowText
        """
        follow_text = self._parent.get_or_change_to_buSzTx()
        self._bullet_size = _BulletSizeFollowText(follow_text)

    def set_points(self, value):
        """
        Sets the BulletSize to _BulletSizePoints and sets the points value
        """
        points = self._parent.get_or_change_to_buSzPts()
        self._bullet_size = _BulletSizePoints(points)
        self.points = value

    def set_percentage(self, value):
        """
        Sets the BulletSize to _BulletSizePercentage and sets the percent value
        """
        percentage = self._parent.get_or_change_to_buSzPct()
        self._bullet_size = _BulletSizePercent(percentage)
        self.percentage = value

    @property
    def points(self):
        """
        Returns the points size
        """
        return self._bullet_size.points
    
    @points.setter
    def points(self, value):
        """ Sets the points size """
        self._bullet_size.points = value    

    @property
    def percentage(self):
        """
        Returns the percentage size
        """
        return self._bullet_size.percentage
    
    @percentage.setter
    def percentage(self, value):
        """ Sets the percentage size """
        self._bullet_size.percentage = value    

class _BulletSize(object):
    """
    Object factory for BulletSize objects
    """

    def __new__(cls, xBulletSize):
        if isinstance(xBulletSize, CT_TextBulletSizeFollowText):
            bullet_size_cls = _BulletSizeFollowText
        elif isinstance(xBulletSize, CT_TextBulletSizePercent):
            bullet_size_cls = _BulletSizePercent
        elif isinstance(xBulletSize, CT_TextBulletSizePoints):
            bullet_size_cls = _BulletSizePoints
        else:
            bullet_size_cls = _BulletSize
        
        return super(_BulletSize, cls).__new__(bullet_size_cls)

    @property
    def points(self):
        """Raise TypeError for types that do not override this property."""
        tmpl = (
            "BulletSize type %s has no points property, call .set_points() first"
        )
        raise TypeError(tmpl % self.__class__.__name__)
    
    @points.setter
    def points(self, value):
        """Raise TypeError for types that do not override this property."""
        tmpl = (
            "BulletSize type %s has no points property, call .set_points() first"
        )
        raise TypeError(tmpl % self.__class__.__name__)

    @property
    def percentage(self):
        """Raise TypeError for types that do not override this property."""
        tmpl = (
            "BulletSize type %s has no percentage property, call .set_percentage() first"
        )
        raise TypeError(tmpl % self.__class__.__name__)

    @percentage.setter
    def percentage(self, value):
        """Raise TypeError for types that do not override this property."""
        tmpl = (
            "BulletSize type %s has no percentage property, call .set_percentage() first"
        )
        raise TypeError(tmpl % self.__class__.__name__)

class _BulletSizeFollowText(_BulletSize):
    """
    Designates that the Bullet Size will match the accompanying paragraph text.
    """
    @property
    def type(self):
        return "BulletSizeFollowText"

class _BulletSizePercent(_BulletSize):
    """ Proxies a `a: buSzPct` element. """

    def __init__(self, bullet_size):
        super(_BulletSizePercent, self).__init__()
        self._bullet_size = bullet_size

    @property
    def type(self):
        return "BulletSizePercent"

    @property
    def percentage(self):
        """ Returns the percentage value for the bullet size """
        return self._bullet_size.val
            
    @percentage.setter
    def percentage(self, value):
        """ Sets the percentage value for the bullet size """
        self._bullet_size.val = value
            

class _BulletSizePoints(_BulletSize):
    """ Proxies a `a: buSzPts` element. """
    
    def __init__(self, bullet_size):
        super(_BulletSizePoints, self).__init__()
        self._bullet_size = bullet_size


    @property
    def type(self):
        return "BulletSizePoints"


    @property
    def points(self):
        """ Returns the point value for the bullet size """
        return Centipoints(self._bullet_size.val)
            
    @points.setter
    def points(self, value):
        """ Sets the points value for the bullet size """
        size = Emu(value).centipoints
        self._bullet_size.val = size
            


class BulletFont(object):
    """
    Provides access to the bullet font typeface options
    """
    def __init__(self, parent, bullet_font_obj):
        super(BulletFont, self).__init__()
        self._parent = parent
        self._bullet_font = bullet_font_obj
    
    @classmethod
    def from_parent(cls, parent):
        """
        Return |BulletFont| object
        """
        bullet_font_elm = parent.eg_textBulletTypeface
        bullet_font = _BulletFont(bullet_font_elm)
        text_bullet_font = cls(parent, bullet_font)
        return text_bullet_font

    def follow_text(self):
        """
        Sets the BulletFont to _BulletFontFollowText
        """
        follow_text = self._parent.get_or_change_to_buFontTx()
        self._bullet_font = _BulletFontFollowText(follow_text)

    def set_typeface(self, typeface="Arial"):
        """
        Sets the BulletFont to _BulletFontSpecific
        """
        bullet_font = self._parent.get_or_change_to_buFont()
        self._bullet_font = _BulletFontSpecific(bullet_font)
        self.typeface = typeface

    @property
    def typeface(self):
        return self._bullet_font.typeface

    @typeface.setter
    def typeface(self, value):
        self._bullet_font.typeface = value
    
    @property
    def pitch_family(self):
        return self._bullet_font.pitchFamily

    @pitch_family.setter
    def pitch_family(self, value):
        self._bullet_font.pitchFamily = value

    @property
    def panose(self):
        return self._bullet_font.panose
    
    @panose.setter
    def panose(self, value):
        self._bullet_font.panose = value
    
    @property
    def charset(self):
        return self._bullet_font.charset
    
    @charset.setter
    def charset(self, value):
        self._bullet_font.charset = value

    @property
    def type(self):
        """ Return a string type """
        return self._bullet_font.type
    
     

class _BulletFont(object):
    """
    Object factory for BulletFont objects
    """

    def __new__(cls, xBulletFont):
        if isinstance(xBulletFont, CT_TextBulletTypefaceFollowText):
            bullet_font_cls = _BulletFontFollowText
        elif isinstance(xBulletFont, CT_TextFont):
            bullet_font_cls = _BulletFontSpecific
        else:
            bullet_font_cls = _BulletFont
        
        return super(_BulletFont, cls).__new__(bullet_font_cls)
        
    @property
    def type(self):
        return "NoBulletFont"

    @property
    def typeface(self):
        """Raise TypeError for types that do not override this property."""
        tmpl = (
            "BulletFont type %s has no typeface property, call .set_typeface() first"
        )
        raise TypeError(tmpl % self.__class__.__name__)

    @typeface.setter
    def typeface(self, value):
        """Raise TypeError for types that do not override this property."""
        tmpl = (
            "BulletFont type %s has no typeface property, call .set_typeface() first"
        )
        raise TypeError(tmpl % self.__class__.__name__)
    
    @property
    def pitch_family(self):
        """Raise TypeError for types that do not override this property."""
        tmpl = (
            "BulletFont type %s has no pitch_family property, call .set_typeface() first"
        )
        raise TypeError(tmpl % self.__class__.__name__)

    @pitch_family.setter
    def pitch_family(self, value):
        """Raise TypeError for types that do not override this property."""
        tmpl = (
            "BulletFont type %s has no pitch_family property, call .set_typeface() first"
        )
        raise TypeError(tmpl % self.__class__.__name__)

    @property
    def panose(self):
        """Raise TypeError for types that do not override this property."""
        tmpl = (
            "BulletFont type %s has no panose property, call .set_typeface() first"
        )
        raise TypeError(tmpl % self.__class__.__name__)
    
    @panose.setter
    def panose(self, value):
        """Raise TypeError for types that do not override this property."""
        tmpl = (
            "BulletFont type %s has no panose property, call .set_typeface() first"
        )
        raise TypeError(tmpl % self.__class__.__name__)
    
    @property
    def charset(self):
        """Raise TypeError for types that do not override this property."""
        tmpl = (
            "BulletFont type %s has no charset property, call .set_typeface() first"
        )
        raise TypeError(tmpl % self.__class__.__name__)
    
    @charset.setter
    def charset(self, value):
        """Raise TypeError for types that do not override this property."""
        tmpl = (
            "BulletFont type %s has no charset property, call .set_typeface() first"
        )
        raise TypeError(tmpl % self.__class__.__name__)


class _BulletFontFollowText(_BulletFont):
    """
    Designates that the Bullet Font will match the accompanying paragraph text.
    """
    @property
    def type(self):
        return "BulletFontFollowText"

class _BulletFontSpecific(_BulletFont):
    """
    Designates the specific Bullet font typeface characteristics
    """
    def __init__(self, bullet_font):
        super(_BulletFontSpecific, self).__init__()
        self._bullet_font = bullet_font

    @property
    def type(self):
        return "BulletFontSpecific"

    @property
    def typeface(self):
        return self._bullet_font.typeface

    @typeface.setter
    def typeface(self, value):
        self._bullet_font.typeface = value
    
    @property
    def pitch_family(self):
        return self._bullet_font.pitchFamily

    @pitch_family.setter
    def pitch_family(self, value):
        self._bullet_font.pitchFamily = value
    

    @property
    def panose(self):
        return self._bullet_font.panose
    
    @panose.setter
    def panose(self, value):
        self._bullet_font.panose = value
    
    @property
    def charset(self):
        return self._bullet_font.charset
    
    @charset.setter
    def charset(self, value):
        self._bullet_font.charset = value

