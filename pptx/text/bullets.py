# encoding: utf-8

""" 
Bullet-related objects such as TextBullet,
TextBulletTypeface, TextBulletColor, and TextBulletSize.
"""

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
from pptx.util import Pt, Centipoints, Emu, lazyproperty

class TextBullet(object):
    """
    Provides acess to the current Text Bullet properties and provides
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
        return self._bullet.char

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
            bullet_cls = _Bullet
        
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

    @property
    def start_at(self):
        return self._autoNumBullet.startAt
    
    @start_at.setter
    def start_at(self, value):
        self._autoNumBullet.startAt = value


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



class TextBulletColor(object):
    """
    Provides access to the bullet color options
    """
    def __init__(self, parent, bullet_color_obj):
        super(TextBulletColor, self).__init__()
        self._parent = parent
        self._bullet_color = bullet_color_obj
    
    @classmethod
    def from_parent(cls, parent):
        """
        Return |TextBulletColor| object
        """
        bullet_color_elm = parent.eg_textBulletColor
        bullet_color = _TextBulletColor(bullet_color_elm)
        text_bullet_color = cls(parent, bullet_color)
        return text_bullet_color

    def follow_text(self):
        """
        Sets the TextBulletColor to _TextBulletColorFollowText
        """
        follow_text = self._parent.get_or_change_to_buClrTx()
        self._bullet_color = _TextBulletColorFollowText(follow_text)

    def set_color(self):
        """
        Sets the TextBulletColor to _TextBulletColorSpecific
        """
        bullet_color = self._parent.get_or_change_to_buClr()
        self._bullet_color = _TextBulletColorSpecific(bullet_color)

    @property
    def type(self):
        """ Return a string type """
        return self._bullet_color.type
    
    @property
    def color(self):
        """ Return the |ColorFormat| object """
        return self._bullet_color.color

class _TextBulletColor(object):
    """
    Object factory for TextBulletColor objects
    """

    def __new__(cls, xTextBulletColor):
        if isinstance(xTextBulletColor, CT_TextBulletColorFollowText):
            bullet_color_cls = _TextBulletColorFollowText
        elif isinstance(xTextBulletColor, CT_Color):
            bullet_color_cls = _TextBulletColorSpecific
        else:
            bullet_color_cls = _TextBulletColor
        
        return super(_TextBulletColor, cls).__new__(bullet_color_cls)
        
    @property
    def type(self):
        tmpl = ".type property must be implmented on %s"
        raise NotImplementedError(tmpl % self.__class__.__name__)

    @property
    def color(self):
        """Raise TypeError for types that do not override this property."""
        tmpl = (
            "TextBulletColor type %s has no color, call .set_color() first"
        )
        raise TypeError(tmpl % self.__class__.__name__)


class _TextBulletColorFollowText(_TextBulletColor):
    """
    Designates that the Bullet Color will match the accompanying paragraph text.
    """
    @property
    def type(self):
        return "TextBulletColorFollowText"

class _TextBulletColorSpecific(_TextBulletColor):
    """
    Designates a specific color for the bullet through a |ColorFormat| object.
    """
    def __init__(self, bullet_color):
        super(_TextBulletColorSpecific, self).__init__()
        self._bullet_color = bullet_color

    @property
    def type(self):
        return "TextBulletColorSpecific"

    @property
    def color(self):
        return ColorFormat.from_colorchoice_parent(self._bullet_color)



class TextBulletSize(object):
    """
    Provides access to the bullet size options
    """
    def __init__(self, parent, bullet_size_obj):
        super(TextBulletSize, self).__init__()
        self._parent = parent
        self._bullet_size = bullet_size_obj
    
    @classmethod
    def from_parent(cls, parent):
        """
        Return |TextBulletSize| object
        """
        bullet_size_elm = parent.eg_textBulletSize
        bullet_size = _TextBulletSize(bullet_size_elm)
        text_bullet_size = cls(parent, bullet_size)
        return text_bullet_size

    @property
    def type(self):
        """ Return a string type """
        return self._bullet_size.type

    def follow_text(self):
        """
        Sets the TextBulletSize to _TextBulletSizeFollowText
        """
        follow_text = self._parent.get_or_change_to_buSzTx()
        self._bullet_size = _TextBulletSizeFollowText(follow_text)

    def set_points(self, value=None):
        """
        Sets the TextBulletSize to _TextBulletSizePoints and sets the points value
        """
        points = self._parent.get_or_change_to_buSzPts()
        self._bullet_size = _TextBulletSizePoints(points)
        if value:
            self.points = value

    def set_percentage(self, value=None):
        """
        Sets the textBulletSize to _TextBulletSizePercent and sets the percent value
        """
        percentage = self._parent.get_or_change_to_buSzPct()
        self._bullet_size = _TextBulletSizePercent(percentage)
        if value:
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

class _TextBulletSize(object):
    """
    Object factory for TextBulletSize objects
    """

    def __new__(cls, xBulletSize):
        if isinstance(xBulletSize, CT_TextBulletSizeFollowText):
            bullet_size_cls = _TextBulletSizeFollowText
        elif isinstance(xBulletSize, CT_TextBulletSizePercent):
            bullet_size_cls = _TextBulletSizePercent
        elif isinstance(xBulletSize, CT_TextBulletSizePoints):
            bullet_size_cls = _TextBulletSizePoints
        else:
            bullet_size_cls = _TextBulletSize
        
        return super(_TextBulletSize, cls).__new__(bullet_size_cls)

    @property
    def points(self):
        """Raise TypeError for types that do not override this property."""
        tmpl = (
            "TextBulletSize type %s has no points property, call .set_points() first"
        )
        raise TypeError(tmpl % self.__class__.__name__)
    
    @points.setter
    def points(self, value):
        """Raise TypeError for types that do not override this property."""
        tmpl = (
            "TextBulletSize type %s has no points property, call .set_points() first"
        )
        raise TypeError(tmpl % self.__class__.__name__)

    @property
    def percentage(self):
        """Raise TypeError for types that do not override this property."""
        tmpl = (
            "TextBulletSize type %s has no percentage property, call .set_percentage() first"
        )
        raise TypeError(tmpl % self.__class__.__name__)

    @percentage.setter
    def percentage(self, value):
        """Raise TypeError for types that do not override this property."""
        tmpl = (
            "TextBulletSize type %s has no percentage property, call .set_percentage() first"
        )
        raise TypeError(tmpl % self.__class__.__name__)

    @property
    def type(self):
        tmpl = ".type property must be implmented on %s"
        raise NotImplementedError(tmpl % self.__class__.__name__)

class _TextBulletSizeFollowText(_TextBulletSize):
    """
    Designates that the Bullet Size will match the accompanying paragraph text.
    """
    @property
    def type(self):
        return "TextBulletSizeFollowText"

class _TextBulletSizePercent(_TextBulletSize):
    """ Proxies a `a: buSzPct` element. """

    def __init__(self, bullet_size):
        super(_TextBulletSizePercent, self).__init__()
        self._bullet_size = bullet_size

    @property
    def type(self):
        return "TextBulletSizePercent"

    @property
    def percentage(self):
        """ Returns the percentage value for the bullet size """
        return self._bullet_size.val
            
    @percentage.setter
    def percentage(self, value):
        """ Sets the percentage value for the bullet size """
        self._bullet_size.val = value
            

class _TextBulletSizePoints(_TextBulletSize):
    """ Proxies a `a: buSzPts` element. """
    
    def __init__(self, bullet_size):
        super(_TextBulletSizePoints, self).__init__()
        self._bullet_size = bullet_size


    @property
    def type(self):
        return "TextBulletSizePoints"


    @property
    def points(self):
        """ Returns the point value for the bullet size """
        if self._bullet_size.val is None:
            return None
        return Centipoints(self._bullet_size.val)
            
    @points.setter
    def points(self, value):
        """ Sets the points value for the bullet size """
        size = Emu(value).centipoints
        self._bullet_size.val = size
            


class TextBulletTypeface(object):
    """
    Provides access to the bullet font typeface options
    """
    def __init__(self, parent, bullet_typeface_obj):
        super(TextBulletTypeface, self).__init__()
        self._parent = parent
        self._bullet_typeface = bullet_typeface_obj
    
    @classmethod
    def from_parent(cls, parent):
        """
        Return |TextBulletTypeface| object
        """
        bullet_typeface_elm = parent.eg_textBulletTypeface
        bullet_typeface = _BulletTypeface(bullet_typeface_elm)
        text_bullet_typeface = cls(parent, bullet_typeface)
        return text_bullet_typeface

    def follow_text(self):
        """
        Sets the TextBulletTypeface to _BulletTypefaceFollowText
        """
        follow_text = self._parent.get_or_change_to_buFontTx()
        self._bullet_typeface = _BulletTypefaceFollowText(follow_text)

    def set_typeface(self, typeface="Arial"):
        """
        Sets the TextBulletTypeface to _BulletTypefaceSpecific
        """
        bullet_typeface = self._parent.get_or_change_to_buFont()
        self._bullet_typeface = _BulletTypefaceSpecific(bullet_typeface)
        self.typeface = typeface

    @property
    def typeface(self):
        return self._bullet_typeface.typeface

    @typeface.setter
    def typeface(self, value):
        self._bullet_typeface.typeface = value
    
    @property
    def pitch_family(self):
        return self._bullet_typeface.pitch_family

    @pitch_family.setter
    def pitch_family(self, value):
        self._bullet_typeface.pitch_family = value

    @property
    def panose(self):
        return self._bullet_typeface.panose
    
    @panose.setter
    def panose(self, value):
        self._bullet_typeface.panose = value
    
    @property
    def charset(self):
        return self._bullet_typeface.charset
    
    @charset.setter
    def charset(self, value):
        self._bullet_typeface.charset = value

    @property
    def type(self):
        """ Return a string type """
        return self._bullet_typeface.type
    
     

class _BulletTypeface(object):
    """
    Object factory for TextBulletTypeface objects
    """

    def __new__(cls, xBulletTypeface):
        if isinstance(xBulletTypeface, CT_TextBulletTypefaceFollowText):
            bullet_typeface_cls = _BulletTypefaceFollowText
        elif isinstance(xBulletTypeface, CT_TextFont):
            bullet_typeface_cls = _BulletTypefaceSpecific
        else:
            bullet_typeface_cls = _BulletTypeface
        
        return super(_BulletTypeface, cls).__new__(bullet_typeface_cls)
        
    @property
    def type(self):
        return "NoBulletTypeface"

    @property
    def typeface(self):
        """Raise TypeError for types that do not override this property."""
        tmpl = (
            "TextBulletTypeface type %s has no typeface property, call .set_typeface() first"
        )
        raise TypeError(tmpl % self.__class__.__name__)

    @typeface.setter
    def typeface(self, value):
        """Raise TypeError for types that do not override this property."""
        tmpl = (
            "TextBulletTypeface type %s has no typeface property, call .set_typeface() first"
        )
        raise TypeError(tmpl % self.__class__.__name__)
    
    @property
    def pitch_family(self):
        """Raise TypeError for types that do not override this property."""
        tmpl = (
            "TextBulletTypeface type %s has no pitch_family property, call .set_typeface() first"
        )
        raise TypeError(tmpl % self.__class__.__name__)

    @pitch_family.setter
    def pitch_family(self, value):
        """Raise TypeError for types that do not override this property."""
        tmpl = (
            "TextBulletTypeface type %s has no pitch_family property, call .set_typeface() first"
        )
        raise TypeError(tmpl % self.__class__.__name__)

    @property
    def panose(self):
        """Raise TypeError for types that do not override this property."""
        tmpl = (
            "TextBulletTypeface type %s has no panose property, call .set_typeface() first"
        )
        raise TypeError(tmpl % self.__class__.__name__)
    
    @panose.setter
    def panose(self, value):
        """Raise TypeError for types that do not override this property."""
        tmpl = (
            "TextBulletTypeface type %s has no panose property, call .set_typeface() first"
        )
        raise TypeError(tmpl % self.__class__.__name__)
    
    @property
    def charset(self):
        """Raise TypeError for types that do not override this property."""
        tmpl = (
            "TextBulletTypeface type %s has no charset property, call .set_typeface() first"
        )
        raise TypeError(tmpl % self.__class__.__name__)
    
    @charset.setter
    def charset(self, value):
        """Raise TypeError for types that do not override this property."""
        tmpl = (
            "TextBulletTypeface type %s has no charset property, call .set_typeface() first"
        )
        raise TypeError(tmpl % self.__class__.__name__)


class _BulletTypefaceFollowText(_BulletTypeface):
    """
    Designates that the BulletTypeface will match the accompanying paragraph text.
    """
    @property
    def type(self):
        return "BulletTypefaceFollowText"

class _BulletTypefaceSpecific(_BulletTypeface):
    """
    Designates the specific BulletTypeface characteristics
    """
    def __init__(self, bullet_typeface):
        super(_BulletTypefaceSpecific, self).__init__()
        self._bullet_typeface = bullet_typeface

    @property
    def type(self):
        return "BulletTypefaceSpecific"

    @property
    def typeface(self):
        return self._bullet_typeface.typeface

    @typeface.setter
    def typeface(self, value):
        self._bullet_typeface.typeface = value
    
    @property
    def pitch_family(self):
        return self._bullet_typeface.pitchFamily

    @pitch_family.setter
    def pitch_family(self, value):
        self._bullet_typeface.pitchFamily = value
    

    @property
    def panose(self):
        return self._bullet_typeface.panose
    
    @panose.setter
    def panose(self, value):
        self._bullet_typeface.panose = value
    
    @property
    def charset(self):
        return self._bullet_typeface.charset
    
    @charset.setter
    def charset(self, value):
        self._bullet_typeface.charset = value

