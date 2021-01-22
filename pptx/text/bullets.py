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
)
from pptx.oxml.dml.color import CT_Color
from pptx.dml.color import ColorFormat, _SchemeColor, RGBColor

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
        buChar = self._parent.get_or_change_to_buChar()
        self._bullet = _CharBullet(buChar)
        self.char = char

    @property
    def char_type(self):
        """ Format of the AutoNumber Bullets """
        if self.type != "AutoNumBullet":
            raise TypeError("TextBullet is not of type AutoNumber")
        return self._bullet.char_type

    @char_type.setter
    def char_type(self, value):
        """ Set's the Format of the AutoNumber Bullets """
        if self.type != "AutoNumBullet":
            raise TypeError("TextBullet is not of type AutoNumber")
        self._bullet.char_type = value

    @property
    def start_at(self):
        """ Starting Value for AutoNumber Bullets """
        if self.type != "AutoNumBullet":
            raise TypeError("TextBullet is not of type AutoNumber")
        return self._bullet.start_at
    
    @start_at.setter
    def start_at(self, value):
        """ Sets starting Value for AutoNumber Bullets """
        if self.type != "AutoNumBullet":
            raise TypeError("TextBullet is not of type AutoNumber")
        self._bullet.start_at = value
    
    @property
    def char(self):
        """ String used as the bullet for a CharBullet"""
        if self.type != "CharBullet":
            raise TypeError("TextBullet is not of type CharBullet")
        return self._bullet.character

    @char.setter
    def char(self, value):
        """ String used as the bullet for a CharBullet"""
        if self.type != "CharBullet":
            raise TypeError("TextBullet is not of type CharBullet")
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

    @property
    def char_type(self):
        """ Raise TypeError for types that do not override this property"""
        tmpl = "TextBullet type %s has no char_type property. call .auto_number() first"
        raise TypeError(tmpl % self.__class__.__name__)

    @property
    def start_at(self):
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
        # buClr = self._bullet_color.get_or_add_buClr()
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
        Return |BulletColor| object
        """

        bullet_size_elm = parent.eg_textBulletSize
        bullet_size = _BulletSize(bullet_sizer_elm)
        text_bullet_size = cls(parent, bullet_size)
        return text_bullet_size


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

class _BulletSizeFollowText(_BulletSize):
    """
    Designates that the Bullet Size will match the accompanying paragraph text.
    """
    @property
    def type(self):
        return "BulletSizeFollowText"

class _BulletSizePercent(_BulletSize):
    """
    """
    @property
    def type(self):
        return "BulletSizePercent"

class _BulletSizePoints(_BulletSize):
    """
    """
    @property
    def type(self):
        return "BulletSizePoints"



