# encoding: utf-8

""" Bullet-related objects such as BulletFont and BulletChar."""

from __future__ import absolute_import, division, print_function, unicode_literals

from pptx.enum.text import AUTO_NUMBER_SCHEME
from pptx.oxml.text import CT_TextBulletAutoNumber, CT_TextNoBullet, \
    CT_TextBlipBullet, CT_TextCharBullet

class TextBullet(object):
    """
    Provides acess ot the current Text Bullet properties and provies
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
        self._bullet.character = char

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
    def character(self):
        """ String used as the bullet for a CharBullet"""
        if self.type != "CharBullet":
            raise TypeError("TextBullet is not of type CharBullet")
        return self._bullet.character

    @character.setter
    def character(self, value):
        """ String used as the bullet for a CharBullet"""
        if self.type != "CharBullet":
            raise TypeError("TextBullet is not of type CharBullet")
        self._bullet.character = value


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
    def character(self, value):
        self._charBullet.char = str(value)


class _PictureBullet(_Bullet):
    """ Picture Bullets are not fully implemented at this time """
    @property
    def type(self):
        return "PictureBullet"
