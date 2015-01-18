.. _image_api:

Image
=====

An image depicted in a |Picture| shape can be accessed using its
:attr:`~.Picture.image` property. The |Image| object provides access to
detailed properties of the image itself, including the bytes of the image
file itself.


|Image| objects
---------------

The |Image| object is encountered as the :attr:`~Picture.image` property of
|Picture|.

.. autoclass:: pptx.parts.image.Image()
   :members:
   :exclude-members: from_blob, from_file
