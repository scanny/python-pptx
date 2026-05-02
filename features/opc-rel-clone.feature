Feature: Cross-part relationship cloning helper
  In order to duplicate slides, shapes, and charts across parts
  As a developer building on python-pptx
  I need a PartRelationshipCloner that copies an XML element plus every
  Part it references (images, media, charts, hyperlinks, OLE objects),
  rewrites its rIds, and returns the clone ready to insert on the target part


  Scenario: PartRelationshipCloner clones a picture element and its image part
    Given a slide containing a picture
      And a second slide with no picture
     When I use PartRelationshipCloner to copy the picture element to the second slide's part
     Then the returned element has a fresh r:embed value
      And the second slide's part is related to a new image part with the same blob
      And the original slide's picture still points at its original image part
