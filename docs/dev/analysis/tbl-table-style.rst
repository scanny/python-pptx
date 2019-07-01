
How table styles work
---------------------

* PowerPoint allows many formatting characteristics of a table to be set at
  once by assigning a *table style* to a table.

* PowerPoint includes an array of built-in table styles that a user can
  browse in the *table style gallery*. These styles are built into the
  PowerPoint *application*, and are only added to a `.pptx` file on first
  use.

* Zero or more table-styles can appear in the `ppt/tableStyles.xml` part.
  Each is keyed with a UUID, by which it is referred to by a particular table
  that uses that style.

* A table may be assigned a table style by placing the matching UUID in the
  `a:tbl/a:tblPr/a:tableStyleId` element text.

* A default table style may be specified by placing its id in
  `a:tblStyleLst/@def` in the `tableStyles.xml` part.
