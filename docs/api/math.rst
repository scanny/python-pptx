Math
====

The following classes provide access to mathematical equations and OMML (Office Math Markup Language) functionality in PowerPoint presentations.

Overview
--------

The math module provides a high-level interface for adding mathematical equations to PowerPoint slides. It handles the complex XML structure required by PowerPoint, including the `mc:AlternateContent` wrapper, `a14:m` extension elements, and automatic formatting with Cambria Math font.

Math objects
-------------

|Math| objects provide access to the mathematical content of a math shape.

.. autoclass:: pptx.math.Math()
   :members:
   :inherited-members:

   The Math class provides high-level access to OMML content within a math shape. It handles automatic formatting of text runs with PowerPoint-compatible properties including Cambria Math font, language settings, and color schemes.

   **Key Methods:**

   * ``get_omml()`` - Get the OMML XML string for the equation
   * ``set_omml(omml_xml)`` - Replace the current OMML with new XML content
   * ``add_omml(omml_xml)`` - Add OMML content with automatic formatting applied

   **Example Usage:**

   .. code-block:: python

      # Direct XML approach (as used in the working demo)
      complete_omml_xml = '''<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
        <mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">
          <p:sp>
            <p:txBody>
              <a:bodyPr wrap="none">
                <a:spAutoFit/>
              </a:bodyPr>
              <a:lstStyle/>
              <a:p>
                <a:pPr/>
                <a14:m>
                  <m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
                    <m:oMathParaPr>
                      <m:jc m:val="centerGroup"/>
                    </m:oMathParaPr>
                    <m:oMath>
                      <m:r>
                        <a:rPr lang="en-US" sz="2800" b="0" i="1" smtClean="0">
                          <a:solidFill>
                            <a:schemeClr val="bg1"/>
                          </a:solidFill>
                          <a:latin typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/>
                        </a:rPr>
                        <m:t>x = </m:t>
                      </m:r>
                      <m:f>
                        <m:num>
                          <m:r>
                            <a:rPr lang="en-US" sz="2800" b="0" i="1" smtClean="0">
                              <a:solidFill>
                                <a:schemeClr val="bg1"/>
                              </a:solidFill>
                              <a:latin typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/>
                            </a:rPr>
                            <m:t>-b ± </m:t>
                          </m:r>
                          <m:rad>
                            <m:radPr/>
                            <m:deg/>
                            <m:e>
                              <m:r>
                                <a:rPr lang="en-US" sz="2800" b="0" i="1" smtClean="0">
                                  <a:solidFill>
                                    <a:schemeClr val="bg1"/>
                                  </a:solidFill>
                                  <a:latin typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/>
                                </a:rPr>
                                <m:t>b² - 4ac</m:t>
                              </m:r>
                            </m:e>
                          </m:rad>
                        </m:num>
                        <m:den>
                          <m:r>
                            <a:rPr lang="en-US" sz="2800" b="0" i="1" smtClean="0">
                              <a:solidFill>
                                <a:schemeClr val="bg1"/>
                              </a:solidFill>
                              <a:latin typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/>
                            </a:rPr>
                            <m:t>2a</m:t>
                          </m:r>
                        </m:den>
                      </m:f>
                    </m:oMath>
                  </m:oMathPara>
                </a14:m>
                <a:endParaRPr lang="en-US" sz="2800" dirty="0">
                  <a:solidFill>
                    <a:schemeClr val="bg1"/>
                  </a:solidFill>
                </a:endParaRPr>
              </a:p>
            </p:txBody>
          </p:sp>
        </mc:Choice>
        <mc:Fallback>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="3" name="TextBox 3">
                <a:extLst>
                  <a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">
                    <a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{6F2CE406-D9A3-8D9C-6BAC-B21E81A4A63E}"/>
                  </a:ext>
                </a:extLst>
              </p:cNvPr>
              <p:cNvSpPr txBox="1">
                <a:spLocks noRot="1" noChangeAspect="1" noMove="1" noResize="1" noEditPoints="1" noAdjustHandles="1" noChangeArrowheads="1" noChangeShapeType="1" noTextEdit="1"/>
              </p:cNvSpPr>
              <p:nvPr/>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm>
                <a:off x="914400" y="1828800"/>
                <a:ext cx="4572000" cy="914400"/>
              </a:xfrm>
              <a:prstGeom prst="rect">
                <a:avLst/>
              </a:prstGeom>
              <a:noFill/>
            </p:spPr>
            <p:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:r>
                    <a:rPr lang="en-US">
                      <a:noFill/>
                    </a:rPr>
                    <a:t> </a:t>
                </a:r>
              </a:p>
            </p:txBody>
          </p:sp>
        </mc:Fallback>
      </mc:AlternateContent>'''

      from pptx.oxml import parse_xml
      alt_content_element = parse_xml(complete_omml_xml)
      sp_tree = slide.element.xpath('.//*[local-name() = "spTree"]')[0]
      sp_tree.append(alt_content_element)

MathShape objects
------------------

|MathShape| objects represent mathematical equations on a slide.

.. autoclass:: pptx.shapes.math.MathShape()
   :members:
   :inherited-members:

   MathShape provides the interface between PowerPoint shapes and mathematical content. Each MathShape contains a Math object that manages the OMML content.

   **Example Usage:**

   .. code-block:: python

      # Access math content
      omml_xml = math_shape.math.get_omml()
      math_shape.math.set_omml(new_omml)

ShapeTree Integration
-------------------

.. automethod:: pptx.shapes.shapetree.ShapeTree.add_math_equation()

   Add a mathematical equation to the slide. This method creates a textbox shape with math extension markers and returns a MathShape object.

   **Parameters:**

   * ``left`` (Length, optional) - Horizontal position (default: 1 inch)
   * ``top`` (Length, optional) - Vertical position (default: 0.75 inch)
   * ``width`` (Length, optional) - Width of equation box (default: 2 inches)
   * ``height`` (Length, optional) - Height of equation box (default: 1 inch)

   **Returns:**

   * ``MathShape`` object for accessing and manipulating the equation

OMML Structure Requirements
-----------------------

PowerPoint requires a specific XML structure for mathematical equations:

**Required Structure:**

.. code-block:: xml

   <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
     <mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">
       <p:sp>
         <!-- PowerPoint shape with txBox="1" -->
         <p:txBody>
           <a:p>
             <a14:m>
               <m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
                 <m:oMath>
                   <!-- Your equation content here -->
                 </m:oMath>
               </m:oMathPara>
             </a14:m>
           </a:p>
         </p:txBody>
       </p:sp>
     </mc:Choice>
     <mc:Fallback>
       <!-- Fallback content for older PowerPoint versions -->
     </mc:Fallback>
   </mc:AlternateContent>

**Key Components:**

* **Note:** Powerpoint and Word use different XML structures for math equations.  They share most of the XML format, but have unique tags for other portions.
* ``mc:AlternateContent`` - Compatibility wrapper for Office 2010+ features
* ``a14:m`` - Office 2010 extension that allows OMML in DrawingML
* ``m:oMathPara`` - Container for equation formatting and alignment
* ``m:oMath`` - The actual mathematical content
* ``a:rPr`` - Run properties for font, color, and language formatting

Limitations
-----------

* Users must provide fully-compliant PowerPoint OMML strings.
* Direct XML manipulation requires understanding of PowerPoint's XML structure.
* Future update should provide OMML builder support which is docx and pptx agnostic.