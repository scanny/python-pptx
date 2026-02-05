#!/usr/bin/env python3
"""
Create a PowerPoint slide with blue background and white quadratic equation.
"""

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml import parse_xml
from lxml import etree

def create_quadratic_equation_slide():
    """Create a slide with blue background and white quadratic equation."""

    # Create a new presentation
    prs = Presentation()

    # Add a blank slide
    slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)

    # Set blue background for slide
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(0, 51, 102)  # Dark blue

    # Add the quadratic formula using a simpler approach
    # Create a textbox shape and add OMML to it directly

    # Create a textbox shape for the equation
    equation_box = slide.shapes.add_textbox(
        left=Inches(1), top=Inches(2.5), width=Inches(8), height=Inches(2)
    )

    # Get the text body and build the complete OMML structure as XML string
    tx_body = equation_box.text_frame

    # Build the complete mc:AlternateContent structure exactly like the working demo
    complete_omml_xml = '''<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="2" name="Math Equation">
                <a:extLst>
                  <a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">
                    <a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{6F2CE406-D9A3-8D9C-6BAC-B21E81A4A63E}"/>
                  </a:ext>
                </a:extLst>
              </p:cNvPr>
              <p:cNvSpPr txBox="1"/>
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

    # Parse the complete XML and add to the slide's spTree
    alt_content_element = parse_xml(complete_omml_xml)

    # Find the spTree element where shapes are stored
    sp_tree = slide.element.xpath('.//*[local-name() = "spTree"]')[0]
    sp_tree.append(alt_content_element)

    # Add a title with white text
    title_box = slide.shapes.add_textbox(
        left=Inches(1), top=Inches(1), width=Inches(8), height=Inches(1)
    )
    title_frame = title_box.text_frame
    title_frame.text = "Quadratic Formula"
    title_para = title_frame.paragraphs[0]
    title_run = title_para.runs[0]
    title_run.font.color.rgb = RGBColor(255, 255, 255)  # White

    # Add subtitle with white text
    subtitle_box = slide.shapes.add_textbox(
        left=Inches(1), top=Inches(4.5), width=Inches(8), height=Inches(1)
    )
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.text = "x = (-b ± √(b² - 4ac)) / 2a"
    subtitle_para = subtitle_frame.paragraphs[0]
    subtitle_run = subtitle_para.runs[0]
    subtitle_run.font.color.rgb = RGBColor(255, 255, 255)  # White

    # Save presentation
    output_path = "/Users/marvin/Desktop/Projects/python-pptx/temp/quadratic_equation_demo.py"
    prs.save("/Users/marvin/Desktop/Projects/python-pptx/temp/quadratic_equation_demo.pptx")
    print(f"Presentation saved to: /Users/marvin/Desktop/Projects/python-pptx/temp/quadratic_equation_demo.pptx")

if __name__ == "__main__":
    create_quadratic_equation_slide()
