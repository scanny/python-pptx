.. _MsoAutoShapeType:

``MSO_AUTO_SHAPE_TYPE``
=======================

Specifies a type of AutoShape, e.g. DOWN_ARROW

Alias: ``MSO_SHAPE``

Example::

    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Inches

    left = top = width = height = Inches(1.0)
    slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
    )

----

ACTION_BUTTON_BACK_OR_PREVIOUS
    Back or Previous button. Supports mouse-click and mouse-over actions

ACTION_BUTTON_BEGINNING
    Beginning button. Supports mouse-click and mouse-over actions.

ACTION_BUTTON_CUSTOM
    Button with no default picture or text. Supports mouse-click and mouse-over
    action.

ACTION_BUTTON_DOCUMENT
    Document button. Supports mouse-click and mouse-over actions.

ACTION_BUTTON_END
    End button. Supports mouse-click and mouse-over actions.

ACTION_BUTTON_FORWARD_OR_NEXT
    Forward or Next button. Supports mouse-click and mouse-over actions.

ACTION_BUTTON_HELP
    Help button. Supports mouse-click and mouse-over actio.

ACTION_BUTTON_HOME
    Home button. Supports mouse-click and mouse-over actio.

ACTION_BUTTON_INFORMATION
    Information button. Supports mouse-click and mouse-over actions.

ACTION_BUTTON_MOVIE
    Movie button. Supports mouse-click and mouse-over actions.

ACTION_BUTTON_RETURN
    Return button. Supports mouse-click and mouse-over actions.

ACTION_BUTTON_SOUND
    Sound button. Supports mouse-click and mouse-over actions.

ARC
    Arc

BALLOON
    Rounded Rectangular Callout

BENT_ARROW
    Block arrow that follows a curved 90-degree angle

BENT_UP_ARROW
    Block arrow that follows a sharp 90-degree angle. Points up by default.

BEVEL
    Bevel

BLOCK_ARC
    Block arc

CAN
    Can

CHART_PLUS
    Chart Plus

CHART_STAR
    Chart Star

CHART_X
    Chart X

CHEVRON
    Chevron

CHORD
    Geometric chord shape

CIRCULAR_ARROW
    Block arrow that follows a curved 180-degree angle

CLOUD
    Cloud

CLOUD_CALLOUT
    Cloud callout

CORNER
    Corner

CORNER_TABS
    Corner Tabs

CROSS
    Cross

CUBE
    Cube

CURVED_DOWN_ARROW
    Block arrow that curves down

CURVED_DOWN_RIBBON
    Ribbon banner that curves down

CURVED_LEFT_ARROW
    Block arrow that curves left

CURVED_RIGHT_ARROW
    Block arrow that curves right

CURVED_UP_ARROW
    Block arrow that curves up

CURVED_UP_RIBBON
    Ribbon banner that curves up

DECAGON
    Decagon

DIAGONAL_STRIPE
    Diagonal Stripe

DIAMOND
    Diamond

DODECAGON
    Dodecagon

DONUT
    Donut

DOUBLE_BRACE
    Double brace

DOUBLE_BRACKET
    Double bracket

DOUBLE_WAVE
    Double wave

DOWN_ARROW
    Block arrow that points down

DOWN_ARROW_CALLOUT
    Callout with arrow that points down

DOWN_RIBBON
    Ribbon banner with center area below ribbon ends

EXPLOSION1
    Explosion

EXPLOSION2
    Explosion

FLOWCHART_ALTERNATE_PROCESS
    Alternate process flowchart symbol

FLOWCHART_CARD
    Card flowchart symbol

FLOWCHART_COLLATE
    Collate flowchart symbol

FLOWCHART_CONNECTOR
    Connector flowchart symbol

FLOWCHART_DATA
    Data flowchart symbol

FLOWCHART_DECISION
    Decision flowchart symbol

FLOWCHART_DELAY
    Delay flowchart symbol

FLOWCHART_DIRECT_ACCESS_STORAGE
    Direct access storage flowchart symbol

FLOWCHART_DISPLAY
    Display flowchart symbol

FLOWCHART_DOCUMENT
    Document flowchart symbol

FLOWCHART_EXTRACT
    Extract flowchart symbol

FLOWCHART_INTERNAL_STORAGE
    Internal storage flowchart symbol

FLOWCHART_MAGNETIC_DISK
    Magnetic disk flowchart symbol

FLOWCHART_MANUAL_INPUT
    Manual input flowchart symbol

FLOWCHART_MANUAL_OPERATION
    Manual operation flowchart symbol

FLOWCHART_MERGE
    Merge flowchart symbol

FLOWCHART_MULTIDOCUMENT
    Multi-document flowchart symbol

FLOWCHART_OFFLINE_STORAGE
    Offline Storage

FLOWCHART_OFFPAGE_CONNECTOR
    Off-page connector flowchart symbol

FLOWCHART_OR
    "Or" flowchart symbol

FLOWCHART_PREDEFINED_PROCESS
    Predefined process flowchart symbol

FLOWCHART_PREPARATION
    Preparation flowchart symbol

FLOWCHART_PROCESS
    Process flowchart symbol

FLOWCHART_PUNCHED_TAPE
    Punched tape flowchart symbol

FLOWCHART_SEQUENTIAL_ACCESS_STORAGE
    Sequential access storage flowchart symbol

FLOWCHART_SORT
    Sort flowchart symbol

FLOWCHART_STORED_DATA
    Stored data flowchart symbol

FLOWCHART_SUMMING_JUNCTION
    Summing junction flowchart symbol

FLOWCHART_TERMINATOR
    Terminator flowchart symbol

FOLDED_CORNER
    Folded corner

FRAME
    Frame

FUNNEL
    Funnel

GEAR_6
    Gear 6

GEAR_9
    Gear 9

HALF_FRAME
    Half Frame

HEART
    Heart

HEPTAGON
    Heptagon

HEXAGON
    Hexagon

HORIZONTAL_SCROLL
    Horizontal scroll

ISOSCELES_TRIANGLE
    Isosceles triangle

LEFT_ARROW
    Block arrow that points left

LEFT_ARROW_CALLOUT
    Callout with arrow that points left

LEFT_BRACE
    Left brace

LEFT_BRACKET
    Left bracket

LEFT_CIRCULAR_ARROW
    Left Circular Arrow

LEFT_RIGHT_ARROW
    Block arrow with arrowheads that point both left and right

LEFT_RIGHT_ARROW_CALLOUT
    Callout with arrowheads that point both left and right

LEFT_RIGHT_CIRCULAR_ARROW
    Left Right Circular Arrow

LEFT_RIGHT_RIBBON
    Left Right Ribbon

LEFT_RIGHT_UP_ARROW
    Block arrow with arrowheads that point left, right, and up

LEFT_UP_ARROW
    Block arrow with arrowheads that point left and up

LIGHTNING_BOLT
    Lightning bolt

LINE_CALLOUT_1
    Callout with border and horizontal callout line

LINE_CALLOUT_1_ACCENT_BAR
    Callout with vertical accent bar

LINE_CALLOUT_1_BORDER_AND_ACCENT_BAR
    Callout with border and vertical accent bar

LINE_CALLOUT_1_NO_BORDER
    Callout with horizontal line

LINE_CALLOUT_2
    Callout with diagonal straight line

LINE_CALLOUT_2_ACCENT_BAR
    Callout with diagonal callout line and accent bar

LINE_CALLOUT_2_BORDER_AND_ACCENT_BAR
    Callout with border, diagonal straight line, and accent bar

LINE_CALLOUT_2_NO_BORDER
    Callout with no border and diagonal callout line

LINE_CALLOUT_3
    Callout with angled line

LINE_CALLOUT_3_ACCENT_BAR
    Callout with angled callout line and accent bar

LINE_CALLOUT_3_BORDER_AND_ACCENT_BAR
    Callout with border, angled callout line, and accent bar

LINE_CALLOUT_3_NO_BORDER
    Callout with no border and angled callout line

LINE_CALLOUT_4
    Callout with callout line segments forming a U-shape

LINE_CALLOUT_4_ACCENT_BAR
    Callout with accent bar and callout line segments forming a U-shape

LINE_CALLOUT_4_BORDER_AND_ACCENT_BAR
    Callout with border, accent bar, and callout line segments forming a
    U-shape

LINE_CALLOUT_4_NO_BORDER
    Callout with no border and callout line segments forming a U-shape.

LINE_INVERSE
    Straight Connector

MATH_DIVIDE
    Division

MATH_EQUAL
    Equal

MATH_MINUS
    Minus

MATH_MULTIPLY
    Multiply

MATH_NOT_EQUAL
    Not Equal

MATH_PLUS
    Plus

MOON
    Moon

NO_SYMBOL
    "No" symbol

NON_ISOSCELES_TRAPEZOID
    Non-isosceles Trapezoid

NOTCHED_RIGHT_ARROW
    Notched block arrow that points right

OCTAGON
    Octagon

OVAL
    Oval

OVAL_CALLOUT
    Oval-shaped callout

PARALLELOGRAM
    Parallelogram

PENTAGON
    Pentagon

PIE
    Pie

PIE_WEDGE
    Pie

PLAQUE
    Plaque

PLAQUE_TABS
    Plaque Tabs

QUAD_ARROW
    Block arrows that point up, down, left, and right

QUAD_ARROW_CALLOUT
    Callout with arrows that point up, down, left, and right

RECTANGLE
    Rectangle

RECTANGULAR_CALLOUT
    Rectangular callout

REGULAR_PENTAGON
    Pentagon

RIGHT_ARROW
    Block arrow that points right

RIGHT_ARROW_CALLOUT
    Callout with arrow that points right

RIGHT_BRACE
    Right brace

RIGHT_BRACKET
    Right bracket

RIGHT_TRIANGLE
    Right triangle

ROUND_1_RECTANGLE
    Round Single Corner Rectangle

ROUND_2_DIAG_RECTANGLE
    Round Diagonal Corner Rectangle

ROUND_2_SAME_RECTANGLE
    Round Same Side Corner Rectangle

ROUNDED_RECTANGLE
    Rounded rectangle

ROUNDED_RECTANGULAR_CALLOUT
    Rounded rectangle-shaped callout

SMILEY_FACE
    Smiley face

SNIP_1_RECTANGLE
    Snip Single Corner Rectangle

SNIP_2_DIAG_RECTANGLE
    Snip Diagonal Corner Rectangle

SNIP_2_SAME_RECTANGLE
    Snip Same Side Corner Rectangle

SNIP_ROUND_RECTANGLE
    Snip and Round Single Corner Rectangle

SQUARE_TABS
    Square Tabs

STAR_10_POINT
    10-Point Star

STAR_12_POINT
    12-Point Star

STAR_16_POINT
    16-point star

STAR_24_POINT
    24-point star

STAR_32_POINT
    32-point star

STAR_4_POINT
    4-point star

STAR_5_POINT
    5-point star

STAR_6_POINT
    6-Point Star

STAR_7_POINT
    7-Point Star

STAR_8_POINT
    8-point star

STRIPED_RIGHT_ARROW
    Block arrow that points right with stripes at the tail

SUN
    Sun

SWOOSH_ARROW
    Swoosh Arrow

TEAR
    Teardrop

TRAPEZOID
    Trapezoid

U_TURN_ARROW
    Block arrow forming a U shape

UP_ARROW
    Block arrow that points up

UP_ARROW_CALLOUT
    Callout with arrow that points up

UP_DOWN_ARROW
    Block arrow that points up and down

UP_DOWN_ARROW_CALLOUT
    Callout with arrows that point up and down

UP_RIBBON
    Ribbon banner with center area above ribbon ends

VERTICAL_SCROLL
    Vertical scroll

WAVE
    Wave
