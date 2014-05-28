import os, glob, sys, time, math, zipfile, MySQLdb
import xml.etree.ElementTree as ET
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# simple class to create objects
class SimpleClass(object):
	pass

# global variables
OVAL_HEIGHT = 950000 # height of oval shape
OVAL_WIDTH = 950000 # width of oval shape
HALF_OVAL_HEIGHT = round(OVAL_HEIGHT / 2)
HALF_OVAL_WIDTH = round(OVAL_WIDTH / 2)
SLIDE_HEIGHT = 6858000 # assume standard 10" x 7.5" slide (4:3 aspect ratio)
SLIDE_WIDTH = 9144000 # assume standard 10" x 7.5" slide (4:3 aspect ratio)
MAX_TOP = SLIDE_HEIGHT - OVAL_HEIGHT # furthest top of oval shape where oval is completely on slide
MAX_LEFT = SLIDE_WIDTH - OVAL_WIDTH # furthest left of oval shape where oval is completely on slide

# check for proper number of command line arguments
if len(sys.argv) < 4:
	print "Usage: python export.py <client_id> <threshold> <output_path>"
	print "Example: python export.py 1 1 /dl/metaphoria-ruby/public/output.pptx"
	exit()

# get client_id from command line
client_id = sys.argv[1]
threshold = int(sys.argv[2])
output_path = sys.argv[3]

# parse path and filename
index = output_path.rfind("/") + 1
output_folder = output_path[0:index]
filename = output_path[index:]

# variables
current_min_top = -1
current_max_top = -1
current_min_left = -1
current_max_left = -1
x_scale = 1
y_scale = 1
constructs = []
links = []
current_link_id = 1000

# remove old .pptx files
old_pptxs = glob.glob("*.pptx")
for p in old_pptxs:
	os.remove(p)

# connect to database and get cursor
db = MySQLdb.connect(host="localhost", user="user", passwd="notreallymypassword", db="database")
cur = db.cursor()

# create a new presentation
pptx = Presentation()

# set up new slide
BLANK_SLIDE_LAYOUT = pptx.slide_layouts[6]
slide = pptx.slides.add_slide(BLANK_SLIDE_LAYOUT)

# query database for construct links
query = "SELECT * FROM database"
cur.execute(query, [client_id, client_id])

# for each link
for row in cur.fetchall():

	# get link and set attributes
	link = SimpleClass()
	link.weight = int(row[0])
	link.start = row[1]
	link.end = row[2]
	link.id = "-1"

	# append if we will draw this link
	if link.weight >= threshold:
		links.append(link)

# query database for constructs
query = "SELECT * FROM database"
cur.execute(query, client_id)

# add constructs to array
for row in cur.fetchall():

	# create construct object
	c = SimpleClass()
	c.id = row[0]
	c.word = row[2]
	c.type = row[3]
	c.top = row[4]
	c.left = row[5]

	# for all links
	for link in links:

		# if link references this construct
		if (link.start == c.id) or (link.end == c.id):

			# add construct to array
			constructs.append(c)

			# find mins and maxs for scaling
			if (current_min_left == -1) or (c.left < current_min_left):
				current_min_left = c.left
			if (current_max_left == -1) or (c.left > current_max_left):
				current_max_left = c.left
			if (current_min_top == -1) or (c.top < current_min_top):
				current_min_top = c.top
			if (current_max_top == -1) or (c.top > current_max_top):
				current_max_top = c.top

			# break from loop
			break

# calculate maxs and scales (because we will offset to force the mins to 0)
current_max_left -= current_min_left
current_max_top -= current_min_top
x_scale = MAX_LEFT / current_max_left
y_scale = MAX_TOP / current_max_top

# scale and add nodes
for c in constructs:

	# calculate shape left and top
	c.left = (c.left - current_min_left) * x_scale
	c.top = (c.top - current_min_top) * y_scale

	# create oval shape and set id
	shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, c.left, c.top, OVAL_WIDTH, OVAL_HEIGHT)
	c.xml_id = shape.id
	current_link_id = shape.id + 1

	# clear current text and add new placeholder
	shape.textframe.clear()
	p = shape.textframe.paragraphs[0]
	run = p.add_run()
	run.text = c.word

	# set font, size, and color
	font = run.font
	font.name = "Calibri"
	font.size = Pt(10.5)
	font.color.rgb = RGBColor(50, 50, 50) # grey

	# set shape fill
	fill = shape.fill
	fill.solid()
	if c.type == "emotional": # blue
		fill.fore_color.rgb = RGBColor(204, 224, 255)
	elif c.type == "psychosocial": # green
		fill.fore_color.rgb = RGBColor(217, 228, 191)
	elif c.type == "functional": # peach
		fill.fore_color.rgb = RGBColor(255, 231, 219)
	elif c.type == "attribute": # purple
		fill.fore_color.rgb = RGBColor(226, 215, 243)
	elif c.type == "other": # white
		fill.fore_color.rgb = RGBColor(255, 255, 255)
	else: # white
		fill.fore_color.rgb = RGBColor(255, 255, 255)

# save file
os.chdir(output_folder)
pptx.save(filename)

# unzip .pptx to temporary space and remove file
timestamp = str(time.time()).replace(".", "")
zip_file = zipfile.ZipFile(filename, "r")
zip_file.extractall(os.path.join("/tmp", timestamp))
zip_file.close
os.remove(filename)

# register necessary namespaces
ET.register_namespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
ET.register_namespace("p", "http://schemas.openxmlformats.org/presentationml/2006/main")
ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

# parse xml document and find shape tree
tree = ET.parse(os.path.join("/tmp", timestamp, "ppt", "slides", "slide1.xml"))
root = tree.getroot()

# get necessary elements
iterator = root.iter("{http://schemas.openxmlformats.org/presentationml/2006/main}spTree")
shapeTree = iterator.next()
iterator = root.iter("{http://schemas.openxmlformats.org/presentationml/2006/main}grpSpPr")
groupShapeProperties = iterator.next()

# add properties
twodTransform = ET.SubElement(groupShapeProperties, "a:xfrm")
offset = ET.SubElement(twodTransform, "a:off")
offset.set("x", "0")
offset.set("y", "0")
extents = ET.SubElement(twodTransform, "a:ext")
extents.set("cx", "0")
extents.set("cy", "0")
childOffset = ET.SubElement(twodTransform, "a:chOff")
childOffset.set("x", "0")
childOffset.set("y", "0")
childExtents = ET.SubElement(twodTransform, "a:chExt")
childExtents.set("cx", "0")
childExtents.set("cy", "0")

# for each link
for link in links:

	# variables
	start_construct = None
	end_construct = None
	flip_v = False
	flip_h = False

	# assign and increment link id
	link.id = current_link_id
	current_link_id += 1

	# get start and end constructs
	for c in constructs:
		if (c.id == link.start):
			start_construct = c
		elif (c.id == link.end):
			end_construct = c
		if start_construct != None and end_construct != None:
			break

	# if start and end constructs can't be found, do not run this iteration
	if start_construct == None or end_construct == None:
		continue

	# calculate x and y differences
	x_diff = abs(start_construct.left - end_construct.left)
	y_diff = abs(start_construct.top - end_construct.top)

	# start node is completely above end node
	if (start_construct.top + OVAL_HEIGHT) < end_construct.top:

		# start node is completely above and completely left of end node
		if (start_construct.left + OVAL_WIDTH) < end_construct.left:

			# further apart left-right than up-down
			if x_diff > y_diff:

				start_idx = "6"
				end_idx = "2"
				cxn_x = start_construct.left + OVAL_WIDTH
				cxn_y = start_construct.top + HALF_OVAL_HEIGHT
				cxn_cx = end_construct.left - start_construct.left - OVAL_WIDTH
				cxn_cy = end_construct.top - start_construct.top

			# further apart up-down than left-right
			else:

				start_idx = "4"
				end_idx = "0"
				cxn_x = start_construct.left + HALF_OVAL_HEIGHT
				cxn_y = start_construct.top + OVAL_HEIGHT
				cxn_cx = end_construct.left - start_construct.left
				cxn_cy = end_construct.top - start_construct.top - OVAL_HEIGHT

		# start node is completely above and completely right of end node
		elif (end_construct.left + OVAL_WIDTH) < start_construct.left:

			flip_h = True
			
			# further apart left-right than up-down
			if x_diff > y_diff:

				start_idx = "2"
				end_idx = "6"
				cxn_x = end_construct.left + OVAL_WIDTH
				cxn_y = start_construct.top + HALF_OVAL_HEIGHT
				cxn_cx = start_construct.left - end_construct.left - OVAL_WIDTH
				cxn_cy = end_construct.top - start_construct.top

			# further apart up-down than left-right
			else:

				start_idx = "4"
				end_idx = "0"
				cxn_x = end_construct.left + HALF_OVAL_HEIGHT
				cxn_y = start_construct.top + OVAL_HEIGHT
				cxn_cx = start_construct.left - end_construct.left
				cxn_cy = end_construct.top - start_construct.top - OVAL_HEIGHT

		# start node is completely above and partially left of end node
		elif start_construct.left < end_construct.left:

			start_idx = "4"
			end_idx = "0"
			cxn_x = start_construct.left + HALF_OVAL_HEIGHT
			cxn_y = start_construct.top + OVAL_HEIGHT
			cxn_cx = end_construct.left - start_construct.left
			cxn_cy = end_construct.top - start_construct.top - OVAL_HEIGHT

		# start node is completely above and partially right of end node
		else:

			flip_h = True
			start_idx = "4"
			end_idx = "0"
			cxn_x = end_construct.left + HALF_OVAL_HEIGHT
			cxn_y = start_construct.top + OVAL_HEIGHT
			cxn_cx = start_construct.left - end_construct.left
			cxn_cy = end_construct.top - start_construct.top - OVAL_HEIGHT

	# start node is completely below end node
	elif (end_construct.top + OVAL_HEIGHT) < start_construct.top:

		flip_v = True

		# start node is completely below and completely left of end node
		if (start_construct.left + OVAL_WIDTH) < end_construct.left:
			
			# further apart left-right than up-down
			if x_diff > y_diff:

				start_idx = "6"
				end_idx = "2"
				cxn_x = start_construct.left + OVAL_WIDTH
				cxn_y = end_construct.top + HALF_OVAL_HEIGHT
				cxn_cx = end_construct.left - start_construct.left - OVAL_WIDTH
				cxn_cy = start_construct.top - end_construct.top

			# further apart up-down than left-right
			else:

				start_idx = "0"
				end_idx = "4"
				cxn_x = start_construct.left + HALF_OVAL_HEIGHT
				cxn_y = end_construct.top + OVAL_HEIGHT
				cxn_cx = end_construct.left - start_construct.left
				cxn_cy = start_construct.top - end_construct.top - OVAL_HEIGHT

		# start node is completely below and completely right of end node
		elif (end_construct.left + OVAL_WIDTH) < start_construct.left:

			flip_h = True
			
			# further apart left-right than up-down
			if x_diff > y_diff:

				start_idx = "2"
				end_idx = "6"
				cxn_x = end_construct.left + OVAL_WIDTH
				cxn_y = end_construct.top + HALF_OVAL_HEIGHT
				cxn_cx = start_construct.left - end_construct.left - OVAL_WIDTH
				cxn_cy = start_construct.top - end_construct.top

			# further apart up-down than left-right
			else:

				start_idx = "0"
				end_idx = "4"
				cxn_x = end_construct.left + HALF_OVAL_HEIGHT
				cxn_y = end_construct.top + OVAL_HEIGHT
				cxn_cx = start_construct.left - end_construct.left
				cxn_cy = start_construct.top - end_construct.top - OVAL_HEIGHT

		# start node is completely below and partially left of end node
		elif start_construct.left < end_construct.left:

			start_idx = "0"
			end_idx = "4"
			cxn_x = start_construct.left + HALF_OVAL_HEIGHT
			cxn_y = end_construct.top + OVAL_HEIGHT
			cxn_cx = end_construct.left - start_construct.left
			cxn_cy = start_construct.top - end_construct.top - OVAL_HEIGHT

		# start node is completely below and partially right of end node
		else:

			flip_h = True
			start_idx = "0"
			end_idx = "4"
			cxn_x = end_construct.left + HALF_OVAL_HEIGHT
			cxn_y = end_construct.top + OVAL_HEIGHT
			cxn_cx = start_construct.left - end_construct.left
			cxn_cy = start_construct.top - end_construct.top - OVAL_HEIGHT

	# start node is partially above end node
	elif start_construct.top < end_construct.top:

		# start node is partially above and completely left of end node
		if (start_construct.left + OVAL_WIDTH) < end_construct.left:

			start_idx = "6"
			end_idx = "2"
			cxn_x = start_construct.left + OVAL_WIDTH
			cxn_y = start_construct.top + HALF_OVAL_HEIGHT
			cxn_cx = end_construct.left - start_construct.left - OVAL_WIDTH
			cxn_cy = end_construct.top - start_construct.top

		# start node is partially above and completely right of end node
		elif (end_construct.left + OVAL_WIDTH) < start_construct.left:

			flip_h = True
			start_idx = "2"
			end_idx = "6"
			cxn_x = end_construct.left + OVAL_WIDTH
			cxn_y = start_construct.top + HALF_OVAL_HEIGHT
			cxn_cx = start_construct.left - end_construct.left - OVAL_WIDTH
			cxn_cy = end_construct.top - start_construct.top

		# start node is partially above and partially left of end node
		elif start_construct.left < end_construct.left:

			# further apart left-right than up-down
			if x_diff > y_diff:
				
				flip_h = True
				start_idx = "6"
				end_idx = "2"
				cxn_x = end_construct.left
				cxn_y = start_construct.top + HALF_OVAL_HEIGHT
				cxn_cx = (start_construct.left + OVAL_WIDTH) - end_construct.left
				cxn_cy = end_construct.top - start_construct.top

			# further apart up-down than left-right
			else:

				flip_v = True
				start_idx = "4"
				end_idx = "0"
				cxn_x = start_construct.left + HALF_OVAL_HEIGHT
				cxn_y = end_construct.top
				cxn_cx = end_construct.left - start_construct.left
				cxn_cy = (start_construct.top + OVAL_HEIGHT) - end_construct.top

		# start node is partially above and partially right of end node
		else:

			# further apart left-right than up-down
			if x_diff > y_diff:
				
				start_idx = "2"
				end_idx = "6"
				cxn_x = start_construct.left
				cxn_y = end_construct.top + HALF_OVAL_HEIGHT
				cxn_cx = (end_construct.left + OVAL_WIDTH) - start_construct.left
				cxn_cy = end_construct.top - start_construct.top

			# further apart up-down than left-right
			else:

				flip_v = True
				flip_h = True
				start_idx = "4"
				end_idx = "0"
				cxn_x = end_construct.left + HALF_OVAL_WIDTH
				cxn_y = end_construct.top
				cxn_cx = start_construct.left - end_construct.left
				cxn_cy = (start_construct.top + OVAL_HEIGHT) - end_construct.top

	# start node is partially below end node
	else:

		# start node is partially below and completely left of end node
		if (start_construct.left + OVAL_WIDTH) < end_construct.left:

			flip_v = True
			start_idx = "6"
			end_idx = "2"
			cxn_x = start_construct.left + OVAL_WIDTH
			cxn_y = end_construct.top + HALF_OVAL_HEIGHT
			cxn_cx = end_construct.left - (start_construct.left + OVAL_WIDTH)
			cxn_cy = start_construct.top - end_construct.top

		# start node is partially below and completely right of end node
		elif (end_construct.left + OVAL_WIDTH) < start_construct.left:

			flip_v = True
			flip_h = True
			start_idx = "2"
			end_idx = "6"
			cxn_x = end_construct.left + OVAL_WIDTH
			cxn_y = end_construct.top + HALF_OVAL_HEIGHT
			cxn_cx = start_construct.left - (end_construct.left + OVAL_WIDTH)
			cxn_cy = start_construct.top - end_construct.top

		# start node is partially below and partially left of end node
		elif start_construct.left < end_construct.left:

			# further apart left-right than up-down
			if x_diff > y_diff:

				flip_v = True
				flip_h = True
				start_idx = "6"
				end_idx = "2"
				cxn_x = end_construct.left
				cxn_y = end_construct.top + HALF_OVAL_HEIGHT
				cxn_cx = (start_construct.left + OVAL_WIDTH) - end_construct.left
				cxn_cy = start_construct.top - end_construct.top

			# further apart up-down than left-right
			else:

				start_idx = "0"
				end_idx = "4"
				cxn_x = start_construct.left + HALF_OVAL_HEIGHT
				cxn_y = start_construct.top
				cxn_cx = end_construct.left - start_construct.left
				cxn_cy = (end_construct.top + OVAL_HEIGHT) - start_construct.top

		# start node is partially below and partially right of end node
		else:

			# further apart left-right than up-down
			if x_diff > y_diff:

				flip_v = True
				start_idx = "2"
				end_idx = "6"
				cxn_x = start_construct.left
				cxn_y = end_construct.top + HALF_OVAL_HEIGHT
				cxn_cx = (end_construct.left + OVAL_WIDTH) - start_construct.left
				cxn_cy = start_construct.top - end_construct.top

			# further apart up-down than left-right
			else:

				flip_h = True
				start_idx = "0"
				end_idx = "4"
				cxn_x = end_construct.left + HALF_OVAL_WIDTH
				cxn_y = start_construct.top
				cxn_cx = start_construct.left - end_construct.left
				cxn_cy = (end_construct.top + OVAL_HEIGHT) - start_construct.top

	connectionShape = ET.SubElement(shapeTree, "p:cxnSp")
	nonVisualConnectorShapeDrawingProperties = ET.SubElement(connectionShape, "p:nvCxnSpPr")
	cNonVisualProperties = ET.SubElement(nonVisualConnectorShapeDrawingProperties, "p:cNvPr")
	cNonVisualProperties.set("id", str(link.id))
	cNonVisualProperties.set("name", "Connector " + str(link.id))
	cNonVisualConnectorShapeDrawingProperties = ET.SubElement(nonVisualConnectorShapeDrawingProperties, "p:cNvCxnSpPr")
	ET.SubElement(nonVisualConnectorShapeDrawingProperties, "p:nvPr")

	connectionStart = ET.SubElement(cNonVisualConnectorShapeDrawingProperties, "a:stCxn")
	connectionStart.set("id", str(start_construct.xml_id)) # shape index from which connector starts (param)
	connectionStart.set("idx", start_idx) # connector spawn point index
	connectionEnd = ET.SubElement(cNonVisualConnectorShapeDrawingProperties, "a:endCxn")
	connectionEnd.set("id", str(end_construct.xml_id)) # shape index at which connector ends (param)
	connectionEnd.set("idx", end_idx) # connector termination point index

	shapeProperties = ET.SubElement(connectionShape, "p:spPr")
	twodTransform = ET.SubElement(shapeProperties, "a:xfrm")
	
	# check flipV and flipH
	if flip_v:
		twodTransform.set("flipV", "1")
	if flip_h:
		twodTransform.set("flipH", "1")

	# location of bounding box (params)
	offset = ET.SubElement(twodTransform, "a:off")
	offset.set("x", str(cxn_x).replace(".0", ""))
	offset.set("y", str(cxn_y).replace(".0", ""))

	# height and width of bounding box (params)
	extents = ET.SubElement(twodTransform, "a:ext")
	extents.set("cx", str(cxn_cx).replace(".0", ""))
	extents.set("cy", str(cxn_cy).replace(".0", ""))

	presetGeometry = ET.SubElement(shapeProperties, "a:prstGeom")
	presetGeometry.set("prst", "line")
	ET.SubElement(presetGeometry, "a:avLst")

	style = ET.SubElement(connectionShape, "p:style")

	lineReference = ET.SubElement(style, "a:lnRef")
	lineReference.set("idx", "1")
	schemeColor = ET.SubElement(lineReference, "a:schemeClr")
	schemeColor.set("val", "dk1")

	fillReference = ET.SubElement(style, "a:fillRef")
	fillReference.set("idx", "0")
	schemeColor = ET.SubElement(fillReference, "a:schemeClr")
	schemeColor.set("val", "dk1")

	effectReference = ET.SubElement(style, "a:effectRef")
	effectReference.set("idx", "0")
	schemeColor = ET.SubElement(effectReference, "a:schemeClr")
	schemeColor.set("val", "dk1")

	fontReference = ET.SubElement(style, "a:fontRef")
	fontReference.set("idx", "minor")
	schemeColor = ET.SubElement(fontReference, "a:schemeClr")
	schemeColor.set("val", "tx1")

# write out the final tree
f = open(os.path.join("/tmp", timestamp, "ppt", "slides", "slide1.xml"), "w")
tree.write(f, encoding="UTF-8", xml_declaration=True)
f.close()

# zip the file
zip_file = zipfile.ZipFile(filename, "w", zipfile.ZIP_DEFLATED)
os.chdir(os.path.join("/tmp", timestamp))
for root, dirs, files in os.walk("."):
	for file in files:
		zip_file.write(os.path.join(root, file))
zip_file.close
