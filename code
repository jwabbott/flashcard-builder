# # Flashcard Builder
# ## Joshua W. Abbott
# ### Description:
# This code creates printable flashcard sheets for learning student names and faces.
# It accepts as input a CSV file containing student information, including thair name and some basic biographical info, along with image files containing the students' headshots.
# It creates a Word file containing cards for 10 students on each sheet, which can be printed, cut out, and used as flashcards.

# ## Install tools

import pandas as pd
import numpy as np
import math
from pathlib import Path
from PIL import Image
import os
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.shared import Inches, Cm, Pt
from docx.enum.table import WD_ROW_HEIGHT_RULE

# Set values for page and card size, orientation, and margins (all measurements in inches unless otherwise indicated)
PAGE_HEIGHT = 8.5
PAGE_WIDTH = 11
PAGE_MARGINS = 0.25
ROWS = 2
COLUMNS = 5
CARDS_PER_PAGE = ROWS * COLUMNS
TABLE_HEIGHT = 0.9 * PAGE_HEIGHT
TABLE_WIDTH = PAGE_WIDTH
CARD_HEIGHT = TABLE_HEIGHT / ROWS
CARD_WIDTH = TABLE_WIDTH / COLUMNS
IMG_MAX_HEIGHT = CARD_HEIGHT * 0.9
IMG_MAX_WIDTH = CARD_WIDTH * 0.9
NAME_FONT_PT = 14
INFO_FONT_PT = 9


# Read student information into a dataframe and verify size and columns
df = pd.read_csv('sample_student_info.csv')
print (df.head())
print(df.columns)
print("Number of students: ", len(df))


# Replace any missing values with an empty string to avoid "NaN" errors
df.fillna('', inplace=True)

# Specify the path to the folder containing the headshot images
image_folder_location = Path("C:/Users/jwabbott/Desktop/CAS-502/flashcards/test_headshots")


def build_image_index(image_folder):

	"""
	Goes through a specified folder and creates an index dictionary of the filenames (key: filename, value: full path).
	"""

	image_dict = {}

	# Iterate through all items in the image folder directory, and create the full path for each file
	for file_name in os.listdir(image_folder):
		full_path = os.path.join(image_folder, file_name)

		# if the item is a file (i.e. not a folder), add its filename to the dictionary
		if os.path.isfile(full_path):
			image_dict[file_name] = full_path

	# Count and display how many headshot images were found and indexed; display first few image filenames
	num_headshots = len(image_dict)
	print("Number of headshot images found: ", num_headshots)
	first_five_keys = list(image_dict.keys())[:5]
	print("Sampling of image filenames: ", first_five_keys)

	# Verify that the number of student entries and the number of headshot images match and print a warning if they don't
	if not (len(df) == num_headshots):
		print("Warning: number of students and number of headshot images do not match.")
	            
	return image_dict

image_dict = build_image_index(image_folder_location)

# Add a column to the dataframe containing the filepaths of the image files for each student
df["image_path"] = df["image_filename"].map(image_dict)

# Check whether the correct path has been added to the dataframe for at least the first row
cell_value = df.at[0, 'image_path']
print(cell_value)


def format_image (image):

	"""
	opens and formats the given image to prepare it for insertion into Word document
	then stores formatting details in a dictionary that can be added to corresponding row in the main dataframe
	"""

	image_formatting_info = {}
	image_formatting_info['image_path'] = image

	with Image.open(image) as img:
		img_size_tuple = img.size
		img_width_px, img_height_px = img_size_tuple
		image_formatting_info['width_px'] = img_width_px			# add image dimensions in pixels to the dictionary
		image_formatting_info['height_px'] = img_height_px

		dpi_val = img.info.get('dpi')								# get the "dots-per-inch" from the image info and add to dictionary
		if dpi_val:													# use 72 as the default dpi if not contained in image info
			dpi = dpi_val[0]
		else:
			dpi = 72
		image_formatting_info['dpi'] = dpi

		img_height = img_height_px / dpi							# calculate the image dimensions in inches
		img_width = img_width_px / dpi
		image_formatting_info['height_in'] = img_height
		image_formatting_info['width_in'] = img_width

		scale_h = IMG_MAX_HEIGHT / img_height						# calculate the scaling factor to fit the image in the table cell
		scale_w = IMG_MAX_WIDTH / img_width
		scale = min(scale_w, scale_h, 1.0)
		image_formatting_info['scale'] = scale
	
		scaled_height = img_height * scale							# calculate scaled image dimensions in inches
		scaled_width = img_width * scale
		image_formatting_info['scaled_height_in'] = scaled_height
		image_formatting_info['scaled_width_in'] = scaled_width

		scaled_height_px = int(img_height_px * scale)				# calculate scaled image dimensions in pixels
		scaled_width_px = int(img_width_px * scale)
		image_formatting_info['scaled_height_px'] = scaled_height_px
		image_formatting_info['scaled_width_px'] = scaled_width_px

	return image_formatting_info									# return dictionary with all image info


# Define column names for and create an empty new dataframe that will contain image formatting data

image_formatting_columns = ['image_path', 'height_px', 'width_px', 'height_in', 'width_in', 'scale', 'scaled_height_in', 'scaled_width_in', 'scaled_height_px', 'scaled_width_px']
image_info_df = pd.DataFrame(columns=image_formatting_columns)


# Populate the new dataframe with image formatting data for each image

for row in range(len(df)):
	image_to_format = df.at[row, 'image_path']
	image_info_df.loc[row] = format_image(image_to_format)


# Merge the new image formatting dataframe with the student information dataframe so that all information needed to inserting into Word is contained in a single dataframe
# Verify size and columns

df = pd.merge(df, image_info_df, on='image_path')

print (df.head())
print(df.columns)


# Create a new blank Word document
document = Document()


# set document's orientation to landscape
section = document.sections[-1] 
new_width, new_height = section.page_height, section.page_width
section.page_width = new_width
section.page_height = new_height

# set document's margins to a quarter inch
section.top_margin = Inches(0.25)
section.bottom_margin = Inches(0.25)
section.left_margin = Inches(0.25)
section.right_margin = Inches(0.25)



def add_front(document, batch_df):

	"""
	Adds front table with student info for current batch to Word document.
	"""

	table = document.add_table(rows=2, cols=5)					# insert table with fixed row and column sizes
	table.style = 'Table Grid' 
	table.allow_autofit = False

	for column in table.columns:
		column.width = Inches(CARD_WIDTH)
		for cell in column.cells:
			cell.width = Inches(CARD_WIDTH)

	for row in table.rows:
		row.height = Inches(CARD_HEIGHT)
		row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
											

	student_index = 0

	for row in table.rows:										# iterate through each cell in the table
		for cell in row.cells:

			if student_index < len(batch_df):					# only add info to cell is there are more students in the batch

				# Center text, set to 14-point font size, enter student name

				para = cell.add_paragraph()
				para.alignment = WD_ALIGN_PARAGRAPH.CENTER
	
				name_text = batch_df.iat[student_index, 0]
				run = para.add_run(name_text + "\n")
				run.bold = True	
				run.font.size = Pt(14)
		

				# Left justify text, reset to 11-pt font size, enter other student info
		
				new_para = cell.add_paragraph()
				new_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
				new_para_format = new_para.paragraph_format
				new_para_format.left_indent = Inches(0.25)


				# iterate through student info and add to card

				for i in range(3):
					line = i + 1
					line_text = batch_df.iat[student_index, line]
					run = new_para.add_run(line_text + "\n")
					run.bold = False	
					run.font.size = Pt(11)


			student_index = student_index + 1						# iterate to next student in the batch


def add_back(document, batch_df):

	"""
	Adds back table with student photos to Word document.
	"""

	table = document.add_table(rows=2, cols=5)						# insert table with fixed row and column sizes
	table.style = 'Table Grid' 
	table.allow_autofit = False

	for column in table.columns:
		column.width = Inches(CARD_WIDTH)
		for cell in column.cells:
			cell.width = Inches(CARD_WIDTH)

	for row in table.rows:
		row.height = Inches(CARD_HEIGHT)
		row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY


	student_index = 0


	for row in table.rows:											# iterate through each cell in the table, 
		for cell in reversed(row.cells):							# but in such a way as to pair each photo on the back to the matching student info on the front

			if student_index < len(batch_df):						# only add photos if there are more in the batch, up to 10

				para = cell.add_paragraph()
				para.alignment = WD_ALIGN_PARAGRAPH.CENTER
				run = para.add_run()

				photo_path = batch_df.iat[student_index, 5]
				photo_width = batch_df.iat[student_index, 12]

				run.add_picture(photo_path, width=Inches(photo_width))
				cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

			student_index = student_index + 1						# iterate to the next studdent in the batch


num_of_batches = math.ceil(len(df) / 10)							# calculate the number of batches of 10 students each, with last batch containing any remaining

for batch_index in range(num_of_batches):							# iterate through each batch of students

	start = batch_index * 10
	end = start + 10
	end = min(end, len(df))

	batch_df = df.iloc[start:end].reset_index(drop=True)			# create a new dataframe containing only the info for the students in the current batch

	add_front(document, batch_df)									# insert student info and photos into Word document by calling the functions to add front and back

	add_back(document, batch_df)

	
	
document.save('flashcards.docx')									# save the Word document
