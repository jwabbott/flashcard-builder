# Flashcard Builder

A Python tool that generates printable double-sided flashcards from a student information CSV file and a folder of headshot images.

**By Joshua W. Abbott**

# Introduction

Hi, I wrote this script to create flashcards for learning the names and faces of new students. Working in higher education, I meet hundreds of new students every year and struggle with remembering their names. The best method I have found for memorizing any kind of information is to use flashcards. However, making flashcards with names, photos, and other information can be very time-intensive and tedious. This script allows you to upload a spreadsheet with student names and other basic info, along with a folder of the students' headshot images, and it builds the flashcards for you. All you have to do then is print out the sheets and cut out the cards.

NOTE: The sample student information contained here is fictional and does not disclose any actual student information, which is protected under FERPA. The headshot images were ai-generated from the website, [thispersondoesnotexist.org](https://thispersonnotexist.org/).

# Overview

## Quick Start

Once you have downloaded flashcard_builder.py and imported the required packages (see INSTALL file for details), you can run the script from your shell prompt.

For example:

**python flashcard_builder.py --csv students.csv --images headshots --out flashcards.docx**

where:
- "students.csv" is the path that specifies the .csv file that contains your student information,
- "headshots" is the path that specifies the folder containing your headshot images, and
- "flashcards.docx" is the path that specifies the Word doc you want to generate with your flashcards

Please feel free to use the sample_headshots and the sample_student_info.csv in this repository. You can download and save them in a file structure such as this:

  flashcard-builder/
  │
  ├── flashcard_builder.py
  ├── students.csv
  └── headshots/
    ├── image001.jpeg
    ├── image002.jpeg
    └── ...


## Inputs

This script requires two inputs: (1) a CSV file with student information, and (2) a folder containing the students' headshot images. 

(1) The CSV file should contain the following columns, with one student for each row:
- name
- hometown
- undergrad_school
- major
- image_filename

(2) The image files in the headshots folder should follow this naming convention: "image001.jpeg", "image002.jpeg", etc., where image001.jpeg contains the headshot image of the student whose name and information are listed on the first row (after the column names) in the CSV file, image002.jpeg contains the headshot of the student on the next row, and so on.  

## Output

With the inputs above, the script will create a two-sided MS Word document containing flashcards for 10 students per page. The front contains the student names and information, and the back holds the students' photos. You can print the document, being sure to set the printer settings to double-sided printing (flip on long edge), and then cut out the cards. 

# Contributing

If you would like to help with this project by reporting bugs or suggesting new features, please feel free to open a GitHub issue. Pull requests are welcome. You can contact me via email at josh dot abbott at asu dot edu.


