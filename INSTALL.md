# Installation Instructions

## Requirements

Prerequisites:
- Python 3.10+ (tested with Python 3.13.9)
- Git (for cloning)
- Anaconda (recommended) or standard Python

Required packages:
- pandas
- pillow
- python-docx
- numpy

## Recommended Installation

### 1. Download code and data

a. Download flashcard_builder.py  
b. (Optional) ownload sample CSV file and headshot images:  
- sample_headshots
- sample_student_info.csv

### 1. Create and activate an environment (Conda - preferred)

Example:  
conda create -n flashcards_py python=3.11  
conda activate flashcards_py

### 2. Install required packages

Examples:  
pip install -r requirements.txt  
OR  
conda install -c conda-forge pandas pillow python-docx pytest  

### 3. Run the script

Generate Word file - example:  
python flashcard_builder.py('.\sample_student_info.csv', '.\sample_headshots', '.\sample_doc.docx')
