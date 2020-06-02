import os
from os.path import basename
import xlsxwriter
from zipfile import ZipFile
import shutil

# Get current directory
dir_path = os.path.dirname(os.path.realpath(__file__))

# Create new directory
report_dir_name = 'reports'
created_dir = os.mkdir(report_dir_name)

# Go to created directory
os.chdir(report_dir_name)

for i in range(0, 3):
  # Create a workbook and add a worksheet.
  workbook = xlsxwriter.Workbook('Expenses%s.xlsx' % i)
  worksheet = workbook.add_worksheet()

  # Some data we want to write to the worksheet.
  expenses = (
      ['Rent', 1000],
      ['Gas',   100],
      ['Food',  300],
      ['Gym',    50],
  )

  # Start from the first cell. Rows and columns are zero indexed.
  row = 0
  col = 0

  # Iterate over the data and write it out row by row.
  for item, cost in (expenses):
      worksheet.write(row, col,     item)
      worksheet.write(row, col + 1, cost)
      row += 1

  # Write a total using a formula.
  worksheet.write(row, 0, 'Total')
  worksheet.write(row, 1, '=SUM(B1:B4)')

  workbook.close()

# Go to previous directory
os.chdir('../')

# create a ZipFile object
with ZipFile('%s.zip' % report_dir_name, 'w') as zipObj:
  # Iterate over all the files in directory
  for folderName, subfolders, filenames in os.walk('%s\\%s' % (dir_path, report_dir_name)):
    for filename in filenames:
        #create complete filepath of file in directory
        filePath = os.path.join(folderName, filename)
        # Add file to zip
        zipObj.write(filePath, basename(filePath))

# Delete created directory
shutil.rmtree('%s\\%s' % (dir_path, report_dir_name))

# # Delete zip file
# os.remove('%s.zip' % (report_dir_name))