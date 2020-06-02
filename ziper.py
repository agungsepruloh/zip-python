from zipfile import ZipFile
import os
from os.path import basename

# Create folder
new_dir = 'reports'
created_dir = os.mkdir(new_dir)

# Show current directory
dir_path = os.path.dirname(os.path.realpath(__file__))

# # create a ZipFile object
# zipObj = ZipFile('sample.zip', 'w')

# # Add multiple files to the zip
# for i in range(0, 3):
#   zipObj.write('Expenses%s.xlsx' % i)

# # close the Zip File
# zipObj.close()

# print(dir_path)

# for root, dirs, files in os.walk('D:/Dev'):
#   print('root:', root)
#   print('dirs:', dirs)
#   print('files:', files)

# create a ZipFile object
with ZipFile(new_dir + '.zip', 'w') as zipObj:
  # Iterate over all the files in directory
  for folderName, subfolders, filenames in os.walk(dir_path + '\\' + new_dir):
    for filename in filenames:
        #create complete filepath of file in directory
        filePath = os.path.join(folderName, filename)
        # Add file to zip
        zipObj.write(filePath, basename(filePath))