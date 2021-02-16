import sys
from os import chdir, environ
from os.path import abspath, dirname, expanduser, join, realpath

if hasattr(sys, '_MEIPASS'):
    # PyInstaller >= 1.6
    chdir(sys._MEIPASS)
    application_path = join(sys._MEIPASS)
elif '_MEIPASS2' in environ:
    # PyInstaller < 1.6 (tested on 1.5 only)
    chdir(environ['_MEIPASS2'])
    application_path = join(environ['_MEIPASS2'])
else:
    chdir(dirname(sys.argv[0]))
    application_path = join(dirname(sys.argv[0]))
	
__location__ = realpath(application_path)

class FilePaths:
	def __init__(self, template=expanduser("~"), csv_file=expanduser("~"), folder=expanduser("~"), filename="Output file"):
		self.template = template
		self.csv_file = csv_file
		self.folder = folder
		self.filename = filename

		@property
		def template(self):
			return self.__template

		@template.setter
		def template(self, file):
			self.__template = abspath(file)

		@property
		def csv_file(self):
			return self.__csv_file

		@csv_file.setter
		def csv_file(self, file):
			self.__csv_file = abspath(file)
		
		@property
		def folder(self):
			return self.__folder

		@folder.setter
		def folder(self, folder):
			self.__folder = abspath(folder)
