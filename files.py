from os.path import abspath, expanduser

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