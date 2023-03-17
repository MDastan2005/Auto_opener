# from __future__ import annotations
from webbrowser import open as open_web
import os
from shutil import rmtree as delete_directory
import win32com.client as wcl
from pathlib import Path
from tkinter.filedialog import askopenfilename
import customtkinter as ctk
from enum import Enum


class OptionAppType(Enum):
	lnk = 1
	web = 2


class OptionApp:

	def __init__(self, path: Path, app_type: OptionAppType) -> None:
		self.path = path
		self.type = app_type
		self.selected = False

	def __str__(self):
		if self.type == OptionAppType.lnk:
			return self.path.name
		return self.path

	def __repr__(self):
		return str(self)

	def run(self) -> None:
		"""
		Checks the type of app.
		1. If it is a .lnk type then runs the program.
		2. If it is a web file then opens the file and opens all urls.
		TODO: handle possible errors if app cannot be launched
		"""
		if self.type == OptionAppType.lnk:
			os.startfile(self.path)
		else:
			open_web(self.path)

	def get_name(self) -> str:
		if self.type == OptionAppType.lnk:
			return self.path.name
		return self.path

	def create_label(self, master):
		self.label: ctk.CTkLabel = ctk.CTkLabel(
			master, text=self.get_name(), bg_color='grey10', width=200)
		self.label.bind('<Button-1>', lambda e: self.select_app())
		self.label.pack(pady=10)

	def select_app(self):
		"""
		Selects or deselects app according to its selected field.
		"""
		if self.selected:
			self.label.configure(bg_color='grey10')
			self.selected = False
		else:
			self.label.configure(bg_color='blue')
			self.selected = True

	@staticmethod
	def create_lnk(exe_path: Path, save_path: Path) -> Path:
		"""
		Creates and saves link to the given program into the given folder.
		Returns the path to the created link
		"""
		shortcut_path = Path.joinpath(save_path, f'{exe_path.stem}.lnk')
		shell = wcl.Dispatch('WScript.Shell')
		shortcut = shell.CreateShortCut(str(shortcut_path))
		shortcut.Targetpath = str(exe_path)
		shortcut.IconLocation = str(exe_path)
		shortcut.save()
		return shortcut_path


class Option:
	__all__: list['Option'] = []

	def __init__(self,
				 window: ctk.CTk,
				 name: str,
				 apps: list[OptionApp] = []) -> None:
		self.__all__.append(self)
		self.name = name
		self.apps: list[OptionApp] = apps
		self.window = window
		self.selected = False
		self.label: ctk.CTkLabel = ctk.CTkLabel(window,
												text=self.name,
												font=('Roboto', 15),
												bg_color='gray19',
												height=30,
												width=300)
		self.label.bind('<Button-1>', lambda e: self.on_click())
		self.label.pack(pady=5, padx=10)

	def delete(self):
		self.deselect()
		self.label.pack_forget()
		self.__all__.remove(self)
		del self

	def add_app(self, app: OptionApp) -> None:
		self.apps.append(app)

	def run_all(self) -> None:
		for app in self.apps:
			app.run()
	
	def run_selected(self) -> None:
		for app in self.apps:
			if app.selected:
				app.run()

	def get_apps(self) -> list[OptionApp]:
		return self.apps

	def select(self) -> None:
		self.selected = True
		self.label.configure(bg_color='blue')

	def deselect(self) -> None:
		self.selected = False
		self.label.configure(bg_color='gray19')

	def on_click(self) -> None:
		try:
			self.get_selected().deselect()
		except AttributeError:
			pass
		self.select()

	@classmethod
	def get_all(cls):
		return cls.__all__

	@classmethod
	def get_selected(cls) -> 'Option':
		"""
		Returns option with field selected equal to True
		or None if such object not found
		"""
		for option in cls.__all__:
			if option.selected:
				return option
		return None

	def __str__(self):
		return str(self.apps)

	def __repr__(self):
		return str(self)

	def __iter__(self):
		return OptionIter(self)


class OptionIter:

	def __init__(self, opt: Option) -> None:
		self.apps = opt.apps
		self._cur_index = 0

	def __iter__(self):
		return self

	def __next__(self):
		if self._cur_index >= len(self.apps):
			raise StopIteration
		app = self.apps[self._cur_index]
		self._cur_index += 1
		return app


class MainPage(ctk.CTkFrame):

	def __init__(self, master: ctk.CTkFrame, **kwargs) -> None:
		super().__init__(master, **kwargs)
		self.master = master
		# Main Frame
		self.rowconfigure((0, 1, 2), weight=0)
		self.columnconfigure((0, 1), weight=1)

		self.option_frame: ctk.CTkFrame = ctk.CTkFrame(self, width=0, height=0)
		self.option_frame.grid(column=0, row=0, pady=10, columnspan=2)

		self.get_options()

		# Option control
		self.entry_option_name: ctk.CTkEntry = ctk.CTkEntry(
			self, placeholder_text="Folder Name", font=("Roboto", 16))
		self.entry_option_name.grid(column=0, row=1)
		# Add
		self.button_add_option: ctk.CTkButton = ctk.CTkButton(
			self,
			text="Create Folder",
			fg_color='green',
			command=lambda: self.add_new_option())
		self.button_add_option.grid(column=1,
									row=1,
									pady=10,
									padx=10,
									sticky='ns')
		# Delete
		self.button_del_option: ctk.CTkButton = ctk.CTkButton(
			self,
			text="Delete Folder",
			fg_color='red',
			command=lambda: self.del_selected_option())
		self.button_del_option.grid(column=0, row=2, pady=10, padx=10)
		# Run
		self.button_run_option: ctk.CTkButton = ctk.CTkButton(
			self,
			text="Run all from folder",
			command=lambda: self.run_selected_option())
		self.button_run_option.grid(column=1, row=2, pady=10, padx=10)

	def __repr__(self):
		return '\n'.join([str(option) for option in self.options])

	def __str__(self):
		return 'MainPage'

	def add_option(self,
				   option_name: str,
				   option_apps: list[OptionApp] | None = None) -> None:
		"""
		Creates appropriate option folder in data folder.
		If option_name is empty then ignores all
		"""
		option_name = option_name.strip()
		if not option_name:
			return
		if any(option_name == option.name for option in self.options):
			return
		if option_apps is None:
			option_apps = []
		new_option = Option(self.option_frame, option_name, option_apps)

		self.options.append(new_option)
		self.master.add_option(new_option)

		option_path = Path.joinpath(self.master.data_folder_path, option_name)
		if not os.path.isdir(option_path):
			os.mkdir(option_path)

			f = open(Path.joinpath(option_path, 'weburls.txt'), 'x')
			f.close()

	def add_new_option(self):
		self.add_option(self.entry_option_name.get())
		self.entry_option_name.delete(0, 'end')

	def get_selected_option(self) -> Option | None:
		"""
		Returns selected option.
		If no option is selected then returns None
		"""
		selected_option = Option.get_selected()
		return selected_option

	def del_option(self, option: Option) -> None:
		"""
		Deletes option and removes from all frames.
		"""
		self.options.remove(option)
		option.delete()

	def del_selected_option(self) -> None:
		"""
		Gets selected option. 
		If no option is selected then does nothing.
		If seleced found then deletes option
		"""
		selected_option = self.get_selected_option()
		if selected_option is not None:
			delete_directory(
				Path.joinpath(self.master.data_folder_path,
							  selected_option.name))
			self.del_option(selected_option)

	def add_site(self, option: Option, site_url: str) -> None:
		"""
		Adds url into weburls.txt file in the corresponding option.
		"""
		self.get_option(option).add_site(site_url)

	def get_options(self) -> list[Option]:
		"""
		Updates options list, and returns list of all options from data folder.
		"""
		self.options: list[Option] = []

		for option_name in os.listdir(self.master.data_folder_path):
			cur_apps: list[OptionApp] = []
			option_path = Path.joinpath(self.master.data_folder_path,
										option_name)

			for app_name in os.listdir(option_path):
				app_path = Path.joinpath(option_path, app_name)
				if app_name == 'weburls.txt':
					with open(app_path, 'r') as file:
						urls = file.readlines()
						for url in urls:
							url = url.strip()
							if not url:
								continue
							cur_apps.append(OptionApp(url, OptionAppType.web))
				else:
					cur_apps.append(OptionApp(app_path, OptionAppType.lnk))

			self.add_option(option_name, cur_apps)

		return self.options

	def get_option(self, option: Option) -> Option | None:
		"""
		Returns option with the give name, or None if not found.
		"""
		for option in self.get_options():
			if option.name == option:
				return option
		return None

	def get_apps(self, option: Option) -> list[OptionApp]:
		"""
		Returns list of all apps, and website urls from given option.
		"""
		return self.get_option(option).get_apps()

	def run_option(self, option: Option) -> None:
		"""
		Runs all apps and opens all websites from the option.
		"""
		option.run_all()

	def run_selected_option(self):
		selected_option = self.get_selected_option()
		if selected_option is not None:
			self.run_option(selected_option)


class OptionPage(ctk.CTkFrame):

	def __init__(self, master: any, option: Option, **kwargs) -> None:
		super().__init__(master, **kwargs)
		self.master = master
		self.option: Option = option
		self.option.label.bind('<Double-Button-1>',
							   lambda e: self.master.show_page(self))

		self.button_return: ctk.CTkButton = ctk.CTkButton(
			self, text="Return back", command=lambda: self.master.show_page(self.master.main_page))
		self.button_return.pack(pady=10)

		self.frame_labels: ctk.CTkFrame = ctk.CTkFrame(self, width=0, height=0)
		self.frame_labels.pack(pady=20, ipadx=20)
		for app in self.option.apps:
			app.create_label(self.frame_labels)

		self.button_add_exe: ctk.CTkButton = ctk.CTkButton(
			self, text="Add new program", command=lambda: self.add_exe())
		self.button_add_exe.pack(pady=10)

		self.frame_add_web: ctk.CTkFrame = ctk.CTkFrame(self, fg_color='transparent')
		self.frame_add_web.pack(pady=10)
		self.entry_web_url: ctk.CTkEntry = ctk.CTkEntry(self.frame_add_web, placeholder_text='Enter website url')
		self.entry_web_url.grid(row=0, column=0, padx=15)
		self.button_add_web: ctk.CTkButton = ctk.CTkButton(self.frame_add_web, text="Add website", command=lambda: self.add_web())
		self.button_add_web.grid(row=0, column=1, padx=15)

		self.frame_run_remove: ctk.CTkFrame = ctk.CTkFrame(self, fg_color='transparent')
		self.frame_run_remove.pack(pady=10)
		self.button_remove_selected: ctk.CTkButton = ctk.CTkButton(self.frame_run_remove, text="Remove all selected", command=lambda: self.remove_selected())
		self.button_remove_selected.grid(row=0, column=0, pady=10, padx=15)
		self.button_run_selected: ctk.CTkButton = ctk.CTkButton(self.frame_run_remove, text="Run all selected", command=lambda: self.run_selected())
		self.button_run_selected.grid(row=0, column=1, pady=10, padx=15)

	def add_exe(self) -> None:
		file = ctk.filedialog.askopenfile()
		if file is not None:
			file_path = Path(file.name)
			save_path = Path.joinpath(self.master.data_folder_path, 
			     					  self.option.name)
			new_app_path = OptionApp.create_lnk(file_path, save_path)
			new_app = OptionApp(new_app_path, OptionAppType.lnk)
			new_app.create_label(self.frame_labels)
			self.option.add_app(new_app)
	
	def add_web(self) -> None:
		"""
		Adds url to weburls.txt file in data folder, and creates new OptionApp object
		"""
		url = self.entry_web_url.get()
		url = url.strip()
		if not url:
			return
		if not url.startswith(('http://', 'https://')):
			url = 'https://' + url
		
		weburls_path = Path.joinpath(self.master.data_folder_path, self.option.name, 'weburls.txt')
		with open(weburls_path, 'a') as file:
			file.write(url + os.linesep)
		
		new_app = OptionApp(url, OptionAppType.web)
		new_app.create_label(self.frame_labels)
		self.option.add_app(new_app)
	
	def remove_selected(self) -> None:
		urls_to_remove: list[str] = []
		for app in self.option.apps:
			if not app.selected:
				continue
			if app.type == OptionAppType.lnk:
				os.remove(app.path)
				app.label.pack_forget()
				self.option.apps.remove(app)
				del app
			else:
				urls_to_remove.append(app.path)
				app.label.pack_forget()
				self.option.apps.remove(app)
				del app
		weburls_path = Path.joinpath(self.master.data_folder_path, self.option.name, 'weburls.txt')
		with open(weburls_path, 'r') as file:
			urls = file.readlines()
		with open(weburls_path, 'w') as file:
			for url in urls:
				url = url.strip()
				if not url:
					continue
				if url not in urls_to_remove:
					file.write(url + os.linesep)
	
	def run_selected(self) -> None:
		self.option.run_selected()


class AutoOpenerApp(ctk.CTk):

	def __init__(self,
				 fg_color: str | tuple[str, str] | None = None,
				 **kwargs) -> None:
		super().__init__(fg_color, **kwargs)

		self.title('Auto opener')
		self.minsize(400, 500)
		self.resizable(True, True)

		self._create_data_folder()

		self.option_pages: list[OptionPage] = []
		for option in Option.__all__:
			self.add_option(option)
		self.main_page = MainPage(self, fg_color='transparent')

		self.cur_page = self.main_page

		self.show_page(self.main_page)

	def _create_data_folder(self) -> None:
		"""
		Assigns path to data folder, and creates data folder if not existed yet.
		The data folder is in the same folder as executable.
		"""
		self.data_folder_path = Path.joinpath(Path(__file__).parent, 'data')
		if not os.path.isdir(self.data_folder_path):
			os.mkdir(self.data_folder_path)

	def show_page(self, new_page: ctk.CTkFrame) -> None:
		self.cur_page.pack_forget()
		self.cur_page = new_page
		self.cur_page.pack()
	
	def add_option(self, option: Option):
		self.option_pages.append(OptionPage(self, option, fg_color='transparent'))


if __name__ == '__main__':
	ctk.set_appearance_mode('dark')
	win = AutoOpenerApp()
	win.mainloop()

# pyinstaller --noconfirm --onedir --windowed --add-data "d:\pyprojects\autoscripts\auto_opener\.venv\lib\site-packages/customtkinter;customtkinter/"  "d:\pyprojects\autoscripts\auto_opener\main.py"