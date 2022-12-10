import tkinter
from tkinter import filedialog
import customtkinter
import webbrowser
from data_etl import *


customtkinter.set_appearance_mode("Dark")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("dark-blue")  # Themes: "blue" (standard), "green", "dark-blue"

class App(customtkinter.CTk):

    WIDTH = 1200
    HEIGHT = 700
    CORNER_RADIUS = 10

    def __init__(self):
        super().__init__()
        # self.iconbitmap(r"assets/icon.ico")
        self.homerooms_path = ""
        self.grades_path = ""
        output_dir = ""

        # Initialize Application
        self._init_root() # highest level settings of appliction

        self._init_homepage()

        self._init_body_pages(row_pos=0, col_pos=1, padx=20, pady=10)        

        self._init_settings_page()

        self._init_navbar() # navbar with buttons to navigate app

        self.mgr_start_page.tkraise() # make homepage the first frame to show

    def _init_root(self):
        self.title("MGR Creator")
        self.geometry(f"{App.WIDTH}x{App.HEIGHT}")
        self.protocol("WM_DELETE_WINDOW", self._e_on_closing)  # call ._e_on_closing() when app gets closed

        # configure grid layout (2x1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

    def _init_navbar(self):
        self.navbar = customtkinter.CTkFrame(master=self,
                                                width=200,
                                                corner_radius=0)
        self.navbar.grid(row=0, column=0, sticky="nswe",)
        
        self.navbar_header = customtkinter.CTkLabel(master=self.navbar,
                                              text="MGR Creator",
                                              text_font=("Roboto Medium", 16))  # font name and size in px
        self.navbar_header.grid(row=1, column=0, pady=10, padx=10)

        self.home_button = customtkinter.CTkButton(master=self.navbar,
                                                text="Home",
                                                command=lambda: self._e_switch_frame(self.mgr_start_page))
        self.home_button.grid(row=2, column=0, pady=10, padx=20)

        # self.mgr_button = customtkinter.CTkButton(master=self.navbar,
        #                                         text="MGR Creator",
        #                                         command=lambda: self._e_switch_frame(self.mgr_start_page))
        # self.mgr_button.grid(row=3, column=0, pady=10, padx=20)

        self.settings_button = customtkinter.CTkButton(master=self.navbar,
                                                text="Settings",
                                                command=lambda: self._e_switch_frame(self.settings_page))
        self.settings_button.grid(row=4, column=0, pady=10, padx=20)

    def _init_body_pages(self, row_pos: int, col_pos: int, padx: int, pady:int):

        self._init_mgr_start_page(row_pos=row_pos, col_pos=col_pos, padx=padx, pady=pady)
        self._init_choose_mgr_files_page(row_pos=row_pos, col_pos=col_pos, padx=padx, pady=pady)
        self._init_choose_output_dir_page(row_pos=row_pos, col_pos=col_pos, padx=padx, pady=pady)

    def _init_mgr_start_page(self, row_pos: int, col_pos: int, padx: int, pady:int):

        # start page for MGR, master is the root
        self.mgr_start_page = customtkinter.CTkFrame(
            master=self
        )

        self.mgr_start_header = customtkinter.CTkLabel(
            master = self.mgr_start_page, text="Generate Your MGR",
            text_font=("Roboto Medium", 32),
            
        )
        self.mgr_start_instructions = customtkinter.CTkLabel(
            master=self.mgr_start_page,
            text="Before we get started, go ahead and log into Powerschool ",
            text_font=("Roboto Medium", 16, "italic")

        )
        
        self.b_next_mgr_start = customtkinter.CTkButton(master=self.mgr_start_page,
            text="Next",
            text_font=("Roboto Medium", 16),
            width=200,
            height=40,
            command=lambda: self._e_switch_frame(self.choose_files_page))


        # place start page onto root grid
        self.mgr_start_page.grid(row=row_pos, column=col_pos, sticky="nswe", padx=20, pady=10)
        
        # place text onto start page
        self.mgr_start_header.pack(pady=self.mgr_start_page._current_height/2)
        self.mgr_start_instructions.pack()
        self.b_next_mgr_start.pack(padx=60, pady=40, anchor="e", side="bottom")

    def _init_choose_mgr_files_page(self, row_pos: int, col_pos: int, padx: int, pady:int):

        self.choose_files_page = customtkinter.CTkFrame(
            master=self,
        )
        self.choose_files_page.grid_rowconfigure(1, weight=1)
        self.choose_files_page.grid_columnconfigure((0,1), weight=1)

        # drag and drop files frames
        self.f_choose_gradesfile = customtkinter.CTkFrame(
            master=self.choose_files_page,
            # fg_color=customtkinter.ThemeManager.theme["color"]["frame_low"],
        )
        self.f_choose_homeroomfile = customtkinter.CTkFrame(
            master=self.choose_files_page
        )


        self.choose_files_header = customtkinter.CTkLabel(
            master = self.choose_files_page, 
            text="Upload Files",
            text_font=("Roboto Medium", 32),
            padx=20,
            pady=10
        )
 
        self.h_choose_gradesfile = customtkinter.CTkLabel(
            master = self.f_choose_gradesfile, 
            text="Choose Grades File",
            text_font=("Roboto Medium", 16),
            padx=20,
            pady=10
        )

        self.h_choose_homeroomfile = customtkinter.CTkLabel(
            master = self.f_choose_homeroomfile, 
            text="Choose Homeroom File",
            text_font=("Roboto Medium", 16),
            padx=20,
            pady=10
        )
 
        self.t_choose_grades_file_inst = customtkinter.CTkLabel(
            master = self.f_choose_gradesfile, 
            text="Download the student grades file from Powerschool as a .csv file",
            text_font=("Roboto Medium", 16),
            padx=20,
            pady=10,
            wraplength=self.f_choose_gradesfile._current_width*3.5,
        )
        self.t_choose_grades_file_inst.bind('<Configure>', lambda e: self.t_choose_grades_file_inst.configure(wraplength=self.t_choose_grades_file_inst.winfo_width()))

        self.t_choose_homerooms_file_inst = customtkinter.CTkLabel(
            master = self.f_choose_homeroomfile, 
            text="Download the student homerooms file from Powerschool as a .csv file",
            text_font=("Roboto Medium", 16),
            padx=20,
            pady=10,
            wraplength=self.f_choose_homeroomfile._current_width*3.5,
        )
        self.t_choose_homerooms_file_inst.bind('<Configure>', lambda e: self.t_choose_homerooms_file_inst.configure(wraplength=self.t_choose_homerooms_file_inst.winfo_width()))

        self.t_choose_homerooms_file_inst_2 = customtkinter.CTkLabel(
            master = self.f_choose_homeroomfile, 
            text="Note: Choose \"Home_Room\" from the \"Select a Field\" drop down",
            text_font=("Roboto Medium", 14, "italic"),
            padx=20,
            pady=10,
            wraplength=self.f_choose_homeroomfile._current_width*2,
        )

        
        self.t_choose_homerooms_file_link = tkinter.Label(
            master = self.f_choose_homeroomfile, 
            text="Link to Download Homerooms",
            bg=customtkinter.ThemeManager.theme["color"]["frame_high"][1],
            font=("Roboto Medium", 16, "underline"),
            fg="#5C85D6"
            )
        self.t_choose_homerooms_file_link.bind("<Button-1>",lambda e: self._callback("https://greatoakscharter.powerschool.com/admin/reports_pscb/studentdataverify/FieldDataVerificationCount.html#"))

        self.t_choose_grades_file_link = tkinter.Label(
            master = self.f_choose_gradesfile, 
            text="Link to Download Grades",
            bg=customtkinter.ThemeManager.theme["color"]["frame_high"][1],
            font=("Roboto Medium", 16, "underline"),
            fg="#5C85D6"
            )
        self.t_choose_grades_file_link.bind("<Button-1>",lambda e: self._callback("https://greatoakscharter.powerschool.com/admin/reports_pscb/grading/ClassGradesComments.html"))
            

            # text_font=("Roboto Medium", 16),
        


        self.b_choose_gradesfile = customtkinter.CTkButton(
            master=self.f_choose_gradesfile,
            text="Upload",
            text_font=("Roboto Medium", 12),
            width=150,
            height=35,
            command=self._e_choose_grades_file)


        self.b_choose_homeroomfile = customtkinter.CTkButton(
            master=self.f_choose_homeroomfile,
            text="Upload",
            text_font=("Roboto Medium", 12),
            width=150,
            height=35,
            command=self._e_choose_homerooms_file)

        self.b_choose_homeroomfile_next = customtkinter.CTkButton(
            master=self.choose_files_page,
            text="Next",
            text_font=("Roboto Medium", 16),
            width=200,
            height=40,
            command=lambda: self._e_switch_frame(self.f_choose_out_dir))


        # grid choose files page frame onto root        
        self.choose_files_page.grid(row=row_pos, column=col_pos, sticky='nswe', padx=padx, pady=pady)

        # drag and drop files frame
        self.f_choose_gradesfile.grid(row=1, column=0, sticky="nswe", padx=20, pady=20)
        self.f_choose_homeroomfile.grid(row=1, column=1, sticky="nswe", padx=10, pady=20)

        # place widgets
        self.choose_files_header.grid(row=0, column=0, pady=25, padx=20, columnspan=4)

        self.h_choose_gradesfile.pack()
        self.h_choose_homeroomfile.pack()
        self.t_choose_homerooms_file_inst.pack(anchor="n")       
        self.t_choose_grades_file_inst.pack(anchor="n")
        self.t_choose_homerooms_file_inst_2.pack(anchor="n", pady=20)


        
        self.b_choose_gradesfile.pack(anchor="s", side="bottom", pady=20)
        self.b_choose_homeroomfile.pack(anchor="s", side="bottom", pady=20)

        self.t_choose_homerooms_file_link.pack( side="bottom", pady=20)
        self.t_choose_grades_file_link.pack( side="bottom", pady=20)
        

        self.b_choose_homeroomfile_next.grid(row=2, column=1, sticky="ne", padx=60, pady=40)

    def _init_choose_output_dir_page(self, row_pos: int, col_pos: int, padx: int, pady:int):
        self.f_choose_out_dir = customtkinter.CTkFrame(
            master=self,
        )

        self.h_choose_out_dir = customtkinter.CTkLabel(
            master = self.f_choose_out_dir, 
            text="Choose Folder for MGR",
            text_font=("Roboto Medium", 32),
            padx=20,
            pady=10
        )

        self.b_submit_button = customtkinter.CTkButton(
            master=self.f_choose_out_dir,
            text="Submit MGR",
            text_font=("Roboto Medium", 14),
            width=200,
            height=40,
            command=self._e_submit_mgr)

        self.b_choose_out_dir = customtkinter.CTkButton(
            master=self.f_choose_out_dir,
            text="Choose Output Folder",
            text_font=("Roboto Medium", 14),
            width=150,
            height=35,
            command=self._e_choose_out_dir)

        self.t_choose_out_dir = customtkinter.CTkLabel(
            master=self.f_choose_out_dir,
            text="Choose your output folder below ",
            text_font=("Roboto Medium", 20, "italic")

        )


        self.f_choose_out_dir.grid(row=row_pos, column=col_pos, padx=padx, pady=pady, sticky='nswe')
        self.h_choose_out_dir.pack(side="top")
        self.t_choose_out_dir.place(relx=0.5, rely=0.4, anchor="center")
        self.b_choose_out_dir.place(relx=0.5, rely=0.5, anchor="center")
        self.b_submit_button.pack(anchor="se", side="bottom", padx=60, pady=40)

    def _init_homepage(self):

        self.homepage = customtkinter.CTkFrame(master=self, )
        self.homepage.grid(row=0, column=1, sticky="nswe", padx=20, pady=20)

        self.homepage.grid_columnconfigure(0, weight=1)
        self.homepage.grid_rowconfigure(0,weight=1)

        self.homepage_label = customtkinter.CTkLabel(
                                            master = self.homepage, text="Homepage",
                                            text_font=("Roboto Medium", 16))  # font name and size in px

        self.homepage_label.grid(row=0, column=0)

    def _init_settings_page(self):
        self.settings_page = customtkinter.CTkFrame(master=self, )
        self.settings_page.grid(row=0, column=1, sticky="nswe", padx=20, pady=20)

        self.settings_page.grid_columnconfigure(0, weight=1)
        self.settings_page.grid_rowconfigure(0,weight=1)

        self.settings_label = customtkinter.CTkLabel(master = self.settings_page, text="Coming Soon",
                                            text_font=("Roboto Medium", 16))  # font name and size in px
        self.settings_label.pack(anchor="n")

    def _e_switch_frame(self, master: customtkinter.CTkFrame):
        master.tkraise()

    def _e_choose_grades_file(self,):
        self.grades_path = filedialog.askopenfilename()
   
    def _e_choose_homerooms_file(self,):
        self.homerooms_path = filedialog.askopenfilename()
    
    def _e_choose_out_dir(self,):
        self.output_dir = filedialog.askdirectory()

    def _callback(self, url):
        webbrowser.open_new_tab(url)

    def _e_change_appearance_mode(self, new_appearance_mode: str):
        customtkinter.set_appearance_mode(new_appearance_mode)

    def _e_on_closing(self, event=0):
        self.destroy()

    def _e_submit_mgr(self,):
        print(self.output_dir)
        new = GradeTables(self.grades_path, self.homerooms_path, output_path = self.output_dir)
        del new

if __name__ == "__main__":
    app = App()
    app.mainloop()