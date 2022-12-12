import pandas as pd
import os
import constants
from utility import *

# remove header styles from pd.to_excel
from pandas.io.formats.excel import ExcelFormatter
ExcelFormatter.header_style = None

class GradeTables():

    def __init__(
        self, 
        grades_path: str = "", 
        homeroom_path: str = "", 
        courses_path: str = "", 
        output_path: str = "",
        resource_homeroom: str = "") -> None:
        """initializes dataframes needed for MGR. Sets them empty if no paths are given

        Args:
            grades_path (str, optional): path of student grades data. Defaults to "".
            homeroom_path (str, optional): path of student homerooms data. Defaults to "".
            courses_path (str, optional): path of student courses data. Defaults to "".
            output_path (str, optional): path to output mgr file. Defaults to "".
        """

        # initialize helper variables
        self.sheets_list = []
        self.resource_homerooms = []
        self.sheets_to_write_list = []
        self.courses = []
        self.output_path = output_path


        # store grades and homerooms dataframes
        self.grades_df = self._store_df_from_file(fpath=grades_path)
        self.homerooms_df = self._store_df_from_file(fpath=homeroom_path)

        # store term and campus name, for output file naming if possible
        try: self.term = self.grades_df.loc[self.grades_df["GR Term"].str.contains("T"), "GR Term"].iloc[0]
        except: self.term = ""
        try: self.campus = self.homerooms_df.loc[0,"School"]
        except: self.campus = ""

        # clean grades and homeroom dataframes
        self._clean_grades_df()
        self._clean_homerooms_df()

        # merge homeroom data into student grades data
        self._aggregate_data()

        # get list of students courses
        self._init_courses()

        # create master df
        self.master_df = self._init_master_df()
        self.master_df = self._create_master_df()

        # determine sheet names for xlsx output
        self._slice_and_store_grade_sheets()
        self._create_mgr()
        

        # write mgr to file
        self._write_mgr_to_file()

     
    def _store_df_from_file(self, fpath: os.path) -> pd.DataFrame:
        
        df = pd.DataFrame({}) # df to return
        file_ext = os.path.splitext(fpath)[1] # to determine pd.read method

        if file_ext == '.csv':
            try: df = pd.read_csv(fpath)
            except: return None

        elif file_ext == '.txt':
            try: df = pd.read_csv(fpath, sep="\t")
            except: return None

        return df

    def _clean_grades_df(self):
        try: self.grades_df = self.grades_df.loc[self.grades_df["Enrollment"]=="active", :]
        except: print("No Enrollment data")
        try: self.grades_df = self.grades_df.loc[self.grades_df["GR Term"]=="T1", :]
        except: print("No Term Data")

        self.grades_df = self.grades_df.loc[:, ["Crs Name", "Crs Num", "Grade", "Student Number", "Student", "Pct",]]

        self.grades_df = self.grades_df.astype({"Grade": str})

        # NOTE: this may be incorrect, but it takes the last occurence of a grade only
        self.grades_df["Duplicates"] = self.grades_df.duplicated(subset=["Student Number", "Crs Name"], keep="last")
        self.grades_df = self.grades_df.loc[self.grades_df["Duplicates"]==False, :]
        self.grades_df = self.grades_df.drop(columns=["Duplicates"])

    def _clean_homerooms_df(self):
        self.homerooms_df = self.homerooms_df.loc[self.homerooms_df["Enroll Status"]=="Active", :]
        self.homerooms_df = self.homerooms_df.loc[:, ["Student Number", "Student", "Grade", "Home_Room"]]
        self.homerooms_df.loc["Home_Room"] = self.homerooms_df["Home_Room"].str.lower()
        self.homerooms_df.loc[self.homerooms_df["Home_Room"].str.lower() == "michigan", "Home_Room"] = "MICHIGAN STATE"
        self.homerooms_df["Home_Room"] = self.homerooms_df["Home_Room"].str.capitalize()

    def _aggregate_data(self):
        self.grades_df = pd.merge(self.grades_df, self.homerooms_df[["Student Number", "Home_Room"]], "left", ["Student Number"], )
        self.grades_df = self.grades_df[["Home_Room"] + self.grades_df.columns[:-1].to_list()]

    def _init_master_df(self) -> pd.DataFrame:
        df = pd.DataFrame({})
        df[["Home_Room", "Student", "Grade"]] = self.grades_df.loc[:, ["Home_Room", "Student", "Grade"]].drop_duplicates()
        return df

    def _init_courses(self):
        self.courses = self.grades_df["Crs Name"].unique().tolist()
        for crs in self.courses:
            iscrs = get_crs(crs)
            if iscrs == -1:
                self.courses.remove(crs)
                self.grades_df = self.grades_df.loc[self.grades_df["Crs Name"]!=crs, :]
        
    def _create_master_df(self):

        for crs in self.courses:
            df = self.grades_df.loc[self.grades_df["Crs Name"] == crs, ["Student", "Pct"]]
            df.rename(columns={"Pct": crs}, inplace=True) 

            self.master_df = pd.merge(self.master_df, df, "left", on="Student")

        return self.master_df

    def _write_mgr_to_file(self):
    
        # set ouput file name
        # fname = create_filename_intelligently(constants.OUTPUT_PATH, self.campus, self.term)
        sheet_names = self._determine_sheet_names(x[0] for x in self.sheets_list)

        fname = create_filename_intelligently(self.output_path, self.campus, self.term) 

        # initialize writer
        writer = pd.ExcelWriter(path=fname, engine="xlsxwriter")

        # create sheet summary, here so that it will be the first sheet in the file
        pd.DataFrame({}).to_excel(writer, sheet_name="Summary", index=False)

        # prep and write grades sheets
        for df, sheet in zip(self.sheets_to_write_list, sheet_names):

            # prep and write grades sheets            
            df = self._clean_grades_df_for_output(df)
            df.to_excel(writer, sheet_name=sheet, index=False) 

            format_grades(df, sheet, writer) # format the grades file for clean look

        # prep and write master df
        self.master_df = self._clean_master_df_for_output(self.master_df)
        self.master_df.to_excel(writer, "Full School MGR", index=False)
        format_grades(self.master_df, "Full School MGR", writer)
        worksheet = writer.sheets['Full School MGR']
        worksheet.hide()

        # prep and write summary sheet
        summary_df = create_summary_sheet(self.sheets_list)
        summary_df.to_excel(writer, "Summary", index=False, )
        honor_roll_df = create_honor_roll_table(self.master_df.columns.get_loc("Honor Roll Status"))
        honor_roll_df.to_excel(writer, "Summary", header=False, index=False, startrow=0, startcol=10)
        format_summary(summary_df, "Summary", writer)

        writer.close()

    def _clean_master_df_for_output(self, df:pd.DataFrame):

        if type(df) == None: return

        # initialize and correct order of master_df columns
        df.loc[:, ["# Failing", "Average", "Honor Roll Status"]] = ""
        grades_col = df.columns[3:-3].sort_values().tolist()
        df = df[
            ["Grade", "Home_Room", "Student", "# Failing"] + 
            grades_col + 
            ["Average", "Honor Roll Status"]
        ].sort_values(["Grade", "Student"])

        df = df.drop_duplicates("Student") # clean master_df of duplicates
        df = write_formulas_grade_sheet(df)

        # replace empty columns with space so formatting doesn't highlight them read in excel
        for col in grades_col: df.loc[df[col].isnull(), col] = " "
        return df

    def _clean_grades_df_for_output(self, df:pd.DataFrame):

        if type(df) == None: return
        df = df.drop_duplicates("Student")
        df = df.drop(columns="Grade")
        df = df.sort_values(["Home_Room", "Student"])

        # Add columns that will show in xlsx file
        df[["Average", "#_Failing", "Honor_Roll_Status"]] = ""
        reordered_cols = ["Home_Room", "Student", "#_Failing"] + df.columns.tolist()[2:-3] + ["Average", "Honor_Roll_Status"]
        df = df[reordered_cols]

        # Remove all extra characters from column names
        for col in df.columns:
            new_col = col.replace("_", " ")
            df = df.rename(columns={col: new_col})

        df = write_formulas_grade_sheet(df) 
        return df

    def _create_mgr(self):
    
        # create each sheet in mgr
        for sheet_name in self.sheets_list:
            col = sheet_name[1]     # column name to reference in df 
            value = sheet_name[0]   # column value to look for in df
            courses = []            # courses to input grades for in df

            if col == "Grade":
                sheet_df = self.master_df.loc[
                    (self.master_df[col] == value) & 
                    (self.master_df["Home_Room"].isin(self.resource_homerooms) == False), 
                    :]
                courses = self.grades_df.loc[
                    (self.grades_df[col] == value) & 
                    (self.grades_df["Home_Room"].isin(self.resource_homerooms) == False), 
                    "Crs Name"].unique().tolist()

            elif col == "Home_Room":
                sheet_df = self.master_df.loc[self.master_df[col] == value, :]
                courses = self.grades_df.loc[self.grades_df[col] == value, "Crs Name"].unique().tolist()


            courses = self._determine_courses(courses)

            sheet_df = sheet_df[["Student", "Home_Room", "Grade"] + courses]

            for column in sheet_df.columns:
                
                crs = get_crs(column)
                
                if crs != -1: 
                    sheet_df = sheet_df.rename(columns={column: crs})
            
            curr_cols = sheet_df.columns.to_list()
            for column in curr_cols:
                if column == "ELA" and "Reading" in curr_cols: 
                    sheet_df = sheet_df.rename(columns={"ELA": "Reading"})


            #https://stackoverflow.com/questions/9835762/how-do-i-find-the-duplicates-in-a-list-and-create-another-list-with-them
            seen = set()
            dupes = [x for x in sheet_df.columns if x in seen or seen.add(x)]
            
            for column in dupes:
                temp_df = sheet_df.loc[:, [column, "Student"]]
                sheet_df[column] = (temp_df.iloc[:, [0,2]].fillna(0) + temp_df.iloc[:, [1,2]].fillna(0))[column]
                
            
            sheet_df = sheet_df.loc[:,~sheet_df.columns.duplicated()].copy()

            self.sheets_to_write_list.append(sheet_df)
        
    def _determine_courses(self, courses) -> list:
        for i, crs in zip(range(0, len(courses)),courses):
            if "homeroom" in crs.lower(): del courses[i] 
            
        return courses

    def _slice_and_store_grade_sheets(self):
        
        # sheet names for creating pandas sheets
        temp_sheets_list = []
        grades_list = self.grades_df["Grade"].unique().tolist()
        grades_list.sort()

        for grade in grades_list:
            temp_sheets_list.append([grade, "Grade", grade])

            # slice from grades df to look at students in this grade
            df = self.grades_df.loc[self.grades_df["Grade"] == grade, :]
            homerooms_list = df.loc[df["Home_Room"].isnull() == False, "Home_Room"].unique().tolist()
            
            
            for hr in homerooms_list:
                if len(self.master_df.loc[self.master_df["Home_Room"]==hr, "Grade"].unique()) > 1:
                    max = self.master_df.loc[self.master_df["Home_Room"] == hr, "Grade"].max()
                    temp_sheets_list.append([hr, "Home_Room", str(max)])
                    self.resource_homerooms.append(hr)

        # drop duplicate sheet names
        t = []
        for i, item in zip(range(0, len(temp_sheets_list)), temp_sheets_list):
            if item[0] in t: 
                del temp_sheets_list[i]
                continue
            
            t.append(item[0])

        
        # sort sheet names to put them in order and store data
        temp_sheets_list = sorted(temp_sheets_list,key=lambda x: (x[2],x[1]))
        self.sheets_list = temp_sheets_list 

    def _determine_sheet_names(self, names:list):
        sheet_names = []
        for name in names:
            temp = str(name)
            if temp == "0": sheet_names.append("Kindergarten")
            elif temp == "1": sheet_names.append("1st Grade")
            elif temp == "2": sheet_names.append("2nd Grade")
            elif temp == "3": sheet_names.append("3rd Grade")
            elif temp == "4": sheet_names.append("4th Grade")
            elif temp == "5": sheet_names.append("5th Grade")
            elif temp == "6": sheet_names.append("6th Grade")
            elif temp == "7": sheet_names.append("7th Grade")
            elif temp == "8": sheet_names.append("8th Grade")
            elif temp == "9": sheet_names.append("9th Grade")
            elif temp == "10": sheet_names.append("10th Grade")
            elif temp == "11": sheet_names.append("11th Grade")
            elif temp == "12": sheet_names.append("12th Grade")
            else: sheet_names.append(temp.upper())
        return sheet_names

    def _clean_courses_df(self):
        self.courses_df = self.courses_df.loc[:, ["Course Name", "Course.Section"]]

# Utility Functions
def create_filename_intelligently(dir: str, campus: str = "", term: str = "") -> str:
    """Creates a filename intelligently without overwriting other MGR"s in the directory

    Args:
        dir (str): String representing the directory to save the file in

    Returns:
        str: returns full filepath of file
    """
    filename = campus + " - " + term + " MGR"
    file_ext = ".xlsx"
    file_copy_num = 0
    # Check if file exists. If so, name a copy
    for i in range(0,1000):
        temp = ""        
        if file_copy_num == 0:
            temp = filename + file_ext
        else:
            temp = filename + " (" + str(file_copy_num) +")" + file_ext
        
        filepath = os.path.join(dir, temp)

        if os.path.exists(filepath) != True:
            return filepath
        
        else:
            file_copy_num += 1

    
def get_crs(crs: str) -> str:
    """Checks for unstandardized names in course name
       and returns the Standardized Course Name for 
       MGR  

    Args:
        crs (str): unstandardized course name

    Returns:
        str: standardized course name
    """
    
    # make string lowercase for case handling
    temp = crs
    crs = crs.lower()

    if "homeroom" in crs: return -1
    if "homework" in crs: return -1
    if "ela intervention" in crs: return "ELA Intervention"
    if "ela" in crs: return "ELA"
    if "kindergarten " in crs: return crs.replace("kindergarten ", "").title()
    if "1st grade " in crs: return crs.replace("1st grade ", "").title()
    if "2nd grade " in crs: return crs.replace("2nd grade ", "").title()
    if "3rd grade " in crs: return crs.replace("3rd grade ", "").title()
    if "4th grade " in crs: return crs.replace("4th grade ", "").title()

    # Courses Checks
    if "math" in crs: return "Math"
    if "algebra" in crs: return "Algebra"
    if "computer science" in crs: return "Computer Science"
    if "science" in crs: return "Science"
    if "humanities" in crs: return "Humanities"
    if "english" in crs: return "ELA"
    if "reading comprehension" in crs: return "Reading Comprehension"
    if "reading" in crs: return "Reading"
    if "writing" in crs: return "Writing"
    if "social studies" in crs: return "Social Studies"
    if "phonics" in crs: return "Phonics"
    if "history" in crs: return "History"
    if "biology" in crs: return "Biology"
    if "geometry" in crs: return "Geometry"
    if "spanish" in crs: return "Spanish"
    if "physed" in crs: return "Phys Ed."
    if "chemistry" in crs: return "Chemistry."
    if "pre-calculus" in crs: return "Pre-Calculus"
    if "physics" in crs: return "Physics"
    if "language" in crs: return "Language"
    if "literacy" in crs: return "Financial Literacy"
    if "senior seminar" in crs: return "Senior Seminar"
    if "calculus" in crs: return "Calculus"
    if "statistics" in crs: return "Statistics"
    if "literature" in crs: return "Literature"
    if "government" in crs: return "Government"

    if "college and career readiness" in crs: return "College and Career Readiness"

    # Physical Education Check
    if "p.e." in crs: return "Phys Ed."
    if "p.e" in crs: return "Phys Ed."
    if "pe" in crs: return "Phys Ed."
    if "pe." in crs: return "Phys Ed."
    if "sel" in crs: return "SEL"



    if "physedhealth" in crs: 
        return "Phys_Ed."

    # return NaN value 
    return -1


if __name__ == "__main__":
    out = r'out'
    g_dir = r"test_files/grades"
    hr_dir = r"test_files/homerooms"
      


    tbl = GradeTables(        
        grades_path="test_files/grades/HES.csv", 
        homeroom_path="test_files/homerooms/HES.csv",
        courses_path=constants.COURSES_PATH, 
        output_path=constants.OUTPUT_PATH)

    
