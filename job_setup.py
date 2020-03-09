import sys
import os
import shutil
import re
import ConfigParser


def main():
    job_details = JobDetails()
    job_details.getUserInput()
        
    folder_mgr = FolderManager(job_details)
    folder_mgr.createJob()
    
    createConfigFile(folder_mgr)
 

class JobDetails(object):
    """ Class to get the Params Module variables.
    Can be obtain either by user input or config file """
    
    def __init__(self):
        
        self.JobNum = None
        self.Maildate = None 
        self.AddBlank = None
        self.recordsPerGroup = "500"

    def provideJobNumber(self):
        """ Get 5 digit Mailshop job ticker number from input """
        self.JobNum = raw_input("\r\n-- Provide DS job number:  ")
        while re.match(r'^\d{5}$', self.JobNum) == None:
            self.JobNum = raw_input("Job number is not valid !!! -->  ")
            
    def provideMailDate(self):
        """ Date on Letter. Format (YYYYMMDD) """
        self.Maildate = raw_input("\r\n-- Provide Letter Date (YYYYMMDD):  ")
        while re.match(r'^20\d{6}$', self.Maildate) == None:
            self.Maildate = raw_input("Date is not valid !!! -->  ")        

    def selectAddBlank(self):
        """ Determine if blanks will be added. Add blanks if printing duplex. """
        add_blank_dict = {"Y" : "True", "N" : "False"}
        add_blank_choice = raw_input("\r\n-- Add Blanks for duplex printing (Y or N)?: ").upper()
        while add_blank_choice not in add_blank_dict.keys():
            add_blank_choice = raw_input("Invalid choice !!! -->  ").upper()            
        self.AddBlank = add_blank_dict[add_blank_choice]
        
    def provideRecordsPerGroup(self):
        """ Determine number of records for each print file. 500 records in the default. """
        recordcount_choice = raw_input("\r\n-- Split print files every 500 Records (Y or N)?: ").upper()
        while recordcount_choice not in ["Y", "N"]:
            recordcount_choice = raw_input("Invalid choice !!! -->  ").upper()
        if recordcount_choice == "N":
            self.recordsPerGroup = raw_input("\r\n-- Provide record count per print file:  ")
            while not self.recordsPerGroup.isnumeric(): 
                self.recordsPerGroup = raw_input("Invalid input !!! -->  ").upper()
        
    def getUserInput(self):
        """ Get input from the user. """
        self.provideJobNumber()
        self.provideMailDate()
        self.selectAddBlank()
        self.provideRecordsPerGroup()
        
    def getConfigInput(self, config_file):
        """ Get Params from Config file. Used when 
        creating samples and print files. """
        config = ConfigParser.ConfigParser()
        with open(config_file, 'rb') as config_handle:
            config.readfp(config_handle)
            self.JobNum = config.get('Params', 'JobNum')
            self.Maildate = config.get('Params', 'Maildate')
            self.AddBlank = config.get('Params', 'AddBlank')
            self.recordsPerGroup = config.get('Params', 'recordsPerGroup')
            self.RecordRange = config.get('Params', 'RecordRange')        
        
        
class FolderManager(object):
    """ Class to set and create common folder structure for all jobs. """

    
    def __init__(self, job_details):
        """ Set the job paths  """
        self.JobNum          = job_details.JobNum
        self.Maildate        = job_details.Maildate 
        self.AddBlank        = job_details.AddBlank
        self.recordsPerGroup = job_details.recordsPerGroup
        self.RecordRange     = ""

        self.ast_folder = r'P:\AST'
        self.project_folder = os.path.join(self.ast_folder, "Aged Loss")
        
        self.job_folder = os.path.join(self.ast_folder, "{}_Aged_Loss".format(self.JobNum))
        self.data_folder = os.path.join(self.job_folder, "Data")
        self.reports_folder = os.path.join(self.job_folder, "Reports")
        self.sample_folder = os.path.join(self.job_folder, "Sample")
        self.print_folder = os.path.join(self.job_folder, "Print")
       
    def createJobFolders(self):
        """ Create the job folders """
        os.makedirs(os.path.join(self.data_folder, "1_original_data"))
        os.mkdir(os.path.join(self.data_folder, "2_combined"))
        os.mkdir(os.path.join(self.data_folder, "3_misc"))
        os.mkdir(os.path.join(self.data_folder, "4_client_excel"))
        os.mkdir(os.path.join(self.data_folder, "5_overnight_labels"))
        os.mkdir(self.reports_folder)
        os.mkdir(self.sample_folder)
        os.makedirs(os.path.join(self.print_folder, "ps"))

    def copy_coversheet_to_folder(self):
        
        """ Copy the print coversheet from the 
        project folder to the the job folder. """
        # Copy over the coversheet
        coversheet = "Aged Loss - Print Files.xlsx"
        cover_src = os.path.join(self.project_folder, "coversheets", coversheet)
        cover_dst = os.path.join(self.job_folder, coversheet)
        shutil.copy2(cover_src, cover_dst)
            
    def createJob(self):
        """ Create the job folders and copy the coversheet """
        if not os.path.exists(self.job_folder):
            self.createJobFolders()
            self.copy_coversheet_to_folder()
        else:
            print '"{}" already exists. No folders created.'.format(self.job_folder)
            print ""
            
            
def createConfigFile(fldr_mgr):
    """ Set the data and params inputs for PrintNet.  Write to 
    config file to use for creating sample and print files. """

    overnight_data = os.path.join(fldr_mgr.data_folder, "Overnight.csv")
    addr_6_sheets = os.path.join(fldr_mgr.data_folder, "Address_6_Sheets.csv")
    addr_7_sheets = os.path.join(fldr_mgr.data_folder, "Address_7_Sheets.csv")
    samples_6_sheets = os.path.join(fldr_mgr.data_folder, "samples 6 Sheets.csv")
    samples_7_sheets = os.path.join(fldr_mgr.data_folder, "samples 7 Sheets.csv")
    static_data = os.path.join(fldr_mgr.data_folder, "StaticData.dat")   
    
    job_config_file = os.path.join(fldr_mgr.data_folder, "{}_config.txt".format(fldr_mgr.JobNum))
    
    with open(job_config_file, 'wb') as configHandle:
        config = ConfigParser.ConfigParser()

        config.add_section('Folder')
        config.set('Folder', 'job_folder', fldr_mgr.job_folder)
        
        config.add_section('DataInput')
        config.set('DataInput', 'Overnight', overnight_data)
        config.set('DataInput', 'Address_6_Sheets', addr_6_sheets)
        config.set('DataInput', 'Address_7_Sheets', addr_7_sheets)
        config.set('DataInput', 'samples_6_sheets', samples_6_sheets)
        config.set('DataInput', 'samples_7_sheets', samples_7_sheets)
        config.set('DataInput', 'StaticData', static_data)

        config.add_section('Params')
        config.set('Params', 'JobNum', fldr_mgr.JobNum)
        config.set('Params', 'Maildate', fldr_mgr.Maildate)
        config.set('Params', 'AddBlank', fldr_mgr.AddBlank)
        config.set('Params', 'recordsPerGroup', fldr_mgr.recordsPerGroup)
        config.set('Params', 'RecordRange', fldr_mgr.RecordRange)

        config.write(configHandle) 

    
  
if __name__ == '__main__':
    main()
