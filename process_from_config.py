'''
Count the number of companies and records per company in the original text data


'''

___author__ = "Shaun Thomas"
___date__ = "July 24, 2019"
__version__ = 1.0
__modified_by__=""
__last_modified__=""



import os
import sys
import csv
import subprocess
import ConfigParser
import job_setup



def main():

    # Get Params variables from config file. 
    config_file = os.path.abspath(sys.argv[1])

    print_params = PrintnetParameters(config_file)  

    # Review the parameters and ensure the correct info is being processed.
    # If data has been presorted, make sure the config file is manually updated. 
    print "\r\n".join(["",
                       "Job to process:",
                       "",
                       "Job Number:  {}".format(print_params.JobNum),
                       "Maildate:    {}".format(print_params.Maildate), 
                       "AddBlank: {}".format(print_params.AddBlank),
                       "recordsPerGroup:   {}".format(print_params.recordsPerGroup),
                       "RecordRange:   {}".format(print_params.RecordRange),
                       "",
                       "Overnight Data:   {}".format(print_params.overnight_data),
                       "6 Sheet Data:   {}".format(print_params.addr_6_sheets),
                       "7 Sheet Data:   {}".format(print_params.addr_7_sheets),
                       ""])
                       
    print_params.generateOutput()
  

class PrintnetParameters(object):
    
    def __init__(self, config_file):
        """ Class to: 
        1) Set paths and parameters for PrintNet.
        2) Set command line arguments for processing
        3) Create samples and print files. """
    
        # Set PrintNet Params module variables from Config
        self.job_details = job_setup.JobDetails()
        self.job_details.getConfigInput(config_file)
        
        self.JobNum = self.job_details.JobNum
        self.Maildate = self.job_details.Maildate 
        self.AddBlank = self.job_details.AddBlank
        self.recordsPerGroup = self.job_details.recordsPerGroup
        self.RecordRange = self.job_details.RecordRange
        
        # Set Data files to process in PrintNet from Config
        self.getDataFilesFromConfig(config_file)
        
        # Job directories
        self.fldr_mgr = job_setup.FolderManager(self.job_details)
        self.data_folder = self.fldr_mgr.data_folder
        self.sample_folder = self.fldr_mgr.sample_folder
        self.print_folder = self.fldr_mgr.print_folder
    
        # Common PrintNet config values
        self.gmc_exe = "G:\PrintNet T Designer\PNetTC.exe"        
        self.printnet_folder = os.path.join(self.fldr_mgr.project_folder, "PrintNet")        
        self.gmc_workflow = os.path.join(self.printnet_folder, "Aged_Loss_Letters.wfd")
        self.gmc_config = os.path.join(self.printnet_folder, "Aged_Loss.job")
        self.printer_config = "Docutech"
        self.driver_config = "Aged_Loss"
        self.engine = "AdobePostScript3"
        self.job_log = os.path.join(self.data_folder, "{}.log".format(self.JobNum))
        
        # Set common PrintNet Command Line arguments
        self.gmc_job_params = self.set_gmc_job_params()
        self.set_gmc_configs = self.set_gmc_configs()

    def getDataFilesFromConfig(self, config_file):
        """ Get the data files to process from the config file. """
        
        config = ConfigParser.ConfigParser()
        with open(config_file, 'rb') as config_handle:
            config.readfp(config_handle)
            self.overnight_data = config.get('DataInput', 'overnight')
            self.addr_6_sheets = config.get('DataInput', 'address_6_sheets')
            self.addr_7_sheets = config.get('DataInput', 'address_7_sheets')
            self.samples_6_sheets = config.get('DataInput', 'samples_6_sheets')
            self.samples_7_sheets = config.get('DataInput', 'samples_7_sheets')
            self.static_data = config.get('DataInput', 'staticdata')

    def set_gmc_job_params(self):
        """ Set the command line arguments common for all processing. """
        self.addr_module = "-difAddrData"
        
        return [self.gmc_exe, self.gmc_workflow, 
        "-JobNumParams", self.JobNum,
        "-MaildateParams", self.Maildate,
        "-AddBlankParams", self.AddBlank,
        "-recordsPerGroupParams", self.recordsPerGroup,
        "-RecordRangeParams", self.RecordRange,
        "-difStaticData",  self.static_data]
        
    def set_gmc_configs(self):
        """ Set the PrintNet configs. Used for sample and print 
        files only. Excluded for creating print count files. """
        return ["-c", self.gmc_config,
        "-pc", self.printer_config,
        "-dc", self.driver_config,
        "-e", self.engine,    
        "-la", self.job_log]
         
    def createSampleOutput(self):
        """ Create sample output for 6 and 7 sheets. 
        Ensure that the sample files have been created. """
        sample_filename = r'P:\AST\%h04_Aged_Loss\Sample\%h04 Aged Loss - %h08 - %h05.%e'
        sample_args = ["-o", "Sample", "-f", sample_filename, "-splitbygroup"]
        printnet_commands = self.gmc_job_params + self.set_gmc_configs + sample_args
        
        subprocess.call(printnet_commands + [self.addr_module, self.samples_6_sheets])
        subprocess.call(printnet_commands + [self.addr_module, self.samples_7_sheets])
        
    def createPrintOutput(self):
        """ Create print output for all groups. """
        print_filename = r'P:\AST\%h04_Aged_Loss\Print\%h08\%h04 Aged Loss - %h08 - %h09.%e'
        print_args = ["-o", "Print", "-f", print_filename, "-splitbygroup"]                      
        printnet_commands = self.gmc_job_params + self.set_gmc_configs + print_args
        
        subprocess.call(printnet_commands + [self.addr_module, self.overnight_data])
        subprocess.call(printnet_commands + [self.addr_module, self.addr_6_sheets])
        subprocess.call(printnet_commands + [self.addr_module, self.addr_7_sheets])
        
    def createPrintCountsOutput(self):                     
        """ Create print count text files for all groups """
        printcount_args = ["-o", "PrintCounts", 
                           "-dataoutputtype", 
                           "CSV", "-datacodec", "UTF-8"]
                           
        printnet_commands = self.gmc_job_params + printcount_args
        counts_overnight = os.path.join(self.data_folder, "3_misc", "counts_overnight.TXT") 
        counts_6_Sheets = os.path.join(self.data_folder, "3_misc", "counts_6_Sheets.TXT") 
        counts_7_Sheets = os.path.join(self.data_folder, "3_misc", "counts_7_Sheets.TXT")
        
        subprocess.call(printnet_commands + [self.addr_module, self.overnight_data, "-f", counts_overnight])
        subprocess.call(printnet_commands + [self.addr_module, self.addr_6_sheets, "-f", counts_6_Sheets])
        subprocess.call(printnet_commands + [self.addr_module, self.addr_7_sheets, "-f", counts_7_Sheets])
        
    def generateOutput(self):   
        """ Select and create sample an/or print output """
        choice_list = ["S", "P", "B"]
        output_choice = raw_input("Output: Sample (s), Print (p) or Both (b): ").upper()
        while output_choice not in choice_list:
            output_choice = raw_input("Invalid input !!! -->  ").upper()    
        
        if output_choice in ["S", "B"]:
            self.createSampleOutput()
            
        if output_choice in ["P", "B"]:
            self.createPrintOutput()
            self.createPrintCountsOutput()
      
   
if __name__ == '__main__':
    main()


