# -*- coding: utf-8 -*-
"""
Created on Wed Sep 18 22:37:07 2019

@author: shreya.aggarwal, visheshtandon1
"""
import os, json, shutil, logging, datetime as dt, tkinter, dropbox, zipfile, win32com, argparse
from tkinter import messagebox
from win32com import client

class FileTransfer:

    def __init__(self):
        self.__log = ""
        pass


    def create_log(self):
        
        ###create the logger
        CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
        logname = 'PluggableFileTransfer_Log_' + str(dt.datetime.now().strftime('%Y_%m_%d_%H%M%S'))

        self.__log = logging.getLogger('Date_Time_Log')
        self.__log.setLevel(logging.INFO)
        
        ###set up console handler with stringio object
        if not os.path.isdir(CURRENT_DIR+'/../PluggableFileTransfer_Logs/'):
            CURRENT_DIR+='/../PluggableFileTransfer_Logs/'
            os.mkdir(CURRENT_DIR)
        else:
            CURRENT_DIR+='/../PluggableFileTransfer_Logs/'
        fh = logging.FileHandler(CURRENT_DIR+'/'+logname)
        
        fh.setLevel(logging.INFO)
        
        ### optional added formatter
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        fh.setFormatter(formatter)
        
        ### add console handler to logger
        self.__log.addHandler(fh)
        
        self.__log.info('Log Created', exc_info=True)
        return self.__log, logname


    def read_config(self, filename):

        try:
            with open(filename,"r") as config_file:
                data = json.load(config_file)
            if(('local machine' or 'dropbox' or 'outlook') in data['movement_type'].keys()):
                self.__log.info('Loaded vars from %s' % (filename), exc_info=True)
                return data
            else: 
                self.__log.error('Unable to load vars from %s' % (filename), exc_info=True)
                return "Error"
                # raise Exception('Unable to load vars from %s' % (filename))
        except:
            self.__log.error('Unable to load vars from %s' % (filename), exc_info=True)
            return "Error"


    def __filesCopyUtility(self, src_dir, dest_dir, file_name, overwrite):
        
        flag = 0
        if(os.path.isdir(dest_dir)):
            #Check if Target/Destination Directory exists or not

            if(os.path.isfile(dest_dir+file_name)):
                #Check if File exists in Target/Destination Directory

                if(overwrite.lower()=='yes'):
                    #Check if Overwrite Flag is set or not.

                    try:
                        os.remove(dest_dir + file_name)
                        self.__log.info("Destination File "+ dest_dir + file_name + " removed from " + dest_dir + " .",exc_info=True)
                        shutil.copy(src_dir+file_name, dest_dir)
                        self.__log.info("Source File "+ src_dir + file_name + " successfully copied to " + dest_dir + " .",exc_info=True)
                        flag=1

                    except:
                        self.__log.error("File "+ src_dir + file_name + " transfer process failed.",exc_info=True)
                        flag=0
                    
                else:
                    
                    flag=0
                    self.__log.error("File Transfer Incomplete! Overwrite = 'No'",exc_info=True) 
            
            else:
                try:
                    os.chmod(dest_dir,777)
                    shutil.copy(src_dir+file_name, dest_dir)
                    self.__log.info("Source File "+ src_dir + file_name + " successfully copied to " + dest_dir + " .",exc_info=True) 
                    flag=1
                
                except:
                    self.__log.error("File "+ src_dir + file_name + " transfer process failed.",exc_info=True)
                    flag=1
        
        else:
            
            os.makedirs(dest_dir)
            shutil.copy(src_dir+file_name, dest_dir)
            self.__log.info("Source File "+ src_dir + file_name + " successfully copied to " + dest_dir + " .",exc_info=True) 
            flag=1
        
        return flag


    def copyFile(self, src_dir, dest_dir, file_name, overwrite):        
        #Copy a single file

        flag=0
        if(not src_dir.endswith('/')):
            src_dir+='/'
        if(not dest_dir.endswith('/')):
            dest_dir+='/'

        src_files =  [i for i in os.listdir(src_dir) if(os.path.isfile(src_dir+i)) and (file_name.lower() in i.lower())]

        if(len(src_files)==0):
            self.__log.error("Source File "+ src_dir + file_name + "doesn't exist.",exc_info=True) 
            messagebox.showinfo("Alert","Source Folder "+ src_dir + file_name + " doesn't exist.")
            flag=0

        elif(len(src_files)==1):
            self.__filesCopyUtility(src_dir, dest_dir, src_files[0], overwrite)
            flag=1

        else:
            status = []
            for file in src_files:
                status.append(self.__filesCopyUtility(src_dir, dest_dir, file, overwrite))
            if(all(status)):
                flag = 1

        return flag      


    def copyFilesFromDirectory(self, src_dir, dest_dir, file_name,overwrite):
        #Copy a single file

        flag=0
        if(src_dir.endswith('/')):
            src_dir=src_dir[:-1]
        if(not dest_dir.endswith('/')):
            dest_dir = dest_dir+ '/'
        #dest_dir+=os.path.basename(os.path.dirname(src_dir))
        dest_dir+=os.path.basename(src_dir)
        if(os.path.isdir(src_dir)):
            #Check if source file exists or not

            if(os.path.isdir(dest_dir)):
                #Check if Target/Destination Directory exists or not

                if(os.path.exists(dest_dir+file_name)):
                    #Check if File exists in Target/Destination Directory

                    if(overwrite.lower()=='yes'):
                        
                        try:
                            shutil.rmtree(dest_dir)
                            self.__log.info("Destination Directory "+ dest_dir + " removed.",exc_info=True)
                            shutil.copytree(src_dir,dest_dir)
                            self.__log.info("Contents of Source Directory "+ src_dir + " successfully copied to " + dest_dir + " .",exc_info=True)
                            flag=1
                        except:
                            self.__log.error("Destination Directory " + dest_dir + " transfer process failed.",exc_info=True)
                            flag=0
                    else:
                        self.__log.error("File Transfer Incomplete! Overwrite = 'No'",exc_info=True) 
                        flag=0
                    
                else:
                    try:
                        #print("print "+src_dir+'\n'+dest_dir+'\n'+file_name)
                        os.chmod(dest_dir,777)
                        shutil.copy(src_dir+file_name, dest_dir)
                        self.__log.info("Source File "+ src_dir + file_name + " successfully copied to " + dest_dir + " .",exc_info=True) 
                        flag=1
                    
                    except:
                        self.__log.error("File "+ src_dir + file_name + " transfer process failed.",exc_info=True)
                        flag=0
            
            else:
                #Destination directory doesn't exist

                shutil.copytree(src_dir, dest_dir)
                self.__log.info("Source Directory "+ src_dir + " successfully copied to " + dest_dir + " .",exc_info=True) 
                flag=1

        else:
            self.__log.error("Source Directory " + src_dir + " doesn't exist.",exc_info=True) 
            messagebox.showinfo("Alert","Source Folder "+ src_dir + file_name + " doesn't exist.")
            flag=0

        return flag      


    def copyDirectoryFromDropbox(self, src_dir, dest_dir, file_name,overwrite,access_token):
        
        flag=0
        #dropbox.Dropbox(access_token).users_get_current_account().email
        #Use it to validate access_token
        try:
            
            session=dropbox.Dropbox(access_token)
            if(session.users_get_current_account().email!=''):
                
                #session=dropbox.Dropbox(access_token)
                
                if(not src_dir.endswith('/')):
                    src_dir+='/'
                if(not src_dir.startswith('/')):
                    src_dir= '/' + src_dir
                if(not dest_dir.endswith('/')):
                    dest_dir+='/'
                #dest_dir=dest_dir + os.path.basename(os.path.dirname(src_dir)) + '/'
            
                try:
                    
                    if(session.files_list_folder(src_dir)):
                        
                        if(os.path.isdir(dest_dir)):
                            
                            if(os.path.exists(dest_dir+os.path.basename(os.path.dirname(src_dir))+'/'+file_name)):
                                
                                if(overwrite.lower()=='yes'):
                                    
                                    try:
                                        shutil.rmtree(dest_dir+os.path.basename(os.path.dirname(src_dir))+'/')
                                        
                                        self.__log.info("Destination Directory "+ dest_dir + " removed.",exc_info=True)
                                        
                                        session.files_download_zip_to_file(dest_dir+'Dropbox Downloads.zip',src_dir)
                                        
                                        # Unzipping folder
                                        with zipfile.ZipFile(dest_dir+"Dropbox Downloads.zip","r") as zip_ref:
                                            zip_ref.extractall(dest_dir)
                                            
                                        # Removing zip files
                                        os.remove(dest_dir+'Dropbox Downloads.zip')
                                        
                                        self.__log.info("Contents of Source Directory "+ src_dir + " successfully copied to " + dest_dir + " .",exc_info=True)
                                        flag=1
                                    except:
                                        #print("Destination Directory " + dest_dir + " transfer process failed.")
                                        self.__log.error("Destination Directory " + dest_dir + " transfer process failed.",exc_info=True)
                                        flag=0
                                else:
                                    #print("File Transfer Incomplete! Overwrite = 'No'")
                                    self.__log.error("File Transfer Incomplete! Overwrite = 'No'",exc_info=True) 
                                    flag=0
                            else:
                                
                                session.files_download_zip_to_file(dest_dir+'Dropbox Downloads.zip',src_dir)
                                        
                                # Unzipping folder
                                with zipfile.ZipFile(dest_dir+"Dropbox Downloads.zip","r") as zip_ref:
                                    zip_ref.extractall(dest_dir)
                                    
                                # Removing zip files
                                os.remove(dest_dir+'Dropbox Downloads.zip')
                                
                                self.__log.info("Source File "+ src_dir + file_name + " successfully copied to " + dest_dir + " .",exc_info=True) 
                                flag=1
                                
                        else:
                            os.makedirs(dest_dir)
                            session.files_download_zip_to_file(dest_dir+'Dropbox Downloads.zip',src_dir)
                            # Unzipping folder
                            with zipfile.ZipFile(dest_dir+"Dropbox Downloads.zip","r") as zip_ref:
                                zip_ref.extractall(dest_dir)
                                
                            # Removing zip files
                            os.remove(dest_dir+'Dropbox Downloads.zip')
                            
                            self.__log.info("Source Directory "+ src_dir + " successfully copied to " + dest_dir + " .",exc_info=True) 
                            flag=1
                            
                    else:
                        
                        self.__log.error("Source Directory " + src_dir + " doesn't exist.",exc_info=True) 
                        flag=0
                        
                except:
                    
                    self.__log.error("Source Directory " + src_dir + " doesn't exist.",exc_info=True) 
                    flag=0
        
            else:
                self.__log.error("Access Token " + access_token + " is invalid.",exc_info=True) 
                flag=0
                
        except:
            self.__log.error("Access Token " + access_token + " is invalid.",exc_info=True) 
            flag=0 
            
        return flag


    def copyFileFromDropbox(self, src_dir, dest_dir, file_name,overwrite,access_token):
        #Copy a single file
                
        flag=0
        
        try:
            
            session=dropbox.Dropbox(access_token)
            if(session.users_get_current_account().email!=''):
                
                #dropbox.Dropbox(access_token).users_get_current_account().email
                #Use it to validate access_token
                if(not src_dir.endswith('/')):
                    src_dir+='/'
                if(not src_dir.startswith('/')):
                    src_dir = '/' + src_dir
                if(not dest_dir.endswith('/')):
                    dest_dir+='/'
            
                try:
                    
                    #if(len(session.files_search(src_dir,file_name).matches)>0):
                    if(len(session.files_get_metadata(src_dir+file_name).name)>0):
                        
                        if(os.path.isdir(dest_dir)):
                            
                            if(os.path.isfile(dest_dir+file_name)):
                                
                                if(overwrite.lower()=='yes'):
                                    # print(os.path.basename(os.path.dirname(src_dir)))
                                    # print(src_dir + "\n"+ dest_dir)
                                    try:
                                        
                                        os.remove(dest_dir + file_name)
                                        self.__log.info("Destination File "+ dest_dir + file_name + " removed from " + dest_dir + " .",exc_info=True)
                                        session.files_download_to_file(dest_dir+file_name,src_dir+file_name)
                                        self.__log.info("Source File "+ src_dir + file_name + " successfully copied to " + dest_dir + " .",exc_info=True)
                                        flag=1
                                    
                                    except:
                                        self.__log.error("Destination Directory " + dest_dir + " transfer process failed.",exc_info=True)
                                        flag=0
                                else:
                                    self.__log.error("File Transfer Incomplete! Overwrite = 'No'",exc_info=True) 
                                    flag=0
                            else:
                                
                                session.files_download_to_file(dest_dir+file_name,src_dir+file_name)
                                self.__log.info("Source File "+ src_dir + file_name + " successfully copied to " + dest_dir + " .",exc_info=True) 
                                flag=1
                                
                        else:
                            
                            os.makedirs(dest_dir)
                            session.files_download_to_file(dest_dir+file_name,src_dir+file_name)
                            self.__log.info("Source File "+ src_dir + file_name + " successfully copied to " + dest_dir + " .",exc_info=True) 
                            flag=1
                            
                    else:
                        
                        self.__log.error("Source Directory " + src_dir + " and file "+file_name+" doesn't exist.",exc_info=True) 
                        flag=0
                        
                except:
                    self.__log.error("Source Directory " + src_dir + " and file "+file_name+" doesn't exist.",exc_info=True) 
                    flag=0
        
            else:
                self.__log.error("Access Token " + access_token + " is invalid.",exc_info=True)
                flag=0
    
        except:
            self.__log.error("Access Token " + access_token + " is invalid.",exc_info=True)
            flag=0 
    
        return flag


    def __find_folder(self, folderName,searchIn):
    # the findFolder function takes the folder you're looking for as folderName,
    # and tries to find it with the MAPIFolder object searchIn
        try:
            lowerAccount = searchIn.Folders
            for x in lowerAccount:
                if str(x) == folderName:
                    objective = x
                    return objective
            return None
        except Exception as error:
            self.__log.error('Cannot find folder', exc_info=True)
            return None
        

    def saveOutlookAttachments(self, user_email, dest_dir, email_saved_folder, subjects):
        
        flag=0
        #ASSUMES ONLY 1 EMAIL ACCOUNT IN OUTLOOK
        #reads through messages matching subject
        #saves to dest_dir
        #moves message to Read folder
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            mapi= outlook.GetNamespace("MAPI")
        except (SystemExit, KeyboardInterrupt):
            raise
        except Exception:
            self.__log.error('Cannot access Outlook', exc_info=True)
            # print("ERROR - Cannot access Outlook")
    
        ###find the inbox and create/select read folder
        ###only grabbing the inbox via hardcoded name
        if(email_saved_folder==""):              
            byd=mapi.Folders[user_email].Folders["Inbox"]
        else:
            byd=mapi.Folders[user_email].Folders[email_saved_folder]
            
            try:
                mapifolder = self.__find_folder(user_email, mapi)
                # print(mapifolder)
                saved_files= self.__find_folder(email_saved_folder, mapifolder)
                # print(saved_files)
                
                if saved_files == None:
                    mapifolder.Folders.Add(email_saved_folder)
                    saved_files = self.__find_folder(email_saved_folder, mapifolder)
            except:
                self.__log.error('ERROR - Unable to find/create save folder', exc_info=True)
        
        self.__log.info('Accessed inbox %s ' % (user_email), exc_info=True)
        
        '''
        try:
            mapifolder = find_folder(user_email, mapi)
            print(mapifolder)
            saved_files= find_folder(email_saved_folder, mapifolder)
            print(saved_files)
            
            if saved_files == None:
                mapifolder.Folders.Add(email_saved_folder)
                saved_files = find_folder(email_saved_folder, mapifolder)
        except:
            log.error('ERROR - Unable to find/create save folder', exc_info=True)
        '''
        if(os.path.isdir(dest_dir)):
            pass
        else:
            os.mkdir(dest_dir)

        if(len(subjects)==0):
            messages = byd.Items
            self.__log.info('Retrieving all %s emails ' % (len(messages)), exc_info=True)
            
            if(len(messages)==0):
                flag=1
                self.__log.info('No new emails !', exc_info=True) 
            
            else:
                flag=1
                ###iterate through ALL messages for 
                for message in messages: 
                    for attachment in message.Attachments:
                        try:
                            attachment.SaveAsFile(dest_dir+'/'+str(attachment))
                            self.__log.info("Saved %s " % (str(attachment)), exc_info=True)
                        except:
                            continue
            
        else:
            for subject in subjects:
                
                # print(subject)
                # sFilter = "[Subject] = '" + subject + "'"
                subFilter = "@SQL= urn:schemas:httpmail:subject like " + "'%"+ "%s" %subject + "%'"
                # Filter = "@SQL= urn:schemas:httpmail:subject like " + "'%"+ "%s" %subject + "%'" + " OR urn:schemas:httpmail:fromemail like " + "'%"+ "%s" %fromemail + "%'"
                messages = byd.Items.Restrict(subFilter)
                # print(messages)
                self.__log.info('Filtered %s emails using filter %s' % (len(messages), subFilter), exc_info=True)

                ###Checking for new retrieved messages
                if(len(messages)==0):
                    flag=1
                    self.__log.info('No new emails !', exc_info=True) 
                    continue
                
                else:
                    flag=1
                    ###iterate through ALL messages for 
                    for message in messages: 
                        
                        for attachment in message.Attachments:
                            
                            attachment.SaveAsFile(dest_dir+'/'+str(attachment))
                            self.__log.info("Saved %s " % (str(attachment)), exc_info=True)
                
                        
            '''
            # Return empty string if no new email is sent
            if(flag):
                return p
            else:
                return ''      
            '''

        return flag


    def _run_local(self, data):
        for i in data['movement_type']['local machine']:
            if(i['file_name']==''):
                if(self.copyFilesFromDirectory(i['src_dir'], i['dest_dir'], i['file_name'], i['overwrite'])):
                    pass
                else:
                    raise Exception('Directory Transfer Error!')
                    
            else:
                if(self.copyFile(i['src_dir'], i['dest_dir'], i['file_name'], i['overwrite'])):
                    pass
                else:
                    raise Exception('File Transfer Error!')

    def _run_dropbox(self, data):        
        for i in data['movement_type']['dropbox']:

            if(i['file_name']==''):
                if(self.copyDirectoryFromDropbox(i['src_dir'], i['dest_dir'], i['file_name'], i['overwrite'],i['dropboxSettings']['access_token'])):
                    pass
                else:
                    raise Exception('Directory Transfer Error!')
                    
            else:
                if(self.copyFileFromDropbox(i['src_dir'], i['dest_dir'], i['file_name'], i['overwrite'],i['dropboxSettings']['access_token'])):
                    pass
                else:
                    raise Exception('File Transfer Error!')

    def _run_outlook(self, data):
        for i in data['movement_type']['outlook']:

            if(self.saveOutlookAttachments(i['emailSettings']['user_email'],i['dest_dir'],i['emailSettings']['email_saved_folder'],i['emailSettings']['subjects'])):
                pass
            
            else:
                raise Exception('Outlook Attachment Download Error!')

def main():

    ft = FileTransfer()
    log, logname = ft.create_log()
    
    root = tkinter.Tk()
    root.withdraw()  
    
    CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))

    parser = argparse.ArgumentParser(description='Arguments selecting the file transfer mode (s)')
    parser.add_argument('--local', help='Local mode for file (s) transfer.', required = False, action = 'store_true')
    parser.add_argument('--dropbox',help='Dropbox mode for file (s) transfer.', required=False, action = 'store_true')
    parser.add_argument('--outlook',help='Outlook mode for file (s) transfer.', required=False, action = 'store_true')

    # Assign values from Configuration File
    var=ft.read_config(CURRENT_DIR+"\\"+'config.json')
    
    # Check if there is an error in loading variables from configuration file
    if(isinstance(var,str)):
        
        # Remove Log Handlers
        for handler in log.handlers:
            log.removeHandler(handler)
         
        log.error("Transfer Process Failed.",exc_info=True)
        # Display Alert for Error in movement
        messagebox.showinfo("Alert","Unsuccessful Transfer.")
    
    
    else:
        
        data = var

        argument = vars(parser.parse_args())

        try:
            # if ((argument.local == False) and (argument.dropbox == False) and (argument.outlook == False)):
            if (len(set(argument.values())) == 1):
                ft._run_local(data)
                log.info("Local Transfer Process Successful.",exc_info=True)
                print("Local Transfer Process Successful.")

                ft._run_outlook(data)
                log.info("Outlook Transfer Process Successful.",exc_info=True)
                print("Outlook Transfer Process Successful.")

                ft._run_dropbox(data)
                log.info("Dropbox Transfer Process Successful.",exc_info=True)
                print("Dropbox Transfer Process Successful.")

                log.info("Transfer Process For All Modes Successful.",exc_info=True)
                print("Transfer Process For All Modes Successful.")

            else:
                if(argument['local']):
                    ft._run_local(data)
                    log.info("Local Transfer Process Successful.",exc_info=True)
                    print("Local Transfer Process Successful.")
                

                if(argument['outlook']):
                    ft._run_outlook(data)
                    log.info("Outlook Transfer Process Successful.",exc_info=True)
                    print("Outlook Transfer Process Successful.")

                if(argument['dropbox']):
                    ft._run_dropbox(data)
                    log.info("Dropbox Transfer Process Successful.",exc_info=True)
                    print("Dropbox Transfer Process Successful.")

                var=1
                log.info("Transfer Process Successful.",exc_info=True)
        
        except:
            log.error("Transfer Process Failed.",exc_info=True)
            print("Transfer Process Failed.")
            var="Error"
            
        # Remove Log Handlers
        for handler in log.handlers:
                log.removeHandler(handler)
        
        if(isinstance(var,str)):
             # Display Alert for Error in movement
             messagebox.showinfo("Alert","Unsuccessful Transfer.")
        
        else:    
            # Display Alert for successful movement
            messagebox.showinfo("Alert","Successful Transfer.")

if __name__ =="__main__":
    main()