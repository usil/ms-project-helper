import os
import re
import win32com.client as w3c
from pathlib import Path
from functions.utils import getChildTasks, getParentTask, getGrandparentTask

def fileExists(path_carpet):
    mppFiles: os._ScandirIterator[str] = []
    masterFile: os.DirEntry[str] = None
    
    with os.scandir(path_carpet) as files:
        for file in files: 
            if file.name.endswith('.mpp'):
                # Example of regular expression : 1. Tarea1
                fileName = Path(file.name).stem
                pattern = r'^\d+\. [0-9A-Za-z ]+$'
                if bool(re.match(pattern, fileName)):
                    mppFiles.append(file)
                else:
                    masterFile = file
    
    return mppFiles, masterFile

def updateMasterFile(masterFile: os.DirEntry[str], files):
    try:
        mpp: w3c.CDispatch = w3c.Dispatch("MSProject.Application")
        try:
            mpp.FileOpen(masterFile.path)
            masterProject = mpp.ActiveProject
            
            for file in files:
                fileName = Path(file.name).stem
                masterTask = next((x for x in masterProject.Tasks 
                                    if x is not None and x.Name is not None and x.Name == fileName), None)

                # master task is summary
                if masterTask is not None and masterTask.Summary:
                    try:
                        mpp.FileSave() 
                        mpp.FileOpen(file.path)
                        childProject = mpp.activeProject
                                    
                        childTask = next((x for x in childProject.Tasks 
                                            if x is not None 
                                            and x.Name is not None 
                                            and x.Name == fileName), 
                                        None)
                        
                        # child task is summary
                        if childTask is not None and childTask.Summary:     
                            print(f"Summary task: {masterTask.Name}")
                            groupedChildTasks = getChildTasks(childTask)
                            
                            mpp.FileSave() 
                            mpp.FileOpen(masterFile.path)
                            masterProject = mpp.ActiveProject
                            
                            masterTask = next((x for x in masterProject.Tasks 
                                                if x is not None 
                                                and x.Name is not None 
                                                and x.Name == fileName),
                                              None)
                            
                            if masterTask is not None and masterTask.Summary:
                                groupedMasterTasks = getChildTasks(masterTask)
                                
                                for task in groupedMasterTasks: 
                                    relatedTask = next((x for x in groupedChildTasks 
                                                            if x is not None 
                                                            and x.Name is not None 
                                                            and x.Name == task.Name 
                                                            and getParentTask(x) == getParentTask(task)
                                                            and getGrandparentTask(x) == getGrandparentTask(task)
                                                            ), 
                                                        None)
    
                                    if relatedTask is not None:
                                        if relatedTask.PercentComplete is not None:                                                    
                                            task.PercentComplete = relatedTask.PercentComplete
                                            print(f"Child task {masterTask.Name}: {relatedTask.Name} ({relatedTask.PercentComplete}% Complete)")
                                        else:
                                            print(f"Child task {masterTask.Name}: {relatedTask.Name} not percent complete")    
                                    else:
                                        print(f"Child task {masterTask.Name}: {task.Name} not found")
                                        
                                print(f"Saving master file {masterFile.name}")                   
                                mpp.FileSave() 
                            
                        print(f"--------------------------------------")
                    except Exception as e:
                        print(str(e))
                else:
                    print(f"Task {file.name} not found")  
            
            mpp.FileSave()    
            mpp.Quit()
        except Exception as e:
            print(str(e))
                
    except Exception as  e:
        print(str(e))
        
